import os
import re
import uuid
import datetime as dt
from typing import Optional, Dict, Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import stripe
import subprocess
import shutil

from openai import OpenAI
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

APP_URL = os.getenv("APP_URL", "")  # ex: https://mycvcopilote.netlify.app
STRIPE_SECRET = os.getenv("STRIPE_SECRET") or os.getenv("STRIPE_SECRET_KEY", "")
STRIPE_WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
PUBLIC_BASE_DOWNLOAD = os.getenv("PUBLIC_BASE_DOWNLOAD", "")  # ex: https://mycvcopilote-api.onrender.com/download

client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

# --- MVP "DB" en mémoire (à remplacer par Postgres plus tard)
# quota[email] = "YYYY-MM" (mois où le gratuit a été consommé)
quota: Dict[str, str] = {}
# jobs[job_id] = {"docx_path":..., "pdf_path":...}
jobs: Dict[str, Dict[str, str]] = {}

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # MVP: ouvrir, plus tard restreindre à ton domaine Netlify
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
@app.get("/")
def root():
    return {"ok": True, "service": "mycvcopilot-backend"}

if STRIPE_SECRET:
    stripe.api_key = STRIPE_SECRET

def month_key(now: Optional[dt.datetime] = None) -> str:
    now = now or dt.datetime.utcnow()
    return f"{now.year:04d}-{now.month:02d}"

def has_free_left(email: str) -> bool:
    import os
    import psycopg2

    conn = psycopg2.connect(os.getenv("DATABASE_URL"))
    cur = conn.cursor()

    cur.execute(
        "SELECT month FROM quota WHERE email = %s",
        (email,)
    )
    row = cur.fetchone()

    cur.close()
    conn.close()

    if not row:
        return True  # jamais utilisé

    return row[0] != month_key()

def consume_free(email: str) -> None:
    import os
    import psycopg2

    conn = psycopg2.connect(os.getenv("DATABASE_URL"))
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO quota (email, month)
        VALUES (%s, %s)
        ON CONFLICT (email)
        DO UPDATE SET month = EXCLUDED.month
        """,
        (email, month_key())
    )

    conn.commit()
    cur.close()
    conn.close()

def sector_to_template(sector: str) -> str:
    s = sector.lower()
    if "finance" in s:
        return "templates/finance.docx"
    if "marketing" in s:
        return "templates/marketing.docx"
    if "ressources" in s or "rh" in s:
        return "templates/rh.docx"
    if "droit" in s:
        return "templates/droit.docx"
    return "templates/finance.docx"

def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^a-zA-Z0-9_-]+", "_", name.strip())
    return name[:50] or "cv"

def build_prompt(payload: Dict[str, Any]) -> str:
    # Prompt “dur” pour produire un CV 1 page ATS + structure
    return f"""
Tu es un expert en recrutement. Tu dois générer un CV FRANÇAIS d'1 page maximum, ultra sobre, ATS-friendly (une seule colonne, pas d'icônes, pas de tableau complexe).
Le CV doit être adapté:
1) au secteur: {payload["sector"]}
2) au poste: {payload["role"]}
3) à l'entreprise: {payload["company"]}
4) à l'offre d'emploi ci-dessous (OBLIGATOIRE)

OFFRE D'EMPLOI (texte brut):
\"\"\"{payload["job_posting"]}\"\"\"

PROFIL UTILISATEUR:
- Nom: {payload["full_name"]}
- Ville: {payload["city"]}
- Email: {payload["email"]}
- Téléphone: {payload["phone"]}
- LinkedIn: {payload.get("linkedin","")}

FORMATION:
{payload["education"]}

EXPERIENCES (brut):
{payload["experiences"]}

COMPETENCES (brut):
{payload["skills"]}

LANGUES:
{payload["languages"]}

CENTRES D’INTERET:
{payload.get("interests","")}

EXIGENCES:
- Tu extraits 10-15 mots-clés ATS de l'offre et tu les intègres naturellement.
- Tu intègres 3-5 soft skills/valeurs visibles dans l'offre, sans surcharger.
- Tu reformules en style pro. Pas de mensonge: si une info manque, reste générique/raisonnable.
- Chaque expérience doit contenir 3-4 bullet points orientés résultats, au moins 1-2 avec chiffres si possible (si pas de chiffres, propose une métrique plausible mais prudente).
- Pas de “profil dynamique/motivé” sans preuve.
- Format final en TEXTE STRUCTURÉ avec sections:
  EN-TÊTE, TITRE, ACCROCHE, COMPETENCES, EXPERIENCES, FORMATION, LANGUES, CENTRES D'INTERET.
- Ne donne PAS d'explications, uniquement le CV.
"""
def build_prompt_finance(payload: Dict[str, Any]) -> str:
    return f"""
Tu es un ancien recruteur en banque d’investissement et en Big 4.
Tu sélectionnes uniquement les 10% meilleurs profils étudiants.
Tu élimines immédiatement les CV vagues, imprécis ou sans résultats chiffrés.

OBJECTIF :
Générer un CV FINANCE français d’1 page maximum, ultra structuré, minimal et stratégique.

Le CV doit être adapté :
- au type de finance visé : {payload.get("finance_type", "Non précisé")}
- au poste : {payload["role"]}
- à l’entreprise : {payload["company"]}
- à l’offre d’emploi

OFFRE D’EMPLOI :
\"\"\"{payload["job_posting"]}\"\"\"

RÈGLES :
- 1 page maximum (ABSOLUMENT aucune 2e page).
- Format de dates homogène, toujours sous la forme "MMM YYYY – MMM YYYY"
  (exemple : "Sept 2023 – Juin 2025") et jamais "09/2023", "2023-2025" ou "au".
- Chaque bullet = Verbe fort + Action + Impact business (sans inventer de chiffres).
- 2 à 3 bullets maximum par expérience (2 par défaut, 3 uniquement pour les expériences les plus pertinentes).
- Interdiction des mots : assisted, helped, worked on.
- Ton professionnel, précis, sobre.
- Classe les expériences de la plus pertinente à la moins pertinente par rapport au poste visé.
- Les expériences de tutorat / soutien scolaire sont plus pertinentes qu’un job de caisse générique et doivent être placées AU-DESSUS des jobs étudiants alimentaires.
- Les expériences en finance / audit / assurance / banque / analyse financière doivent être tout en haut, même si elles sont plus anciennes.
- Les jobs étudiants génériques (supermarché, baby-sitting, barista, etc.) doivent toujours être en bas de la section EXPÉRIENCES, même s’ils sont plus récents.
- Si le contenu commence à être trop long pour tenir sur une page, tu SUPPRIMES d’abord les expériences les moins pertinentes (jobs étudiants génériques) et tu raccourcis les bullets les moins importantes.
- Le CV doit être rédigé intégralement en français (même si l’offre ou les intitulés sont en anglais).
- Tous les bullet points doivent être écrits en français.

RÈGLES STRICTES :
Ces règles priment sur toutes les autres instructions.
- Tu n’inventes AUCUN chiffre.
- Tu n’inventes AUCUNE mission.
- Tu n’inventes AUCUN outil.
- Si une information est absente, tu restes général sans ajouter de détails fictifs.
- Si aucun résultat chiffré n’est fourni, tu reformules sans métriques.
- Tu utilises uniquement les informations présentes dans le profil utilisateur.
- Interdiction totale d’inventer pour “améliorer” le CV.
- Si une expérience contient trop peu d'informations, tu la rends professionnelle mais concise, sans extrapolation.
- Évite les verbes faibles (participé, aidé, effectué, travaillé sur).
- Privilégie des verbes orientés impact et responsabilité.
- Chaque bullet doit refléter une contribution concrète.

BDE / ASSOCIATIONS / PROJETS ÉTUDIANTS :
- Tu DOIS les mettre dans "EXPÉRIENCES PROFESSIONNELLES" (même si ce n’est pas une entreprise).
- Tu les écris comme une expérience (titre + dates si disponibles + 2-3 bullets).
- INTERDICTION ABSOLUE d’inventer des chiffres : aucun %, aucun volume, aucun "5 sponsors", aucun "100 participants" si ce n’est pas fourni.

SECTION SKILLS (COMPÉTENCES & OUTILS) :
- Tu produis EXACTEMENT 2 à 3 lignes sous "SKILLS:" :
  1) "Certifications : ..."
  2) "Maîtrise des logiciels : ..."
  3) "Capacités professionnelles : ..." (facultatif si peu d'infos)
- Dans chaque ligne, les éléments sont séparés par des virgules (PAS de "|").
- "Certifications" : tests ou validations concrètes (Excel, PIX, etc.).
- "Maîtrise des logiciels" : Excel, PowerPoint, VBA, outils spécifiques.
- "Capacités professionnelles" : 3–4 compétences en lien direct avec l’offre (ex : analyse financière, reporting, communication client, gestion des priorités).
- Ne pas mettre ici les langues ni les tests de langues (IELTS, TOEIC, etc.).

SECTION LANGUAGES :
- Tu indiques toutes les langues + les tests officiels (IELTS, TOEIC, etc.).
- Exemple : Français (natif), Anglais (C1 – IELTS 8.0).

SECTION ACTIVITIES (CENTRES D’INTÉRÊT) :
- Tu n’y mets QUE des centres d’intérêt / activités personnelles (sport, voyages, engagements associatifs non listés en expérience, hobbies).
- INTERDICTION d’y mettre BDE / associations / projets déjà listés dans EXPÉRIENCES.
- Pas de doublons : si c’est dans EXPÉRIENCES, tu ne le répètes pas ailleurs.
- Tu n’utilises JAMAIS de Markdown (**texte**, *texte*). Tu écris simplement le texte brut.
- Format de chaque activité sur UNE LIGNE :
  Nom de l’activité en gras, suivi de ":" puis une phrase :
  - ce que la personne a fait concrètement (niveau / fréquence / contexte),
  - ce que ça développe comme qualités utiles en finance / environnement exigeant.
- Exemples de structure (à adapter aux infos réelles) :
  - Équitation (niveau national) : calendrier d’entraînement ajusté aux études, renforçant discipline, résilience et gestion du stress.
  - Course à pied & charity runs : participation régulière à des courses caritatives, développant endurance, persévérance et sens de l’engagement.
  - Voyages en Asie : voyages prolongés dans plusieurs pays, renforçant adaptabilité et sensibilité aux environnements multiculturels.

IMPORTANT :
- Toute la sortie (EDUCATION, EXPERIENCES, SKILLS, LANGUAGES, ACTIVITIES)
  doit être rédigée EN FRANÇAIS.
- Si tu écris une phrase en anglais, tu la traduis immédiatement en français.
- Seuls les noms propres (noms d’écoles, diplômes officiels, logiciels, intitulés exacts de postes)
  peuvent rester en anglais.

RÈGLES DE SORTIE (TRÈS IMPORTANT) :
- Ne génère PAS de titre de section.
- Ne génère PAS le nom.
- Ne génère PAS les coordonnées.
- Ne génère PAS d'accroche.
- Génère uniquement le contenu brut des sections.

FORMAT EXACT À RESPECTER :

EDUCATION:
<contenu>

EXPERIENCES:
ROLE: <intitulé exact>
COMPANY: <nom exact>
DATES: <MMM YYYY – MMM YYYY ou MMM YYYY – Present>
LOCATION: <Ville, Pays>
TYPE: <Internship / Apprenticeship / CDI / etc. si fourni sinon vide>
BULLETS:
- ...
- ...

ROLE: ...
COMPANY: ...
DATES: ...
LOCATION: ...
TYPE: ...
BULLETS:
- ...
- ...

SKILLS:
<2 à 3 lignes, chacune commençant par "Certifications :", "Maîtrise des logiciels :" ou "Capacités professionnelles :">

LANGUAGES:
<contenu>

ACTIVITIES:
<contenu>

CONTRAINTE LONGUEUR (INTELLIGENTE) :

Le CV doit tenir STRICTEMENT sur UNE SEULE page Word.
Tu dois viser une densité pro optimale :
- ni trop vide
- ni surchargé
- une page pleine mais aérée.

RÈGLE STRUCTURELLE DE BASE :
- Maximum 9 bullet points au total (jamais plus de 9).
- 2 bullet points par défaut par expérience.
- 3 bullet points uniquement pour les 1 ou 2 expériences **les plus pertinentes** pour l’offre.
- Tu ne crées pas plus de 4 expériences au total (hors éventuellement 1 job étudiant très court).
- Tu ne supprimes **jamais** une expérience en finance / audit / banque / BDE / projet important, sauf si le profil en contient vraiment trop.

RÈGLE D’AJUSTEMENT AUTOMATIQUE :

1️⃣ Si le contenu devient trop long :
- Tu réduis d’abord les expériences les **moins pertinentes** (jobs étudiants génériques, etc.).
- Tu limites à 2 bullet points maximum pour les expériences secondaires.
- Tu raccourcis les formulations (phrases plus directes, une seule idée par bullet).
- Tu supprimes **uniquement en dernier recours** un job étudiant générique (caisse, vente, etc.), jamais une expérience en finance / audit / BDE / projet sérieux.
- Tu gardes toujours au moins 3 expériences au total si possible.

2️⃣ Si le CV semble trop court (moins d’une page) :
- Tu passes à 3 bullet points pour les expériences les plus pertinentes.
- Tu détailles davantage l’impact concret (toujours sans inventer de chiffres).
- Tu enrichis légèrement la section EDUCATION (matières clés, spécialisation, classement si fourni).
- Tu développes un peu les ACTIVITÉS les plus fortes (sport intensif, voyages marquants, engagement régulier).

RÈGLES D’ÉCRITURE :
- Phrases courtes, une seule idée par bullet.
- Tu évites les répétitions entre les bullets et entre les expériences.

PROFIL :
Nom : {payload["full_name"]}
Ville : {payload["city"]}

FORMATION :
{payload["education"]}

EXPÉRIENCES :
{payload["experiences"]}

COMPÉTENCES :
{payload["skills"]}

LANGUES :
{payload["languages"]}

CENTRES D’INTÉRÊT :
{payload.get("interests","")}

Génère uniquement le CV structuré.
"""
    
def generate_cv_text(payload: Dict[str, Any]) -> str:
    if not client:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY manquante sur le serveur.")

    sector = (payload.get("sector") or "").lower()

    if "finance" in sector:
        prompt = build_prompt_finance(payload)
    else:
        prompt = build_prompt(payload)

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )

    print("=== RAW CV TEXT ===")
    print(resp.choices[0].message.content)
    print("=== END RAW CV TEXT ===")

    return resp.choices[0].message.content.strip()

from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def translate_months_fr(text: str) -> str:
    """
    Normalise les mois :
    - Anglais complet ou abrégé -> abréviation FR (Janv, Fév, Mars, Avr, Mai, Juin, Juil, Août, Sept, Oct, Nov, Déc)
    - Français complet -> abréviation FR
    On évite l'effet 'Septt' en ne remplaçant que des mots entiers.
    """
    # Normaliser la casse (Sept au lieu de SEPT)
    text = re.sub(r"\b(SEPT|OCT|NOV|DÉC|DEC|JANV|FÉV|FEV|AVR|JUIN|JUIL|AOÛT|AOUT)\b",
                  lambda m: m.group(0).capitalize(),
                  text)
    if not text:
        return text

    patterns = {
        # EN full
        r"(?i)\bJanuary\b": "Janv",
        r"(?i)\bFebruary\b": "Fév",
        r"(?i)\bMarch\b": "Mars",
        r"(?i)\bApril\b": "Avr",
        r"(?i)\bMay\b": "Mai",
        r"(?i)\bJune\b": "Juin",
        r"(?i)\bJuly\b": "Juil",
        r"(?i)\bAugust\b": "Août",
        r"(?i)\bSeptember\b": "Sept",
        r"(?i)\bOctober\b": "Oct",
        r"(?i)\bNovember\b": "Nov",
        r"(?i)\bDecember\b": "Déc",

        # EN short
        r"(?i)\bJan\b": "Janv",
        r"(?i)\bFeb\b": "Fév",
        r"(?i)\bMar\b": "Mars",
        r"(?i)\bApr\b": "Avr",
        r"(?i)\bJun\b": "Juin",
        r"(?i)\bJul\b": "Juil",
        r"(?i)\bAug\b": "Août",
        r"(?i)\bSep\b": "Sept",
        r"(?i)\bOct\b": "Oct",
        r"(?i)\bNov\b": "Nov",
        r"(?i)\bDec\b": "Déc",

        # FR full
        r"(?i)\bJanvier\b": "Janv",
        r"(?i)\bFévrier\b": "Fév",
        r"(?i)\bFevrier\b": "Fév",
        r"(?i)\bMars\b": "Mars",
        r"(?i)\bAvril\b": "Avr",
        r"(?i)\bMai\b": "Mai",
        r"(?i)\bJuin\b": "Juin",
        r"(?i)\bJuillet\b": "Juil",
        r"(?i)\bAoût\b": "Août",
        r"(?i)\bAout\b": "Août",
        r"(?i)\bSeptembre\b": "Sept",
        r"(?i)\bOctobre\b": "Oct",
        r"(?i)\bNovembre\b": "Nov",
        r"(?i)\bDécembre\b": "Déc",
        r"(?i)\bDecembre\b": "Déc",
    }

    for pattern, repl in patterns.items():
        text = re.sub(pattern, repl, text)

    return text
def _remove_paragraph(p: Paragraph):
    p._element.getparent().remove(p._element)
    p._p = p._element = None

def _add_table_after(paragraph: Paragraph, rows: int, cols: int):
    """
    Ajoute un tableau juste après le paragraphe.

    Objectifs :
    - 2 colonnes : texte formation à gauche, dates à droite
    - Largeur TOTALE légèrement réduite pour éviter l'effet "dates collées à la marge"
    - Largeurs forcées sur les colonnes (Word + LibreOffice)
    """
    doc = paragraph.part.document
    table = doc.add_table(rows=rows, cols=cols)

    # On ne laisse pas Word/LibreOffice recalculer les largeurs
    table.autofit = False

    if cols == 2:
        try:
        
            # 15,1 cm de texte + 3,9 cm pour les dates
            # → texte bien large + plus de place pour la date (évite qu'elle casse)
            # Largeur totale ≈ 19 cm : très proche du bord mais sans dépasser
            widths = [Cm(15.1), Cm(3.9)]

            # Largeur sur les colonnes
            for col, w in zip(table.columns, widths):
                col.width = w

            # Sécurité : on force aussi la largeur sur chaque cellule
            for row in table.rows:
                for j, w in enumerate(widths):
                    row.cells[j].width = w
        except Exception:
            pass

    # On aligne le tableau à gauche pour qu'il commence au même endroit que le texte normal
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Insérer le tableau juste après le paragraphe ancre
    paragraph._p.addnext(table._tbl)
    return table

def parse_finance_experiences(lines: list[str]) -> list[dict]:
    exps = []
    cur = None
    mode = None

    def push():
        nonlocal cur
        if cur and (cur.get("role") or cur.get("bullets")):
            exps.append(cur)
        cur = None

    for raw in lines:
        line = (raw or "").strip()
        if not line:
            continue

        if line.startswith("ROLE:"):
            push()
            cur = {
                "role": line.replace("ROLE:", "").strip(),
                "company": "",
                "dates": "",
                "location": "",
                "type": "",
                "bullets": [],
            }
            mode = None
            continue

        if not cur:
            continue

        if line.startswith("COMPANY:"):
            cur["company"] = line.replace("COMPANY:", "").strip()
        elif line.startswith("DATES:"):
            cur["dates"] = line.replace("DATES:", "").strip()
        elif line.startswith("LOCATION:"):
            cur["location"] = line.replace("LOCATION:", "").strip()
        elif line.startswith("TYPE:"):
            cur["type"] = line.replace("TYPE:", "").strip()
        elif line.startswith("BULLETS:"):
            mode = "bullets"
        elif mode == "bullets" and line.startswith("-"):
            cur["bullets"].append(line[1:].strip())

    push()
    return exps


PLACEHOLDERS = [
    "%%FULL_NAME%%",
    "%%CV_TITLE%%",
    "%%CONTACT_LINE%%",
    "%%EDUCATION%%",
    "%%EXPERIENCE%%",
    "%%SKILLS%%",
    "%%LANGUAGES%%",
    "%%INTERESTS%%",
]

def trim_finance_experiences(
    exps: list[dict],
    max_experiences: int = 4,
    max_total_bullets: int = 8,
    min_experiences: int = 2,
) -> list[dict]:
    """
    Objectif : être SÛR que la section EXPÉRIENCES ne déborde pas.
    - On garde max 4 expériences (en pratique souvent 3).
    - On vise ~8 bullets au total.
    - 1ère expérience = la plus développée.
    - On raccourcit fort la 3ᵉ / 4ᵉ et on supprime la dernière si ça reste trop long.

    Heuristique "intelligente" :
    - Tant que le volume de texte est raisonnable → on peut garder 4 expériences.
    - Si le volume est trop gros (beaucoup de bullets très longues) → on passe à 3 expériences.
    """

    # 1) Nettoyage des expériences vides
    cleaned: list[dict] = []
    for e in exps:
        role = (e.get("role") or "").strip()
        bullets = [b for b in (e.get("bullets") or []) if (b or "").strip()]
        if not role and not bullets:
            continue
        e["role"] = role
        e["bullets"] = bullets
        cleaned.append(e)

    if not cleaned:
        return []

    # 2) On limite quand même à max_experiences (ordre supposé trié par pertinence dans le prompt)
    if len(cleaned) > max_experiences:
        cleaned = cleaned[:max_experiences]

    # 2bis) Heuristique volume de texte :
    # - si on a 4 expériences ET que les bullets sont très verbeuses,
    #   on supprime la 4ᵉ (la moins prioritaire).
    if len(cleaned) == 4:
        # Score "volume" très simple : nb de caractères dans titres + bullets
        total_chars = 0
        for e in cleaned:
            total_chars += len(e.get("role", ""))
            total_chars += len(e.get("company", ""))
            total_chars += sum(len(b) for b in e.get("bullets", []))

        # Seuil assez haut pour ne pas supprimer la 4ᵉ pour rien.
        # Au-delà, on considère que la section risque de faire déborder la page.
        if total_chars > 550:
            cleaned = cleaned[:3]

    # 3) Limite de bullets par expérience :
    #    - exp 0 : max 3 bullets
    #    - exp 1 : max 2 bullets
    #    - exp 2 et 3 : max 1 bullet
    for idx, e in enumerate(cleaned):
        if idx == 0:
            max_b = 3
        elif idx == 1:
            max_b = 2
        else:
            max_b = 1
        e["bullets"] = e["bullets"][:max_b]

    def total_bullets(exps_list: list[dict]) -> int:
        return sum(len(e.get("bullets", [])) for e in exps_list)

    # 4) Si on est encore trop long → on enlève des bullets en partant du bas,
    #    puis en dernier recours on supprime la DERNIÈRE expérience.
    while total_bullets(cleaned) > max_total_bullets and cleaned:
        changed = False

        # a) On enlève une bullet à la dernière expérience qui en a encore
        for idx in range(len(cleaned) - 1, -1, -1):
            b_list = cleaned[idx].get("bullets", [])
            if len(b_list) > 0:
                b_list.pop()
                changed = True
                break

        # b) Si plus aucune bullet à enlever et qu'on a > min_experiences,
        #    on supprime la DERNIÈRE expérience (la moins prioritaire).
        if not changed and len(cleaned) > min_experiences:
            cleaned.pop()
            changed = True

        if not changed:
            break

    return cleaned

def trim_activities(
    lines: list[str],
    ideal_max: int = 3,
    hard_max_when_long: int = 2,
    max_chars_per_activity: int = 140,
    long_total_threshold: int = 260,
) -> list[str]:
    """
    - Idéalement on garde jusqu'à 3 activités.
    - Si elles sont très verbeuses (beaucoup de texte au total), on n'en garde que 2
      en supprimant la moins développée (la plus courte).
    - Chaque activité est éventuellement raccourcie proprement (une phrase max).
    """

    # 1) Nettoyage des lignes vides
    cleaned = [(l or "").strip() for l in (lines or []) if (l or "").strip()]
    if not cleaned:
        return []

    # 2) On garde au maximum 3 bruts (avant filtrage long/court)
    cleaned = cleaned[:ideal_max]

    def shorten(text: str) -> str:
        # Si la phrase est courte → on ne touche à rien
        if len(text) <= max_chars_per_activity:
            return text

        # Sinon on coupe proprement près d'un séparateur AVANT la limite
        candidates = []
        for sep in [".", ";", ",", " et "]:
            idx = text.rfind(sep, 0, max_chars_per_activity)
            if idx != -1:
                candidates.append(idx)

        if candidates:
            cut = max(candidates)
            return text[:cut].rstrip(" ,;et") + "…"

        # Fallback : coupe brutale mais courte
        return text[: max_chars_per_activity - 1].rstrip() + "…"

    # 3) On raccourcit chaque activité si besoin
    shortened = [shorten(t) for t in cleaned]

    # 4) Heuristique "CV serré" :
    #    - Si on a 3 activités
    #    - ET que le total de texte est élevé
    #    -> on n'en garde que 2 (en supprimant la moins développée)
    if len(shortened) == ideal_max:
        total_len = sum(len(t) for t in shortened)
        if total_len > long_total_threshold and hard_max_when_long < ideal_max:
            # index de la phrase la plus courte => moins développée
            idx_to_drop = min(range(len(shortened)), key=lambda i: len(shortened[i]))
            shortened.pop(idx_to_drop)

    return shortened

def _find_paragraph_containing(doc: Document, needle: str):
    for p in doc.paragraphs:
        if needle in (p.text or ""):
            return p
    return None


def _clear_paragraph(p):
    p.text = ""


def _insert_paragraph_after(paragraph, text="", style=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)

    if text:
        new_para.add_run(text)

    if style:
        try:
            new_para.style = style
        except Exception:
            pass

    return new_para


def _insert_lines_after(paragraph, lines, make_bullets=False):
    last = paragraph
    for line in lines:
        line = (line or "").rstrip()

        if not line:
            last = _insert_paragraph_after(last, "")
            continue

        if make_bullets and line.lstrip().startswith("-"):
            text = line.lstrip()[1:].strip()
            last = _insert_paragraph_after(last, text, style="List Bullet")
        else:
            last = _insert_paragraph_after(last, line)

    return last


def _split_sections(cv_text: str) -> dict:
    t = (cv_text or "").replace("\r\n", "\n").strip()

    tags = [
        "EDUCATION:", "FORMATION:",
        "EXPERIENCES:", "EXPÉRIENCES:", "EXPERIENCE:",
        "SKILLS:", "COMPETENCES:", "COMPÉTENCES:",
        "LANGUAGES:", "LANGUES:",
        "INTERESTS:", "ACTIVITIES:", "ACTIVITÉS:"
    ]
    pos = {tag: t.find(tag) for tag in tags}

    if all(pos[tag] == -1 for tag in tags):
        return {
            "EDUCATION": t.splitlines(),
            "EXPERIENCES": [],
            "SKILLS": [],
            "LANGUAGES": [],
            "INTERESTS": [],
            "ACTIVITIES": [],
        }

    present = [(tag, pos[tag]) for tag in tags if pos[tag] != -1]
    present.sort(key=lambda x: x[1])

    sections = {}
    for i, (tag, start) in enumerate(present):
        end = present[i + 1][1] if i + 1 < len(present) else len(t)
        block = t[start:end].strip().splitlines()

        if block and block[0].strip() == tag:
            block = block[1:]

        while block and not block[0].strip():
            block = block[1:]
        while block and not block[-1].strip():
            block = block[:-1]

        sections[tag.replace(":", "")] = block

    if not sections.get("SKILLS"):
        sections["SKILLS"] = sections.get("COMPETENCES") or sections.get("COMPÉTENCES") or []

    if not sections.get("LANGUAGES"):
        sections["LANGUAGES"] = sections.get("LANGUES") or []

    if not sections.get("EXPERIENCES"):
        sections["EXPERIENCES"] = sections.get("EXPÉRIENCES") or sections.get("EXPERIENCE") or []

    # 🔴 IMPORTANT : si le modèle écrit "FORMATION:" au lieu de "EDUCATION:"
    if not sections.get("EDUCATION"):
        sections["EDUCATION"] = sections.get("FORMATION") or sections.get("EDUCATION") or []

    return sections
def _render_education(anchor: Paragraph, lines: list[str]):
    """
    Rend la section FORMATION de façon un peu plus premium :
    - Première ligne de chaque bloc en gras
    - 'Cours pertinents' -> 'Matières fondamentales'
    - 'Matières fondamentales :' souligné
    - Dans la section EDUCATION, chaque diplôme ou programme est sur son propre paragraphe, séparé par UNE LIGNE VIDE du suivant (ex : Programme Grande École, ligne vide, puis Baccalauréat, etc.
    """
    last = anchor
    first_in_block = True

    for raw in (lines or []):
        line = (raw or "").strip()

        # ligne vide = séparation entre deux formations
        if not line:
            last = _insert_paragraph_after(last, "")
            first_in_block = True
            continue

        # Remplace le texte
        if "Cours pertinents" in line or "Key coursework" in line:
            line = line.replace("Cours pertinents", "Matières fondamentales")
            line = re.sub(r"(?i)key coursework", "Matières fondamentales", line)

        # Première ligne du bloc = nom d'école / programme -> gras
        if first_in_block:
            para = _insert_paragraph_after(last, "")
            run = para.add_run(line)
            run.bold = True
            para.paragraph_format.space_after = Pt(0)
            last = para
            first_in_block = False
            continue

        # Ligne "Matières fondamentales : ..." avec le label souligné
        if "Matières fondamentales" in line:
            para = _insert_paragraph_after(last, "")
            before, sep, after = line.partition(":")
            label = before + sep  # "Matières fondamentales:"
            r1 = para.add_run(label + " ")
            r1.underline = True
            if after:
                para.add_run(after.strip())
            last = para
            continue

        # Autres lignes normales
        last = _insert_paragraph_after(last, line)

    return last

def _render_interests(anchor: Paragraph, lines: list[str]):
    """
    Rend la section ACTIVITIES / CENTRES D'INTÉRÊT :
    - Chaque ligne -> puce
    - Nom de l'activité en gras avant ':' ou ' - '
    """
    last = anchor

    for raw in (lines or []):
        text = (raw or "").strip()
        if not text:
            last = _insert_paragraph_after(last, "")
            continue

        # Nouveau paragraphe en mode liste à puces
        new_p = _insert_paragraph_after(last, "")
        try:
            new_p.style = "List Bullet"
        except Exception:
            pass

        head = text
        tail = ""

        if ":" in text:
            head, tail = text.split(":", 1)
        elif " - " in text:
            left, right = text.split(" - ", 1)
            # On considère que la partie gauche est le "nom" si elle est courte
            if len(left.split()) <= 4:
                head, tail = left, right
            else:
                head, tail = text, ""

        head = head.strip()
        tail = tail.strip()

        # Nettoyage des éventuels **...** ou *...* venant du modèle
        while head.startswith("*") and head.endswith("*") and len(head) > 2:
            head = head[1:-1].strip()

        r_head = new_p.add_run(head)
        r_head.bold = True

        if tail:
            new_p.add_run(" : " + tail)

        last = new_p

    return last

def _render_skills(anchor: Paragraph, lines: list[str]):
    """
    Rend la section COMPÉTENCES & OUTILS :
    - Pas de puces
    - Sous-titres en gras (Certifications, Maîtrise des logiciels, Capacités professionnelles)
    - Les éléments sont séparés par des virgules
    """
    last = anchor

    for raw in (lines or []):
        text = (raw or "").strip()
        if not text:
            last = _insert_paragraph_after(last, "")
            continue

        # On remplace les ' | ' par des virgules si jamais le modèle en met encore
        text = text.replace(" | ", ", ")

        new_p = _insert_paragraph_after(last, "")
        head = text
        tail = ""

        if ":" in text:
            head, tail = text.split(":", 1)
        elif " - " in text:
            left, right = text.split(" - ", 1)
            if len(left.split()) <= 4:
                head, tail = left, right
            else:
                head, tail = text, ""

        head = head.strip()
        tail = tail.strip()

        r_head = new_p.add_run(head)
        r_head.bold = True

        if tail:
            new_p.add_run(" : " + tail)

        last = new_p

    return last
    
def _education_end_year(block: list[str]) -> int:
    """
    Récupère l'année de fin à partir de la première ligne du bloc.
    On prend simplement la DERNIÈRE année à 4 chiffres trouvée dans la ligne.
    Ex :
      'Programme Grande École – ESCP — Sept 2022 – Juin 2026' -> 2026
      'Classe préparatoire ECG – Lycée du Parc (2020-2022)'   -> 2022
    """

    if not block:
        return 0

    first_line = (block[0] or "").strip()

    # On cherche toutes les années à 4 chiffres dans la ligne complète
    years = re.findall(r"(?:19|20)\d{2}", first_line)

    if not years:
        return 0

    try:
        # Dernière année = année de fin
        return int(years[-1])
    except ValueError:
        return 0

def _is_bac_block(block: list[str]) -> bool:
    """Retourne True si le bloc correspond à un baccalauréat classique."""
    if not block:
        return False
    first = (block[0] or "").lower()
    return "baccalauréat" in first or "baccalaureat" in first


def _keep_bac_block(block: list[str]) -> bool:
    """
    On garde le bac UNIQUEMENT si :
    1) lycée d'exception (Henri IV, Louis-le-Grand, lycée international, etc.)
    2) bac / diplôme international (IB, Abibac, maturité suisse, etc.)
    3) mention d'honneur type 'félicitations du jury'
    """
    text = " ".join(block).lower()
    # Cas spécifiques : honneurs / honeurs du jury
    if "honneurs du jury" in text or "honeurs du jury" in text:
        return True

    elite_keywords = [
        "henri iv", "henri-iv", "henry iv",
        "louis-le-grand", "louis le grand",
        "lycée international", "lycee international",
        "lycée du parc", "lycee du parc",
        "stanislas", "lycée stanislas",
        "janson de sailly",
        "franklin", "lycée franklin",
        "fénelon", "fenelon",
        "charlemagne",
        "buffon",
        "condorcet",
        "sainte-geneviève", "sainte genevieve", "ginette",
        "le parc",
        "masséna", "massena",
        "thiers",
        "hoche",
        "kléber", "kleber",
        "clemenceau",
        "du parc",
        "chateaubriand",
        "berthelot",
        "pierre de fermat",
        "montaigne",
        "descartes",
        "champollion",
    ]

    intl_keywords = [
        "baccalauréat international", "baccalaureat international",
        "international baccalaureate", "ib diploma", "ib programme",
        "abibac", "esabac",
        "maturité suisse", "maturite suisse", "maturité gymnasiale",
        "matura",
        " ib ",
        "cess",  # Belgique
        "certificat d'enseignement secondaire supérieur",
        "certificat d'enseignement secondaire superieur",
    ]

    honours_keywords = [
        "félicitations du jury", "felicitations du jury", "honneurs du jury"
    ]

    if any(k in text for k in elite_keywords):
        return True
    if any(k in text for k in intl_keywords):
        return True
    if any(k in text for k in honours_keywords):
        return True

    return False

def normalize_contract_type(t: str) -> str:
    if not t:
        return ""

    original = t.strip()
    t_clean = original.lower().strip()

    base_mapping = {
        "internship": "Stage",
        "intern": "Stage",
        "traineeship": "Stage",
        "apprenticeship": "Alternance",
        "full-time": "CDI",
        "full time": "CDI",
        "part-time": "Temps partiel",
        "part time": "Temps partiel",
        "part-time job": "Job étudiant",
        "student job": "Job étudiant",
        "summer job": "Job d'été",
        "temporary": "CDD",
        "contract": "CDD",
        "volunteering": "Volontariat",
        "volunteer": "Volontariat",
    }

    # Match exact
    if t_clean in base_mapping:
        return base_mapping[t_clean]

    # Match préfixe (ex: "part-time job - barista")
    for key, value in base_mapping.items():
        if t_clean.startswith(key + " "):
            suffix = original[len(key):].lstrip(" -–—")
            return value + (f" – {suffix}" if suffix else "")

    return original

def _split_education_block_on_degree_titles(block: list[str]) -> list[list[str]]:
    """
    Si l'IA enchaîne plusieurs diplômes dans un même bloc (sans ligne vide),
    on découpe dès qu'une ligne commence par un mot typique de diplôme.
    Exemple :
      Master 2 Finance ...
      Licence Finance ...
      Baccalauréat ES ...

    devient 3 blocs distincts.
    """
    if not block:
        return []

    DEGREE_STARTERS = (
        "Master", "Master 1", "Master 2",
        "Programme Grande École", "Programme Grande Ecole", "Programme",
        "Licence", "License",
        "Baccalauréat", "Baccalaureat",
        "Classe préparatoire", "Classe préparatoire ECG",
        "Classe preparatoire", "Classe preparatoire ECG",
        "CPGE", "Prépa", "Prepa",
    )

    new_blocks: list[list[str]] = []
    current: list[str] = []

    for raw in block:
        line = (raw or "").strip()
        if not line:
            # Lignes vides : on ferme le bloc en cours
            if current:
                new_blocks.append(current)
                current = []
            continue

        # Si on tombe sur une nouvelle ligne qui ressemble à un début de diplôme
        # on démarre un nouveau bloc
        if current and any(line.startswith(prefix) for prefix in DEGREE_STARTERS):
            new_blocks.append(current)
            current = [line]
        else:
            current.append(line)

    if current:
        new_blocks.append(current)

    return new_blocks
def write_docx_from_template(template_path: str, cv_text: str, out_path: str, payload: dict = None) -> None:
    doc = Document(template_path)

    # Marges plus petites pour mieux utiliser la largeur
    for section in doc.sections:
        section.left_margin = Cm(1.0)
        section.right_margin = Cm(1.0)
        section.top_margin = Cm(1.2)      
        section.bottom_margin = Cm(1.2)   

    # ------- Données générales -------
    payload = payload or {}
    full_name = payload.get("full_name", "").strip() or "NOM Prénom"
    role = payload.get("role", "").strip()
    finance_type = payload.get("finance_type", "").strip()
    cv_title = finance_type if finance_type else role

    contact_line = " | ".join([
        x.strip()
        for x in [
            payload.get("phone", ""),
            payload.get("email", ""),
            payload.get("linkedin", ""),
        ]
        if x and x.strip()
    ])

    sections = _split_sections(cv_text)

    # SKILLS : on garde plusieurs lignes, on nettoie juste les tirets éventuels
    if isinstance(sections.get("SKILLS"), list):
        cleaned = [x.strip().lstrip("-").strip() for x in sections["SKILLS"] if x.strip()]
        sections["SKILLS"] = cleaned
    
    # ⬇️ Langues intégrées dans Compétences & Outils
    languages_raw = sections.get("LANGUAGES") or []
    if isinstance(languages_raw, list):
        lang_text = ", ".join(x.strip() for x in languages_raw if x.strip())
    else:
        lang_text = str(languages_raw).strip()
    
    if lang_text:
        skills_list = sections.get("SKILLS") or []
        skills_list.append(f"Langues : {lang_text}")
        sections["SKILLS"] = skills_list
    
    sections["LANGUAGES"] = []
    
    interests_raw = sections.get("INTERESTS", []) or sections.get("ACTIVITIES", [])

    if isinstance(interests_raw, list):
        interests_value = trim_activities(interests_raw)
    else:
        interests_value = interests_raw
    mapping = {
        "%%FULL_NAME%%": full_name,
        "%%CONTACT_LINE%%": contact_line,
        "%%CV_TITLE%%": cv_title,
        "%%EDUCATION%%": sections.get("EDUCATION", []),
        "%%EXPERIENCE%%": sections.get("EXPERIENCES", []),
        "%%SKILLS%%": sections.get("SKILLS", []),
        "%%LANGUAGES%%": sections.get("LANGUAGES", []),
        "%%INTERESTS%%": interests_value,
    }

    for ph, value in mapping.items():
        p = _find_paragraph_containing(doc, ph)
        if not p:
            continue

        _clear_paragraph(p)

        # ------- COMPÉTENCES & OUTILS -------
        if ph == "%%SKILLS%%" and isinstance(value, list):
            _render_skills(p, value or [])
            _remove_paragraph(p)
            continue

        # ------- FORMATION -------
        if ph == "%%EDUCATION%%" and isinstance(value, list):
            anchor = p

            # 1) Regrouper les lignes par formation (blocs séparés par ligne vide)
            blocks = []
            current_block = []
            for line in value:
                text = (line or "").rstrip()
                if not text:
                    if current_block:
                        blocks.append(current_block)
                        current_block = []
                else:
                    current_block.append(text)
            if current_block:
                blocks.append(current_block)

            # 2) Découper les blocs s'il y a plusieurs diplômes collés
            split_blocks = []
            for b in blocks:
                split_blocks.extend(_split_education_block_on_degree_titles(b))

            # 3) Tri du plus récent au plus ancien
            blocks_sorted = sorted(split_blocks, key=_education_end_year, reverse=True)

            # 4) Gestion du bac (on peut le masquer)
            non_bac_blocks = [b for b in blocks_sorted if not _is_bac_block(b)]
            if len(non_bac_blocks) <= 1:
                filtered_blocks = blocks_sorted[:]
            else:
                filtered_blocks = []
                for b in blocks_sorted:
                    if _is_bac_block(b) and not _keep_bac_block(b):
                        continue
                    filtered_blocks.append(b)

            # 5) Pour chaque formation -> tableau 1 ligne / 2 colonnes
            for block in filtered_blocks:
                if not block:
                    continue

                first_line = block[0]

                # Normalisation des termes d'échange
                lower_first = first_line.lower()
                if "exchange semester" in lower_first or "exchange program" in lower_first:
                    first_line = re.sub(r"(?i)exchange semester", "Échange académique", first_line)
                    first_line = re.sub(r"(?i)exchange program", "Échange académique", first_line)
                if "study abroad" in lower_first:
                    first_line = re.sub(r"(?i)study abroad", "Échange académique", first_line)

                # Séparation Titre / Dates en cherchant un VRAI intervalle de dates en fin de ligne
                title_part = first_line
                date_part = ""

                # On cherche d'abord un pattern du type "Sept 2022 – Juin 2026"
                date_range_patterns = [
                    # Ex : "Sept 2022 – Juin 2026"
                    r"(Janv|Fév|Fev|Mars|Avr|Mai|Juin|Juil|Août|Aout|Sept|Oct|Nov|Déc|Dec)\s+\d{4}\s*[–-]\s*(Janv|Fév|Fev|Mars|Avr|Mai|Juin|Juil|Août|Aout|Sept|Oct|Nov|Déc|Dec)\s+\d{4}\s*$",
                    # Ex : "09/2023 – 06/2025"
                    r"(0[1-9]|1[0-2])/\d{4}\s*[–-]\s*(0[1-9]|1[0-2])/\d{4}\s*$",
                    # Ex : "2020 – 2023"
                    r"(19|20)\d{4}?\s*[–-]\s*(19|20)\d{4}?\s*$"
                ]

                for pat in date_range_patterns:
                    m = re.search(pat, first_line)
                    if m:
                        # Toute la plage de dates part à droite
                        date_part = m.group(0).strip()
                        # Tout ce qui est AVANT la plage reste dans le titre
                        title_part = first_line[:m.start()].rstrip(" ,–-").strip()
                        break

                # Si on n'a toujours pas trouvé, on retombe sur l'ancien fallback : dernier séparateur
                if not date_part:
                    for sep in ("–", "—", "-"):
                        idx = first_line.rfind(sep)
                        if idx != -1:
                            title_part = first_line[:idx].strip()
                            date_part = first_line[idx + 1:].strip()
                            break

                # Dernière sécurité : si une année traîne encore à la fin du titre, on la coupe
                if date_part:
                    m = re.search(r"(19|20)\d{2}\s*$", title_part)
                    if m:
                        title_part = title_part[:m.start()].rstrip(" ,–-")

                # Création du tableau
                table = _add_table_after(anchor, rows=1, cols=2)
                left = table.cell(0, 0)
                right = table.cell(0, 1)
                left.text = ""
                right.text = ""

                # ---- Colonne gauche : titre + détails ----
                lp = left.paragraphs[0]
                try:
                    lp.style = doc.styles["Normal"]
                except Exception:
                    pass
                lp.paragraph_format.left_indent = Pt(0)
                lp.paragraph_format.first_line_indent = Pt(0)

                title_run = lp.add_run(title_part)
                title_run.bold = True
                title_run.font.size = Pt(11)

                # On repère la ligne "ville, pays" pour ne pas la répéter à gauche
                location = ""
                location_index = -1
                for idx_line, raw in enumerate(block):
                    t = (raw or "").strip()
                    lower_t = t.lower()
                    if not t:
                        continue

                    candidate = None
                    if "," in t:
                        parts = [pp.strip() for pp in t.split(",")]
                        if len(parts) == 2 and len(parts[0].split()) <= 3:
                            candidate = t
                    else:
                        if len(t.split()) <= 3:
                            bad_tokens = [
                                "cours", "course", "key", "ranked",
                                "mention", "option", "majeure",
                                "matières", "matieres", "gpa"
                            ]
                            if not any(bt in lower_t for bt in bad_tokens):
                                candidate = t

                    if candidate:
                        location = candidate
                        location_index = idx_line
                        break

                # Détails sous le titre (on saute la ligne du lieu si détectée)
                for idx_line, line in enumerate(block[1:], start=1):
                    if idx_line == location_index:
                        continue

                    text = (line or "").strip()
                    if not text:
                        continue

                    para = left.add_paragraph()
                    try:
                        para.style = doc.styles["Normal"]
                    except Exception:
                        pass
                    para.paragraph_format.left_indent = Pt(0)
                    para.paragraph_format.first_line_indent = Pt(0)

                    label_text = None
                    after_text = None

                    if ":" in text:
                        before, sep, after = text.partition(":")
                        before_clean = before.strip()
                        lower_before = before_clean.lower()

                        if "cours pertinents" in lower_before:
                            label_text = "Matières fondamentales"
                            after_text = after or ""
                        else:
                            word_count = len(before_clean.split())
                            keywords = [
                                "gpa", "hl", "matières", "matieres",
                                "option", "majeure",
                                "spécialité", "specialite",
                            ]
                            if word_count <= 4 or any(k in lower_before for k in keywords):
                                label_text = before_clean
                                after_text = after or ""

                    if label_text:
                        r1 = para.add_run(label_text + " :")
                        r1.underline = True
                        r1.font.size = Pt(10)
                        if after_text and after_text.strip():
                            r2 = para.add_run(" " + after_text.strip())
                            r2.font.size = Pt(10)
                    else:
                        run = para.add_run(text)
                        run.font.size = Pt(10)

                # ---- Colonne droite : dates + lieu ----
                rp = right.paragraphs[0]
                rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                rp.paragraph_format.space_after = Pt(0)

                if date_part:
                    clean_date = date_part.replace("\r", " ").replace("\n", " ")
                    clean_date = re.sub(r"\s+", " ", clean_date.strip())
                    clean_date = translate_months_fr(clean_date)
                    clean_date = clean_date.replace(" - ", " – ")
                    clean_date = clean_date.replace(" ", "\u00A0")
                    r_date = rp.add_run(clean_date)
                    r_date.italic = True
                    r_date.font.size = Pt(9)

                if location:
                    rp.add_run("\n")
                    r_loc = rp.add_run(location.strip())
                    r_loc.italic = True
                    r_loc.font.size = Pt(9)
                    rp.paragraph_format.space_after = Pt(0)

                # Paragraphe vide pour ancrer la prochaine formation
                new_p_elt = OxmlElement("w:p")
                table._tbl.addnext(new_p_elt)
                anchor = Paragraph(new_p_elt, p._parent)

            # On supprime le dernier paragraphe vide utilisé comme ancre
            try:
                if anchor is not None:
                    _remove_paragraph(anchor)
            except Exception:
                pass

            _remove_paragraph(p)
            continue

        # ------- ACTIVITIES / INTERESTS -------
        if ph == "%%INTERESTS%%" and isinstance(value, list):
            _render_interests(p, value or [])
            _remove_paragraph(p)
            continue

        if ph == "%%EXPERIENCE%%":
            exps = parse_finance_experiences(value or [])
            exps = trim_finance_experiences(exps)  # 💡 NOUVEAU
            anchor = p

            # Si le modèle ne respecte pas le format, on retombe sur une liste simple
            if not exps:
                _insert_lines_after(p, value or [], make_bullets=True)
                continue

            CONTRACT_PREFIXES = [
                "stagiaire", "stage",
                "summer job", "part-time job", "student job",
                "volunteering", "volunteer",
                "internship", "intern", "traineeship",
                "apprenticeship",
                "full-time", "full time",
                "part-time", "part time",
            ]

            for exp in exps:
                raw_role = (exp.get("role") or "").strip()
                role = raw_role

                # 1) Cas du type "Stage en audit financier" -> on vire "Stage + en/dans/au/aux"
                role = re.sub(
                    r"^(stage|stagiaire|internship|intern|traineeship)\s+(en|dans|au|aux)\s+",
                    "",
                    role,
                    flags=re.IGNORECASE,
                ).strip()

                lower_role = role.lower()

                # 2) Si le rôle commence encore par un type de contrat (hors "en ..."), on enlève juste ce préfixe
                for key in CONTRACT_PREFIXES:
                    if lower_role.startswith(key + " "):
                        role = role[len(key):].lstrip(" -–—")
                        lower_role = role.lower()
                        break

                # 3) Cas particulier "Student tutor"
                if "student tutor" in lower_role:
                    role = role.replace("Student tutor", "Tuteur bénévole").replace("student tutor", "Tuteur bénévole")

                # 4) On force une majuscule au début du rôle si besoin
                if role and role[0].islower():
                    role = role[0].upper() + role[1:]

                company = (exp.get("company") or "").strip()
                title_parts = [x for x in [role, company] if x]
                title_line = " - ".join(title_parts)

                # Tableau 2 colonnes
                table = _add_table_after(anchor, rows=1, cols=2)
                left = table.cell(0, 0)
                right = table.cell(0, 1)
                left.text = ""
                right.text = ""

                # Colonne gauche : titre + bullets
                lp = left.paragraphs[0]
                try:
                    lp.style = doc.styles["Normal"]
                except Exception:
                    pass
                lp.paragraph_format.left_indent = Pt(0)
                lp.paragraph_format.first_line_indent = Pt(0)

                if title_line:
                    title_run = lp.add_run(title_line)
                    title_run.bold = True
                    title_run.font.size = Pt(11)

                bullets = (exp.get("bullets") or [])[:3]
                for b in bullets:
                    if not b:
                        continue
                    b_clean = b.strip().lower()
                    if b_clean in {"n/a", "na", "not applicable", "non applicable", "non-applicable"}:
                        continue
                    bp = left.add_paragraph()
                    try:
                        bp.style = "List Bullet"
                        bp.add_run(b)
                    except Exception:
                        bp.text = f"• {b}"
                    bp.paragraph_format.space_after = Pt(0)

                # Colonne droite : dates / lieu / type
                rp = right.paragraphs[0]
                rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                rp.paragraph_format.space_after = Pt(0)

                dates_raw = (exp.get("dates") or "").strip()
                if dates_raw:
                    clean_date = dates_raw.replace("\r", " ").replace("\n", " ")
                    clean_date = re.sub(r"\s+", " ", clean_date.strip())
                    clean_date = translate_months_fr(clean_date)
                    clean_date = clean_date.replace(" - ", " – ")
                    clean_date = clean_date.replace(" ", "\u00A0")
                    r_date = rp.add_run(clean_date)
                    r_date.italic = True
                    r_date.font.size = Pt(9)

                location = (exp.get("location") or "").strip()
                if location:
                    rp.add_run("\n")
                    r_loc = rp.add_run(location)
                    r_loc.italic = True
                    r_loc.font.size = Pt(9)

                type_raw = (exp.get("type") or "").strip()
                type_ = normalize_contract_type(type_raw)
                if type_:
                    rp.add_run("\n")
                    r_type = rp.add_run(type_)
                    r_type.italic = True
                    r_type.font.size = Pt(9)

                # Paragraphe vide pour ancrer l'expérience suivante
                new_p_elt = OxmlElement("w:p")
                table._tbl.addnext(new_p_elt)
                anchor = Paragraph(new_p_elt, p._parent)

            # On supprime le dernier paragraphe vide utilisé comme ancre
            try:
                if anchor is not None:
                    _remove_paragraph(anchor)
            except Exception:
                pass

            _remove_paragraph(p)
            continue

        # ------- Texte simple (nom, titre, contact) -------
        if isinstance(value, str):
            run = p.add_run(value)
            if ph == "%%FULL_NAME%%":
                run.bold = True
                run.font.size = Pt(20)
            elif ph == "%%CV_TITLE%%":
                run.bold = True
                run.font.size = Pt(12)
            elif ph == "%%CONTACT_LINE%%":
                run.font.size = Pt(10)
            continue

        # ------- Listes classiques (si jamais) -------
        _insert_lines_after(p, value or [], make_bullets=True)

    doc.save(out_path)

def write_pdf_simple(cv_text: str, out_path: str) -> None:
    c = canvas.Canvas(out_path, pagesize=A4)
    width, height = A4
    x = 45
    y = height - 55
    line_height = 14

    for raw in cv_text.splitlines():
        line = raw.strip("\n")
        if y < 60:
            c.showPage()
            y = height - 55
        c.drawString(x, y, line[:120])  # coupe sécurité
        y -= line_height

    c.save()

def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> None:
    """
    Convertit un DOCX en PDF via LibreOffice.
    Rend le PDF identique au template Word.
    """
    import os
    import subprocess
    import shutil

    out_dir = os.path.dirname(pdf_path) or "."
    os.makedirs(out_dir, exist_ok=True)

    # Sur Linux/Docker, la commande peut être "soffice" ou "libreoffice"
    cmd = None
    for candidate in ["soffice", "libreoffice"]:
        if shutil.which(candidate):
            cmd = candidate
            break
    if not cmd:
        raise RuntimeError("LibreOffice/soffice introuvable dans l'environnement.")

    subprocess.run(
        [
            cmd,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to", "pdf",
            "--outdir", out_dir,
            docx_path,
        ],
        check=True,
    )

    generated_pdf = os.path.join(
        out_dir,
        os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    )

    if generated_pdf != pdf_path:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        os.rename(generated_pdf, pdf_path)

def make_download_urls(job_id: str) -> Dict[str, str]:
    return {
        "pdf": f"{PUBLIC_BASE_DOWNLOAD}/download/{job_id}/cv.pdf",
        "docx": f"{PUBLIC_BASE_DOWNLOAD}/download/{job_id}/cv.docx",
    }

@app.get("/quota")
def quota_check(email: str):
    email = email.strip().lower()
    if not email:
        raise HTTPException(status_code=400, detail="Email manquant.")
    if has_free_left(email):
        return {"ok": True, "message": "✅ Tu as encore ton CV gratuit ce mois-ci."}
    return {"ok": True, "message": "ℹ️ Ton CV gratuit du mois est déjà utilisé. Le prochain sera payant."}

@app.post("/start")
async def start(payload: Dict[str, Any]):

    required = ["email", "sector", "company", "role", "job_posting", "full_name", "city", "phone"]

    for k in required:
        if not payload.get(k):
            raise HTTPException(status_code=400, detail=f"Champ manquant: {k}")

    email = payload["email"].strip().lower()
    current_month = month_key()

    # Vérifie si CV gratuit disponible
    if has_free_left(email):

        with db_conn() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """INSERT INTO quota (email, month)
                       VALUES (%s, %s)
                       ON CONFLICT (email)
                       DO UPDATE SET month = EXCLUDED.month""",
                    (email, current_month)
                )
            conn.commit()

        job_id = await generate_and_store(payload)
        return {"mode": "free", "downloads": make_download_urls(job_id)}

    # Sinon paiement obligatoire
    raise HTTPException(
        status_code=402,
        detail="CV gratuit déjà utilisé. Paiement requis."
    )

@app.post("/confirm_paid")
async def confirm_paid(payload: Dict[str, Any]):
    # appelé par le front après retour Stripe success
    job_id = payload.get("job_id")
    if not job_id or job_id not in jobs:
        raise HTTPException(status_code=400, detail="job_id invalide.")
    if jobs[job_id].get("pdf_path"):
        return {"ok": True, "downloads": make_download_urls(job_id)}

    stored = jobs[job_id].get("payload")
    if not stored:
        raise HTTPException(status_code=400, detail="Payload introuvable.")
    job_id = await generate_and_store(stored, job_id=job_id)
    return {"ok": True, "downloads": make_download_urls(job_id)}

@app.get("/download/{job_id}/{filename}")
def download(job_id: str, filename: str):
    from fastapi.responses import FileResponse
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Inconnu.")
    if filename == "cv.pdf":
        path = jobs[job_id].get("pdf_path")
    elif filename == "cv.docx":
        path = jobs[job_id].get("docx_path")
    else:
        raise HTTPException(status_code=404, detail="Fichier inconnu.")
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Fichier non prêt.")
    return FileResponse(path, filename=filename)

async def generate_and_store(payload: Dict[str, Any], job_id: Optional[str] = None) -> str:
    job_id = job_id or str(uuid.uuid4())
    os.makedirs("out", exist_ok=True)

    cv_text = generate_cv_text(payload)

    safe = sanitize_filename(payload["full_name"])
    docx_path = os.path.join("out", f"{safe}_{job_id}.docx")
    pdf_path = os.path.join("out", f"{safe}_{job_id}.pdf")

    tpl = sector_to_template(payload["sector"])
    write_docx_from_template(tpl, cv_text, docx_path, payload=payload)
    convert_docx_to_pdf(docx_path, pdf_path)

    jobs[job_id] = {"docx_path": docx_path, "pdf_path": pdf_path, "payload": payload}
    return job_id
import psycopg2
from psycopg2.extras import RealDictCursor
import psycopg2
import os

@app.get("/_setup_db")
def setup_db():
    DATABASE_URL = os.getenv("DATABASE_URL")
    if not DATABASE_URL:
        return {"error": "DATABASE_URL not configured"}

    conn = psycopg2.connect(DATABASE_URL)
    cur = conn.cursor()

    cur.execute("""
    DROP TABLE IF EXISTS quota;

    CREATE TABLE quota (
        email TEXT PRIMARY KEY,
        month TEXT NOT NULL
    );
    """)

    conn.commit()
    cur.close()
    conn.close()

    return {"ok": True, "message": "Table quota créée proprement"}
import os
import psycopg2
from fastapi import HTTPException

DATABASE_URL = os.getenv("DATABASE_URL", "")

def db_conn():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL manquant")
    return psycopg2.connect(DATABASE_URL)

@app.get("/_debug_quota_columns")
def debug_quota_columns():
    try:
        with db_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT column_name, data_type
                    FROM information_schema.columns
                    WHERE table_name = 'quota'
                    ORDER BY ordinal_position;
                """)
                rows = cur.fetchall()
        return {"columns": [{"name": r[0], "type": r[1]} for r in rows]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
