import os
import re
import uuid
import datetime as dt
from typing import Optional, Dict, Any
import glob 
import json

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, FileResponse
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

from pypdf import PdfReader

def clean_cv_output(cv_text: str) -> str:
    if not cv_text:
        return ""
    lines = cv_text.replace("\r\n", "\n").split("\n")
    out = []
    for ln in lines:
        s = ln.strip()
        if not s:
            out.append(ln)
            continue
        if s in {".", "..", "...", "\"", "''", "\"\"", "\"\"\""}:
            continue
        low = s.lower()
        if low.startswith("cette version") or low.startswith("ce cv") or low.startswith("note :"):
            continue
        out.append(ln)
    return "\n".join(out).strip()

REQUIRED_SECTIONS = ["EDUCATION:", "EXPERIENCES:", "SKILLS:", "LANGUAGES:", "ACTIVITIES:"]

def has_all_sections(cv_text: str) -> bool:
    t = (cv_text or "")
    return all(sec in t for sec in REQUIRED_SECTIONS)

def safe_apply_llm_edit(old_text: str, new_text: str) -> str:
    """
    Si l'IA renvoie un CV cassé (sections manquantes, etc.),
    on garde l'ancien pour éviter de tout péter.
    """
    new_clean = clean_cv_output(new_text)
    if not has_all_sections(new_clean):
        return old_text  # on refuse la sortie cassée
    return new_clean

def pdf_page_count(pdf_path: str) -> int:
    reader = PdfReader(pdf_path)
    return len(reader.pages)

def pdf_fill_ratio_first_page(pdf_path: str) -> float:
    """
    Heuristique simple : nombre de lignes non vides extraites de la page 1.
    Sert à détecter "trop vide" (beaucoup d'espace en bas).
    """
    reader = PdfReader(pdf_path)
    if len(reader.pages) == 0:
        return 0.0
    page = reader.pages[0]
    try:
        text = page.extract_text() or ""
    except Exception:
        text = ""
    if not text.strip():
        return 0.0
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    n = len(lines)

    # calibrage simple
    if n <= 22:
        return 0.60
    if n >= 55:
        return 0.95
    return 0.60 + (n - 22) * (0.35 / (55 - 22))

def llm_shrink_cv(cv_text: str) -> str:
    if not client:
        return cv_text

    prompt = f"""
Tu dois rendre ce CV PLUS COURT pour tenir sur 1 page Word, SANS le casser.

Règles ABSOLUES :
- Tu gardes exactement les sections : EDUCATION:, EXPERIENCES:, SKILLS:, LANGUAGES:, ACTIVITIES:
- Tu ne rajoutes AUCUN commentaire ni phrase méta.
- Tu ne coupes JAMAIS une phrase.
- Tu n'utilises JAMAIS "..." ni de guillemets triples.
- Tu n'inventes rien : pas de nouvelles missions, chiffres, outils.
- Tu peux uniquement :
  1) raccourcir les bullets (phrases plus directes),
  2) réduire DETAILS dans EDUCATION (1-2 lignes max par diplôme),
  3) réduire ACTIVITIES (max 2 activités, une ligne chacune),
  4) limiter à 2 bullets les expériences secondaires (garder 3 bullets pour l'expérience la plus pertinente).

Sortie : UNIQUEMENT le CV complet.

CV :
\"\"\"{cv_text}\"\"\"
"""
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.choices[0].message.content.strip()

def llm_expand_cv(cv_text: str) -> str:
    if not client:
        return cv_text

    prompt = f"""
Tu dois rendre ce CV PLUS DENSE pour remplir correctement 1 page Word (éviter un grand vide en bas).

Règles ABSOLUES :
- Tu gardes exactement les sections : EDUCATION:, EXPERIENCES:, SKILLS:, LANGUAGES:, ACTIVITIES:
- Tu ne rajoutes AUCUN commentaire ni phrase méta.
- Tu ne coupes JAMAIS une phrase.
- Tu n'utilises JAMAIS "..." ni de guillemets triples.
- Tu n'inventes rien : pas de nouvelles missions, chiffres, outils.
- Tu peux uniquement :
  1) ajouter 1 bullet à l'expérience la plus pertinente (si elle n'en a que 2),
  2) préciser légèrement 1-2 bullets (sans inventer),
  3) ajouter 1 ligne utile dans EDUCATION (si déjà suggérée dans les infos),
  4) enrichir 1 activité forte (toujours une ligne).

Sortie : UNIQUEMENT le CV complet.

CV :
\"\"\"{cv_text}\"\"\"
"""
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.choices[0].message.content.strip()

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

    try:
        conn = psycopg2.connect(os.getenv("DATABASE_URL"))
        cur = conn.cursor()
        cur.execute("SELECT month FROM quota WHERE email = %s", (email,))
        row = cur.fetchone()
        cur.close()
        conn.close()
    except Exception as e:
        raise HTTPException(status_code=503, detail="DB unavailable")

    if not row:
        return True
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
- Chaque expérience doit contenir 3-4 bullet points orientés résultats. N’invente jamais de chiffres : si aucun chiffre n’est fourni, reste qualitatif.
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
- 2 à 3 bullets maximum par expérience (3 par défaut, 2 uniquement pour les expériences les moins pertinentes).
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
- Tu DOIS les mettre dans "EXPERIENCES" (même si ce n’est pas une entreprise).
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
  Nom de l’activité en gras (nous ferons le gras côté Word), suivi de ":" puis une phrase :
  - ce que la personne a fait concrètement (niveau / fréquence / contexte),
  - ce que ça développe comme qualités utiles en finance / environnement exigeant.

IMPORTANT :
- Toute la sortie (EDUCATION, EXPERIENCES, SKILLS, LANGUAGES, ACTIVITIES)
  doit être rédigée EN FRANÇAIS.
- Si tu écris une phrase en anglais, tu la traduis immédiatement en français.
- Seuls les noms propres (noms d’écoles, diplômes officiels, logiciels, intitulés exacts de postes)
  peuvent rester en anglais.

RÈGLES DE SORTIE (TRÈS IMPORTANT) :
- Tu génères UNIQUEMENT les sections suivantes, dans cet ordre :
  EDUCATION:
  EXPERIENCES:
  SKILLS:
  LANGUAGES:
  ACTIVITIES:
- Tu ne génères PAS de titre de section supplémentaire.
- Tu ne génères PAS le nom.
- Tu ne génères PAS les coordonnées.
- Tu ne génères PAS d'accroche.

FORMAT EXACT À RESPECTER :

1️⃣ TU DOIS ABSOLUMENT PRODUIRE CES 5 BLOCS DANS CET ORDRE EXACT,
   CHAQUE EN-TÊTE SUR SA PROPRE LIGNE :

EDUCATION:
<contenu éducation>

EXPERIENCES:
<contenu expériences>

SKILLS:
<contenu compétences>

LANGUAGES:
<contenu langues>

ACTIVITIES:
<contenu activités>

2️⃣ TU NE DOIS RIEN ÉCRIRE AVANT "EDUCATION:" NI APRÈS LA DERNIÈRE LIGNE D’ACTIVITIES.
   PAS DE COMMENTAIRES, PAS DE TEXTE EXPLICATIF, PAS D’INTRODUCTION, RIEN.

3️⃣ FORMAT PRÉCIS DE CHAQUE BLOC :

EDUCATION:
DEGREE: <intitulé du diplôme ou programme>
SCHOOL: <école ou université>
LOCATION: <Ville, Pays>
DATES: <MMM YYYY – MMM YYYY ou MMM YYYY – Present>
DETAILS:
- <ligne de détail 1 (ex : Matières fondamentales : ... )>
- <ligne de détail 2>
- <ligne de détail 3>

DEGREE: <autre diplôme ou échange académique>
SCHOOL: <école ou université>
LOCATION: <Ville, Pays>
DATES: <MMM YYYY – MMM YYYY ou MMM YYYY – Present>
DETAILS:
- <détail 1>
- <détail 2>

EXPERIENCES:
ROLE: <intitulé exact du poste>
COMPANY: <nom exact de l’entreprise ou de l’association>
DATES: <MMM YYYY – MMM YYYY ou MMM YYYY – Present>
LOCATION: <Ville, Pays>
TYPE: <Internship / Alternance / CDI / etc. si fourni sinon vide>
BULLETS:
- <bullet 1>
- <bullet 2>
- <bullet 3>

ROLE: <autre poste>
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
<une seule ligne ou plusieurs, mais sans puces>

ACTIVITIES:
<une activité par ligne, sans puces, sous la forme "Nom de l’activité : description">

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
- Tu enrichis un peu la section EDUCATION (matières clés, spécialisation, classement si fourni).
- Tu développes un peu les ACTIVITIES les plus fortes (sport intensif, voyages marquants, engagement régulier).

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

    # 1er appel : génération normale
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )

    cv_text = resp.choices[0].message.content.strip()

    # Nettoyage final robuste (enlève les phrases meta, les """ etc.)
    cv_text = clean_cv_output(cv_text)

    print("=== RAW CV TEXT ===")
    print(cv_text)
    print("=== END RAW CV TEXT ===")

    return cv_text

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
    if p is None:
        return
    el = getattr(p, "_element", None)
    if el is None:
        return
    parent = el.getparent()
    if parent is None:
        return
    parent.remove(el)
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
        elif mode == "bullets" and line.lstrip().startswith("-"):
            # On tolère les espaces avant le tiret ("  - bullet")
            stripped = line.lstrip()
            bullet_text = stripped[1:].strip()
            if bullet_text:
                bullet_text = re.sub(r"(?i)^participé\s+à\s+", "Contribué à ", bullet_text)
                bullet_text = re.sub(r"(?i)^aidé\s+à\s+", "Soutenu ", bullet_text)
                
                cur["bullets"].append(bullet_text)

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
def parse_education_structured(lines: list[str]) -> list[dict]:
    """
    Parse une section EDUCATION structurée avec les tags :
    DEGREE:, SCHOOL:, LOCATION:, DATES:, DETAILS:
    """
    programs = []
    cur = None
    mode = None

    def push():
        nonlocal cur
        if cur and (cur.get("degree") or cur.get("school")):
            # On s'assure d'avoir toujours une liste de détails
            cur.setdefault("details", [])
            programs.append(cur)
        cur = None

    for raw in (lines or []):
        line = (raw or "").strip()
        if not line:
            continue

        if line.startswith("DEGREE:"):
            push()
            cur = {
                "degree": line.replace("DEGREE:", "").strip(),
                "school": "",
                "location": "",
                "dates": "",
                "details": [],
            }
            mode = None
            continue

        if not cur:
            continue

        if line.startswith("SCHOOL:"):
            cur["school"] = line.replace("SCHOOL:", "").strip()
        elif line.startswith("LOCATION:"):
            cur["location"] = line.replace("LOCATION:", "").strip()
        elif line.startswith("DATES:"):
            cur["dates"] = line.replace("DATES:", "").strip()
        elif line.startswith("DETAILS:"):
            mode = "details"
        elif mode == "details" and line.lstrip().startswith("-"):
            txt = line.lstrip()[1:].strip()
            if txt:
                cur["details"].append(txt)

    push()
    return programs
def _no_space_len(s: str) -> int:
    """Longueur d'un texte sans compter les espaces."""
    return len(re.sub(r"\s+", "", s or ""))

def shorten_experience_bullets_with_llm(
    exps: list[dict],
    max_no_space_per_bullet: int = 90,
) -> list[dict]:
    """
    RÉÉCRIT les bullets via l'API pour qu'elles soient plus courtes,
    SANS changer le sens, SANS inventer, SANS '...'
    et SANS JAMAIS changer le nombre de bullets.

    Si l'IA ne respecte pas ça -> on garde la version ORIGINALE.
    """
    if not client:
        return exps  # pas d'API dispo -> on ne touche rien

    # On prépare une version simplifiée pour l'IA
    simple_exps = []
    for e in exps:
        simple_exps.append({
            "role": e.get("role", ""),
            "company": e.get("company", ""),
            "bullets": e.get("bullets", []),
        })

    payload = {
        "max_no_space": max_no_space_per_bullet,
        "experiences": simple_exps,
    }

    prompt = f"""
Tu es recruteur en finance.

On te donne des expériences avec leurs bullet points au format JSON.

POUR CHAQUE BULLET :
- tu réécris la phrase en français,
- tu gardes exactement le même sens (aucune nouvelle mission, aucun nouveau chiffre, aucun nouvel outil),
- tu gardes la structure : verbe d'action + moyen + résultat,
- INTERDIT de commencer par : "Participé", "Aidé", "Effectué", "Travaillé",
- la phrase est complète et se termine par un point,
- maximum {max_no_space_per_bullet} caractères SANS espaces,
- JAMAIS de points de suspension ("...").

INTERDIT :
- changer le nombre d'expériences,
- changer le nombre de bullets,
- réordonner les bullets,
- inventer des éléments.

Tu dois renvoyer UNIQUEMENT un JSON de la forme :

{{"experiences": [
  {{"bullets": ["...", "..."]}},
  {{"bullets": ["...", "..."]}}
]}}

Voici le JSON d'entrée :

{json.dumps(payload, ensure_ascii=False)}
"""

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    content = resp.choices[0].message.content.strip()

    try:
        data = json.loads(content)
        new_exps = data.get("experiences", [])
    except Exception:
        # Si l'IA ne répond pas en JSON -> on garde TOUT tel quel
        return exps

    # Sécurité maximale : si la longueur ne colle pas, on ne change RIEN
    if not isinstance(new_exps, list) or len(new_exps) != len(exps):
        return exps

    result: list[dict] = []

    for old, new in zip(exps, new_exps):
        if not isinstance(new, dict):
            result.append(old)
            continue

        old_bullets = old.get("bullets") or []
        new_bullets = new.get("bullets")

        if not isinstance(new_bullets, list):
            result.append(old)
            continue

        # ⚠️ SI le nombre de bullets ne correspond pas -> on garde l'ancien
        if len(new_bullets) != len(old_bullets):
            result.append(old)
            continue

        cleaned_bullets = []
        invalid = False
        for b in new_bullets:
            txt = (b or "").strip()
            if not txt:
                invalid = True
                break
            cleaned_bullets.append(txt)
        if invalid:
            result.append(old)
            continue

        updated = dict(old)
        updated["bullets"] = cleaned_bullets
        result.append(updated)

    # Par sécurité : si on a un souci, on renvoie l'original
    if len(result) != len(exps):
        return exps

    return result

def trim_finance_experiences(
    exps: list[dict],
    is_cv_long: bool = True,
    max_experiences: int = 4,
    max_total_bullets: int = 8,
    min_experiences: int = 2,
    max_no_space_per_bullet: int = 90,
) -> list[dict]:
    """
    NOUVELLE VERSION :
    - NE SUPPRIME PLUS JAMAIS d'expérience.
    - NE SUPPRIME PLUS JAMAIS de bullet.
    - Se contente de nettoyer les vides et, si le CV est long,
      de faire RÉÉCRIRE les bullets via l'API pour les raccourcir.
    """

    # 1) Nettoyage des expériences VRAIMENT vides
    cleaned: list[dict] = []
    for e in exps:
        role = (e.get("role") or "").strip()
        bullets = [b for b in (e.get("bullets") or []) if (b or "").strip()]
        if not role and not bullets:
            continue  # là c'est du vide total, ça ne sert à rien
        e["role"] = role
        e["bullets"] = bullets
        cleaned.append(e)

    if not cleaned:
        return []

    # 2) Si le CV n'est pas long -> on NE TOUCHE À RIEN
    if not is_cv_long:
        return cleaned

    # 3) Si le CV est long -> on raccourcit PAR RÉÉCRITURE (pas par suppression)
    cleaned = shorten_experience_bullets_with_llm(
        cleaned,
        max_no_space_per_bullet=max_no_space_per_bullet,
    )

    return cleaned

def shorten_activities_with_llm(
    lines: list[str],
    max_no_space_per_activity: int = 90,
) -> list[str]:
    """
    Réécrit chaque activité pour qu'elle tienne sur une ligne,
    phrase complète, sans '...', SANS jamais changer le NOMBRE d'activités.

    Si l'IA ne respecte pas ça -> on garde les lignes d'origine.
    """
    if not client:
        return lines

    activities = [(l or "").strip() for l in lines if (l or "").strip()]
    if not activities:
        return []

    payload = {
        "max_no_space": max_no_space_per_activity,
        "activities": activities,
    }

    prompt = f"""
Tu es recruteur en finance.

On te donne une liste d'activités / centres d'intérêt.

POUR CHAQUE ACTIVITÉ :
- tu gardes UNE activité par ligne (pas de fusion),
- tu réécris en français en gardant le sens,
- style CV (PAS de "je", PAS de "nous", PAS de phrase à la première personne),
- formulation orientée finance : pratique + discipline / exigence / rigueur,
- tu fais une phrase complète qui se termine par un point,
- tu ne mets JAMAIS de points de suspension ("..."),
- la phrase doit faire au maximum {max_no_space_per_activity} caractères SANS espaces.

INTERDIT :
- changer le nombre d'activités,
- fusionner plusieurs activités en une seule.

Réponds UNIQUEMENT avec un JSON de la forme :
{{"activities": ["Activité 1 : ...", "Activité 2 : ...", ...]}}

Voici les activités :

{json.dumps(payload, ensure_ascii=False)}
"""

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    content = resp.choices[0].message.content.strip()

    try:
        data = json.loads(content)
        new_acts = data.get("activities", [])
    except Exception:
        # L'IA n'a pas répondu comme il faut -> on garde tout
        return activities

    # Sécurité : si le nombre ne correspond pas -> on garde l'original
    if not isinstance(new_acts, list) or len(new_acts) != len(activities):
        return activities

    out: list[str] = []
    for old, new in zip(activities, new_acts):
        txt = (new or "").strip()
        if not txt:
            # si une activité disparaît -> on annule tout
            return activities
        out.append(txt)

    return out


def trim_activities(
    lines: list[str],
    cv_is_long: bool,
    ideal_max: int = 2,
    max_no_space_per_activity: int = 90,
) -> list[str]:
    """
    ACTIVITÉS / CENTRES D'INTÉRÊT :

    - On nettoie les lignes vides.
    - Si le CV n'est pas long -> on garde tout.
    - Si le CV est long -> on garde au maximum `ideal_max` activités
      (les premières, normalement les plus pertinentes),
      puis on les fait reformuler plus court par l'IA.
    """
    cleaned = [(l or "").strip() for l in (lines or []) if (l or "").strip()]
    if not cleaned:
        return []

    if not cv_is_long:
        return cleaned

    # CV long : on raccourcit les activités, mais on ne supprime pas
    return shorten_activities_with_llm(
        cleaned,
        max_no_space_per_activity=max_no_space_per_activity,
    )

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
        text = re.sub(r"(?i)^je\s+", "", text).strip()
        text = re.sub(r"(?i)^j['’]\s*", "", text).strip()
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
    is_first = True  # ✅ pour ajouter un petit espace avant la 1ère ligne

    for raw in (lines or []):
        text = (raw or "").strip()
        if not text:
            last = _insert_paragraph_after(last, "")
            continue

        # On remplace les ' | ' par des virgules si jamais le modèle en met encore
        text = text.replace(" | ", ", ")

        new_p = _insert_paragraph_after(last, "")

        # ✅ petit espace juste au début de la section
        if is_first:
            is_first = False
        
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
        "Échange académique", "Echange académique", "Exchange programme", "Exchange program",
        "BBA", "Bachelor"
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

def collapse_blank_paragraphs(doc: Document, max_consecutive: int = 1):
    """
    Supprime les paragraphes vraiment vides,
    MAIS conserve ceux qui servent d'espacement (space_before/space_after > 0).
    """
    blanks = 0

    for p in list(doc.paragraphs):
        txt = (p.text or "")
        fmt = p.paragraph_format

        has_spacing = bool(
            (fmt.space_before and fmt.space_before.pt > 0) or
            (fmt.space_after and fmt.space_after.pt > 0)
        )

        is_blank_text = (txt.strip() == "")
        is_blank = is_blank_text and not has_spacing

        if is_blank:
            blanks += 1
            if blanks > max_consecutive:
                _remove_paragraph(p)
        else:
            blanks = 0
            
def write_docx_from_template(template_path: str, cv_text: str, out_path: str, payload: dict = None) -> None:
    doc = Document(template_path)

    # On mesure la longueur du texte pour savoir si on doit "tailler" ou pas.
    raw_text = cv_text or ""
    nb_lines = raw_text.count("\n") + 1  # nombre de lignes brutes

    # Longueur SANS espaces (celle que tu mesures dans Word)
    chars_no_space = len(re.sub(r"\s+", "", raw_text))

    # Au-delà d’environ 2225 caractères sans espaces → CV considéré comme "long"
    cv_is_long = (chars_no_space > 2225) or (nb_lines > 85)

    # Marges plus petites pour mieux utiliser la largeur
    for section in doc.sections:
        section.left_margin = Cm(1.0)
        section.right_margin = Cm(1.0)
        section.top_margin = Cm(1.0)      
        section.bottom_margin = Cm(1.0)   

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
        interests_value = trim_activities(interests_raw, cv_is_long=cv_is_long)
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
            # paragraphe vide "réel" pour espacer le titre de section et le contenu
            spacer = _insert_paragraph_after(p, "")
            spacer.paragraph_format.space_after = Pt(2)
            spacer.paragraph_format.space_before = Pt(0)
            
            _render_skills(spacer, value or [])
            _remove_paragraph(p)
            continue

        # ------- ACTIVITÉS / CENTRES D'INTÉRÊT -------
        if ph == "%%INTERESTS%%" and isinstance(value, list):
            _render_interests(p, value or [])
            _remove_paragraph(p)
            continue

        # ------- FORMATION -------
        if ph == "%%EDUCATION%%" and isinstance(value, list):

            # 🔹 CAS 1 : format structuré avec DEGREE:/SCHOOL:/LOCATION:/DATES:/DETAILS:
            if any((line or "").strip().startswith("DEGREE:") for line in value):
                programs = parse_education_structured(value)
                anchor = p
                first_edu = True

                for idx, edu in enumerate(programs):
                    degree = (edu.get("degree") or "").strip()
                    school = (edu.get("school") or "").strip()
                    location = (edu.get("location") or "").strip()
                    dates = (edu.get("dates") or "").strip()
                    details = edu.get("details") or []

                    # Création du tableau 2 colonnes
                    table = _add_table_after(anchor, rows=1, cols=2)
                    
                    left = table.cell(0, 0)
                    right = table.cell(0, 1)
                    left.text = ""
                    right.text = ""

                    # ---- Colonne gauche : diplôme + école + détails ----
                    lp = left.paragraphs[0]
                    try:
                        lp.style = doc.styles["Normal"]
                    except Exception:
                        pass
                    lp.paragraph_format.left_indent = Pt(0)
                    lp.paragraph_format.first_line_indent = Pt(0)

                    mention_value = ""
                    deg_low = (degree or "").lower()
                    if "mention" in deg_low:
                        parts = [p.strip() for p in degree.split("–")]
                        kept = []
                        for part in parts:
                            if part.lower().startswith("mention"):
                                mention_value = part.replace("Mention", "").strip()
                            else:
                                kept.append(part)
                        degree = " – ".join(kept).strip()

                    title_parts = [x for x in [degree, school] if x]
                    title_line = " – ".join(title_parts) if title_parts else (degree or school)

                    if title_line:
                        r_title = lp.add_run(title_line)
                        r_title.bold = True
                        r_title.font.size = Pt(11)

                    if mention_value:
                        para_m = left.add_paragraph()
                        try:
                            para_m.style = doc.styles["Normal"]
                        except Exception:
                            pass
                        r1 = para_m.add_run("Mention :")
                        r1.underline = True
                        r1.font.size = Pt(10)
                        r2 = para_m.add_run(" " + mention_value)
                        r2.font.size = Pt(10)

                    # Détails sous le titre
                    for d in details:
                        text = (d or "").strip()
                        if not text:
                            continue
                    
                        # ✅ On supprime BDE/Association dans EDUCATION (car ça va dans EXPERIENCES)
                        low = text.lower()
                        if "bde" in low or low.startswith("association"):
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
                        lower = text.lower()

                        # Fix orthographe/accord
                        text = text.replace("Analyse financières", "Analyse financière")
                        lower = text.lower()
                        
                        # ✅ 1) Projets (avec ou sans ":") => label souligné
                        if lower.startswith("projets"):
                            label_text = "Projets"
                            after_text = re.sub(r"(?i)^projets(\s+de\s+groupe)?\s*", "", text).strip()
                            # enlève les ponctuations type ": :"
                            after_text = re.sub(r"^[\s:–-]+", "", after_text).strip()
                        
                        # ✅ 2) "Cours en ..." => Matières fondamentales
                        elif re.match(r"(?i)^cours\s+en\s+", text):
                            label_text = "Matières fondamentales"
                            after_text = re.sub(r"(?i)^cours\s+en\s+", "", text).strip().rstrip(".")
                        elif lower.startswith("cours") and ":" in text:
                            label_text = "Matières fondamentales"
                            _, _, after = text.partition(":")
                            after_text = after.strip().rstrip(".")
                        
                        # ✅ 3) Matières fondamentales / cours pertinents / key coursework
                        elif "matières fondamentales" in lower or "cours pertinents" in lower or "key coursework" in lower:
                            label_text = "Matières fondamentales"
                            if ":" in text:
                                _, _, after = text.partition(":")
                                after_text = after.strip()
                        
                        # ✅ 4) Autres labels courts "X: Y"
                        elif ":" in text:
                            before, sep, after = text.partition(":")
                            before_clean = before.strip()
                            if len(before_clean.split()) <= 4:
                                label_text = before_clean
                                after_text = after.strip()

                        if label_text:
                            r1 = para.add_run(label_text + " :")
                            r1.underline = True
                            r1.font.size = Pt(10)
                            if after_text and after_text.strip():
                                r2 = para.add_run(" " + after_text.strip())
                                r2.font.size = Pt(10)
                        else:
                            r = para.add_run(text)
                            r.font.size = Pt(10)

                    # ---- Colonne droite : dates + lieu ----
                    rp = right.paragraphs[0]
                    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    rp.paragraph_format.space_after = Pt(0)

                    if dates:
                        clean_date = dates.replace("\r", " ").replace("\n", " ")
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

                    # Paragraphe vide pour ancrer la prochaine formation
                    new_p_elt = OxmlElement("w:p")
                    table._tbl.addnext(new_p_elt)
                    anchor = Paragraph(new_p_elt, p._parent)
                    
                    # ✅ spacer uniquement ENTRE les formations (pas après la dernière)
                    if idx < len(programs) - 1:
                        anchor.paragraph_format.space_after = Pt(2)
                        anchor.paragraph_format.space_before = Pt(0)

                # ✅ pas de paragraphe vide supplémentaire, on garde juste l'ancre
                anchor.paragraph_format.space_after = Pt(2)
                
                _remove_paragraph(p)
                continue

            # 🔹 CAS 2 : ancien format libre (on garde ton ancien comportement)
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

                date_range_patterns = [
                    r"(Janv|Fév|Fev|Mars|Avr|Mai|Juin|Juil|Août|Aout|Sept|Oct|Nov|Déc|Dec)\s+\d{4}\s*[–-]\s*(Janv|Fév|Fev|Mars|Avr|Mai|Juin|Juil|Août|Aout|Sept|Oct|Nov|Déc|Dec)\s+\d{4}\s*$",
                    r"(0[1-9]|1[0-2])/\d{4}\s*[–-]\s*(0[1-9]|1[0-2])/\d{4}\s*$",
                    r"(19|20)\d{4}?\s*[–-]\s*(19|20)\d{4}?\s*$"
                ]

                for pat in date_range_patterns:
                    m = re.search(pat, first_line)
                    if m:
                        date_part = m.group(0).strip()
                        title_part = first_line[:m.start()].rstrip(" ,–-").strip()
                        break

                if not date_part:
                    for sep in ("–", "—", "-"):
                        idx = first_line.rfind(sep)
                        if idx != -1:
                            title_part = first_line[:idx].strip()
                            date_part = first_line[idx + 1:].strip()
                            break

                if date_part:
                    m = re.search(r"(19|20)\d{2}\s*$", title_part)
                    if m:
                        title_part = title_part[:m.start()].rstrip(" ,–-")

                table = _add_table_after(anchor, rows=1, cols=2)
                try:
                    _remove_paragraph(anchor)
                except Exception:
                    pass
                left = table.cell(0, 0)
                right = table.cell(0, 1)
                left.text = ""
                right.text = ""

                lp = left.paragraphs[0]
                try:
                    lp.style = doc.styles["Normal"]
                except Exception:
                    pass
                lp.paragraph_format.left_indent = Pt(0)
                lp.paragraph_format.first_line_indent = Pt(0)

                # ✅ Si "Mention ..." est dans le titre, on la sort pour la mettre en dessous
                mention_value = ""
                if "mention" in title_part.lower():
                    parts = [p.strip() for p in title_part.split("–")]
                    kept = []
                    for part in parts:
                        if part.lower().startswith("mention"):
                            mention_value = part.replace("Mention", "").strip()
                        else:
                            kept.append(part)
                    title_part = " – ".join(kept).strip()
                
                title_run = lp.add_run(title_part)
                title_run.bold = True
                title_run.font.size = Pt(11)
                
                # ✅ Ligne en dessous : Mention : (souligné)
                if mention_value:
                    para = left.add_paragraph()
                    try:
                        para.style = doc.styles["Normal"]
                    except Exception:
                        pass
                    r1 = para.add_run("Mention :")
                    r1.underline = True
                    r1.font.size = Pt(10)
                    r2 = para.add_run(" " + mention_value)
                    r2.font.size = Pt(10)

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

                new_p_elt = OxmlElement("w:p")
                table._tbl.addnext(new_p_elt)
                anchor = Paragraph(new_p_elt, p._parent)
                
                # ✅ spacer "réel" (pas vide) => espacement visible entre formations
                anchor.paragraph_format.space_after = Pt(2)
                anchor.paragraph_format.space_before = Pt(0)

            try:
                if anchor is not None:
                    _remove_paragraph(anchor)
            except Exception:
                pass

            _remove_paragraph(p)
            continue

        # ------- EXPÉRIENCES PROFESSIONNELLES -------
        if ph == "%%EXPERIENCE%%":
            # On parse les expériences au format structuré ROLE/COMPANY/DATES/...
            exps = parse_finance_experiences(value or [])
            exps = trim_finance_experiences(exps, is_cv_long=cv_is_long)
            anchor = p
            first_table = True

            # Si jamais le modèle n'a pas respecté le format structuré,
            # on retombe sur un simple rendu en liste pour ne pas tout casser.
            if not exps:
                _insert_lines_after(p, value or [], make_bullets=True)
                continue

            # Mots-clés qui correspondent plutôt à un type de contrat qu'à un vrai rôle
            CONTRACT_PREFIXES = [
                "stagiaire", "stage",
                "summer job", "part-time job", "student job",
                "volunteering", "volunteer",
                "internship", "intern", "traineeship",
                "apprenticeship",
                "full-time", "full time",
                "part-time", "part time",
            ]

            for idx, exp in enumerate(exps):
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

                # 2) Si le rôle commence encore par un type de contrat (hors "en ..."),
                #    on enlève juste ce préfixe, mais on garde la suite.
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

                # ✅ spacer AVANT la 1ère expérience (sépare le titre section du contenu)
               if first_table:
                    spacer = _insert_paragraph_after(anchor, "")
                    spacer.paragraph_format.space_after = Pt(2)
                    spacer.paragraph_format.space_before = Pt(0)
                    anchor_for_table = spacer
                else:
                    anchor_for_table = anchor
                
                # Tableau 2 colonnes (mêmes tailles qu'avant via _add_table_after)
                table = _add_table_after(anchor_for_table, rows=1, cols=2)
                
                # ✅ On supprime UNIQUEMENT le placeholder la première fois
                if first_table:
                    try:
                        _remove_paragraph(anchor)
                    except Exception:
                        pass
                    first_table = False
                    
                    
                left = table.cell(0, 0)
                right = table.cell(0, 1)
                left.text = ""
                right.text = ""

                # ----- Colonne gauche : rôle + bullets -----
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

                lp.paragraph_format.space_after = Pt(1)

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

                # ----- Colonne droite : dates, lieu, type -----
                rp = right.paragraphs[0]
                rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                rp.paragraph_format.space_after = Pt(0)

                dates_raw = (exp.get("dates") or "").strip()
                if dates_raw:
                    clean_date = dates_raw.replace("\r", " ").replace("\n", " ")
                    clean_date = re.sub(r"\s+", " ", clean_date.strip())
                    clean_date = translate_months_fr(clean_date)
                    clean_date = clean_date.replace(" - ", " – ")
                    clean_date = clean_date.replace(" ", "\u00A0")  # espaces insécables
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
                
                # ✅ spacer uniquement ENTRE les expériences (pas après la dernière)
                if idx < len(exps) - 1:
                    anchor.paragraph_format.space_after = Pt(2)
                    anchor.paragraph_format.space_before = Pt(0)

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
        
    # Nettoyage des paragraphes vides en fin de document pour éviter la page blanche
    try:
        for p in reversed(doc.paragraphs):
            if (p.text or "").strip():
                break
            _remove_paragraph(p)
    except Exception:
        pass
    collapse_blank_paragraphs(doc, max_consecutive=1)
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

    safe = sanitize_filename(payload["full_name"])
    docx_path = os.path.join("out", f"{safe}_{job_id}.docx")
    pdf_path = os.path.join("out", f"{safe}_{job_id}.pdf")

    tpl = sector_to_template(payload["sector"])

    # 1) baseline
    cv_text = generate_cv_text(payload)

    # 2) boucle max 3 essais (baseline + 2 corrections)
    for attempt in range(3):
        write_docx_from_template(tpl, cv_text, docx_path, payload=payload)
        convert_docx_to_pdf(docx_path, pdf_path)

        pages = pdf_page_count(pdf_path)
        fill = pdf_fill_ratio_first_page(pdf_path) if pages == 1 else 0.0
        print("attempt", attempt, "pages", pages, "fill", round(fill, 2))

        # Trop long => reformulation + courte
        if pages > 1:
            cv_text = safe_apply_llm_edit(cv_text, llm_shrink_cv(cv_text))
            continue

        # 1 page mais trop vide => reformulation + dense
        if fill < 0.78:
            cv_text = safe_apply_llm_edit(cv_text, llm_expand_cv(cv_text))
            continue

        # ✅ 1 page et suffisamment rempli
        break

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
