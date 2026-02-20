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
- 1 page maximum.
- Format de dates homogène (MMM YYYY – MMM YYYY).
- Chaque bullet = Verbe fort + Action + Chiffre + Impact business.
- 3 à 5 bullets maximum par expérience.
- Interdiction des mots : assisted, helped, worked on.
- Ton professionnel, précis, sobre.
RÈGLES STRICTES :
Ces règles priment sur toutes les autres instructions.
- Tu n’inventes AUCUN chiffre.
- Tu n’inventes AUCUNE mission.
- Tu n’inventes AUCUN outil.
- Si une information est absente, tu restes général sans ajouter de détails fictifs.
- Si aucun résultat chiffré n’est fourni, tu reformules sans métriques.
- Tu utilises uniquement les informations présentes dans le profil utilisateur.
- Interdiction totale d’inventer pour “améliorer” le CV.
- Si une expérience contient trop peu d'informations,tu la rends professionnelle mais concise,sans extrapolation.
INTERDICTION ABSOLUE D'INVENTER DES CHIFFRES.
Tu n'écris un nombre (% , €, k€, volumes, "5 sponsors", "100 participants", etc.) QUE s'il est présent dans les infos utilisateur.
Si aucun chiffre n'est fourni : reformule sans métrique.

BDE / ASSOCIATIONS / PROJETS ÉTUDIANTS :
- Tu DOIS les mettre dans "EXPÉRIENCES PROFESSIONNELLES" (même si ce n’est pas une entreprise).
- Tu les écris comme une expérience (titre + dates si disponibles + 2-3 bullets).
- INTERDICTION ABSOLUE d’inventer des chiffres : aucun %, aucun volume, aucun "5 sponsors", aucun "100 participants" si ce n’est pas fourni.

ACTIVITIES (CENTRES D’INTÉRÊT) :
- Tu n’y mets QUE des centres d’intérêt / activités personnelles (sport, langues, certifications, hobbies).
- INTERDICTION d’y mettre BDE / associations / projets / expériences (ils vont uniquement dans EXPERIENCES).
- Pas de doublons : si c’est dans EXPERIENCES, tu ne le répètes pas ailleurs.

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
<une seule ligne, avec séparateur " | ">

LANGUAGES:
<contenu>

ACTIVITIES:
<contenu>

CONTRAINTE LONGUEUR :
- Maximum 12 bullet points au total.
- Maximum 4 bullet points par expérience.
- Format concis.
- Pas de phrases longues.

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


def _remove_paragraph(p: Paragraph):
    p._element.getparent().remove(p._element)
    p._p = p._element = None


def _add_table_after(paragraph: Paragraph, rows: int, cols: int):
    # Get the real Document object (works in body, tables, etc.)
    doc = paragraph.part.document
    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Move the table right after the anchor paragraph
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
        if "Cours pertinents" in line:
            line = line.replace("Cours pertinents", "Matières fondamentales")

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

def write_docx_from_template(template_path: str, cv_text: str, out_path: str, payload: dict = None) -> None:
    doc = Document(template_path)

    # Fix margins (some templates have invalid decimal margins that break python-docx tables)
    try:
        for s in doc.sections:
            s.left_margin = Cm(2)
            s.right_margin = Cm(2)
            s.top_margin = Cm(1.5)
            s.bottom_margin = Cm(1.5)
    except Exception:
        pass

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

    # SKILLS en une seule ligne
    if isinstance(sections.get("SKILLS"), list):
        cleaned = [x.strip().lstrip("-").strip() for x in sections["SKILLS"] if x.strip()]
        sections["SKILLS"] = [" | ".join(cleaned)] if cleaned else [""]

    mapping = {
        "%%FULL_NAME%%": full_name,
        "%%CONTACT_LINE%%": contact_line,
        "%%CV_TITLE%%": cv_title,
        "%%EDUCATION%%": sections.get("EDUCATION", []),
        "%%EXPERIENCE%%": sections.get("EXPERIENCES", []),
        "%%SKILLS%%": sections.get("SKILLS", []),
        "%%LANGUAGES%%": sections.get("LANGUAGES", []),
        "%%INTERESTS%%": sections.get("INTERESTS", []) or sections.get("ACTIVITIES", []),
    }

    for ph, value in mapping.items():
        p = _find_paragraph_containing(doc, ph)
        if not p:
            continue

        _clear_paragraph(p)

        # ------- FORMATION : style spécial -------
    if ph == "%%EDUCATION%%" and isinstance(value, list):

        anchor = p
        current_block = []

        # On regroupe les lignes par formation (séparées par ligne vide)
        blocks = []
        for line in value:
            if not line.strip():
                if current_block:
                    blocks.append(current_block)
                    current_block = []
            else:
                current_block.append(line.strip())
        if current_block:
            blocks.append(current_block)

        for block in blocks:

            # --- 1️⃣ Extraire infos principales ---
            first_line = block[0]

            # Tentative d'extraction date + lieu
            # Format attendu: "Master Finance – Université X — Sep 2023 – Jun 2025"
            date_part = ""
            title_part = first_line

            if "—" in first_line:
                parts = first_line.split("—")
                title_part = parts[0].strip()
                date_part = parts[-1].strip()

            # --- 2️⃣ Créer tableau 2 colonnes ---
            table = _add_table_after(anchor, rows=2, cols=2)

            left_top = table.cell(0, 0)
            right_top = table.cell(0, 1)

            left_bottom = table.cell(1, 0)
            right_bottom = table.cell(1, 1)

            # Nettoyer
            left_top.text = ""
            right_top.text = ""
            left_bottom.text = ""
            right_bottom.text = ""

            # --- Ligne 1 droite : dates (italique) ---
            rp = right_top.paragraphs[0]
            rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run_date = rp.add_run(date_part)
            run_date.italic = True

            # --- Ligne 2 droite : ville, pays (si présent dans 2e ligne du bloc) ---
            location = ""
            for line in block:
                if "," in line and "Matières" not in line:
                    location = line
                    break

            if location:
                rp2 = right_bottom.paragraphs[0]
                rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                rp2.add_run(location)

            # --- Colonne gauche : titre en gras ---
            lp = left_top.paragraphs[0]
            run_title = lp.add_run(title_part)
            run_title.bold = True

            # --- Lignes suivantes sous le tableau ---
            last_anchor = table

            for line in block[1:]:

                # Remplacement label
                if "Cours pertinents" in line:
                    after = ""
                    if ":" in line:
                        after = line.split(":", 1)[1]

                    para = _insert_paragraph_after(last_anchor, "")
                    r1 = para.add_run("Matières fondamentales :")
                    r1.underline = True
                    if after.strip():
                        para.add_run(after)
                    last_anchor = para
                    continue

                # Ne pas réafficher la ligne ville si déjà mise à droite
                if line == location:
                    continue

                para = _insert_paragraph_after(last_anchor, line)
                last_anchor = para

            anchor = last_anchor

        _remove_paragraph(p)
        continue

        # ------- EXPERIENCE : tableau premium -------
        if ph == "%%EXPERIENCE%%":
            exps = parse_finance_experiences(value or [])
            anchor = p

            # Si pas au format ROLE/COMPANY/etc, on met en bullets classiques
            if not exps:
                _insert_lines_after(p, value or [], make_bullets=True)
                continue

            for exp in exps:
                title = exp.get("role", "")
                if exp.get("company"):
                    title += f" - {exp['company']}"

                tpara = _insert_paragraph_after(anchor, title)
                tpara.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if tpara.runs:
                    tpara.runs[0].bold = True
                    tpara.runs[0].font.size = Pt(11)

                table = _add_table_after(tpara, rows=1, cols=2)
                left = table.cell(0, 0)
                right = table.cell(0, 1)

                # Colonne gauche : bullets
                left.text = ""
                for b in exp.get("bullets", []):
                    bp = left.add_paragraph()  # pas de style direct
                    try:
                        bp.style = "List Bullet"
                        bp.add_run(b)
                    except Exception:
                        # si le style n'existe pas dans le template
                        bp.text = f"• {b}"
                    bp.paragraph_format.space_after = Pt(0)

                # Colonne droite : dates + lieu/type
                right.text = ""
                rp = right.paragraphs[0]
                rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                block_lines = [x for x in [exp.get("dates", "")] if x]
                second = " - ".join(
                    [x for x in [exp.get("location", ""), exp.get("type", "")] if x]
                ).strip()
                if second:
                    block_lines.append(second)

                rr = rp.add_run("\n".join(block_lines))
                rr.italic = True
                rr.font.size = Pt(9)

                anchor = tpara

            _remove_paragraph(p)
            continue

        # ------- Texte simple -------
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

        # ------- Listes classiques (bullets) -------
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
