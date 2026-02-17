import os
import re
import uuid
import datetime as dt
from typing import Optional, Dict, Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import stripe
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
- bullet
- bullet

SKILLS:
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
    elif "droit" in sector:
        prompt = build_prompt_droit(payload)
    elif "rh" in sector or "ressources" in sector:
        prompt = build_prompt_rh(payload)
    elif "business" in sector:
        prompt = build_prompt_business(payload)
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

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
        line = line.rstrip()

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

    tags = ["EDUCATION:", "EXPERIENCES:", "SKILLS:", "ACTIVITIES:"]
    pos = {tag: t.find(tag) for tag in tags}

    # fallback si le modèle ne respecte pas le format
    if any(pos[tag] == -1 for tag in tags):
        return {
            "EDUCATION": t.splitlines(),
            "EXPERIENCES": [],
            "SKILLS": [],
            "ACTIVITIES": [],
        }

    order = sorted([(tag, pos[tag]) for tag in tags], key=lambda x: x[1])

    sections = {}
    for i, (tag, start) in enumerate(order):
        end = order[i + 1][1] if i + 1 < len(order) else len(t)
        block = t[start:end].strip().splitlines()

        if block and block[0].strip() == tag:
            block = block[1:]

        while block and not block[0].strip():
            block = block[1:]
        while block and not block[-1].strip():
            block = block[:-1]

        sections[tag.replace(":", "")] = block

    return sections

def write_docx_from_template(template_path: str, cv_text: str, out_path: str, payload: dict = None) -> None:
    doc = Document(template_path)

    payload = payload or {}
    full_name = payload.get("full_name", "").strip() or "NOM Prénom"
    role = payload.get("role", "").strip()
    finance_type = payload.get("finance_type", "").strip()
    cv_title = finance_type if finance_type else role

    contact_line = " | ".join([x for x in [
        payload.get("phone", "").strip(),
        payload.get("email", "").strip(),
        payload.get("linkedin", "").strip(),
    ] if x])

    sections = _split_sections(cv_text)

    mapping = {
    "%%FULL_NAME%%": full_name,
    "%%CONTACT_LINE%%": contact_line,
    "%%CV_TITLE%%": cv_title,

    "%%EDUCATION%%": sections.get("FORMATION", []),
    "%%EXPERIENCE%%": sections.get("EXPÉRIENCES PROFESSIONNELLES", []),
    "%%SKILLS%%": sections.get("COMPÉTENCES & OUTILS", []),
    "%%LANGUAGES%%": sections.get("LANGUES", []),
    "%%INTERESTS%%": sections.get("ACTIVITÉS & CENTRES D'INTÉRÊT", []),
}

    for ph, lines in mapping.items():
        p = _find_paragraph_containing(doc, ph)
        if not p:
            continue

        _clear_paragraph(p)

        if ph in ("%%FULL_NAME%%", "%%CV_TITLE%%", "%%CONTACT_LINE%%"):
            txt = lines[0] if lines else ""
            run = p.add_run(txt)
            if ph == "%%FULL_NAME%%":
                run.bold = True
                run.font.size = Pt(20)
            elif ph == "%%CV_TITLE%%":
                run.bold = True
                run.font.size = Pt(12)
            else:
                run.font.size = Pt(10)
            continue

        _insert_lines_after(p, lines, make_bullets=True)

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
    write_pdf_simple(cv_text, pdf_path)

    jobs[job_id] = {"docx_path": docx_path, "pdf_path": pdf_path}
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
