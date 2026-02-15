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
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

APP_URL = os.getenv("APP_URL", "")  # ex: https://mycvcopilote.netlify.app
STRIPE_SECRET = os.getenv("STRIPE_SECRET_KEY", "")
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
    return quota.get(email) != month_key()

def consume_free(email: str) -> None:
    quota[email] = month_key()

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

def generate_cv_text(payload: Dict[str, Any]) -> str:
    if not client:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY manquante sur le serveur.")

    prompt = build_prompt(payload)

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "user", "content": prompt}
        ],
    )

    return resp.choices[0].message.content.strip()

def write_docx_from_template(template_path: str, cv_text: str, out_path: str) -> None:
    doc = Document(template_path)
    # MVP: on remplace le contenu principal par le texte du CV
    # Astuce simple: on vide le doc et on écrit le contenu ligne par ligne
    for p in doc.paragraphs:
        p.clear()

    lines = cv_text.splitlines()
    for line in lines:
        line = line.rstrip()
        if not line:
            doc.add_paragraph("")
            continue
        # Headers simples
        if line.isupper() and len(line) <= 40:
            para = doc.add_paragraph(line)
            para.runs[0].bold = True
        else:
            doc.add_paragraph(line)

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
    # validations minimum
    required = ["email", "sector", "company", "role", "job_posting", "full_name", "city", "phone", "education", "experiences"]
    for k in required:
        if not payload.get(k):
            raise HTTPException(status_code=400, detail=f"Champ manquant: {k}")

    email = payload["email"].strip().lower()

    # === FREE : 1 seul CV à vie (par email) ===
    if email not in quota:
        quota[email] = "used"
        job_id = await generate_and_store(payload)
        return {"mode": "free", "downloads": make_download_urls(job_id)}

    # === Sinon : Stripe Checkout (paiement à l'unité) ===
    if not STRIPE_SECRET:
        raise HTTPException(status_code=500, detail="Stripe non configuré sur le serveur.")

    job_id = str(uuid.uuid4())
    jobs[job_id] = {"pending": "1"}  # marqueur
    jobs[job_id]["payload"] = payload  # on garde le payload pour le générer après paiement

    session = stripe.checkout.Session.create(
        mode="payment",
        line_items=[{
            "price_data": {
                "currency": "eur",
                "product_data": {"name": "MyCVCopilote – CV sur-mesure"},
                "unit_amount": 499,  # 4,99€
            },
            "quantity": 1,
        }],
        success_url=f"{APP_URL}/app.html?paid=1&job_id={job_id}",
        cancel_url=f"{APP_URL}/app.html?cancel=1",
        metadata={
            "job_id": job_id,
            "email": email,
        },
    )

    return {"mode": "stripe", "checkout_url": session.url}

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
    write_docx_from_template(tpl, cv_text, docx_path)
    write_pdf_simple(cv_text, pdf_path)

    jobs[job_id] = {"docx_path": docx_path, "pdf_path": pdf_path}
    return job_id
