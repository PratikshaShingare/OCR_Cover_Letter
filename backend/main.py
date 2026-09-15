"""
Khanna Travels & Holidays — Visa Document Automation System
Main FastAPI Application Entrypoint
"""

import os
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from .api.ocr import router as ocr_router
from .api.applications import router as applications_router
from .api.templates import router as templates_router
from .api.documents import router as documents_router
from .services.storage import init_db

app = FastAPI(
    title="Khanna Travels & Holidays — Visa Document Automation System",
    description="Production-ready Visa Document Automation Prototype with single source of truth master client data.",
    version="1.0.0"
)

# CORS middleware for local Vite frontend dev server
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include API routers
app.include_router(ocr_router)
app.include_router(applications_router)
app.include_router(templates_router)
app.include_router(documents_router)


@app.on_event("startup")
async def on_startup():
    init_db()


@app.get("/api/health")
async def health_check():
    return {
        "status": "healthy",
        "system": "Khanna Travels & Holidays — Visa Document Automation System",
        "version": "1.0.0"
    }


# Static assets & modular frontend serving
frontend_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "frontend"))

css_dir = os.path.join(frontend_dir, "css")
if os.path.exists(css_dir):
    app.mount("/css", StaticFiles(directory=css_dir), name="css")

js_dir = os.path.join(frontend_dir, "js")
if os.path.exists(js_dir):
    app.mount("/js", StaticFiles(directory=js_dir), name="js")

pages_dir = os.path.join(frontend_dir, "pages")
if os.path.exists(pages_dir):
    app.mount("/pages", StaticFiles(directory=pages_dir), name="pages")

assets_dir = os.path.join(frontend_dir, "assets")
if os.path.exists(assets_dir):
    app.mount("/assets", StaticFiles(directory=assets_dir), name="assets")


@app.get("/logo.png")
@app.get("/khanna travels logo.png")
async def serve_logo():
    logo_file = os.path.join(frontend_dir, "logo.png")
    if not os.path.exists(logo_file):
        logo_file = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "khanna travels logo.png"))
    if os.path.exists(logo_file):
        return FileResponse(logo_file, media_type="image/png")
    return None


@app.get("/")
async def serve_root():
    index_file = os.path.join(frontend_dir, "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return {"message": "Frontend not found"}


@app.get("/{full_path:path}")
async def serve_spa(request: Request, full_path: str):
    if full_path.startswith("api") or full_path.startswith("css") or full_path.startswith("js"):
        return None
    index_file = os.path.join(frontend_dir, "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return {"message": "Frontend not found"}
