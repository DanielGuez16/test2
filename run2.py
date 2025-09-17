#!/usr/bin/env python3
# run.py - Chatbot T&E pour région APAC
"""
T&E Chatbot System - APAC Region
=================================

Application FastAPI pour analyser automatiquement les tickets T&E
contre les politiques internes et fournir des recommandations.
"""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request, Depends, Cookie
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
import uvicorn
import uuid
import logging
from pathlib import Path
from datetime import datetime
import tempfile
import os
import json
from typing import Dict, Any, Optional, List
import hashlib
import secrets
from PIL import Image
import pytesseract
import io
import re
from decimal import Decimal
import PyPDF2
from docx import Document
import docx2txt
from striprtf.striprtf import rtf_to_text
import xlrd
from context_builder import TEContextBuilder
from ticket_analyzer import TicketAnalyzer
# Import du système RAG
from rag_system import TERAGSystem

# Initialiser le système RAG
rag_system = TERAGSystem()

from sharepoint_connector import SharePointClient
from context_builder import TEContextBuilder
from ticket_analyzer import TicketAnalyzer

# Imports internes
from llm_connector import LLMConnector
from te_document_processor import TEDocumentProcessor
from user_management import authenticate_user, log_activity, get_logs, USERS_DB

# Configuration du logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Création de l'application FastAPI
app = FastAPI(title="T&E Chatbot - APAC Region", version="1.0.0")

# Session utilisateur global
active_sessions = {}

# Variables globales pour les documents T&E et analyses
te_documents = {
    "excel_rules": None,
    "word_policies": None,
    "last_loaded": None
}

# Session chatbot global
chatbot_session = {
    "messages": [],
    "context_data": {},
    "analysis_history": [],
    "feedback_data": []
}

# Initialiser les composants
llm_connector = LLMConnector()
te_processor = TEDocumentProcessor()

def generate_session_token():
    return secrets.token_urlsafe(32)

def get_current_user_from_session(session_token: Optional[str] = Cookie(None)):
    """Récupère l'utilisateur depuis le token de session"""
    if not session_token or session_token not in active_sessions:
        return None
    return active_sessions[session_token]

# Configuration CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Création des dossiers requis
required_dirs = ["data", "templates", "static", "static/js", "static/css", "static/images", "uploads", "te_documents"]
for directory in required_dirs:
    Path(directory).mkdir(exist_ok=True)

# Configuration des fichiers statiques et templates
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

#######################################################################################################################################
#                           API ROUTES
#######################################################################################################################################

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    """Page d'accueil - redirige vers login si non connecté"""
    session_token = request.cookies.get("session_token")
    current_user = get_current_user_from_session(session_token)
    
    if not current_user:
        return templates.TemplateResponse("login.html", {
            "request": request,
            "title": "Login - T&E Chatbot"
        })
    
    log_activity(current_user["username"], "ACCESS", "Accessed T&E Chatbot main page")
    
    # Charger automatiquement les documents si pas encore fait
    if not te_documents["excel_rules"]:
        load_te_documents_from_sharepoint()
    
    return templates.TemplateResponse("te_index.html", {
        "request": request,
        "title": "T&E Analysis Chatbot",
        "version": "1.0.0",
        "timestamp": datetime.now().isoformat(),
        "user": current_user,
        "documents_loaded": bool(te_documents["excel_rules"]),
        "last_loaded": te_documents["last_loaded"],
        "documents_source": te_documents.get("source", "Manual")
    })

@app.get("/health")
async def health_check():
    """Endpoint de vérification de l'état de l'application"""
    return {
        "status": "healthy",
        "service": "te-chatbot-apac",
        "version": "1.0.0",
        "timestamp": datetime.now().isoformat(),
        "documents_loaded": bool(te_documents["excel_rules"]),
        "templates_available": Path("templates/te_index.html").exists(),
        "static_available": Path("static/js/te_main.js").exists()
    }

@app.post("/api/login")
async def login(request: Request):
    """Authentification utilisateur"""
    try:
        data = await request.json()
        username = data.get("username")
        password = data.get("password")
        
        user = authenticate_user(username, password)
        if not user:
            raise HTTPException(status_code=401, detail="Invalid credentials")
        
        session_token = generate_session_token()
        active_sessions[session_token] = user
        
        log_activity(username, "LOGIN", "Successful login to T&E Chatbot")
        
        response = JSONResponse({
            "success": True,
            "message": f"Welcome {user['full_name']}!",
            "redirect": "/"
        })
        
        response.set_cookie(
            key="session_token",
            value=session_token,
            httponly=True,
            max_age=24*60*60,
            samesite="lax"
        )
        
        return response
        
    except Exception as e:
        return JSONResponse(
            {"success": False, "message": str(e)}, 
            status_code=401
        )

@app.post("/api/logout")
async def logout(request: Request):
    """Déconnexion utilisateur"""
    session_token = request.cookies.get("session_token")
    
    if session_token and session_token in active_sessions:
        user = active_sessions[session_token]
        log_activity(user["username"], "LOGOUT", "User logged out from T&E Chatbot")
        del active_sessions[session_token]
    
    response = JSONResponse({"success": True, "redirect": "/"})
    response.delete_cookie("session_token")
    return response
 
@app.post("/api/analyze-ticket")
async def analyze_ticket(
    ticket_file: UploadFile = File(...),
    question: str = Form(""),
    session_token: Optional[str] = Cookie(None)
):
    analyzer = TicketAnalyzer(rag_system, llm_connector)

    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if not te_documents["excel_rules"]:
        raise HTTPException(status_code=400, detail="T&E documents not loaded. Please load Excel and Word files first.")
    
    try:
        log_activity(current_user["username"], "TICKET_ANALYSIS", f"Analyzing ticket: {ticket_file.filename}")
        
        # Lire le ticket (garder ta logique existante)
        ticket_content = await ticket_file.read()
        ticket_info = extract_ticket_information(ticket_content, ticket_file.filename)
        
        analyzer = TicketAnalyzer(rag_system, llm_connector)
        
        analysis_result = analyzer.analyze_ticket(ticket_info, question)
        
        # Sauvegarder l'analyse
        analysis_record = {
            "timestamp": datetime.now().isoformat(),
            "user": current_user["username"],
            "ticket_filename": ticket_file.filename,
            "ticket_info": ticket_info,
            "analysis_result": analysis_result,
            "question": question
        }
        
        chatbot_session["analysis_history"].append(analysis_record)
        
        return {
            "success": True,
            "ticket_info": ticket_info,
            "analysis_result": analysis_result,
            "timestamp": analysis_record["timestamp"]
        }
        
    except Exception as e:
        logger.error(f"Erreur analyse ticket: {e}")
        raise HTTPException(status_code=500, detail=f"Error analyzing ticket: {str(e)}")
    
@app.post("/api/chat")
async def chat_with_ai(request: Request, session_token: Optional[str] = Cookie(None)):
    analyzer = TicketAnalyzer(rag_system, llm_connector)

    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        data = await request.json()
        user_message = data.get("message", "")
        
        log_activity(current_user["username"], "CHAT_MESSAGE", f"Chat message: {user_message[:100]}...")
        
        analyzer = TicketAnalyzer(rag_system, llm_connector)
        
        # Préparer résumé des règles
        te_rules_summary = {}
        if te_documents["excel_rules"]:
            for sheet_name, rules in te_documents["excel_rules"].items():
                te_rules_summary[sheet_name] = len(rules)
        
        response_data = analyzer.answer_general_question(user_message, te_rules_summary)
        ai_response = response_data['ai_response']
        
        # Sauvegarder la conversation
        chatbot_session["messages"].append({
            "type": "user",
            "message": user_message,
            "timestamp": datetime.now().isoformat(),
            "user": current_user["username"]
        })
        
        chatbot_session["messages"].append({
            "type": "assistant",
            "message": ai_response,
            "timestamp": datetime.now().isoformat()
        })
        
        return {
            "success": True,
            "response": ai_response,
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Erreur chatbot: {e}")
        raise HTTPException(status_code=500, detail=f"Erreur chatbot: {str(e)}")
      
@app.post("/api/feedback")
async def submit_feedback(
    request: Request,
    session_token: Optional[str] = Cookie(None)
):
    """Soumet un feedback sur une réponse du chatbot"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        data = await request.json()
        
        feedback_record = {
            "timestamp": datetime.now().isoformat(),
            "user": current_user["username"],
            "analysis_id": data.get("analysis_id", ""),
            "rating": data.get("rating", 0),
            "comment": data.get("comment", ""),
            "issue_type": data.get("issue_type", "")
        }
        
        # Sauvegarder en CSV local
        save_feedback_to_csv(feedback_record)
        
        # Ajouter à la session
        chatbot_session["feedback_data"].append(feedback_record)
        
        log_activity(current_user["username"], "FEEDBACK", f"Rating: {feedback_record['rating']}, Issue: {feedback_record['issue_type']}")
        
        return {
            "success": True,
            "message": "Feedback submitted successfully"
        }
        
    except Exception as e:
        logger.error(f"Erreur feedback: {e}")
        raise HTTPException(status_code=500, detail=f"Error submitting feedback: {str(e)}")

@app.get("/api/analysis-history")
async def get_analysis_history(session_token: Optional[str] = Cookie(None)):
    """Récupère l'historique des analyses"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    # Filtrer par utilisateur si non admin
    if current_user["role"] == "admin":
        history = chatbot_session["analysis_history"]
    else:
        history = [
            record for record in chatbot_session["analysis_history"]
            if record.get("user") == current_user["username"]
        ]
    
    return {
        "success": True,
        "history": history[-20:],  # Derniers 20 enregistrements
        "total": len(history)
    }

@app.get("/api/te-status")
async def get_te_status():
    excel_count = 0
    if te_documents.get("excel_rules"):
        excel_count = sum(len(rules) for rules in te_documents["excel_rules"].values())
    
    return {
        "documents_loaded": bool(te_documents.get("excel_rules")),
        "last_loaded": te_documents.get("last_loaded"),
        "excel_rules_count": excel_count,
        "word_policies_available": bool(te_documents.get("word_policies")),
        "word_policies_length": len(te_documents.get("word_policies", "")),
        "timestamp": datetime.now().isoformat(),
        "debug_keys": list(te_documents.keys())  # Pour debug
    }

@app.post("/api/analyze-multiple-tickets")
async def analyze_multiple_tickets(
    request: Request,
    session_token: Optional[str] = Cookie(None)
):
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    form = await request.form()
    files = form.getlist("files")
    question = form.get("question", "")
    
    analyzer = TicketAnalyzer(rag_system, llm_connector)
    results = []
    
    for file in files:
        if hasattr(file, 'read'):
            try:
                file_content = await file.read()
                ticket_info = extract_ticket_information(file_content, file.filename)
                analysis_result = analyzer.analyze_ticket(ticket_info, question)
                
                results.append({
                    "filename": file.filename,
                    "ticket_info": ticket_info,
                    "analysis_result": analysis_result
                })
                
            except Exception as e:
                results.append({
                    "filename": file.filename,
                    "error": str(e)
                })
    
    return {
        "success": True,
        "results": results,
        "total_files": len(files)
    }

@app.get("/api/view-excel-rules")
async def view_excel_rules(session_token: Optional[str] = Cookie(None)):
    """Visualise les règles Excel chargées"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if not te_documents.get("excel_rules"):  
        raise HTTPException(status_code=404, detail="No Excel rules loaded")
    
    return {
        "success": True,
        "rules": te_documents["excel_rules"],
        "last_loaded": te_documents["last_loaded"]
    }

@app.get("/api/view-word-policies")
async def view_word_policies(session_token: Optional[str] = Cookie(None)):
    """Visualise les politiques Word chargées"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if not te_documents.get("word_policies"):
        raise HTTPException(status_code=404, detail="No Word policies loaded")
    
    return {
        "success": True,
        "policies_text": te_documents["word_policies"],
        "last_loaded": te_documents["last_loaded"]
    }

@app.post("/api/refresh-te-documents")
async def refresh_te_documents(session_token: Optional[str] = Cookie(None)):
    """Recharge les documents T&E depuis SharePoint"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    log_activity(current_user["username"], "DOCUMENT_REFRESH", "Refreshing T&E documents from SharePoint")
    
    success = load_te_documents_from_sharepoint()
    
    if success:
        return {
            "success": True,
            "message": "Documents rechargés depuis SharePoint",
            "last_loaded": te_documents["last_loaded"]
        }
    else:
        raise HTTPException(status_code=500, detail="Error refreshing documents from SharePoint")
    

#######################################################################################################################################
#                           UTILITY FUNCTIONS
#######################################################################################################################################

def load_te_documents_from_sharepoint():
    """Charge automatiquement les documents T&E depuis SharePoint au démarrage"""
    try:
        logger.info("Chargement automatique des documents T&E depuis SharePoint...")
        
        # Initialiser le client SharePoint
        sharepoint_client = SharePointClient()
        
        # Chemins des fichiers sur SharePoint
        excel_path = "Chatbot/sources/Consolidated Limits.xlsx"
        docx_path = "Chatbot/sources/APAC Travel Entertainment Procedure Mar2025_Clean.docx"
        
        # Charger le fichier Excel
        logger.info("Chargement du fichier Excel...")
        excel_binary = sharepoint_client.read_binary_file(excel_path)
        excel_dict = sharepoint_client.read_excel_file_as_dict(excel_binary)
        excel_rules = te_processor.process_excel_rules_from_dict(excel_dict, "Consolidated Limits.xlsx")
        
        # Charger le fichier DOCX
        logger.info("Chargement du fichier DOCX...")
        docx_binary = sharepoint_client.read_binary_file(docx_path)
        docx_text = sharepoint_client.read_docx_file_as_text(docx_binary)
        word_policies = te_processor.process_word_policies_from_text(docx_text, "APAC Travel Entertainment Procedure.docx")
        
        # Stocker globalement
        te_documents["excel_rules"] = excel_rules
        te_documents["word_policies"] = docx_text  # Stocker le texte brut, pas l'objet processé
        te_documents["last_loaded"] = datetime.now().isoformat()
        te_documents["source"] = "SharePoint"
        
        # Debug logs
        logger.info(f"Excel rules keys: {list(excel_rules.keys()) if excel_rules else 'None'}")
        logger.info(f"Word policies length: {len(docx_text) if docx_text else 0} characters")
        logger.info(f"te_documents keys: {list(te_documents.keys())}")

        # Indexer dans le système RAG
        logger.info("Indexation dans le système RAG...")
        rag_system.index_excel_rules(excel_rules)
        rag_system.index_word_policies(word_policies)
        
        logger.info("Documents T&E chargés avec succès depuis SharePoint")
        return True
        
    except Exception as e:
        logger.error(f"Erreur chargement documents SharePoint: {e}")
        return False
    

def preprocess_image_for_ocr(image):
    """Préprocessing d'image pour améliorer l'OCR"""
    try:
        import cv2
        import numpy as np
        from PIL import Image, ImageEnhance, ImageFilter
        
        # Convertir en niveaux de gris
        if image.mode != 'L':
            image = image.convert('L')
        
        # Améliorer le contraste
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)
        
        # Améliorer la netteté
        image = image.filter(ImageFilter.SHARPEN)
        
        # Redimensionner si trop petit
        width, height = image.size
        if width < 1000 or height < 1000:
            scale_factor = max(1000/width, 1000/height)
            new_size = (int(width * scale_factor), int(height * scale_factor))
            image = image.resize(new_size, Image.Resampling.LANCZOS)
        
        # Convertir en numpy pour OpenCV
        img_array = np.array(image)
        
        # Seuillage adaptatif
        img_thresh = cv2.adaptiveThreshold(
            img_array, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
        )
        
        # Débruitage
        img_denoised = cv2.medianBlur(img_thresh, 3)
        
        # Reconvertir en PIL
        processed_image = Image.fromarray(img_denoised)
        
        return processed_image
        
    except ImportError:
        logger.warning("OpenCV non disponible, preprocessing basique")
        return image.convert('L')
    except Exception as e:
        logger.warning(f"Erreur preprocessing: {e}")
        return image
    
def extract_ticket_information(file_content: bytes, filename: str) -> dict:
    """Extrait les informations d'un ticket avec OCR amélioré et parsing intelligent"""
    try:
        file_ext = Path(filename).suffix.lower()
        text = ""
        
        if file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.gif', '.webp']:
            # Images - utiliser OCR avec preprocessing
            try:
                import cv2
                import numpy as np
                from PIL import Image, ImageEnhance, ImageFilter
                
                # Convertir en PIL Image
                image = Image.open(io.BytesIO(file_content))
                
                # Preprocessing pour améliorer l'OCR
                image = preprocess_image_for_ocr(image)
                
                # OCR avec Tesseract
                import pytesseract
                
                # Configuration OCR optimisée pour tickets
                custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,-€$£¥₹/: '
                text = pytesseract.image_to_string(image, config=custom_config)
                
                logger.info(f"OCR extrait {len(text)} caractères de {filename}")
                
            except Exception as e:
                logger.warning(f"Erreur OCR image: {e}")
                text = f"Image file {filename} - OCR extraction failed"
                
        elif file_ext == '.pdf':
            # PDF - extraire le texte avec multiple méthodes
            try:
                # Méthode 1: PyPDF2
                pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_content))
                text_parts = []
                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text.strip():
                        text_parts.append(page_text)
                
                text = "\n".join(text_parts)
                
                # Si pas de texte extrait, essayer OCR sur PDF
                if not text.strip():
                    try:
                        import fitz  # PyMuPDF
                        pdf_doc = fitz.open(stream=file_content, filetype="pdf")
                        
                        for page_num in range(len(pdf_doc)):
                            page = pdf_doc[page_num]
                            # Convertir en image puis OCR
                            mat = fitz.Matrix(2, 2)  # zoom factor
                            pix = page.get_pixmap(matrix=mat)
                            img_data = pix.tobytes("png")
                            
                            image = Image.open(io.BytesIO(img_data))
                            image = preprocess_image_for_ocr(image)
                            
                            page_text = pytesseract.image_to_string(image)
                            if page_text.strip():
                                text_parts.append(page_text)
                        
                        text = "\n".join(text_parts)
                        logger.info(f"PDF OCR extrait {len(text)} caractères")
                        
                    except ImportError:
                        logger.warning("PyMuPDF non disponible pour OCR PDF")
                    except Exception as e:
                        logger.warning(f"Erreur OCR PDF: {e}")
                
            except Exception as e:
                logger.warning(f"Erreur lecture PDF: {e}")
                text = f"PDF file {filename} - text extraction failed"
        
        elif file_ext in ['.docx', '.doc']:
            # Documents Word
            try:
                if file_ext == '.docx':
                    doc = Document(io.BytesIO(file_content))
                    text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
                else:
                    # Pour .doc, utiliser python-docx2txt si disponible
                    try:
                        import docx2txt
                        text = docx2txt.process(io.BytesIO(file_content))
                    except ImportError:
                        text = f"DOC file {filename} - docx2txt not available, please convert to DOCX"
            except Exception as e:
                logger.warning(f"Erreur lecture Word: {e}")
                text = f"Word file {filename} - text extraction failed"
                
        elif file_ext in ['.xlsx', '.xls']:
            # Fichiers Excel
            try:
                import pandas as pd
                if file_ext == '.xlsx':
                    df = pd.read_excel(io.BytesIO(file_content), sheet_name=None)
                else:
                    df = pd.read_excel(io.BytesIO(file_content), sheet_name=None, engine='xlrd')
                
                # Convertir toutes les sheets en texte
                text_parts = []
                for sheet_name, sheet_df in df.items():
                    text_parts.append(f"Sheet: {sheet_name}")
                    text_parts.append(sheet_df.to_string())
                text = "\n".join(text_parts)
            except Exception as e:
                logger.warning(f"Erreur lecture Excel: {e}")
                text = f"Excel file {filename} - text extraction failed"
                
        elif file_ext in ['.txt', '.csv']:
            # Fichiers texte et CSV
            try:
                text = file_content.decode('utf-8', errors='ignore')
            except Exception as e:
                try:
                    text = file_content.decode('latin-1', errors='ignore')
                except:
                    text = f"Text file {filename} - encoding detection failed"
                    
        elif file_ext in ['.rtf']:
            # Rich Text Format
            try:
                from striprtf.striprtf import rtf_to_text
                text = rtf_to_text(file_content.decode('utf-8', errors='ignore'))
            except ImportError:
                text = f"RTF file {filename} - striprtf not available"
            except Exception as e:
                logger.warning(f"Erreur lecture RTF: {e}")
                text = f"RTF file {filename} - text extraction failed"
                
        else:
            # Type non supporté - essayer comme texte brut
            try:
                text = file_content.decode('utf-8', errors='ignore')
                if not text.strip():
                    text = f"Unknown file type {filename} - content unreadable"
            except:
                text = f"Unsupported file type {filename} ({file_ext})"
    
        # Parser les informations du ticket avec extraction améliorée
        ticket_info = parse_ticket_text_enhanced(text)
        ticket_info["filename"] = filename
        ticket_info["raw_text"] = text[:1500]  # Augmenté pour plus de contexte
        ticket_info["file_type"] = file_ext
        
        return ticket_info
        
    except Exception as e:
        logger.error(f"Erreur extraction ticket: {e}")
        return {
            "filename": filename,
            "error": str(e),
            "raw_text": "",
            "file_type": Path(filename).suffix.lower(),
            "amount": None,
            "currency": None,
            "date": None,
            "vendor": None,
            "category": "unknown",
            "location": None,
            "confidence": 0.0
        }
    
def parse_ticket_text_enhanced(text: str) -> dict:
    """Parse amélioré du texte d'un ticket avec regex robustes"""
    info = {
        "amount": None,
        "currency": None,
        "date": None,
        "vendor": None,
        "category": "unknown",
        "location": None,
        "country_code": None,
        "description": "",
        "confidence": 0.0
    }
    
    confidence_factors = []
    
    # === EXTRACTION DES MONTANTS AMÉLIORÉE ===
    amount_patterns = [
        # Patterns avec devise avant
        r'(EUR|USD|AED|CHF|AUD|GBP|JPY|SGD|HKD|CNY|INR|THB|MYR|KRW|TWD)\s*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2})?)',
        # Patterns avec devise après
        r'([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2})?)\s*(EUR|USD|AED|CHF|AUD|GBP|JPY|SGD|HKD|CNY|INR|THB|MYR|KRW|TWD)',
        # Patterns avec symboles
        r'([€$£¥₹])\s*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2})?)',
        r'([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2})?)\s*([€$£¥₹])',
        # Total/Amount labels
        r'(?:Total|Amount|Price|Prix|Montant)[\s:]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2})?)\s*(EUR|USD|AED|CHF|AUD|GBP|JPY|SGD|HKD)?',
        # Patterns numériques seuls (moins fiables)
        r'([0-9]{1,4}[.,][0-9]{2})\b'
    ]
    
    amount_found = False
    for i, pattern in enumerate(amount_patterns):
        matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
        if matches:
            for match in matches:
                try:
                    if len(match) == 2:
                        # Déterminer quel élément est le montant
                        if match[0].replace(',', '.').replace('.', '').isdigit():
                            amount_str = match[0]
                            currency_str = match[1]
                        else:
                            amount_str = match[1]
                            currency_str = match[0]
                        
                        # Nettoyer et convertir le montant
                        cleaned_amount = clean_amount_string(amount_str)
                        if cleaned_amount and 0.01 <= cleaned_amount <= 100000:  # Plage raisonnable
                            info["amount"] = cleaned_amount
                            info["currency"] = (currency_str)
                            confidence_factors.append(0.3 - (i * 0.05))  # Plus de confiance pour les premiers patterns
                            amount_found = True
                            break
                    elif len(match) == 1:
                        # Montant seul
                        cleaned_amount = clean_amount_string(match)
                        if cleaned_amount and 1 <= cleaned_amount <= 100000:
                            info["amount"] = cleaned_amount
                            confidence_factors.append(0.1)
                            amount_found = True
                            break
                except (ValueError, IndexError):
                    continue
        if amount_found:
            break
    
    # === EXTRACTION DES DATES AMÉLIORÉE ===
    date_patterns = [
        r'(\d{1,2}[/.-]\d{1,2}[/.-]\d{4})',      # DD/MM/YYYY ou MM/DD/YYYY
        r'(\d{4}[/.-]\d{1,2}[/.-]\d{1,2})',      # YYYY/MM/DD
        r'(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})',  # DD Month YYYY
        r'((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4})', # Month DD, YYYY
        r'(\d{1,2}[/.-]\d{1,2}[/.-]\d{2})'       # DD/MM/YY
    ]
    
    for pattern in date_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            info["date"] = matches[0]
            confidence_factors.append(0.2)
            break
    
    # === DÉTECTION INTELLIGENTE DU PAYS/LIEU ===
    location_info = detect_location_from_text(text)
    info.update(location_info)
    if location_info.get("country_code"):
        confidence_factors.append(0.2)
    
    # === CATÉGORISATION AMÉLIORÉE ===
    category_info = categorize_expense_enhanced(text)
    info["category"] = category_info["category"]
    confidence_factors.append(category_info["confidence"])
    
    # === EXTRACTION DU VENDEUR ===
    vendor_info = extract_vendor_info(text)
    info["vendor"] = vendor_info["name"]
    confidence_factors.append(vendor_info["confidence"])
    
    # === CALCUL DE LA CONFIANCE GLOBALE ===
    info["confidence"] = min(sum(confidence_factors), 1.0)
    
    info["description"] = text[:300]  # Description plus longue
    
    return info

def parse_ticket_text(text: str) -> dict:
    """Parse le texte d'un ticket pour extraire les informations clés"""
    info = {
        "amount": None,
        "currency": None,
        "date": None,
        "vendor": None,
        "category": "unknown",
        "location": None,
        "description": ""
    }
    
    # Extraction des montants
    amount_patterns = [
        r'(\d+[,.]?\d*)\s*(EUR|USD|AED|CHF|AUD|GBP|JPY)',
        r'(EUR|USD|AED|CHF|AUD|GBP|JPY)\s*(\d+[,.]?\d*)',
        r'Total[\s:]*(\d+[,.]?\d*)',
        r'Amount[\s:]*(\d+[,.]?\d*)'
    ]
    
    for pattern in amount_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            match = matches[0]
            if isinstance(match, tuple) and len(match) == 2:
                if match[0].replace(',', '.').replace('.', '').isdigit():
                    info["amount"] = float(match[0].replace(',', '.'))
                    info["currency"] = match[1].upper()
                else:
                    info["amount"] = float(match[1].replace(',', '.'))
                    info["currency"] = match[0].upper()
            break
    
    # Extraction des dates
    date_patterns = [
        r'(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})',
        r'(\d{2,4}[/.-]\d{1,2}[/.-]\d{1,2})'
    ]
    
    for pattern in date_patterns:
        matches = re.findall(pattern, text)
        if matches:
            info["date"] = matches[0]
            break
    
    # Catégorisation basique
    text_lower = text.lower()
    if any(word in text_lower for word in ['hotel', 'accommodation', 'lodging']):
        info["category"] = "hotel"
    elif any(word in text_lower for word in ['restaurant', 'meal', 'food', 'dining']):
        info["category"] = "meal"
    elif any(word in text_lower for word in ['taxi', 'uber', 'transport', 'metro', 'bus']):
        info["category"] = "transport"
    elif any(word in text_lower for word in ['flight', 'airline', 'airport']):
        info["category"] = "flight"
    
    # Extraction du vendeur/établissement
    lines = text.split('\n')
    if lines:
        info["vendor"] = lines[0].strip()[:50]  # Premier ligne comme vendeur potentiel
    
    info["description"] = text[:200]  # Premiers 200 caractères comme description
    
    return info

def analyze_against_te_rules(ticket_info: dict, excel_rules: dict) -> dict:
    """Analyse un ticket contre les règles T&E chargées"""
    try:
        analysis = {
            "is_valid": False,
            "confidence": 0.0,
            "issues": [],
            "recommendations": [],
            "matching_rules": [],
            "status": "pending_review"
        }
        
        amount = ticket_info.get("amount")
        currency = ticket_info.get("currency", "EUR")
        category = ticket_info.get("category", "unknown")
        
        if not amount:
            analysis["issues"].append("Amount not found in ticket")
            analysis["recommendations"].append("Please provide a clear ticket with visible amount")
            return analysis
        
        # Rechercher les règles applicables
        applicable_rules = find_applicable_rules(excel_rules, currency, category)
        
        if not applicable_rules:
            analysis["issues"].append(f"No rules found for {category} in {currency}")
            analysis["recommendations"].append("Check if this expense category is covered by policy")
            return analysis
        
        # Vérifier les limites
        for rule in applicable_rules:
            limit = rule.get("limit", 0)
            rule_type = rule.get("type", "")
            
            analysis["matching_rules"].append({
                "type": rule_type,
                "limit": limit,
                "currency": currency,
                "description": rule.get("description", "")
            })
            
            if amount <= limit:
                analysis["is_valid"] = True
                analysis["confidence"] = 0.9
                analysis["status"] = "approved"
            else:
                analysis["issues"].append(f"{category.title()} amount {amount} {currency} exceeds limit of {limit} {currency}")
                analysis["recommendations"].append("Get manager approval for amount exceeding policy limit")
                analysis["confidence"] = 0.3
                analysis["status"] = "requires_approval"
        
        # Vérifications additionnelles
        if ticket_info.get("date"):
            # Vérifier si la date n'est pas trop ancienne (ex: > 30 jours)
            try:
                from datetime import datetime, timedelta
                # Cette vérification pourrait être implémentée selon vos règles
                pass
            except:
                pass
        
        # Si aucun problème majeur
        if not analysis["issues"] and analysis["is_valid"]:
            analysis["recommendations"].append("Expense appears compliant with T&E policy")
        
        return analysis
        
    except Exception as e:
        logger.error(f"Erreur analyse règles T&E: {e}")
        return {
            "is_valid": False,
            "confidence": 0.0,
            "issues": [f"Analysis error: {str(e)}"],
            "recommendations": ["Please contact support"],
            "matching_rules": [],
            "status": "error"
        }

def find_applicable_rules(excel_rules: dict, currency: str, category: str) -> list:
    """Trouve les règles applicables selon la devise et catégorie"""
    applicable = []
    
    # Mapper les catégories aux sheets Excel
    category_mapping = {
        "meal": "Internal staff Meal",
        "hotel": "Hotel", 
        "breakfast": "Breakfast & Lunch & Dinner",
        "lunch": "Breakfast & Lunch & Dinner",
        "dinner": "Breakfast & Lunch & Dinner",
        "transport": "Internal staff Meal"  # Fallback
    }
    
    sheet_name = category_mapping.get(category, "Internal staff Meal")
    
    if sheet_name in excel_rules:
        sheet_data = excel_rules[sheet_name]
        
        # Rechercher par devise (CRN_KEY)
        for rule in sheet_data:
            if rule.get("CRN_KEY") == currency:
                applicable.append({
                    "type": rule.get("TYPE", category),
                    "limit": rule.get("AMOUNT1", 0),
                    "currency": currency,
                    "country": rule.get("ID_01", ""),
                    "description": f"{sheet_name} - {rule.get('TYPE', '')}"
                })
    
    return applicable

def prepare_te_context(ticket_info: dict, analysis_result: dict, te_documents: dict) -> str:
    """Prépare le contexte T&E pour l'IA"""
    context_parts = []
    
    # Contexte métier
    context_parts.append("CONTEXTE T&E (Travel & Expense):")
    context_parts.append("- Analyse automatique des tickets de frais selon les politiques APAC")
    context_parts.append("- Validation des montants contre les limites approuvées")
    context_parts.append("- Vérification de conformité avec les règles internes")
    
    # Informations du ticket
    context_parts.append(f"\nTICKET ANALYSÉ:")
    context_parts.append(f"- Fichier: {ticket_info.get('filename', 'N/A')}")
    context_parts.append(f"- Montant: {ticket_info.get('amount', 'N/A')} {ticket_info.get('currency', 'N/A')}")
    context_parts.append(f"- Catégorie: {ticket_info.get('category', 'N/A')}")
    context_parts.append(f"- Date: {ticket_info.get('date', 'N/A')}")
    context_parts.append(f"- Vendeur: {ticket_info.get('vendor', 'N/A')}")
    
    # Résultat de l'analyse
    context_parts.append(f"\nRÉSULTAT ANALYSE:")
    context_parts.append(f"- Statut: {analysis_result.get('status', 'N/A')}")
    context_parts.append(f"- Valide: {analysis_result.get('is_valid', False)}")
    context_parts.append(f"- Confiance: {analysis_result.get('confidence', 0):.1%}")
    
    if analysis_result.get("issues"):
        context_parts.append("- Problèmes identifiés:")
        for issue in analysis_result["issues"]:
            context_parts.append(f"  • {issue}")
    
    if analysis_result.get("matching_rules"):
        context_parts.append("- Règles applicables:")
        for rule in analysis_result["matching_rules"]:
            context_parts.append(f"  • {rule['description']}: limite {rule['limit']} {rule['currency']}")
    
    # Règles générales depuis les documents
    if te_documents.get("word_policies"):
        context_parts.append(f"\nPOLITIQUES T&E:")
        # Extraire les premiers paragraphes des politiques Word
        policies_text = te_documents["word_policies"][:500] + "..."
        context_parts.append(policies_text)
    
    return "\n".join(context_parts)

def prepare_general_te_context() -> str:
    """Prépare un contexte général T&E pour les questions sans ticket"""
    context_parts = []
    
    context_parts.append("SYSTÈME T&E CHATBOT - RÉGION APAC:")
    context_parts.append("- Aide aux collaborateurs pour comprendre les politiques T&E")
    context_parts.append("- Validation automatique des tickets de frais")
    context_parts.append("- Règles par pays/devise selon les grilles tarifaires")
    
    if te_documents.get("excel_rules"):
        context_parts.append("\nRÈGLES DISPONIBLES:")
        for sheet_name, rules in te_documents["excel_rules"].items():
            context_parts.append(f"- {sheet_name}: {len(rules)} règles")
            if rules:
                currencies = set(rule.get("CRN_KEY", "") for rule in rules)
                context_parts.append(f"  Devises: {', '.join(sorted(currencies))}")
    
    context_parts.append("\nCATÉGORIES COUVERTES:")
    context_parts.append("- Hôtels et hébergement")
    context_parts.append("- Repas (petit-déjeuner, déjeuner, dîner)")
    context_parts.append("- Transport et déplacements")
    context_parts.append("- Frais divers selon politique")
    
    if te_documents.get("word_policies"):
        context_parts.append(f"\nEXTRAIT POLITIQUES:")
        context_parts.append(te_documents["word_policies"][:300] + "...")
    
    return "\n".join(context_parts)

def save_feedback_to_csv(feedback_record: dict):
    """Sauvegarde le feedback dans un fichier CSV local"""
    try:
        import csv
        import os
        
        feedback_file = "data/feedback.csv"
        
        # Créer le header si le fichier n'existe pas
        file_exists = os.path.exists(feedback_file)
        
        with open(feedback_file, mode='a', newline='', encoding='utf-8') as file:
            fieldnames = ['timestamp', 'user', 'analysis_id', 'rating', 'comment', 'issue_type']
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            
            if not file_exists:
                writer.writeheader()
            
            writer.writerow(feedback_record)
        
        logger.info(f"Feedback sauvegardé: {feedback_record['rating']}/5")
        
    except Exception as e:
        logger.error(f"Erreur sauvegarde feedback: {e}")

def clean_amount_string(amount_str: str) -> float:
    """Nettoie et convertit une chaîne de montant en float"""
    try:
        # Supprimer espaces et caractères non numériques sauf , et .
        cleaned = re.sub(r'[^\d.,]', '', amount_str.strip())
        
        if not cleaned:
            return None
        
        # Gérer les différents formats de séparateurs
        if ',' in cleaned and '.' in cleaned:
            # Format: 1,234.56 ou 1.234,56
            comma_pos = cleaned.rfind(',')
            dot_pos = cleaned.rfind('.')
            
            if dot_pos > comma_pos:
                # Format anglais: 1,234.56
                cleaned = cleaned.replace(',', '')
            else:
                # Format européen: 1.234,56
                cleaned = cleaned.replace('.', '').replace(',', '.')
        elif ',' in cleaned:
            # Déterminer si c'est un séparateur de milliers ou décimal
            comma_parts = cleaned.split(',')
            if len(comma_parts) == 2 and len(comma_parts[1]) <= 2:
                # Séparateur décimal
                cleaned = cleaned.replace(',', '.')
            else:
                # Séparateur de milliers
                cleaned = cleaned.replace(',', '')
        
        return float(cleaned)
        
    except (ValueError, AttributeError):
        return None

def normalize_currency(currency_str: str) -> str:
    """Normalise les codes de devise"""
    currency_map = {
        '€': 'EUR', '$': 'USD', '£': 'GBP', '¥': 'JPY', '₹': 'INR',
        'euro': 'EUR', 'dollar': 'USD', 'pound': 'GBP', 'yen': 'JPY'
    }
    
    currency_clean = currency_str.strip().upper()
    return currency_map.get(currency_clean, currency_clean)

def detect_location_from_text(text: str) -> dict:
    """Détection intelligente du pays/lieu à partir du texte"""
    location_info = {
        "location": None,
        "country_code": None,
        "city": None
    }
    
    # Base de données de correspondances géographiques
    location_mappings = {
        # Villes -> Pays
        'paris': {'country': 'FR', 'city': 'Paris'},
        'lyon': {'country': 'FR', 'city': 'Lyon'},
        'marseille': {'country': 'FR', 'city': 'Marseille'},
        'london': {'country': 'GB', 'city': 'London'},
        'berlin': {'country': 'DE', 'city': 'Berlin'},
        'munich': {'country': 'DE', 'city': 'Munich'},
        'frankfurt': {'country': 'DE', 'city': 'Frankfurt'},
        'sydney': {'country': 'AU', 'city': 'Sydney'},
        'melbourne': {'country': 'AU', 'city': 'Melbourne'},
        'tokyo': {'country': 'JP', 'city': 'Tokyo'},
        'osaka': {'country': 'JP', 'city': 'Osaka'},
        'singapore': {'country': 'SG', 'city': 'Singapore'},
        'hongkong': {'country': 'HK', 'city': 'Hong Kong'},
        'hong kong': {'country': 'HK', 'city': 'Hong Kong'},
        'dubai': {'country': 'AE', 'city': 'Dubai'},
        'abu dhabi': {'country': 'AE', 'city': 'Abu Dhabi'},
        'zurich': {'country': 'CH', 'city': 'Zurich'},
        'geneva': {'country': 'CH', 'city': 'Geneva'},
        'new york': {'country': 'US', 'city': 'New York'},
        'chicago': {'country': 'US', 'city': 'Chicago'},
        'los angeles': {'country': 'US', 'city': 'Los Angeles'},
        'bangkok': {'country': 'TH', 'city': 'Bangkok'},
        'kuala lumpur': {'country': 'MY', 'city': 'Kuala Lumpur'},
        'seoul': {'country': 'KR', 'city': 'Seoul'},
        'taipei': {'country': 'TW', 'city': 'Taipei'},
        'mumbai': {'country': 'IN', 'city': 'Mumbai'},
        'delhi': {'country': 'IN', 'city': 'Delhi'},
        'beijing': {'country': 'CN', 'city': 'Beijing'},
        'shanghai': {'country': 'CN', 'city': 'Shanghai'}
    }
    
    # Codes pays directs
    country_codes = {
        'france': 'FR', 'germany': 'DE', 'australia': 'AU', 'japan': 'JP',
        'singapore': 'SG', 'uae': 'AE', 'switzerland': 'CH', 'usa': 'US',
        'united states': 'US', 'united kingdom': 'GB', 'uk': 'GB',
        'thailand': 'TH', 'malaysia': 'MY', 'south korea': 'KR', 'korea': 'KR',
        'taiwan': 'TW', 'india': 'IN', 'china': 'CN'
    }
    
    text_lower = text.lower()
    
    # Rechercher les villes
    for city, info in location_mappings.items():
        if city in text_lower:
            location_info["city"] = info["city"]
            location_info["country_code"] = info["country"]
            location_info["location"] = info["city"]
            return location_info
    
    # Rechercher les pays
    for country, code in country_codes.items():
        if country in text_lower:
            location_info["country_code"] = code
            location_info["location"] = country.title()
            return location_info
    
    # Rechercher les codes de pays (FR, DE, etc.)
    country_pattern = r'\b([A-Z]{2})\b'
    matches = re.findall(country_pattern, text)
    valid_codes = ['FR', 'DE', 'AU', 'JP', 'SG', 'AE', 'CH', 'US', 'GB', 'TH', 'MY', 'KR', 'TW', 'IN', 'CN']
    for match in matches:
        if match in valid_codes:
            location_info["country_code"] = match
            return location_info
    
    return location_info

def categorize_expense_enhanced(text: str) -> dict:
    """Catégorisation améliorée des dépenses"""
    text_lower = text.lower()
    
    # Définitions de catégories avec scores de confiance
    categories = {
        "hotel": {
            "keywords": ["hotel", "accommodation", "lodging", "room", "night", "stay", "resort", "inn", "motel"],
            "patterns": [r"hotel\s+\w+", r"room\s+\d+", r"\d+\s+night"],
            "confidence": 0.3
        },
        "meal": {
            "keywords": ["restaurant", "meal", "food", "dining", "cafe", "bistro", "brasserie", "eatery"],
            "patterns": [r"restaurant\s+\w+", r"table\s+\d+", r"menu"],
            "confidence": 0.25
        },
        "breakfast": {
            "keywords": ["breakfast", "petit dejeuner", "morning", "coffee", "croissant"],
            "patterns": [r"breakfast\s+menu", r"petit\s+dejeuner"],
            "confidence": 0.3
        },
        "transport": {
            "keywords": ["taxi", "uber", "metro", "bus", "train", "transport", "ride", "fare"],
            "patterns": [r"taxi\s+\w+", r"uber\s+trip", r"metro\s+ticket"],
            "confidence": 0.3
        },
        "flight": {
            "keywords": ["flight", "airline", "airport", "boarding", "gate", "seat"],
            "patterns": [r"flight\s+\w+", r"gate\s+\w+", r"seat\s+\w+"],
            "confidence": 0.35
        }
    }
    
    best_category = "unknown"
    best_confidence = 0.0
    
    for category, data in categories.items():
        confidence = 0.0
        
        # Vérifier les mots-clés
        keyword_matches = sum(1 for keyword in data["keywords"] if keyword in text_lower)
        if keyword_matches > 0:
            confidence += data["confidence"] * (keyword_matches / len(data["keywords"]))
        
        # Vérifier les patterns
        for pattern in data["patterns"]:
            if re.search(pattern, text_lower):
                confidence += 0.1
        
        if confidence > best_confidence:
            best_confidence = confidence
            best_category = category
    
    return {
        "category": best_category,
        "confidence": min(best_confidence, 0.3)  # Plafonner à 0.3
    }

def extract_vendor_info(text: str) -> dict:
    """Extraction des informations du vendeur/établissement"""
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    if not lines:
        return {"name": None, "confidence": 0.0}
    
    # La première ligne non-vide est souvent le nom du vendeur
    first_line = lines[0]
    
    # Nettoyer les artefacts OCR courants
    cleaned_line = re.sub(r'^[^\w]+|[^\w]+$', '', first_line)
    
    # Vérifier si ça ressemble à un nom d'établissement
    confidence = 0.1
    if len(cleaned_line) > 3 and len(cleaned_line) < 50:
        confidence = 0.2
        
        # Bonus si contient des mots d'établissement
        establishment_words = ["hotel", "restaurant", "cafe", "bistro", "store", "shop", "market"]
        if any(word in cleaned_line.lower() for word in establishment_words):
            confidence = 0.25
    
    return {
        "name": cleaned_line[:50] if cleaned_line else None,
        "confidence": confidence
    }


#######################################################################################################################################
#                           ADMIN ROUTES
#######################################################################################################################################

@app.get("/api/logs")
async def get_activity_logs(session_token: Optional[str] = Cookie(None), limit: int = 100):
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
        
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    logs = get_logs(limit)
    return {
        "success": True,
        "logs": logs,
        "total": len(logs)
    }

@app.get("/api/users")
async def get_users_list(session_token: Optional[str] = Cookie(None)):
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
        
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    users = [
        {
            "username": user["username"],
            "full_name": user["full_name"], 
            "role": user["role"],
            "created_at": user["created_at"]
        }
        for user in USERS_DB.values()
    ]
    
    return {
        "success": True,
        "users": users
    }

@app.get("/api/feedback-stats")
async def get_feedback_stats(session_token: Optional[str] = Cookie(None)):
    """Statistiques sur les feedbacks (admins seulement)"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
        
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    try:
        import csv
        import os
        
        feedback_file = "data/feedback.csv"
        stats = {
            "total_feedback": 0,
            "average_rating": 0.0,
            "rating_distribution": {1: 0, 2: 0, 3: 0, 4: 0, 5: 0},
            "common_issues": {}
        }
        
        if os.path.exists(feedback_file):
            with open(feedback_file, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                ratings = []
                
                for row in reader:
                    stats["total_feedback"] += 1
                    
                    # Rating distribution
                    rating = int(row.get('rating', 0))
                    if 1 <= rating <= 5:
                        stats["rating_distribution"][rating] += 1
                        ratings.append(rating)
                    
                    # Common issues
                    issue_type = row.get('issue_type', 'unknown')
                    stats["common_issues"][issue_type] = stats["common_issues"].get(issue_type, 0) + 1
                
                # Average rating
                if ratings:
                    stats["average_rating"] = round(sum(ratings) / len(ratings), 2)
        
        return {
            "success": True,
            "stats": stats
        }
        
    except Exception as e:
        logger.error(f"Erreur stats feedback: {e}")
        return {
            "success": False,
            "error": str(e)
        }

if __name__ == "__main__":
    print("🚀 T&E Chatbot - APAC Region")
    print("📊 Interface: http://localhost:8000")
    print("📋 Templates: templates/te_index.html")
    print("🎨 Styles: static/js/te_main.js")
    print("ℹ️  Ctrl+C pour arrêter")
    
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        reload=False,
        log_level="info",
        timeout_keep_alive=600,
        limit_max_requests=100,
        workers=1
    )