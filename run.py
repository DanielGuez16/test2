#!/usr/bin/env python3
# run.py - Chatbot T&E pour région APAC
"""
T&E Chatbot System - APAC Region
=================================

Application FastAPI pour analyser automatiquement les tickets T&E
contre les politiques internes et fournir des recommandations.
"""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request, Depends, Cookie, APIRouter
from pydantic import BaseModel
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

# Imports internes
from llm_connector import LLMConnector
from te_document_processor import TEDocumentProcessor
from user_management import (
    authenticate_user, 
    log_activity, 
    get_logs, 
    get_logs_stats,
    USERS_DB, 
    get_analysis_history, 
    get_feedback_stats, 
    save_feedback,
    save_analysis_history,
    clear_all_logs
)
from sharepoint_connector import SharePointClient
from embedding_connector import EMBEDDINGConnector

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
    "excel_binary": None, 
    "word_binary": None,   
    "last_loaded": None,
    "load_status": "not_loaded",
    "error_message": None
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

async def load_te_documents_from_sharepoint():
    """Charge automatiquement les documents T&E depuis SharePoint"""
    global te_documents
    
    try:
        te_documents["load_status"] = "loading"
        logger.info("Chargement automatique des documents T&E depuis SharePoint...")
        
        # Initialiser SharePoint client
        sharepoint_client = SharePointClient()
        
        # Chemins SharePoint
        excel_path = "Chatbot/sources/Consolidated Limits.xlsx"
        word_path = "Chatbot/sources/APAC Travel Entertainment Procedure Mar2025_Clean.docx"
        
        # Charger Excel
        excel_binary = sharepoint_client.read_binary_file(excel_path)
        excel_dict = sharepoint_client.read_excel_file_as_dict(excel_binary)
        excel_rules = te_processor.process_excel_rules_from_dict(excel_dict, "Consolidated Limits.xlsx")
        
        # Charger Word
        word_binary = sharepoint_client.read_binary_file(word_path)
        word_text = sharepoint_client.read_docx_file_as_text(word_binary)
        word_policies = te_processor.process_word_policies_from_text(word_text, "APAC Travel Entertainment Procedure.docx")

        # Stocker les résultats
        te_documents.update({
            "excel_rules": excel_rules,
            "word_policies": word_policies,
            "excel_binary": excel_binary,
            "word_binary": word_binary,
            "last_loaded": datetime.now().isoformat(),
            "load_status": "loaded",
            "error_message": None
        })
        
        logger.info(f"Documents T&E chargés avec succès - {len(excel_rules)} feuilles Excel")
        return True
        
    except Exception as e:
        logger.error(f"Erreur chargement documents SharePoint: {e}")
        te_documents.update({
            "load_status": "error",
            "error_message": str(e)
        })
        return False

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

    return templates.TemplateResponse("te_index.html", {
        "request": request,
        "title": "T&E Analysis Chatbot",
        "version": "1.0.0",
        "timestamp": datetime.now().isoformat(),
        "user": current_user,
        "te_documents": te_documents,
        "documents_loaded": bool(te_documents["excel_rules"]),
        "last_loaded": te_documents["last_loaded"]
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

@app.post("/api/load-te-documents")
async def load_te_documents(
    excel_file: UploadFile = File(...),
    word_file: UploadFile = File(...),
    session_token: Optional[str] = Cookie(None)
):
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        log_activity(current_user["username"], "DOCUMENT_LOAD", f"Loading T&E documents: {excel_file.filename}, {word_file.filename}")
        
        # Lire les fichiers
        excel_content = await excel_file.read()
        word_content = await word_file.read()
        
        # Traiter les documents
        excel_rules = te_processor.process_excel_rules(excel_content, excel_file.filename)
        word_policies = te_processor.process_word_policies(word_content, word_file.filename)
        
        # Stocker en mémoire globale
        te_documents["excel_rules"] = excel_rules
        te_documents["word_policies"] = word_policies
        te_documents["last_loaded"] = datetime.now().isoformat()
        
        # Indexer dans le système RAG
        logger.info("Indexation des documents dans le système RAG...")
        rag_system.index_excel_rules(excel_rules)
        rag_system.index_word_policies(word_policies)
        
        # Obtenir les stats RAG
        rag_stats = rag_system.get_stats()
        logger.info(f"RAG indexé: {rag_stats}")
        
        total_rules = sum(len(rules) for rules in excel_rules.values())
        
        return {
            "success": True,
            "message": "T&E documents loaded and indexed successfully",
            "excel_rules_count": len(excel_rules),
            "total_rules": total_rules,
            "word_policies_length": len(word_policies),
            "rag_stats": rag_stats,
            "loaded_at": te_documents["last_loaded"]
        }
        
    except Exception as e:
        logger.error(f"Erreur chargement documents T&E: {e}")
        raise HTTPException(status_code=500, detail=f"Error loading T&E documents: {str(e)}")
    
@app.post("/api/ticket-preview")
async def ticket_preview(
    ticket_file: UploadFile = File(...),
    session_token: Optional[str] = Cookie(None)
):
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        log_activity(current_user["username"], "TICKET_PREVIEW", f"Preview ticket: {ticket_file.filename}")
        
        # Extraire le texte du fichier
        ticket_content = await ticket_file.read()
        text_info = extract_ticket_information(ticket_content, ticket_file.filename)
        
        if text_info.get("error") or not text_info.get("raw_text"):
            return {
                "success": False,
                "error": "Cannot extract text from this document"
            }
        
        # Extraction IA des données
        analyzer = TicketAnalyzer(rag_system, llm_connector)
        ticket_info = analyzer.ai_extract_ticket_info(text_info["raw_text"], ticket_file.filename)
        
        return {
            "success": True,
            "ticket_info": ticket_info,
            "extraction_confidence": ticket_info.get("confidence", 0.5),
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Erreur preview ticket: {e}")
        raise HTTPException(status_code=500, detail=f"Error previewing ticket: {str(e)}")
    

@app.get("/api/view-excel")
async def view_excel_document(session_token: Optional[str] = Cookie(None)):
    """Visualise le document Excel chargé"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if not te_documents["excel_rules"]:
        raise HTTPException(status_code=404, detail="No Excel document loaded")
    
    try:
        # Convertir les règles en format d'affichage
        sheets = {}
        for sheet_name, rules in te_documents["excel_rules"].items():
            if rules:
                # Créer les colonnes et lignes pour l'affichage
                columns = ['Country', 'Currency', 'Type', 'Amount Limit']
                rows = []
                for rule in rules:
                    rows.append([
                        rule.get('country', 'N/A'),
                        rule.get('currency', 'N/A'), 
                        rule.get('type', 'N/A'),
                        rule.get('amount_limit', 0)
                    ])
                
                sheets[sheet_name] = {
                    'columns': columns,
                    'rows': rows
                }
        
        return {
            "success": True,
            "filename": "Consolidated Limits.xlsx",
            "sheets": sheets,
            "total_rules": sum(len(rules) for rules in te_documents["excel_rules"].values()),
            "last_loaded": te_documents["last_loaded"]
        }
        
    except Exception as e:
        logger.error(f"Erreur visualisation Excel: {e}")
        raise HTTPException(status_code=500, detail=f"Error viewing Excel: {str(e)}")


@app.get("/api/view-word")  
async def view_word_document(session_token: Optional[str] = Cookie(None)):
    """Visualise le document Word chargé"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    # DEBUG: Vérifier le contenu des documents T&E
    logger.info(f"DEBUG - te_documents keys: {list(te_documents.keys())}")
    logger.info(f"DEBUG - word_policies exists: {bool(te_documents.get('word_policies'))}")
    if te_documents.get("word_policies"):
        logger.info(f"DEBUG - word_policies length: {len(te_documents['word_policies'])}")
        logger.info(f"DEBUG - word_policies preview: {te_documents['word_policies'][:200]}...")
    
    if not te_documents["word_policies"]:
        raise HTTPException(status_code=404, detail="No Word document loaded")
    
    try:
        # Diviser le texte en sections pour l'affichage
        text = te_documents["word_policies"]
        sections = []
        
        # DEBUG: Informations sur le traitement
        logger.info(f"DEBUG - Processing text of length: {len(text)}")
        
        # Diviser par paragraphes
        paragraphs = text.split('\n\n')
        logger.info(f"DEBUG - Found {len(paragraphs)} paragraphs")
        
        current_section = {"title": "T&E Policies", "content": ""}
        
        for i, para in enumerate(paragraphs[:50]):  # Limiter à 50 paragraphes
            if len(para.strip()) > 0:
                if len(current_section["content"]) > 2000:  # Nouvelle section tous les 2000 chars
                    sections.append(current_section)
                    current_section = {"title": f"Section {len(sections) + 1}", "content": para}
                else:
                    current_section["content"] += "\n\n" + para
        
        if current_section["content"]:
            sections.append(current_section)
        
        logger.info(f"DEBUG - Created {len(sections)} sections")
        for i, section in enumerate(sections):
            logger.info(f"DEBUG - Section {i}: title='{section['title']}', content_length={len(section['content'])}")
        
        result = {
            "success": True,
            "filename": "APAC Travel Entertainment Procedure.docx",
            "sections": sections,
            "total_sections": len(sections),
            "last_loaded": te_documents["last_loaded"]
        }
        
        logger.info(f"DEBUG - Returning result with {len(sections)} sections")
        return result
        
    except Exception as e:
        logger.error(f"Erreur visualisation Word: {e}")
        raise HTTPException(status_code=500, detail=f"Error viewing Word: {str(e)}")

@app.post("/api/analyze-ticket")
async def analyze_ticket(
    ticket_file: UploadFile = File(...),
    question: str = Form(""),
    session_token: Optional[str] = Cookie(None)
):
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if not te_documents["excel_rules"]:
        raise HTTPException(status_code=400, detail="T&E documents not loaded. Please load Excel and Word files first.")
    
    try:
        log_activity(current_user["username"], "TICKET_ANALYSIS", f"Analyzing ticket: {ticket_file.filename}")
        
        # 1. Extraire le texte du fichier
        ticket_content = await ticket_file.read()
        text_info = extract_ticket_information(ticket_content, ticket_file.filename)
        
        # 2. Si extraction de texte a échoué
        if text_info.get("error") or not text_info.get("raw_text"):
            return {
                "success": True,
                "ticket_info": text_info,
                "analysis_result": {
                    "result": "FAIL",
                    "expense_type": "Erreur d'extraction",
                    "justification": "Impossible d'extraire le texte du document",
                    "comment": "Unable to extract text from this document. Please ensure the file is readable.",
                    "confidence_score": 0.0
                },
                "timestamp": datetime.now().isoformat()
            }
        
        # 3. Extraction IA complète
        analyzer = TicketAnalyzer(rag_system, llm_connector)
        ticket_info = analyzer.ai_extract_ticket_info(text_info["raw_text"], ticket_file.filename)
        
        # 4. Analyse avec les règles
        analysis_result = analyzer.analyze_ticket(ticket_info, te_documents, question)

        # 5. Préparer la réponse
        response_data = {
            "success": True,
            "ticket_info": ticket_info,
            "analysis_result": analysis_result,
            "timestamp": datetime.now().isoformat()
        }
        
        # 6. Sauvegarder l'analyse - CORRECTION ICI
        analysis_record = {
            "timestamp": response_data["timestamp"],
            "user": current_user["username"],
            "ticket_filename": ticket_file.filename,
            "ticket_info": ticket_info,
            "analysis_result": analysis_result,
            "question": question
        }
        
        # Sauvegarder dans SharePoint 
        try:
            save_analysis_history(
                current_user["username"],
                ticket_file.filename,
                question,
                analysis_result,
                ticket_info
            )
        except Exception as e:
            logger.warning(f"Erreur sauvegarde analyse: {e}")
        
        return response_data
        
    except Exception as e:
        logger.error(f"Erreur analyse ticket: {e}")
        raise HTTPException(status_code=500, detail=f"Error analyzing ticket: {str(e)}")


@app.post("/api/chat")
async def chat_with_ai(request: Request, session_token: Optional[str] = Cookie(None)):
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        data = await request.json()
        user_message = data.get("message", "")
        
        log_activity(current_user["username"], "CHAT_MESSAGE", f"Chat message: {user_message[:100]}...")
        
        analyzer = TicketAnalyzer(rag_system, llm_connector)
        
        response_data = analyzer.answer_general_question(user_message, te_documents)
        ai_response = response_data['ai_response']
        print(ai_response)
        
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


@app.get("/api/analysis-history")
async def get_analysis_history_api(session_token: Optional[str] = Cookie(None)):
    """Récupère l'historique des analyses - Version simplifiée"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        history = get_analysis_history(limit=50)
        
        # Filtrer par utilisateur si non admin
        if current_user["role"] != "admin":
            history = [
                record for record in history
                if record.get("user") == current_user["username"]
            ]
        
        return {
            "success": True,
            "history": history,
            "total": len(history)
        }
        
    except Exception as e:
        logger.error(f"Erreur récupération historique: {e}")
        raise HTTPException(status_code=500, detail=f"Error loading history: {str(e)}")

@app.get("/api/te-status")
async def get_te_status():
    """Vérifie le statut des documents T&E"""
    return {
        "documents_loaded": bool(te_documents["excel_rules"]),
        "last_loaded": te_documents["last_loaded"],
        "excel_rules_count": len(te_documents["excel_rules"]) if te_documents["excel_rules"] else 0,
        "word_policies_available": bool(te_documents["word_policies"]),
        "timestamp": datetime.now().isoformat()
    }

@app.post("/api/refresh-documents")
async def refresh_te_documents(session_token: Optional[str] = Cookie(None)):
    """Force le rechargement des documents T&E depuis SharePoint"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    try:
        log_activity(current_user["username"], "REFRESH_DOCUMENTS", "Manual refresh of T&E documents")
        
        success = await load_te_documents_from_sharepoint()
        
        if success:
            return {
                "success": True,
                "message": "Documents refreshed successfully from SharePoint",
                "last_loaded": te_documents["last_loaded"],
                "status": te_documents["load_status"]
            }
        else:
            return {
                "success": False,
                "message": f"Failed to refresh documents: {te_documents.get('error_message', 'Unknown error')}",
                "status": te_documents["load_status"]
            }
            
    except Exception as e:
        logger.error(f"Erreur refresh documents: {e}")
        raise HTTPException(status_code=500, detail=f"Error refreshing documents: {str(e)}")

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

@app.get("/api/logs-stats")
async def get_logs_statistics(session_token: Optional[str] = Cookie(None)):
    """COPIE EXACTE de run2.py"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
        
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    stats = get_logs_stats()
    
    return {
        "success": True,
        "stats": stats
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

@app.post("/api/feedback")
async def submit_feedback(
    request: Request,
    session_token: Optional[str] = Cookie(None)
):
    """Soumet un feedback - Version simplifiée identique à l'ancien projet"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        data = await request.json()
        
        rating = int(data.get("rating", 0))
        comment = data.get("comment", "")
        issue_type = data.get("issue_type", "")
        analysis_id = data.get("analysis_id", "current")
        
        # Utiliser la fonction simplifiée
        save_feedback(current_user["username"], rating, comment, issue_type, analysis_id)

        
        # Aussi logger dans les logs généraux
        log_activity(current_user["username"], "FEEDBACK", 
                    f"Rating: {rating}/5, Issue: {issue_type}")
        
        return {
            "success": True,
            "message": "Feedback submitted successfully"
        }
        
    except Exception as e:
        logger.error(f"Erreur feedback: {e}")
        raise HTTPException(status_code=500, detail=f"Error submitting feedback: {str(e)}")

@app.get("/api/feedback-stats")
async def get_feedback_stats_api(session_token: Optional[str] = Cookie(None)):
    """Statistiques sur les feedbacks (admins seulement)"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
        
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    try:
        stats = get_feedback_stats()
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
    
@app.post("/api/clear-logs")
async def clear_all_logs_api(session_token: Optional[str] = Cookie(None)):
    """Nettoie tous les logs - ADMIN SEULEMENT"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
        
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    try:
        success = clear_all_logs()
        
        if success:
            log_activity(current_user["username"], "CLEAR_LOGS", "Nettoyage complet des logs")
            return {
                "success": True,
                "message": "Tous les logs ont été nettoyés"
            }
        else:
            return {
                "success": False,
                "message": "Erreur lors du nettoyage des logs"
            }
            
    except Exception as e:
        logger.error(f"Erreur nettoyage logs: {e}")
        raise HTTPException(status_code=500, detail=f"Error clearing logs: {str(e)}")


@app.on_event("startup")
async def startup_event():
    """Événements au démarrage de l'application"""
    logger.info("🚀 Démarrage T&E Chatbot - Chargement des documents...")
    await load_te_documents_from_sharepoint()

#######################################################################################################################################
#                           UTILITY FUNCTIONS - EXTRACTION TEXTE SEULEMENT
#######################################################################################################################################

def preprocess_image_for_ocr(image):
    """Préprocessing d'image pour améliorer l'OCR"""
    try:
        import cv2
        import numpy as np
        from PIL import Image, ImageEnhance, ImageFilter
        
        if image.mode != 'L':
            image = image.convert('L')
        
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)
        
        image = image.filter(ImageFilter.SHARPEN)
        
        width, height = image.size
        if width < 1000 or height < 1000:
            scale_factor = max(1000/width, 1000/height)
            new_size = (int(width * scale_factor), int(height * scale_factor))
            image = image.resize(new_size, Image.Resampling.LANCZOS)
        
        img_array = np.array(image)
        img_thresh = cv2.adaptiveThreshold(
            img_array, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
        )
        img_denoised = cv2.medianBlur(img_thresh, 3)
        processed_image = Image.fromarray(img_denoised)
        
        return processed_image
        
    except ImportError:
        logger.warning("OpenCV non disponible, preprocessing basique")
        return image.convert('L')
    except Exception as e:
        logger.warning(f"Erreur preprocessing: {e}")
        return image

def extract_ticket_information(file_content: bytes, filename: str) -> dict:
    """Extraction avancée avec pipeline OCR robuste sans piexif"""
    try:
        file_ext = Path(filename).suffix.lower()
        
        if file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.gif', '.webp']:
            # Utiliser la nouvelle pipeline OCR
            from advanced_ocr import AdvancedOCRProcessor
            
            processor = AdvancedOCRProcessor()
            result = processor.process_ticket_image(file_content, filename)
            
            return {
                "filename": filename,
                "raw_text": result.get("raw_text", ""),
                "file_type": file_ext,
                "extraction_method": result.get("extraction_method", "paddleocr_advanced"),
                "ocr_confidence": result.get("average_confidence", 0.0),
                "ocr_lines": result.get("ocr_lines", []),
                "lines_detected": result.get("lines_detected", 0),
                # Données structurées extraites
                "amount": result.get("total"),
                "currency": result.get("currency", "EUR"),
                "vendor": result.get("merchant"),
                "date": result.get("date"),
                "category": "unknown"  # À déterminer par l'IA
            }
                
        elif file_ext == '.pdf':
            # PDF - essayer d'extraire le texte d'abord
            try:
                pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_content))
                text_parts = []
                
                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text.strip():
                        text_parts.append(page_text)
                
                text = "\n".join(text_parts)
                
                # Si pas de texte extrait, c'est probablement un PDF-image
                if not text.strip():
                    logger.info(f"PDF sans texte détecté, tentative conversion image: {filename}")
                    
                    # Essayer de convertir le PDF en image avec Pillow + PyPDF2
                    try:
                        # Méthode simple : utiliser fitz si disponible, sinon fallback
                        try:
                            import fitz  # PyMuPDF
                            pdf_doc = fitz.open(stream=file_content, filetype="pdf")
                            
                            # Prendre la première page
                            page = pdf_doc[0]
                            
                            # Convertir en image haute résolution
                            mat = fitz.Matrix(2.0, 2.0)  # Zoom x2 pour meilleure qualité
                            pix = page.get_pixmap(matrix=mat)
                            img_data = pix.tobytes("png")
                            
                            pdf_doc.close()
                            
                            # Traiter comme une image avec notre pipeline OCR
                            from advanced_ocr import AdvancedOCRProcessor
                            processor = AdvancedOCRProcessor()
                            result = processor.process_ticket_image(img_data, filename)
                            
                            return {
                                "filename": filename,
                                "raw_text": result.get("raw_text", ""),
                                "file_type": file_ext,
                                "extraction_method": "pdf_to_image_ocr",
                                "ocr_confidence": result.get("average_confidence", 0.0),
                                "amount": result.get("total"),
                                "currency": result.get("currency", "EUR"),
                                "vendor": result.get("merchant"),
                                "date": result.get("date"),
                                "category": "unknown"
                            }
                            
                        except ImportError:
                            logger.warning("PyMuPDF non disponible, tentative alternative...")
                            
                            # Alternative avec pdf2image si disponible
                            try:
                                from pdf2image import convert_from_bytes
                                
                                # Convertir première page en image
                                images = convert_from_bytes(file_content, first_page=1, last_page=1, dpi=200)
                                
                                if images:
                                    # Convertir PIL en bytes pour notre pipeline
                                    img_buffer = io.BytesIO()
                                    images[0].save(img_buffer, format='PNG')
                                    img_data = img_buffer.getvalue()
                                    
                                    # Traiter avec OCR
                                    from advanced_ocr import AdvancedOCRProcessor
                                    processor = AdvancedOCRProcessor()
                                    result = processor.process_ticket_image(img_data, filename)
                                    
                                    return {
                                        "filename": filename,
                                        "raw_text": result.get("raw_text", ""),
                                        "file_type": file_ext,
                                        "extraction_method": "pdf2image_ocr",
                                        "ocr_confidence": result.get("average_confidence", 0.0),
                                        "amount": result.get("total"),
                                        "currency": result.get("currency", "EUR"),
                                        "vendor": result.get("merchant"),
                                        "date": result.get("date"),
                                        "category": "unknown"
                                    }
                                    
                            except ImportError:
                                logger.warning("pdf2image non disponible")
                                
                                # Dernier fallback : message explicatif
                                return {
                                    "filename": filename,
                                    "raw_text": "",
                                    "file_type": file_ext,
                                    "extraction_method": "pdf_image_unsupported",
                                    "error": "PDF contains images but no conversion library available. Please install PyMuPDF (pip install PyMuPDF) or convert to JPG/PNG manually."
                                }
                    
                    except Exception as e:
                        logger.error(f"Erreur conversion PDF-image: {e}")
                        return {
                            "filename": filename,
                            "error": f"PDF image conversion failed: {str(e)}",
                            "raw_text": "",
                            "file_type": file_ext
                        }
                else:
                    # PDF avec texte extractible
                    return {
                        "filename": filename,
                        "raw_text": text,
                        "file_type": file_ext,
                        "extraction_method": "pdf_text"
                    }
                    
            except Exception as e:
                logger.warning(f"Erreur lecture PDF: {e}")
                return {
                    "filename": filename,
                    "error": f"PDF processing failed: {str(e)}",
                    "raw_text": "",
                    "file_type": file_ext
                }

        elif file_ext in ['.docx', '.doc']:
            # Documents Word
            try:
                if file_ext == '.docx':
                    doc = Document(io.BytesIO(file_content))
                    text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
                else:
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
                if file_ext == '.xlsx':
                    df = pd.read_excel(io.BytesIO(file_content), sheet_name=None)
                else:
                    df = pd.read_excel(io.BytesIO(file_content), sheet_name=None, engine='xlrd')
                
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
    
        # RETOURNER SEULEMENT LE TEXTE - PAS D'ANALYSE
        return {
            "filename": filename,
            "raw_text": text[:2000],  # Plus de texte pour l'IA
            "file_type": file_ext,
            "extraction_method": "text_only"
        }

    except Exception as e:
        logger.error(f"Erreur extraction: {e}")
        return {
            "filename": filename,
            "error": str(e),
            "raw_text": "",
            "file_type": Path(filename).suffix.lower()
        }


if __name__ == "__main__":
    print("🚀 T&E Chatbot - APAC Region")
    print("📊 Interface: http://localhost:8001")
    print("📋 Templates: templates/te_index.html")
    print("🎨 Styles: static/js/te_main.js")
    print("ℹ️  Ctrl+C pour arrêter")
    
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8001,
        reload=False,
        log_level="info",
        timeout_keep_alive=600,
        limit_max_requests=100,
        workers=1
    )
