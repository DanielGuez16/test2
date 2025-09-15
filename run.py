#!/usr/bin/env python3
# main.py - Chatbot T&E pour région APAC
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
    
    return templates.TemplateResponse("te_index.html", {
        "request": request,
        "title": "T&E Analysis Chatbot",
        "version": "1.0.0",
        "timestamp": datetime.now().isoformat(),
        "user": current_user,
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
    """Charge les documents T&E (Excel + Word) en mémoire"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        log_activity(current_user["username"], "DOCUMENT_LOAD", f"Loading T&E documents: {excel_file.filename}, {word_file.filename}")
        
        # Lire les fichiers en mémoire
        excel_content = await excel_file.read()
        word_content = await word_file.read()
        
        # Traiter les documents avec le processeur T&E
        excel_rules = te_processor.process_excel_rules(excel_content, excel_file.filename)
        word_policies = te_processor.process_word_policies(word_content, word_file.filename)
        
        # Stocker en mémoire globale
        te_documents["excel_rules"] = excel_rules
        te_documents["word_policies"] = word_policies
        te_documents["last_loaded"] = datetime.now().isoformat()
        
        total_rules = sum(len(rules) for rules in excel_rules.values())
        logger.info(f"Documents T&E chargés: {total_rules} règles Excel, Policies Word: {len(word_policies)} chars")
        
        return {
            "success": True,
            "message": "T&E documents loaded successfully",
            "excel_rules_count": len(excel_rules),
            "word_policies_length": len(word_policies),
            "loaded_at": te_documents["last_loaded"]
        }
        
    except Exception as e:
        logger.error(f"Erreur chargement documents T&E: {e}")
        raise HTTPException(status_code=500, detail=f"Error loading T&E documents: {str(e)}")

@app.post("/api/analyze-ticket")
async def analyze_ticket(
    ticket_file: UploadFile = File(...),
    question: str = Form(""),
    session_token: Optional[str] = Cookie(None)
):
    """Analyse un ticket contre les règles T&E"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if not te_documents["excel_rules"]:
        raise HTTPException(status_code=400, detail="T&E documents not loaded. Please load Excel and Word files first.")
    
    try:
        log_activity(current_user["username"], "TICKET_ANALYSIS", f"Analyzing ticket: {ticket_file.filename}")
        
        # Lire le ticket
        ticket_content = await ticket_file.read()
        
        # Extraire les informations du ticket (OCR si nécessaire)
        ticket_info = extract_ticket_information(ticket_content, ticket_file.filename)
        
        # Analyser contre les règles T&E
        analysis_result = analyze_against_te_rules(ticket_info, te_documents["excel_rules"])
        
        # Préparer le contexte pour l'IA
        context = prepare_te_context(ticket_info, analysis_result, te_documents)
        
        # Obtenir la réponse IA
        user_prompt = f"Question: {question}\n\nTicket Information: {json.dumps(ticket_info, indent=2)}"
        ai_response = llm_connector.get_llm_response(user_prompt, context)
        
        # Sauvegarder l'analyse
        analysis_record = {
            "timestamp": datetime.now().isoformat(),
            "user": current_user["username"],
            "ticket_filename": ticket_file.filename,
            "ticket_info": ticket_info,
            "analysis_result": analysis_result,
            "question": question,
            "ai_response": ai_response
        }
        
        chatbot_session["analysis_history"].append(analysis_record)
        
        return {
            "success": True,
            "ticket_info": ticket_info,
            "analysis_result": analysis_result,
            "ai_response": ai_response,
            "timestamp": analysis_record["timestamp"]
        }
        
    except Exception as e:
        logger.error(f"Erreur analyse ticket: {e}")
        raise HTTPException(status_code=500, detail=f"Error analyzing ticket: {str(e)}")

@app.post("/api/chat")
async def chat_with_ai(request: Request, session_token: Optional[str] = Cookie(None)):
    """Endpoint pour le chatbot IA sans ticket"""
    current_user = get_current_user_from_session(session_token)
    if not current_user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        data = await request.json()
        user_message = data.get("message", "")
        
        logger.info(f"Chat message reçu de {current_user['username']}: {user_message[:100]}")
        
        if not user_message.strip():
            raise HTTPException(status_code=400, detail="Message vide")
        
        log_activity(current_user["username"], "CHAT_MESSAGE", f"Chat message: {user_message[:100]}...")
        
        # Préparer le contexte T&E
        context = prepare_general_te_context()
        
        # Obtenir la réponse de l'IA
        ai_response = llm_connector.get_llm_response(user_message, context)
        logger.info(f"Réponse IA générée: {len(ai_response)} caractères")
        
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
    """Vérifie le statut des documents T&E"""
    return {
        "documents_loaded": bool(te_documents["excel_rules"]),
        "last_loaded": te_documents["last_loaded"],
        "excel_rules_count": len(te_documents["excel_rules"]) if te_documents["excel_rules"] else 0,
        "word_policies_available": bool(te_documents["word_policies"]),
        "timestamp": datetime.now().isoformat()
    }

#######################################################################################################################################
#                           UTILITY FUNCTIONS
#######################################################################################################################################

def extract_ticket_information(file_content: bytes, filename: str) -> dict:
    """Extrait les informations d'un ticket (OCR + parsing)"""
    try:
        file_ext = Path(filename).suffix.lower()
        
        if file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']:
            # Image - utiliser OCR
            image = Image.open(io.BytesIO(file_content))
            text = pytesseract.image_to_string(image)
        elif file_ext == '.pdf':
            # PDF - extraire le texte
            if PyPDF2:
                try:
                    pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_content))
                    text = ""
                    for page in pdf_reader.pages:
                        text += page.extract_text()
                except Exception as e:
                    logger.warning(f"Erreur lecture PDF: {e}")
                    text = f"PDF file {filename} - text extraction failed"
            else:
                text = f"PDF file {filename} - PyPDF2 not available"
        else:
            # Fichier texte
            text = file_content.decode('utf-8', errors='ignore')
        
        # Parser les informations du ticket
        ticket_info = parse_ticket_text(text)
        ticket_info["filename"] = filename
        ticket_info["raw_text"] = text
        
        return ticket_info
        
    except Exception as e:
        logger.error(f"Erreur extraction ticket: {e}")
        return {
            "filename": filename,
            "error": str(e),
            "raw_text": "",
            "amount": None,
            "currency": None,
            "date": None,
            "vendor": None,
            "category": "unknown"
        }

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