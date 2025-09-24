from typing import Dict, List, Optional
from datetime import datetime
import hashlib
import json
import os
import pandas as pd
from io import BytesIO
from sharepoint_connector import SharePointClient

# Configuration SharePoint pour les différents fichiers
SHAREPOINT_LOGS_PATH = "Chatbot/logs/activity_logs.xlsx"
SHAREPOINT_ANALYSIS_PATH = "Chatbot/logs/analysis_history.xlsx"
SHAREPOINT_FEEDBACK_PATH = "Chatbot/logs/feedback_data.xlsx"

# Base de données utilisateurs
USERS_DB = {
    "daniel.guez@natixis.com": {
        "username": "daniel.guez@natixis.com",
        "password_hash": hashlib.sha256("admin123".encode()).hexdigest(),
        "full_name": "Daniel GUEZ",
        "role": "admin",
        "created_at": "2024-01-01T00:00:00"
    },
    "franck.pokou-ext@natixis.com": {
        "username": "franck.pokou-ext@natixis.com",
        "password_hash": hashlib.sha256("admin123".encode()).hexdigest(),
        "full_name": "Franck POKOU",
        "role": "admin",
        "created_at": "2024-01-01T00:00:00"
    },
    "juvenalamos.ido@natixis.com": {
        "username": "juvenalamos.ido@natixis.com",
        "password_hash": hashlib.sha256("admin123".encode()).hexdigest(),
        "full_name": "Juvenal Amos IDO",
        "role": "admin",
        "created_at": "2024-01-01T00:00:00"
    },
    "user.te@natixis.com": {
        "username": "user.te@natixis.com",
        "password_hash": hashlib.sha256("user123".encode()).hexdigest(),
        "full_name": "User ALM",
        "role": "user",
        "created_at": "2024-01-01T00:00:00"
    }
}

# Fichiers de fallback local
LOGS_FILE = "data/activity_logs.json"
ANALYSIS_FILE = "data/analysis_history.json"
FEEDBACK_FILE = "data/feedback_data.json"

def authenticate_user(username: str, password: str) -> Optional[Dict]:
    """Authentifie un utilisateur"""
    user = USERS_DB.get(username)
    if user:
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        if user["password_hash"] == password_hash:
            return user
    return None

def get_sharepoint_client():
    """Initialise le client SharePoint"""
    return SharePointClient()

def log_activity(username: str, action: str, details: str = ""):
    """Enregistre une activité utilisateur dans SharePoint Excel"""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "username": username,
        "action": action,
        "details": details
    }
    
    try:
        client = get_sharepoint_client()
        
        # Tenter de lire le fichier existant
        try:
            binary_content = client.read_binary_file(SHAREPOINT_LOGS_PATH)
            existing_data = client.read_excel_file_as_dict(binary_content)
            logs_df = pd.DataFrame(existing_data)
        except:
            # Si le fichier n'existe pas, créer un DataFrame vide
            logs_df = pd.DataFrame(columns=["timestamp", "username", "action", "details"])
        
        # Ajouter le nouveau log
        new_log_df = pd.DataFrame([log_entry])
        logs_df = pd.concat([logs_df, new_log_df], ignore_index=True)
        
        # Limiter à 1000 logs maximum
        if len(logs_df) > 1000:
            logs_df = logs_df.tail(1000)
        
        # Sauvegarder dans SharePoint
        client.save_dataframe_in_sharepoint(logs_df, SHAREPOINT_LOGS_PATH, False)
        
        print(f"LOG: {username} - {action} - {details}")
        
    except Exception as e:
        print(f"ERREUR sauvegarde log SharePoint: {e}")
        # Fallback vers le fichier local en cas d'erreur SharePoint
        _log_activity_fallback(username, action, details)

def _log_activity_fallback(username: str, action: str, details: str = ""):
    """Sauvegarde locale de secours en cas d'erreur SharePoint"""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "username": username,
        "action": action,
        "details": details
    }
    
    logs = []
    if os.path.exists(LOGS_FILE):
        try:
            with open(LOGS_FILE, "r", encoding="utf-8") as f:
                logs = json.load(f)
        except:
            logs = []
    
    logs.append(log_entry)
    if len(logs) > 1000:
        logs = logs[-1000:]
    
    try:
        with open(LOGS_FILE, "w", encoding="utf-8") as f:
            json.dump(logs, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"ERREUR sauvegarde fallback: {e}")

def save_analysis_to_sharepoint(analysis_record: dict):
    """Sauvegarde l'analyse dans SharePoint - ANALYSES SEULEMENT"""
    try:
        client = get_sharepoint_client()
        
        # Aplatir les données pour Excel avec structure claire
        flat_record = {
            "timestamp": analysis_record["timestamp"],
            "user": analysis_record["user"],
            "ticket_filename": analysis_record["ticket_filename"],
            "question": analysis_record.get("question", ""),
            "result": analysis_record["analysis_result"]["result"],
            "expense_type": analysis_record["analysis_result"]["expense_type"],
            "amount": analysis_record["ticket_info"].get("amount", ""),
            "currency": analysis_record["ticket_info"].get("currency", ""),
            "vendor": analysis_record["ticket_info"].get("vendor", ""),
            "confidence": analysis_record["ticket_info"].get("confidence", 0),
            "type": "analysis"  # Identificateur pour éviter les mélanges
        }
        
        # Lire existant ou créer nouveau
        try:
            binary_content = client.read_binary_file(SHAREPOINT_ANALYSIS_PATH)
            existing_data = client.read_excel_file_as_dict(binary_content)
            df = pd.DataFrame(existing_data)
        except:
            df = pd.DataFrame(columns=[
                "timestamp", "user", "ticket_filename", "question", 
                "result", "expense_type", "amount", "currency", 
                "vendor", "confidence"
            ])
        
        # Ajouter nouvelle ligne
        new_df = pd.DataFrame([{k: v for k, v in flat_record.items() if k != "type"}])
        df = pd.concat([df, new_df], ignore_index=True)
        
        # Limiter à 1000 analyses
        if len(df) > 1000:
            df = df.tail(1000)
            
        # Sauvegarder
        client.save_dataframe_in_sharepoint(df, SHAREPOINT_ANALYSIS_PATH, False)
        print(f"✓ ANALYSIS SAVED: {analysis_record['ticket_filename']}")
        
    except Exception as e:
        print(f"ERREUR sauvegarde analyse SharePoint: {e}")
        _save_to_local_file(analysis_record, ANALYSIS_FILE)

def save_feedback_to_sharepoint(feedback_record: dict):
    """Sauvegarde le feedback dans SharePoint - FEEDBACKS SEULEMENT"""
    try:
        client = get_sharepoint_client()
        
        # Structure claire pour les feedbacks
        feedback_data = {
            "timestamp": feedback_record["timestamp"],
            "user": feedback_record["user"],
            "analysis_id": feedback_record.get("analysis_id", ""),
            "rating": feedback_record["rating"],
            "comment": feedback_record.get("comment", ""),
            "issue_type": feedback_record.get("issue_type", ""),
            "type": "feedback"  # Identificateur
        }
        
        # Lire existant ou créer nouveau
        try:
            binary_content = client.read_binary_file(SHAREPOINT_FEEDBACK_PATH)
            existing_data = client.read_excel_file_as_dict(binary_content)
            df = pd.DataFrame(existing_data)
        except:
            df = pd.DataFrame(columns=[
                "timestamp", "user", "analysis_id", "rating", 
                "comment", "issue_type"
            ])
        
        # Ajouter nouveau feedback
        new_df = pd.DataFrame([{k: v for k, v in feedback_data.items() if k != "type"}])
        df = pd.concat([df, new_df], ignore_index=True)
        
        # Limiter à 500 feedbacks
        if len(df) > 500:
            df = df.tail(500)
            
        # Sauvegarder
        client.save_dataframe_in_sharepoint(df, SHAREPOINT_FEEDBACK_PATH, False)
        print(f"✓ FEEDBACK SAVED: {feedback_record['rating']}/5")
        
    except Exception as e:
        print(f"ERREUR sauvegarde feedback SharePoint: {e}")
        _save_to_local_file(feedback_record, FEEDBACK_FILE)

def get_logs(limit: int = 500) -> List[Dict]:
    """Récupère les logs depuis SharePoint"""
    try:
        client = get_sharepoint_client()
        binary_content = client.read_binary_file(SHAREPOINT_LOGS_PATH)
        logs_data = client.read_excel_file_as_dict(binary_content)
        
        # Convertir en DataFrame pour manipulation
        logs_df = pd.DataFrame(logs_data)
        
        # Trier par timestamp et limiter
        if not logs_df.empty and 'timestamp' in logs_df.columns:
            logs_df = logs_df.sort_values('timestamp', ascending=False)
            limited_logs = logs_df.head(limit)
            return limited_logs.to_dict('records')
        
        return []
        
    except Exception as e:
        print(f"ERREUR lecture logs SharePoint: {e}")
        # Fallback vers fichier local
        return _get_logs_fallback(limit)

def _get_logs_fallback(limit: int = 100) -> List[Dict]:
    """Lecture locale de secours"""
    if not os.path.exists(LOGS_FILE):
        return []
    
    try:
        with open(LOGS_FILE, "r", encoding="utf-8") as f:
            logs = json.load(f)
        return logs[-limit:] if logs else []
    except Exception as e:
        print(f"ERREUR lecture logs fallback: {e}")
        return []


def get_analysis_history(limit: int = 100) -> List[Dict]:
    """Récupère l'historique des analyses depuis SharePoint"""
    try:
        client = get_sharepoint_client()
        binary_content = client.read_binary_file(SHAREPOINT_ANALYSIS_PATH)
        analysis_data = client.read_excel_file_as_dict(binary_content)
        
        if not analysis_data:
            return []
        
        # Convertir en DataFrame
        df = pd.DataFrame(analysis_data)
        
        # Reconstituer la structure attendue
        analyses = []
        for _, row in df.iterrows():
            analysis = {
                "timestamp": row.get("timestamp", ""),
                "user": row.get("user", ""),
                "ticket_filename": row.get("ticket_filename", ""),
                "question": row.get("question", ""),
                "analysis_result": {
                    "result": row.get("result", ""),
                    "expense_type": row.get("expense_type", "")
                },
                "ticket_info": {
                    "amount": row.get("amount", ""),
                    "currency": row.get("currency", ""),
                    "vendor": row.get("vendor", ""),
                    "confidence": row.get("confidence", 0)
                }
            }
            analyses.append(analysis)
        
        # Trier par timestamp décroissant
        analyses.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
        
        return analyses[:limit]
        
    except Exception as e:
        print(f"ERREUR lecture analyses SharePoint: {e}")
        return _get_from_local_file(ANALYSIS_FILE, limit)

def get_feedback_stats() -> Dict:
    """Récupère les statistiques des feedbacks depuis SharePoint"""
    try:
        client = get_sharepoint_client()
        binary_content = client.read_binary_file(SHAREPOINT_FEEDBACK_PATH)
        feedback_data = client.read_excel_file_as_dict(binary_content)
        
        if not feedback_data:
            return {"total_feedback": 0, "average_rating": 0.0, 
                   "rating_distribution": {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}, 
                   "common_issues": {}}
        
        df = pd.DataFrame(feedback_data)
        
        # Calculer les stats
        total_feedback = len(df)
        ratings = pd.to_numeric(df.get('rating', []), errors='coerce').dropna()
        average_rating = round(ratings.mean(), 2) if not ratings.empty else 0.0
        
        # Distribution des ratings
        rating_distribution = {i: 0 for i in range(1, 6)}
        for rating in ratings:
            if 1 <= rating <= 5:
                rating_distribution[int(rating)] += 1
        
        # Issues communes
        issue_counts = df.get('issue_type', pd.Series()).value_counts().to_dict()
        common_issues = {k: v for k, v in issue_counts.items() if k and k != ""}
        
        return {
            "total_feedback": total_feedback,
            "average_rating": average_rating,
            "rating_distribution": rating_distribution,
            "common_issues": common_issues
        }
        
    except Exception as e:
        print(f"ERREUR stats feedback SharePoint: {e}")
        return {"total_feedback": 0, "average_rating": 0.0, 
               "rating_distribution": {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}, 
               "common_issues": {}}

def _save_to_local_file(record: dict, filepath: str):
    """Sauvegarde locale de secours"""
    try:
        # Créer le dossier si nécessaire
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        
        # Lire les données existantes
        data = []
        if os.path.exists(filepath):
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except:
                data = []
        
        # Ajouter le nouveau record
        data.append(record)
        
        # Limiter la taille selon le type
        if "activity_logs" in filepath:
            data = data[-2000:]
        elif "analysis_history" in filepath:
            data = data[-1000:]
        elif "feedback_data" in filepath:
            data = data[-500:]
        
        # Sauvegarder
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
            
        print(f"✓ Fallback local: {filepath}")
        
    except Exception as e:
        print(f"ERREUR sauvegarde fallback: {e}")

def _get_from_local_file(filepath: str, limit: int = 100) -> List[Dict]:
    """Lecture locale de secours"""
    try:
        if not os.path.exists(filepath):
            return []
        
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # Trier par timestamp décroissant
        if isinstance(data, list) and data:
            data.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
        
        return data[:limit] if data else []
        
    except Exception as e:
        print(f"ERREUR lecture fallback: {e}")
        return []

def clear_all_logs():
    """Fonction utilitaire pour nettoyer tous les logs (admin seulement)"""
    try:
        client = get_sharepoint_client()
        
        # Créer des DataFrames vides
        empty_logs = pd.DataFrame(columns=["timestamp", "username", "action", "details"])
        empty_analysis = pd.DataFrame(columns=[
            "timestamp", "user", "ticket_filename", "question", 
            "result", "expense_type", "amount", "currency", "vendor", "confidence"
        ])
        empty_feedback = pd.DataFrame(columns=[
            "timestamp", "user", "analysis_id", "rating", "comment", "issue_type"
        ])
        
        # Sauvegarder les fichiers vides
        client.save_dataframe_in_sharepoint(empty_logs, SHAREPOINT_LOGS_PATH, False)
        client.save_dataframe_in_sharepoint(empty_analysis, SHAREPOINT_ANALYSIS_PATH, False)
        client.save_dataframe_in_sharepoint(empty_feedback, SHAREPOINT_FEEDBACK_PATH, False)
        
        print("✓ Tous les logs ont été nettoyés")
        return True
        
    except Exception as e:
        print(f"ERREUR nettoyage logs: {e}")
        return False
