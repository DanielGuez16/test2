from typing import Dict, List, Optional
from datetime import datetime
import hashlib
import json
import os
import pandas as pd
from sharepoint_connector import SharePointClient

SHAREPOINT_ANALYSIS_PATH = "Chatbot/logs/analysis_history.xlsx"
SHAREPOINT_LOGS_PATH = "Chatbot/logs/activity_logs.xlsx"
SHAREPOINT_FEEDBACK_PATH = "Chatbot/logs/feedback_data.xlsx"

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

# Fichier local de fallback
LOGS_FILE = "activity_logs.json"

def authenticate_user(username: str, password: str) -> Optional[Dict]:
    """Identique à l'ancien projet"""
    user = USERS_DB.get(username)
    if user:
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        if user["password_hash"] == password_hash:
            return user
    return None

def get_sharepoint_client():
    """Identique à l'ancien projet"""
    return SharePointClient()

def log_activity(username: str, action: str, details: str = ""):
    """COPIE EXACTE de user2.py - Version qui fonctionnait"""
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

def get_logs(limit: int = 100) -> List[Dict]:
    """COPIE EXACTE de user2.py"""
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

def get_logs_stats() -> Dict:
    """COPIE EXACTE de user2.py"""
    try:
        client = get_sharepoint_client()
        binary_content = client.read_binary_file(SHAREPOINT_LOGS_PATH)
        logs_data = client.read_excel_file_as_dict(binary_content)
        
        if not logs_data:
            return {"total": 0, "users": 0, "actions": 0}
        
        logs_df = pd.DataFrame(logs_data)
        
        users_count = logs_df['username'].nunique() if 'username' in logs_df.columns else 0
        actions_count = logs_df['action'].nunique() if 'action' in logs_df.columns else 0
        
        first_log = None
        last_log = None
        
        if 'timestamp' in logs_df.columns and not logs_df.empty:
            sorted_logs = logs_df.sort_values('timestamp')
            first_log = sorted_logs.iloc[0]['timestamp']
            last_log = sorted_logs.iloc[-1]['timestamp']
        
        return {
            "total": len(logs_df),
            "users": users_count,
            "actions": actions_count,
            "first_log": first_log,
            "last_log": last_log
        }
        
    except Exception as e:
        print(f"ERREUR stats SharePoint: {e}")
        return _get_logs_stats_fallback()

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

def _get_logs_stats_fallback() -> Dict:
    """Stats locales de secours"""
    logs = _get_logs_fallback(10000)
    
    if not logs:
        return {"total": 0, "users": 0, "actions": 0}
    
    users = set(log.get("username", "") for log in logs)
    actions = set(log.get("action", "") for log in logs)
    
    return {
        "total": len(logs),
        "users": len(users),
        "actions": len(actions),
        "first_log": logs[0].get("timestamp") if logs else None,
        "last_log": logs[-1].get("timestamp") if logs else None
    }


def save_feedback(username: str, rating: int, comment: str = "", issue_type: str = "", analysis_id: str = ""):
    """Sauvegarde un feedback dans SharePoint Excel - Pattern identique aux logs"""
    feedback_entry = {
        "timestamp": datetime.now().isoformat(),
        "username": username,
        "rating": rating,
        "comment": comment,
        "issue_type": issue_type,
        "analysis_id": analysis_id
    }
    
    try:
        client = get_sharepoint_client()
        
        # Tenter de lire le fichier existant
        try:
            binary_content = client.read_binary_file(SHAREPOINT_FEEDBACK_PATH)
            existing_data = client.read_excel_file_as_dict(binary_content)
            feedback_df = pd.DataFrame(existing_data)
        except:
            # Si le fichier n'existe pas, créer un DataFrame vide
            feedback_df = pd.DataFrame(columns=["timestamp", "username", "rating", "comment", "issue_type", "analysis_id"])
        
        # Ajouter le nouveau feedback
        new_feedback_df = pd.DataFrame([feedback_entry])
        feedback_df = pd.concat([feedback_df, new_feedback_df], ignore_index=True)
        
        # Limiter à 500 feedbacks maximum
        if len(feedback_df) > 500:
            feedback_df = feedback_df.tail(500)
        
        # Sauvegarder dans SharePoint
        client.save_dataframe_in_sharepoint(feedback_df, SHAREPOINT_FEEDBACK_PATH, False)
        
        print(f"FEEDBACK: {username} - Rating {rating}/5 - {issue_type}")
        
    except Exception as e:
        print(f"ERREUR sauvegarde feedback SharePoint: {e}")
        # Fallback vers le fichier local
        _save_feedback_fallback(username, rating, comment, issue_type, analysis_id)

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
        
        feedback_df = pd.DataFrame(feedback_data)
        
        # Calculer les stats
        total_feedback = len(feedback_df)
        ratings = pd.to_numeric(feedback_df.get('rating', []), errors='coerce').dropna()
        average_rating = round(ratings.mean(), 2) if not ratings.empty else 0.0
        
        # Distribution des ratings
        rating_distribution = {i: 0 for i in range(1, 6)}
        for rating in ratings:
            if 1 <= rating <= 5:
                rating_distribution[int(rating)] += 1
        
        # Issues communes
        issue_counts = feedback_df.get('issue_type', pd.Series()).value_counts().to_dict()
        common_issues = {k: v for k, v in issue_counts.items() if k and k != ""}
        
        return {
            "total_feedback": total_feedback,
            "average_rating": average_rating,
            "rating_distribution": rating_distribution,
            "common_issues": common_issues
        }
        
    except Exception as e:
        print(f"ERREUR stats feedback SharePoint: {e}")
        return _get_feedback_stats_fallback()

# Fonctions de fallback pour les feedbacks
FEEDBACK_FILE = "feedback_data.json"

def _save_feedback_fallback(username: str, rating: int, comment: str = "", issue_type: str = "", analysis_id: str = ""):
    """Sauvegarde locale de secours pour les feedbacks"""
    feedback_entry = {
        "timestamp": datetime.now().isoformat(),
        "username": username,
        "rating": rating,
        "comment": comment,
        "issue_type": issue_type,
        "analysis_id": analysis_id
    }
    
    feedbacks = []
    if os.path.exists(FEEDBACK_FILE):
        try:
            with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
                feedbacks = json.load(f)
        except:
            feedbacks = []
    
    feedbacks.append(feedback_entry)
    if len(feedbacks) > 500:
        feedbacks = feedbacks[-500:]
    
    try:
        with open(FEEDBACK_FILE, "w", encoding="utf-8") as f:
            json.dump(feedbacks, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"ERREUR sauvegarde feedback fallback: {e}")

def _get_feedback_stats_fallback() -> Dict:
    """Stats locales de secours pour feedbacks"""
    if not os.path.exists(FEEDBACK_FILE):
        return {"total_feedback": 0, "average_rating": 0.0, 
               "rating_distribution": {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}, 
               "common_issues": {}}
    
    try:
        with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
            feedbacks = json.load(f)
    except:
        return {"total_feedback": 0, "average_rating": 0.0, 
               "rating_distribution": {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}, 
               "common_issues": {}}
    
    if not feedbacks:
        return {"total_feedback": 0, "average_rating": 0.0, 
               "rating_distribution": {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}, 
               "common_issues": {}}
    
    total_feedback = len(feedbacks)
    ratings = [f.get("rating", 0) for f in feedbacks if f.get("rating")]
    average_rating = round(sum(ratings) / len(ratings), 2) if ratings else 0.0
    
    rating_distribution = {i: 0 for i in range(1, 6)}
    for rating in ratings:
        if 1 <= rating <= 5:
            rating_distribution[int(rating)] += 1
    
    issues = [f.get("issue_type", "") for f in feedbacks if f.get("issue_type")]
    from collections import Counter
    issue_counts = Counter(issues)
    common_issues = {k: v for k, v in issue_counts.items() if k}
    
    return {
        "total_feedback": total_feedback,
        "average_rating": average_rating,
        "rating_distribution": rating_distribution,
        "common_issues": common_issues
    }



def save_analysis_history(username: str, ticket_filename: str, question: str, analysis_result: dict, ticket_info: dict):
    """Sauvegarde une analyse dans l'historique SharePoint Excel - Pattern identique"""
    analysis_entry = {
        "timestamp": datetime.now().isoformat(),
        "username": username,
        "ticket_filename": ticket_filename,
        "question": question,
        "result": analysis_result.get("result", "UNKNOWN"),
        "expense_type": analysis_result.get("expense_type", "Unknown"),
        "amount": ticket_info.get("amount", ""),
        "currency": ticket_info.get("currency", ""),
        "vendor": ticket_info.get("vendor", ""),
        "confidence": ticket_info.get("confidence", 0)
    }
    
    try:
        client = get_sharepoint_client()
        
        # Tenter de lire le fichier existant
        try:
            binary_content = client.read_binary_file(SHAREPOINT_ANALYSIS_PATH)
            existing_data = client.read_excel_file_as_dict(binary_content)
            analysis_df = pd.DataFrame(existing_data)
        except:
            # Si le fichier n'existe pas, créer un DataFrame vide
            analysis_df = pd.DataFrame(columns=[
                "timestamp", "username", "ticket_filename", "question", 
                "result", "expense_type", "amount", "currency", "vendor", "confidence"
            ])
        
        # Ajouter la nouvelle analyse
        new_analysis_df = pd.DataFrame([analysis_entry])
        analysis_df = pd.concat([analysis_df, new_analysis_df], ignore_index=True)
        
        # Limiter à 1000 analyses maximum
        if len(analysis_df) > 1000:
            analysis_df = analysis_df.tail(1000)
        
        # Sauvegarder dans SharePoint
        client.save_dataframe_in_sharepoint(analysis_df, SHAREPOINT_ANALYSIS_PATH, False)
        
        print(f"ANALYSIS: {username} - {ticket_filename} - {analysis_result.get('result', 'UNKNOWN')}")
        
    except Exception as e:
        print(f"ERREUR sauvegarde analysis SharePoint: {e}")
        # Fallback vers le fichier local
        _save_analysis_fallback(username, ticket_filename, question, analysis_result, ticket_info)

def get_analysis_history(limit: int = 100) -> List[Dict]:
    """Récupère l'historique des analyses depuis SharePoint"""
    try:
        client = get_sharepoint_client()
        binary_content = client.read_binary_file(SHAREPOINT_ANALYSIS_PATH)
        analysis_data = client.read_excel_file_as_dict(binary_content)
        
        if not analysis_data:
            return []
        
        # Convertir en DataFrame
        analysis_df = pd.DataFrame(analysis_data)
        
        # Trier par timestamp décroissant et limiter
        if not analysis_df.empty and 'timestamp' in analysis_df.columns:
            analysis_df = analysis_df.sort_values('timestamp', ascending=False)
            limited_analysis = analysis_df.head(limit)
            
            # Reconstituer la structure attendue par le frontend
            analyses = []
            for _, row in limited_analysis.iterrows():
                analysis = {
                    "timestamp": row.get("timestamp", ""),
                    "user": row.get("username", ""),
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
            
            return analyses
        
        return []
        
    except Exception as e:
        print(f"ERREUR lecture analysis SharePoint: {e}")
        return _get_analysis_fallback(limit)

# Fonctions de fallback pour l'historique des analyses
ANALYSIS_FILE = "analysis_history.json"

def _save_analysis_fallback(username: str, ticket_filename: str, question: str, analysis_result: dict, ticket_info: dict):
    """Sauvegarde locale de secours pour les analyses"""
    analysis_entry = {
        "timestamp": datetime.now().isoformat(),
        "user": username,
        "ticket_filename": ticket_filename,
        "question": question,
        "analysis_result": analysis_result,
        "ticket_info": ticket_info
    }
    
    analyses = []
    if os.path.exists(ANALYSIS_FILE):
        try:
            with open(ANALYSIS_FILE, "r", encoding="utf-8") as f:
                analyses = json.load(f)
        except:
            analyses = []
    
    analyses.append(analysis_entry)
    if len(analyses) > 1000:
        analyses = analyses[-1000:]
    
    try:
        with open(ANALYSIS_FILE, "w", encoding="utf-8") as f:
            json.dump(analyses, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"ERREUR sauvegarde analysis fallback: {e}")

def _get_analysis_fallback(limit: int = 100) -> List[Dict]:
    """Lecture locale de secours pour les analyses"""
    if not os.path.exists(ANALYSIS_FILE):
        return []
    
    try:
        with open(ANALYSIS_FILE, "r", encoding="utf-8") as f:
            analyses = json.load(f)
        
        # Trier par timestamp décroissant
        analyses.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
        return analyses[:limit] if analyses else []
    except Exception as e:
        print(f"ERREUR lecture analysis fallback: {e}")
        return []