# user_management.py
"""
User Management System for T&E Chatbot
=====================================

Système simple d'authentification et de gestion des sessions
pour le chatbot T&E région APAC.
"""

from typing import Dict, List, Optional
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path

# Base de données utilisateurs simple (en production: vraie BDD)
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

# Fichier de logs persistant
LOGS_FILE = "data/te_activity_logs.json"

def authenticate_user(username: str, password: str) -> Optional[Dict]:
    """
    Authentifie un utilisateur
    
    Args:
        username: Nom d'utilisateur (email)
        password: Mot de passe en clair
        
    Returns:
        Dict des informations utilisateur si succès, None sinon
    """
    user = USERS_DB.get(username)
    if user:
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        if user["password_hash"] == password_hash:
            # Retourner une copie sans le hash du mot de passe
            user_copy = user.copy()
            del user_copy["password_hash"]
            return user_copy
    return None

def log_activity(username: str, action: str, details: str = ""):
    """
    Enregistre une activité utilisateur dans un fichier JSON persistant
    
    Args:
        username: Nom d'utilisateur
        action: Type d'action effectuée
        details: Détails supplémentaires
    """
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "username": username,
        "action": action,
        "details": details,
        "system": "te_chatbot"
    }
    
    # Charger les logs existants
    logs = []
    if os.path.exists(LOGS_FILE):
        try:
            with open(LOGS_FILE, "r", encoding="utf-8") as f:
                logs = json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            logs = []
    
    # Ajouter le nouveau log
    logs.append(log_entry)
    
    # Limiter à 2000 logs maximum
    if len(logs) > 2000:
        logs = logs[-2000:]
    
    # Sauvegarder dans le fichier
    try:
        Path(LOGS_FILE).parent.mkdir(exist_ok=True, parents=True)
        
        with open(LOGS_FILE, "w", encoding="utf-8") as f:
            json.dump(logs, f, indent=2, ensure_ascii=False)
        
        print(f"TE LOG: {username} - {action} - {details}")
    except Exception as e:
        print(f"ERREUR sauvegarde log TE: {e}")

def get_logs(limit: int = 100) -> List[Dict]:
    """
    Récupère les logs d'activité depuis le fichier JSON
    
    Args:
        limit: Nombre maximum de logs à retourner
        
    Returns:
        List des logs récents
    """
    if not os.path.exists(LOGS_FILE):
        return []
    
    try:
        with open(LOGS_FILE, "r", encoding="utf-8") as f:
            logs = json.load(f)
        
        # Retourner les derniers logs (limité)
        return logs[-limit:] if logs else []
    
    except (json.JSONDecodeError, FileNotFoundError, Exception) as e:
        print(f"ERREUR lecture logs TE: {e}")
        return []

def get_logs_stats() -> Dict:
    """
    Statistiques sur les logs d'activité
    
    Returns:
        Dict avec les statistiques
    """
    logs = get_logs(10000)  # Charger beaucoup pour les stats
    
    if not logs:
        return {"total": 0, "users": 0, "actions": 0}
    
    users = set(log.get("username", "") for log in logs)
    actions = set(log.get("action", "") for log in logs)
    
    # Stats par action
    action_counts = {}
    for log in logs:
        action = log.get("action", "unknown")
        action_counts[action] = action_counts.get(action, 0) + 1
    
    # Stats par utilisateur
    user_counts = {}
    for log in logs:
        user = log.get("username", "unknown")
        user_counts[user] = user_counts.get(user, 0) + 1
    
    return {
        "total": len(logs),
        "unique_users": len(users),
        "unique_actions": len(actions),
        "first_log": logs[0].get("timestamp") if logs else None,
        "last_log": logs[-1].get("timestamp") if logs else None,
        "top_actions": sorted(action_counts.items(), key=lambda x: x[1], reverse=True)[:5],
        "top_users": sorted(user_counts.items(), key=lambda x: x[1], reverse=True)[:5]
    }

def get_user_info(username: str) -> Optional[Dict]:
    """
    Récupère les informations d'un utilisateur
    
    Args:
        username: Nom d'utilisateur
        
    Returns:
        Dict des informations utilisateur (sans mot de passe)
    """
    user = USERS_DB.get(username)
    if user:
        user_copy = user.copy()
        del user_copy["password_hash"]
        return user_copy
    return None

def get_all_users() -> List[Dict]:
    """
    Récupère la liste de tous les utilisateurs
    
    Returns:
        List des utilisateurs (sans mots de passe)
    """
    users = []
    for username, user_data in USERS_DB.items():
        user_copy = user_data.copy()
        del user_copy["password_hash"]
        users.append(user_copy)
    
    return sorted(users, key=lambda x: x["full_name"])

def add_user(username: str, password: str, full_name: str, role: str = "user", 
             region: str = "APAC", department: str = "Finance") -> bool:
    """
    Ajoute un nouvel utilisateur
    
    Args:
        username: Nom d'utilisateur (email)
        password: Mot de passe en clair
        full_name: Nom complet
        role: Rôle (admin/user)
        region: Région
        department: Département
        
    Returns:
        bool: True si ajouté avec succès
    """
    if username in USERS_DB:
        return False  # Utilisateur existe déjà
    
    USERS_DB[username] = {
        "username": username,
        "password_hash": hashlib.sha256(password.encode()).hexdigest(),
        "full_name": full_name,
        "role": role,
        "region": region,
        "department": department,
        "created_at": datetime.now().isoformat()
    }
    
    log_activity("system", "USER_ADDED", f"New user added: {username} ({full_name})")
    return True

def change_user_password(username: str, new_password: str) -> bool:
    """
    Change le mot de passe d'un utilisateur
    
    Args:
        username: Nom d'utilisateur
        new_password: Nouveau mot de passe
        
    Returns:
        bool: True si changé avec succès
    """
    if username not in USERS_DB:
        return False
    
    USERS_DB[username]["password_hash"] = hashlib.sha256(new_password.encode()).hexdigest()
    log_activity("system", "PASSWORD_CHANGED", f"Password changed for: {username}")
    return True

def update_user_role(username: str, new_role: str) -> bool:
    """
    Met à jour le rôle d'un utilisateur
    
    Args:
        username: Nom d'utilisateur
        new_role: Nouveau rôle
        
    Returns:
        bool: True si mis à jour avec succès
    """
    if username not in USERS_DB or new_role not in ["admin", "user"]:
        return False
    
    old_role = USERS_DB[username]["role"]
    USERS_DB[username]["role"] = new_role
    log_activity("system", "ROLE_UPDATED", f"Role changed for {username}: {old_role} -> {new_role}")
    return True

def get_user_activity_summary(username: str, days: int = 30) -> Dict:
    """
    Résumé d'activité pour un utilisateur
    
    Args:
        username: Nom d'utilisateur
        days: Nombre de jours à analyser
        
    Returns:
        Dict avec le résumé d'activité
    """
    from datetime import datetime, timedelta
    
    # Date limite
    cutoff_date = datetime.now() - timedelta(days=days)
    
    logs = get_logs(5000)  # Récupérer plus de logs pour l'analyse
    
    # Filtrer par utilisateur et date
    user_logs = []
    for log in logs:
        if log.get("username") == username:
            try:
                log_date = datetime.fromisoformat(log["timestamp"])
                if log_date >= cutoff_date:
                    user_logs.append(log)
            except:
                continue
    
    # Analyser les activités
    activities = {}
    for log in user_logs:
        action = log.get("action", "unknown")
        activities[action] = activities.get(action, 0) + 1
    
    # Analyser les jours actifs
    active_days = set()
    for log in user_logs:
        try:
            log_date = datetime.fromisoformat(log["timestamp"])
            active_days.add(log_date.date())
        except:
            continue
    
    return {
        "username": username,
        "period_days": days,
        "total_actions": len(user_logs),
        "active_days": len(active_days),
        "activities": sorted(activities.items(), key=lambda x: x[1], reverse=True),
        "last_activity": user_logs[-1]["timestamp"] if user_logs else None,
        "first_activity_in_period": user_logs[0]["timestamp"] if user_logs else None
    }

def export_logs_csv(filepath: str = "data/te_logs_export.csv") -> bool:
    """
    Exporte les logs au format CSV
    
    Args:
        filepath: Chemin du fichier CSV
        
    Returns:
        bool: True si exporté avec succès
    """
    try:
        import csv
        
        logs = get_logs(10000)  # Tous les logs
        
        if not logs:
            return False
        
        # Créer le dossier si nécessaire
        Path(filepath).parent.mkdir(exist_ok=True, parents=True)
        
        with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['timestamp', 'username', 'action', 'details', 'system']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for log in logs:
                writer.writerow(log)
        
        print(f"Logs exportés vers: {filepath}")
        return True
        
    except Exception as e:
        print(f"Erreur export CSV: {e}")
        return False

# Fonctions utilitaires pour les tests
def create_test_users():
    """Crée des utilisateurs de test supplémentaires"""
    test_users = [
        ("test.tokyo@company.com", "test123", "Tokyo Test User", "user", "APAC", "IT"),
        ("test.sydney@company.com", "test123", "Sydney Test User", "user", "APAC", "HR"),
        ("admin.test@company.com", "admin123", "Test Admin", "admin", "APAC", "Finance")
    ]
    
    for username, password, full_name, role, region, department in test_users:
        if add_user(username, password, full_name, role, region, department):
            print(f"Utilisateur test créé: {username}")
        else:
            print(f"Utilisateur test existe déjà: {username}")

def test_user_management():
    """Teste le système de gestion des utilisateurs"""
    print("=== Test User Management ===\n")
    
    # Test d'authentification
    print("1. Test authentification:")
    valid_user = authenticate_user("admin.te@company.com", "admin123")
    invalid_user = authenticate_user("admin.te@company.com", "wrongpass")
    
    print(f"   Utilisateur valide: {valid_user is not None}")
    print(f"   Utilisateur invalide: {invalid_user is None}")
    
    if valid_user:
        print(f"   Nom: {valid_user['full_name']}, Rôle: {valid_user['role']}")
    
    # Test des logs
    print("\n2. Test logs:")
    log_activity("test_user", "TEST_ACTION", "Test d'activité")
    log_activity("test_user", "LOGIN", "Connexion test")
    
    recent_logs = get_logs(5)
    print(f"   Logs récents: {len(recent_logs)}")
    
    # Test des statistiques
    print("\n3. Test statistiques:")
    stats = get_logs_stats()
    print(f"   Total logs: {stats['total']}")
    print(f"   Utilisateurs uniques: {stats['unique_users']}")
    print(f"   Actions uniques: {stats['unique_actions']}")
    
    # Test liste utilisateurs
    print("\n4. Test liste utilisateurs:")
    all_users = get_all_users()
    print(f"   Nombre d'utilisateurs: {len(all_users)}")
    for user in all_users[:3]:  # Afficher les 3 premiers
        print(f"   - {user['username']} ({user['role']})")
    
    # Test résumé d'activité
    print("\n5. Test résumé d'activité:")
    if valid_user:
        summary = get_user_activity_summary(valid_user['username'], 7)
        print(f"   Activités sur 7 jours: {summary['total_actions']}")
        print(f"   Jours actifs: {summary['active_days']}")

if __name__ == "__main__":
    test_user_management()