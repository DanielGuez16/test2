
console.log('=== TE_MAIN.JS CHARGÉ ===');

// Test immédiat des fonctions critiques
window.testFunctions = function() {
    console.log('sendMessage:', typeof window.sendMessage);
    console.log('showAdminPanel:', typeof window.showAdminPanel);
    console.log('logout:', typeof window.logout);
};

// Appel automatique
setTimeout(() => {
    console.log('Test auto des fonctions:');
    window.testFunctions();
}, 1000);

// te_main.js - JavaScript principal pour T&E Chatbot
"use strict";

// Variables globales
let chatHistory = [];
let currentTicketFile = null;
let currentAnalysisId = null;
let selectedRating = 0;

// Initialisation quand le DOM est chargé
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    setupEventListeners();
    checkTEStatus();
});

function initializeApp() {
    console.log('T&E Chatbot initialized');
    
    // Vérifier si l'utilisateur est connecté
    if (!document.querySelector('.navbar-custom')) {
        window.location.href = '/';
        return;
    }
    
    // Focus sur l'input de chat
    const chatInput = document.getElementById('chat-input');
    if (chatInput) {
        chatInput.focus();
    }
    
    // Charger l'historique récent
    loadRecentHistory();
}

function setupEventListeners() {
    console.log('Setting up event listeners...');
    
    // Upload de documents T&E
    const excelInput = document.getElementById('excel-file');
    const wordInput = document.getElementById('word-file');
    
    if (excelInput) {
        excelInput.addEventListener('change', function() {
            handleDocumentSelect(this, 'excel');
        });
    }
    
    if (wordInput) {
        wordInput.addEventListener('change', function() {
            handleDocumentSelect(this, 'word');
        });
    }
    
    // Boutons principaux
    const sendBtn = document.querySelector('.btn-chat');
    if (sendBtn) {
        sendBtn.addEventListener('click', sendMessage);
    }
    
    const analyzeBtn = document.getElementById('analyze-btn');
    if (analyzeBtn) {
        analyzeBtn.addEventListener('click', analyzeTicket);
    }
    
    // Upload de ticket
    const ticketUpload = document.getElementById('ticket-upload');
    if (ticketUpload) {
        ticketUpload.addEventListener('change', function() {
            if (this.files && this.files[0]) {
                handleTicketUpload(this.files[0]);
            }
        });
    }
    
    // Chat input enter key
    const chatInput = document.getElementById('chat-input');
    if (chatInput) {
        chatInput.addEventListener('keypress', handleChatKeyPress);
    }
    
    // Drag & drop pour les zones d'upload
    setupDragAndDrop();
    
    // Rating stars
    setupRatingStars();
    
    console.log('Event listeners setup complete');
}

function setupDragAndDrop() {
    // Zone upload ticket
    const ticketUploadArea = document.getElementById('ticket-upload-area');
    if (ticketUploadArea) {
        ticketUploadArea.addEventListener('dragover', function(e) {
            e.preventDefault();
            this.classList.add('drag-over');
        });
        
        ticketUploadArea.addEventListener('dragleave', function(e) {
            e.preventDefault();
            this.classList.remove('drag-over');
        });
        
        ticketUploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            this.classList.remove('drag-over');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleTicketUpload(files[0]);
            }
        });
    }
    
    // Zones upload documents T&E
    const uploadAreas = document.querySelectorAll('.upload-area');
    uploadAreas.forEach(area => {
        area.addEventListener('dragover', function(e) {
            e.preventDefault();
            this.classList.add('drag-over');
        });
        
        area.addEventListener('dragleave', function(e) {
            e.preventDefault();
            this.classList.remove('drag-over');
        });
        
        area.addEventListener('drop', function(e) {
            e.preventDefault();
            this.classList.remove('drag-over');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                // Déterminer le type basé sur l'ID du parent
                const isExcel = this.closest('.col-md-6').querySelector('#excel-file') !== null;
                const targetInput = isExcel ? document.getElementById('excel-file') : document.getElementById('word-file');
                
                // Simuler la sélection de fichier
                const dt = new DataTransfer();
                dt.items.add(files[0]);
                targetInput.files = dt.files;
                
                handleDocumentSelect(targetInput, isExcel ? 'excel' : 'word');
            }
        });
    });
}

function setupRatingStars() {
    const stars = document.querySelectorAll('#rating-stars .star');
    stars.forEach(star => {
        star.addEventListener('click', function() {
            selectedRating = parseInt(this.dataset.rating);
            updateStarsDisplay();
        });
        
        star.addEventListener('mouseover', function() {
            const rating = parseInt(this.dataset.rating);
            highlightStars(rating);
        });
    });
    
    const ratingContainer = document.getElementById('rating-stars');
    if (ratingContainer) {
        ratingContainer.addEventListener('mouseleave', function() {
            updateStarsDisplay();
        });
    }
}

function highlightStars(rating) {
    const stars = document.querySelectorAll('#rating-stars .star');
    stars.forEach((star, index) => {
        if (index < rating) {
            star.classList.add('active');
        } else {
            star.classList.remove('active');
        }
    });
}

function updateStarsDisplay() {
    highlightStars(selectedRating);
}

// ===== FONCTIONS DE CHAT =====

function handleChatKeyPress(event) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        sendMessage();
    }
}

function sendMessage() {
    const chatInput = document.getElementById('chat-input');
    const message = chatInput ? chatInput.value.trim() : '';
    
    if (!message) {
        console.log('Message vide, aucune action');
        return;
    }
    
    console.log('Envoi du message:', message);
    
    // Ajouter le message utilisateur immédiatement
    addMessageToChat('user', message);
    chatInput.value = '';
    
    // Envoyer à l'API
    sendChatMessage(message);
}
function addMessageToChat(type, message, timestamp = null) {
    const chatContainer = document.getElementById('chat-container');
    
    if (!timestamp) {
        timestamp = new Date().toLocaleTimeString();
    }
    
    const messageDiv = document.createElement('div');
    messageDiv.className = `chat-message ${type}`;
    
    if (type === 'user') {
        messageDiv.innerHTML = `
            <div class="d-flex justify-content-end">
                <div>
                    <div class="message-content">${escapeHtml(message)}</div>
                    <small class="text-light opacity-75">${timestamp}</small>
                </div>
            </div>
        `;
    } else {
        messageDiv.innerHTML = `
            <div>
                <strong><i class="fas fa-robot me-2"></i>T&E Assistant:</strong>
                <div class="message-content mt-2">${formatAIResponse(message)}</div>
                <small class="text-muted">${timestamp}</small>
            </div>
        `;
    }
    
    chatContainer.appendChild(messageDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
    
    // Sauvegarder dans l'historique
    chatHistory.push({
        type: type,
        message: message,
        timestamp: new Date().toISOString()
    });
}

async function sendChatMessage(message) {
    try {
        // Afficher un indicateur de loading
        const loadingDiv = document.createElement('div');
        loadingDiv.className = 'chat-message assistant';
        loadingDiv.innerHTML = `
            <div>
                <strong><i class="fas fa-robot me-2"></i>T&E Assistant:</strong>
                <div class="mt-2">
                    <div class="spinner-border spinner-border-sm me-2"></div>
                    Analyzing your question...
                </div>
            </div>
        `;
        
        const chatContainer = document.getElementById('chat-container');
        chatContainer.appendChild(loadingDiv);
        chatContainer.scrollTop = chatContainer.scrollHeight;
        
        const response = await fetch('/api/chat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ message: message })
        });
        
        const result = await response.json();
        
        // Supprimer l'indicateur de loading
        loadingDiv.remove();
        
        if (result.success) {
            addMessageToChat('assistant', result.response);
        } else {
            addMessageToChat('assistant', 'Sorry, I encountered an error processing your request. Please try again.');
        }
        
    } catch (error) {
        console.error('Error sending chat message:', error);
        
        // Supprimer l'indicateur de loading s'il existe
        const loadingMsg = document.querySelector('.chat-message.assistant:last-child');
        if (loadingMsg && loadingMsg.innerHTML.includes('spinner-border')) {
            loadingMsg.remove();
        }
        
        addMessageToChat('assistant', 'Sorry, I cannot process your request right now. Please check your connection and try again.');
    }
}

function clearChat() {
    if (confirm('Clear all chat messages?')) {
        const chatContainer = document.getElementById('chat-container');
        chatContainer.innerHTML = `
            <div class="chat-message assistant">
                <strong>T&E Assistant:</strong>
                <p>Chat cleared. How can I help you with T&E analysis?</p>
            </div>
        `;
        chatHistory = [];
    }
}

// ===== UPLOAD ET ANALYSE DE TICKETS =====

function handleTicketUpload(file) {
    if (!file) {
        console.log('Aucun fichier fourni');
        return;
    }
    
    console.log('Upload de ticket:', file.name, file.type, file.size);
    
    // Vérifications de base
    const allowedTypes = ['application/pdf', 'image/jpeg', 'image/jpg', 'image/png', 'text/plain'];
    const maxSize = 10 * 1024 * 1024; // 10MB
    
    if (!allowedTypes.includes(file.type) && !file.name.match(/\.(pdf|jpg|jpeg|png|txt|doc|docx)$/i)) {
        alert('Type de fichier non supporté. Utilisez PDF, images ou documents.');
        return;
    }
    
    if (file.size > maxSize) {
        alert('Fichier trop volumineux. Maximum 10MB.');
        return;
    }
    
    currentTicketFile = file;
    
    // Mettre à jour l'interface immédiatement
    const ticketStatus = document.getElementById('ticket-status');
    if (ticketStatus) {
        ticketStatus.innerHTML = `
            <div class="alert alert-success">
                <i class="fas fa-check-circle me-2"></i>
                <strong>Fichier chargé:</strong><br>
                ${file.name}<br>
                <small>${formatFileSize(file.size)}</small>
            </div>
        `;
    }
    
    // Activer le bouton d'analyse
    const analyzeBtn = document.getElementById('analyze-btn');
    if (analyzeBtn) {
        analyzeBtn.disabled = false;
        analyzeBtn.classList.remove('btn-secondary');
        analyzeBtn.classList.add('btn-success');
        console.log('Bouton analyse activé');
    }
    
    // Message dans le chat
    addMessageToChat('assistant', `Fichier "${file.name}" chargé avec succès. Vous pouvez maintenant l'analyser.`);
}

async function analyzeTicket() {
    if (selectedFiles.length === 0) return;
    
    const formData = new FormData();
    selectedFiles.forEach(file => {
        formData.append('files', file);
    });
    
    const question = document.getElementById('question-input').value;
    formData.append('question', question);
    
    try {
        const response = await fetch('/api/analyze-multiple-tickets', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        displayAnalysisResults(data.results);
        
    } catch (error) {
        console.error('Erreur analyse:', error);
    }
}

function displayAnalysisResults(results) {
    const container = document.getElementById('ticket-status');
    let html = '<div class="analysis-results">';
    
    results.forEach((result, index) => {
        if (result.error) {
            html += `
                <div class="alert alert-warning mb-2">
                    <strong>${result.filename}:</strong> ${result.error}
                </div>
            `;
        } else {
            html += `
                <div class="analysis-result mb-3">
                    <h6>${result.filename}</h6>
                    <div class="result-content">
                        <!-- Contenu de l'analyse -->
                    </div>
                </div>
            `;
        }
    });
    
    html += '</div>';
    container.innerHTML = html;
}
// ===== GESTION DES DOCUMENTS T&E =====

function showDocumentUpload() {
    console.log('Opening document upload modal...');
    try {
        const modalElement = document.getElementById('documentUploadModal');
        if (modalElement) {
            const modal = new bootstrap.Modal(modalElement);
            modal.show();
            
            // Reset des status
            const excelStatus = document.getElementById('excel-status');
            const wordStatus = document.getElementById('word-status');
            const uploadBtn = document.getElementById('upload-documents-btn');
            
            if (excelStatus) excelStatus.innerHTML = '';
            if (wordStatus) wordStatus.innerHTML = '';
            if (uploadBtn) uploadBtn.disabled = true;
        } else {
            console.error('Modal documentUploadModal non trouvé');
            alert('Erreur : Modal de chargement de documents non disponible');
        }
    } catch (error) {
        console.error('Erreur ouverture modal documents:', error);
        alert('Erreur lors de l\'ouverture du modal de chargement de documents');
    }
}

function handleDocumentSelect(input, type) {
    if (!input || !input.files || !input.files[0]) return;
    
    const file = input.files[0];
    const statusDiv = document.getElementById(`${type}-status`);
    const uploadBtn = document.getElementById('upload-documents-btn');
    
    if (!statusDiv) {
        console.error(`Status div ${type}-status not found`);
        return;
    }
    
    // Vérifier le type de fichier
    const isValidType = (type === 'excel' && file.name.match(/\.(xlsx|xls)$/i)) ||
                       (type === 'word' && file.name.match(/\.(docx|doc)$/i));
    
    if (!isValidType) {
        statusDiv.innerHTML = `
            <div class="alert alert-danger">
                <i class="fas fa-times me-2"></i>
                Invalid file type for ${type} document
            </div>
        `;
        return;
    }
    
    statusDiv.innerHTML = `
        <div class="alert alert-success">
            <i class="fas fa-check me-2"></i>
            ${file.name} (${formatFileSize(file.size)})
        </div>
    `;
    
    // Vérifier si les deux fichiers sont sélectionnés
    const excelFile = document.getElementById('excel-file').files[0];
    const wordFile = document.getElementById('word-file').files[0];
    
    if (excelFile && wordFile && uploadBtn) {
        uploadBtn.disabled = false;
    }
}

async function uploadTEDocuments() {
    const excelFile = document.getElementById('excel-file').files[0];
    const wordFile = document.getElementById('word-file').files[0];
    const uploadBtn = document.getElementById('upload-documents-btn');
    
    if (!excelFile || !wordFile) {
        alert('Please select both Excel and Word files.');
        return;
    }
    
    // État de loading
    const originalText = uploadBtn.innerHTML;
    uploadBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>Loading...';
    uploadBtn.disabled = true;
    
    try {
        const formData = new FormData();
        formData.append('excel_file', excelFile);
        formData.append('word_file', wordFile);
        
        const response = await fetch('/api/load-te-documents', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success) {
            // Fermer le modal
            const modalElement = document.getElementById('documentUploadModal');
            const modal = bootstrap.Modal.getInstance(modalElement);
            if (modal) modal.hide();
            
            // Mettre à jour le status des documents
            updateDocumentsStatus(true, result.loaded_at);
            
            // Message de succès
            addMessageToChat('assistant', `T&E documents loaded successfully! Found ${result.excel_rules_count} rules in Excel file. Ready for expense analysis.`);
            
        } else {
            throw new Error(result.detail || 'Failed to load documents');
        }
        
    } catch (error) {
        console.error('Error uploading documents:', error);
        alert('Error loading documents: ' + error.message);
    } finally {
        uploadBtn.innerHTML = originalText;
        uploadBtn.disabled = false;
    }
}

async function checkTEStatus() {
    try {
        const response = await fetch('/api/te-status');
        const result = await response.json();
        
        updateDocumentsStatus(result.documents_loaded, result.last_loaded);
        
    } catch (error) {
        console.error('Error checking T&E status:', error);
    }
}

function updateDocumentsStatus(loaded, lastLoaded) {
    const statusDiv = document.querySelector('.documents-status');
    if (!statusDiv) return;
    
    if (loaded) {
        statusDiv.classList.add('loaded');
        statusDiv.innerHTML = `
            <div class="row align-items-center">
                <div class="col-md-8">
                    <h5 class="mb-1">
                        <i class="fas fa-folder-open me-2"></i>
                        T&E Documents Status
                    </h5>
                    <p class="mb-0">
                        <i class="fas fa-check-circle text-success me-1"></i>
                        Documents loaded successfully
                        ${lastLoaded ? `(${formatDateTime(lastLoaded)})` : ''}
                    </p>
                </div>
                <div class="col-md-4 text-end">
                    <button class="btn btn-outline-primary" onclick="showDocumentUpload()">
                        <i class="fas fa-sync me-2"></i>
                        Reload Documents
                    </button>
                </div>
            </div>
        `;
    }
}

// ===== HISTORIQUE =====

function showHistory() {
    const modalElement = document.getElementById('historyModal');
    if (modalElement) {
        const modal = new bootstrap.Modal(modalElement);
        modal.show();
        loadFullHistory();
    }
}

async function loadFullHistory() {
    const historyDiv = document.getElementById('full-history');
    if (!historyDiv) return;
    
    try {
        const response = await fetch('/api/analysis-history');
        const result = await response.json();
        
        if (result.success && result.history.length > 0) {
            historyDiv.innerHTML = result.history.map(item => `
                <div class="history-item">
                    <div class="d-flex justify-content-between align-items-start">
                        <div>
                            <h6 class="mb-1">${item.ticket_filename}</h6>
                            <p class="mb-1">${item.question}</p>
                            <small class="text-muted">
                                <i class="fas fa-user me-1"></i>${item.user} •
                                <i class="fas fa-clock me-1"></i>${formatDateTime(item.timestamp)}
                            </small>
                        </div>
                        <div class="text-end">
                            <span class="status-indicator status-${item.analysis_result.is_valid ? 'approved' : 'pending'}">
                                ${item.analysis_result.status}
                            </span>
                        </div>
                    </div>
                </div>
            `).join('');
        } else {
            historyDiv.innerHTML = `
                <div class="text-center text-muted p-4">
                    <i class="fas fa-inbox fa-2x mb-3"></i>
                    <p>No analysis history found</p>
                </div>
            `;
        }
        
    } catch (error) {
        console.error('Error loading history:', error);
        historyDiv.innerHTML = `
            <div class="alert alert-danger">
                <i class="fas fa-exclamation-triangle me-2"></i>
                Error loading history
            </div>
        `;
    }
}

async function loadRecentHistory() {
    const recentDiv = document.getElementById('recent-history');
    if (!recentDiv) return;
    
    try {
        const response = await fetch('/api/analysis-history');
        const result = await response.json();
        
        if (result.success && result.history.length > 0) {
            const recent = result.history.slice(-5).reverse(); // 5 derniers
            
            recentDiv.innerHTML = recent.map(item => `
                <div class="p-2 border-bottom">
                    <div class="d-flex justify-content-between align-items-center">
                        <div>
                            <small class="fw-bold">${item.ticket_filename}</small><br>
                            <small class="text-muted">${formatDateTime(item.timestamp)}</small>
                        </div>
                        <span class="badge bg-${item.analysis_result.is_valid ? 'success' : 'warning'}">
                            ${item.analysis_result.is_valid ? 'Valid' : 'Review'}
                        </span>
                    </div>
                </div>
            `).join('');
        }
        
    } catch (error) {
        console.error('Error loading recent history:', error);
    }
}

function addToRecentHistory(result) {
    const recentDiv = document.getElementById('recent-history');
    if (!recentDiv || !currentTicketFile) return;
    
    const newItem = document.createElement('div');
    newItem.className = 'p-2 border-bottom';
    newItem.innerHTML = `
        <div class="d-flex justify-content-between align-items-center">
            <div>
                <small class="fw-bold">${currentTicketFile.name}</small><br>
                <small class="text-muted">Just now</small>
            </div>
            <span class="badge bg-${result.analysis_result.is_valid ? 'success' : 'warning'}">
                ${result.analysis_result.is_valid ? 'Valid' : 'Review'}
            </span>
        </div>
    `;
    
    // Ajouter en haut de la liste
    if (recentDiv.children.length > 0) {
        recentDiv.insertBefore(newItem, recentDiv.firstChild);
    } else {
        recentDiv.innerHTML = '';
        recentDiv.appendChild(newItem);
    }
    
    // Limiter à 5 éléments
    while (recentDiv.children.length > 5) {
        recentDiv.removeChild(recentDiv.lastChild);
    }
}

// ===== FEEDBACK =====

function showFeedbackModal() {
    if (!currentAnalysisId) {
        alert('No analysis to provide feedback for.');
        return;
    }
    
    const modalElement = document.getElementById('feedbackModal');
    if (modalElement) {
        const modal = new bootstrap.Modal(modalElement);
        modal.show();
        
        // Reset du formulaire
        selectedRating = 0;
        updateStarsDisplay();
        const issueSelect = document.getElementById('issue-type');
        const commentText = document.getElementById('feedback-comment');
        if (issueSelect) issueSelect.value = '';
        if (commentText) commentText.value = '';
    }
}

async function submitFeedback() {
    if (selectedRating === 0) {
        alert('Please select a rating.');
        return;
    }
    
    const issueType = document.getElementById('issue-type').value;
    const comment = document.getElementById('feedback-comment').value;
    
    try {
        const response = await fetch('/api/feedback', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                analysis_id: currentAnalysisId,
                rating: selectedRating,
                issue_type: issueType,
                comment: comment
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            // Fermer le modal
            const modalElement = document.getElementById('feedbackModal');
            const modal = bootstrap.Modal.getInstance(modalElement);
            if (modal) modal.hide();
            
            // Message de confirmation
            addMessageToChat('assistant', 'Thank you for your feedback! It helps improve the T&E analysis system.');
        } else {
            throw new Error(result.detail || 'Failed to submit feedback');
        }
        
    } catch (error) {
        console.error('Error submitting feedback:', error);
        alert('Error submitting feedback: ' + error.message);
    }
}

// ===== ADMIN PANEL =====

function showAdminPanel() {
    console.log('Opening admin panel...');
    
    // Vérifier que l'utilisateur est admin
    const adminBtn = document.querySelector('[onclick="showAdminPanel()"]');
    if (!adminBtn) {
        console.error('Bouton admin non trouvé - utilisateur pas admin?');
        return;
    }
    
    try {
        const modalElement = document.getElementById('adminModal');
        if (!modalElement) {
            console.error('Modal adminModal non trouvé dans le DOM');
            alert('Panel administrateur non disponible pour cet utilisateur');
            return;
        }
        if (modalElement) {
            const modal = new bootstrap.Modal(modalElement);
            modal.show();
            
            // Charger les données admin avec un délai
            setTimeout(() => {
                loadAdminLogs();
                loadAdminUsers(); 
                loadAdminFeedback();
            }, 100);
        } else {
            console.error('Modal adminModal non trouvé - vérifiez les permissions admin');
            alert('Panel administrateur non disponible');
        }
    } catch (error) {
        console.error('Erreur ouverture panel admin:', error);
        alert('Erreur lors de l\'ouverture du panel administrateur');
    }
}

async function loadAdminLogs() {
    const logsDiv = document.getElementById('admin-logs');
    if (!logsDiv) return;
    
    try {
        const response = await fetch('/api/logs');
        const result = await response.json();
        
        if (result.success) {
            logsDiv.innerHTML = `
                <div class="table-responsive">
                    <table class="table table-sm">
                        <thead>
                            <tr>
                                <th>Time</th>
                                <th>User</th>
                                <th>Action</th>
                                <th>Details</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${result.logs.slice(-50).reverse().map(log => `
                                <tr>
                                    <td><small>${formatDateTime(log.timestamp)}</small></td>
                                    <td><small>${log.username}</small></td>
                                    <td><span class="badge bg-secondary">${log.action}</span></td>
                                    <td><small>${log.details}</small></td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            `;
        }
        
    } catch (error) {
        console.error('Error loading admin logs:', error);
        logsDiv.innerHTML = '<div class="alert alert-danger">Error loading logs</div>';
    }
}

async function loadAdminUsers() {
    const usersDiv = document.getElementById('admin-users');
    if (!usersDiv) return;
    
    try {
        const response = await fetch('/api/users');
        const result = await response.json();
        
        if (result.success) {
            usersDiv.innerHTML = `
                <div class="row">
                    ${result.users.map(user => `
                        <div class="col-md-6 mb-3">
                            <div class="card">
                                <div class="card-body">
                                    <h6 class="card-title">${user.full_name}</h6>
                                    <p class="card-text">
                                        <small class="text-muted">${user.username}</small><br>
                                        <span class="badge bg-${user.role === 'admin' ? 'danger' : 'primary'}">${user.role}</span>
                                    </p>
                                </div>
                            </div>
                        </div>
                    `).join('')}
                </div>
            `;
        }
        
    } catch (error) {
        console.error('Error loading admin users:', error);
        usersDiv.innerHTML = '<div class="alert alert-danger">Error loading users</div>';
    }
}

async function loadAdminFeedback() {
    const feedbackDiv = document.getElementById('admin-feedback');
    if (!feedbackDiv) return;
    
    try {
        const response = await fetch('/api/feedback-stats');
        const result = await response.json();
        
        if (result.success) {
            const stats = result.stats;
            feedbackDiv.innerHTML = `
                <div class="row">
                    <div class="col-md-4">
                        <div class="card bg-light">
                            <div class="card-body text-center">
                                <h4>${stats.total_feedback}</h4>
                                <p class="mb-0">Total Feedback</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card bg-light">
                            <div class="card-body text-center">
                                <h4>${stats.average_rating}/5</h4>
                                <p class="mb-0">Average Rating</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card bg-light">
                            <div class="card-body">
                                <h6>Rating Distribution:</h6>
                                ${Object.entries(stats.rating_distribution).map(([rating, count]) => 
                                    `<div>★${rating}: ${count}</div>`
                                ).join('')}
                            </div>
                        </div>
                    </div>
                </div>
                
                ${stats.common_issues && Object.keys(stats.common_issues).length > 0 ? `
                <div class="mt-3">
                    <h6>Common Issues:</h6>
                    <div class="row">
                        ${Object.entries(stats.common_issues).slice(0, 5).map(([issue, count]) => `
                            <div class="col-md-6 mb-2">
                                <div class="d-flex justify-content-between">
                                    <span class="text-truncate">${issue}</span>
                                    <span class="badge bg-warning">${count}</span>
                                </div>
                            </div>
                        `).join('')}
                    </div>
                </div>
                ` : ''}
            `;
        } else {
            feedbackDiv.innerHTML = `
                <div class="alert alert-info">
                    <i class="fas fa-info-circle me-2"></i>
                    No feedback data available yet
                </div>
            `;
        }
        
    } catch (error) {
        console.error('Error loading admin feedback:', error);
        feedbackDiv.innerHTML = '<div class="alert alert-danger">Error loading feedback</div>';
    }
}

// ===== AUTHENTIFICATION =====

async function logout() {
    if (confirm('Are you sure you want to logout?')) {
        try {
            const response = await fetch('/api/logout', {
                method: 'POST'
            });
            
            const result = await response.json();
            
            if (result.success) {
                window.location.href = result.redirect;
            }
            
        } catch (error) {
            console.error('Error during logout:', error);
            window.location.href = '/';
        }
    }
}

// ===== UTILITAIRES =====

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatAIResponse(text) {
    // Convertir le markdown en HTML si la bibliothèque marked est disponible
    if (typeof marked !== 'undefined') {
        return DOMPurify.sanitize(marked.parse(text));
    }
    
    // Sinon, formatage basique
    return text
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/\*(.*?)\*/g, '<em>$1</em>')
        .replace(/\n/g, '<br>')
        .replace(/`(.*?)`/g, '<code>$1</code>');
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function formatDateTime(isoString) {
    try {
        const date = new Date(isoString);
        return date.toLocaleString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch (error) {
        return isoString.substring(0, 16);
    }
}

let selectedFiles = [];

// Gestion drag & drop multiple
function setupDragAndDrop() {
    const uploadArea = document.getElementById('ticket-upload-area');
    
    uploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.classList.add('drag-over');
    });
    
    uploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        this.classList.remove('drag-over');
    });
    
    uploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        this.classList.remove('drag-over');
        
        const files = Array.from(e.dataTransfer.files);
        handleMultipleTicketUpload(files);
    });
}

function handleMultipleTicketUpload(files) {
    const supportedExtensions = [
        'pdf', 'jpg', 'jpeg', 'png', 'bmp', 'tiff', 'gif', 'webp',
        'doc', 'docx', 'txt', 'csv', 'xlsx', 'xls', 'rtf'
    ];
    
    const validFiles = Array.from(files).filter(file => {
        const extension = file.name.split('.').pop().toLowerCase();
        return supportedExtensions.includes(extension);
    });
    
    if (validFiles.length !== files.length) {
        const invalidCount = files.length - validFiles.length;
        alert(`${invalidCount} fichier(s) ignoré(s) - types non supportés`);
    }
    
    selectedFiles = validFiles;
    displaySelectedFiles();
    document.getElementById('analyze-btn').disabled = selectedFiles.length === 0;
}

function displaySelectedFiles() {
    const container = document.getElementById('ticket-status');
    
    if (selectedFiles.length === 0) {
        container.innerHTML = '';
        return;
    }
    
    let html = '<div class="uploaded-files-list">';
    selectedFiles.forEach((file, index) => {
        html += `
            <div class="file-item">
                <span>${file.name}</span>
                <span class="file-remove" onclick="removeFile(${index})">×</span>
            </div>
        `;
    });
    html += '</div>';
    
    container.innerHTML = html;
}

function removeFile(index) {
    selectedFiles.splice(index, 1);
    displaySelectedFiles();
    document.getElementById('analyze-btn').disabled = selectedFiles.length === 0;
}

// Initialiser au chargement
document.addEventListener('DOMContentLoaded', function() {
    setupDragAndDrop();
});

// ===== FONCTIONS GLOBALES POUR LES ONCLICK =====

// Ces fonctions sont exposées globalement pour les onclick dans le HTML
window.showDocumentUpload = showDocumentUpload;
window.showAdminPanel = showAdminPanel;
window.showHistory = showHistory;
window.showFeedbackModal = showFeedbackModal;
window.submitFeedback = submitFeedback;
window.uploadTEDocuments = uploadTEDocuments;
window.logout = logout;
window.clearChat = clearChat;
window.sendMessage = sendMessage;
window.analyzeTicket = analyzeTicket;

// ===== GESTION D'ERREURS =====

window.addEventListener('error', function(event) {
    console.error('JavaScript Error:', event.error);
});

window.addEventListener('unhandledrejection', function(event) {
    console.error('Unhandled Promise Rejection:', event.reason);
});

// ===== FONCTIONS UTILITAIRES SUPPLÉMENTAIRES =====

// Améliorer la gestion des erreurs réseau
async function makeAPICall(url, options = {}) {
    try {
        const response = await fetch(url, options);
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        return await response.json();
    } catch (error) {
        console.error(`API call failed for ${url}:`, error);
        throw error;
    }
}

// Fonction manquante pour gérer la sélection de fichier ticket
function handleTicketFileSelect() {
    const fileInput = document.getElementById('ticket-upload');
    if (fileInput && fileInput.files && fileInput.files[0]) {
        handleTicketUpload(fileInput.files[0]);
    }
}

async function viewExcelRules() {
    try {
        const response = await fetch('/api/view-excel-rules');
        const result = await response.json();
        
        if (result.success) {
            // Afficher dans un modal ou nouvelle page
            showRulesModal(result.rules);
        } else {
            alert('No Excel rules loaded');
        }
    } catch (error) {
        console.error('Error viewing Excel rules:', error);
        alert('Error loading Excel rules');
    }
}

async function viewWordPolicies() {
    try {
        const response = await fetch('/api/view-word-policies');
        const result = await response.json();
        
        if (result.success) {
            // Afficher dans un modal
            showPoliciesModal(result.policies_text);
        } else {
            alert('No Word policies loaded');
        }
    } catch (error) {
        console.error('Error viewing Word policies:', error);
        alert('Error loading Word policies');
    }
}

async function refreshDocuments() {
    const refreshBtn = document.querySelector('button[onclick="refreshDocuments()"]');
    const originalText = refreshBtn.innerHTML;
    
    refreshBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i>Refreshing...';
    refreshBtn.disabled = true;
    
    try {
        const response = await fetch('/api/refresh-te-documents', { method: 'POST' });
        const result = await response.json();
        
        if (result.success) {
            addMessageToChat('assistant', 'Documents refreshed from SharePoint successfully!');
            // Recharger la page pour mettre à jour le statut
            setTimeout(() => location.reload(), 1000);
        } else {
            throw new Error('Refresh failed');
        }
    } catch (error) {
        console.error('Error refreshing documents:', error);
        addMessageToChat('assistant', 'Error refreshing documents from SharePoint');
    } finally {
        refreshBtn.innerHTML = originalText;
        refreshBtn.disabled = false;
    }
}

function showRulesModal(rules) {
    // Créer un modal pour afficher les règles Excel sous forme de tableau
    console.log('Excel Rules:', rules);
    alert('Excel rules loaded - implement table display in modal');
}

function showPoliciesModal(policies) {
    // Créer un modal pour afficher les politiques Word
    console.log('Word Policies:', policies);
    alert('Word policies loaded - implement text display in modal');
}


// Debug: Log des événements pour diagnostiquer les problèmes
function debugEventListeners() {
    console.log('=== DEBUG: Event Listeners Status ===');
    console.log('Excel input:', document.getElementById('excel-file') ? 'Found' : 'Missing');
    console.log('Word input:', document.getElementById('word-file') ? 'Found' : 'Missing');
    console.log('Chat input:', document.getElementById('chat-input') ? 'Found' : 'Missing');
    console.log('Analyze button:', document.getElementById('analyze-btn') ? 'Found' : 'Missing');
    console.log('Ticket upload:', document.getElementById('ticket-upload') ? 'Found' : 'Missing');
    console.log('Upload area:', document.getElementById('ticket-upload-area') ? 'Found' : 'Missing');
    console.log('Document modal:', document.getElementById('documentUploadModal') ? 'Found' : 'Missing');
    console.log('Admin modal:', document.getElementById('adminModal') ? 'Found' : 'Missing');
    console.log('==========================================');
}

// Appeler le debug après l'initialisation
setTimeout(debugEventListeners, 1000);

// ===== EXPORT POUR TESTS =====

// Exposer les fonctions principales pour les tests
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        sendMessage,
        handleTicketUpload,
        analyzeTicket,
        formatFileSize,
        formatDateTime,
        escapeHtml,
        showDocumentUpload,
        showAdminPanel
    };
}