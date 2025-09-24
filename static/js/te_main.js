
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
    console.log('T&E Chatbot initialized with new interface');
    
    // Vérifier si l'utilisateur est connecté
    if (!document.querySelector('.header-fixed')) {
        window.location.href = '/';
        return;
    }
    
    // Focus sur l'input de chat si disponible
    const chatInput = document.getElementById('chat-input');
    if (chatInput) {
        setTimeout(() => chatInput.focus(), 500);
    }
    
    // Charger l'historique récent si la fonction existe
    if (typeof loadRecentHistory === 'function') {
        loadRecentHistory();
    }
}

// Ajoutez ce code dans te_main.js, dans la fonction initializeApp()

function initializeApp() {
    console.log('T&E Chatbot initialized with new interface');
    
    // Vérifier si l'utilisateur est connecté
    if (!document.querySelector('.header-fixed')) {
        window.location.href = '/';
        return;
    }
    
    // Focus sur l'input de chat si disponible
    const chatInput = document.getElementById('chat-input');
    if (chatInput) {
        setTimeout(() => chatInput.focus(), 500);
    }
    
    // Charger l'historique récent si la fonction existe
    if (typeof loadRecentHistory === 'function') {
        loadRecentHistory();
    }
    
    // Ajouter le bouton de retour en haut
    addScrollTopButton();
    
    // Gérer le scroll pour afficher/cacher le bouton
    handleScrollButton();
}

function addScrollTopButton() {
    // Créer le bouton de retour en haut s'il n'existe pas déjà
    if (document.getElementById('scroll-top-btn')) return;
    
    const scrollBtn = document.createElement('button');
    scrollBtn.id = 'scroll-top-btn';
    scrollBtn.className = 'scroll-top-btn';
    scrollBtn.innerHTML = '<i class="fas fa-chevron-up"></i>';
    scrollBtn.setAttribute('aria-label', 'Retour en haut');
    scrollBtn.onclick = scrollToTop;
    
    document.body.appendChild(scrollBtn);
}

function handleScrollButton() {
    const scrollBtn = document.getElementById('scroll-top-btn');
    if (!scrollBtn) return;
    
    window.addEventListener('scroll', function() {
        if (window.pageYOffset > 300) {
            scrollBtn.classList.add('show');
        } else {
            scrollBtn.classList.remove('show');
        }
    });
}

// Également ajouter cette fonction pour détecter quand l'utilisateur scroll dans la zone principale
function addScrollDetection() {
    const mainContent = document.querySelector('.main-content');
    if (!mainContent) return;
    
    mainContent.addEventListener('scroll', function() {
        const scrollBtn = document.getElementById('scroll-top-btn');
        if (scrollBtn) {
            if (this.scrollTop > 200) {
                scrollBtn.classList.add('show');
            } else {
                scrollBtn.classList.remove('show');
            }
        }
    });
}

function setupEventListeners() {
    console.log('Setting up event listeners...');
    
    // Upload de ticket - CORRECTION ICI
    const ticketUpload = document.getElementById('ticket-upload');
    if (ticketUpload) {
        ticketUpload.addEventListener('change', function() {
            if (this.files && this.files[0]) {
                handleTicketUpload(this.files[0]);
            }
        });
    }
    
    // Chat input avec auto-resize
    const chatInput = document.getElementById('chat-input');
    if (chatInput) {
        chatInput.addEventListener('keypress', handleChatKeyPress);
        chatInput.addEventListener('input', function() {
            this.style.height = 'auto';
            this.style.height = Math.min(this.scrollHeight, 100) + 'px';
        });
    }
    
    // Question input avec auto-resize
    const questionInput = document.getElementById('question-input');
    if (questionInput) {
        questionInput.addEventListener('input', function() {
            this.style.height = 'auto';
            this.style.height = Math.min(this.scrollHeight, 150) + 'px';
        });
    }
    
    // Ajouter la détection de scroll
    addScrollDetection();

    console.log('Event listeners setup complete');
}

function handleTicketFileSelect() {
    const fileInput = document.getElementById('ticket-upload');
    if (fileInput && fileInput.files && fileInput.files[0]) {
        handleTicketUpload(fileInput.files[0]);
    }
}

// 7. Ajouter variable globale isAnalyzing au début du fichier:
// À ajouter après les autres variables globales existantes
let isAnalyzing = false;

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
    console.log('=== handleTicketUpload called ===');
    console.log('File:', file ? file.name : 'null');
    console.log('isAnalyzing:', isAnalyzing);
    
    if (!file || isAnalyzing) {
        console.log('No file provided or already analyzing, exiting');
        return;
    }
    
    // Vérifications
    const maxSize = 10 * 1024 * 1024;
    if (file.size > maxSize) {
        showTicketStatus('error', 'File too large. Maximum 10MB allowed.');
        return;
    }
    
    currentTicketFile = file;
    console.log('File stored:', currentTicketFile.name);
    
    // Mettre à jour l'interface
    updateUploadUI(file);
    
    // Activer le bouton d'analyse
    const analyzeBtn = document.getElementById('analyze-btn');
    if (analyzeBtn) {
        console.log('Activating analyze button');
        analyzeBtn.disabled = false;
        analyzeBtn.classList.remove('disabled');
        analyzeBtn.innerHTML = '<i class="fas fa-magic me-2"></i>Analyze Ticket';
    } else {
        console.error('Analyze button not found!');
    }
}

function updateUploadUI(file) {
    const uploadCard = document.getElementById('ticket-upload-area');
    if (!uploadCard) return;
    
    // Mettre à jour l'icône
    const uploadIcon = uploadCard.querySelector('.upload-icon i');
    if (uploadIcon) {
        uploadIcon.className = 'fas fa-file-check';
        uploadIcon.style.color = 'var(--success)';
    }
    
    // Mettre à jour le titre - chercher h4 au lieu de h3
    const title = uploadCard.querySelector('h4');
    if (title) {
        title.textContent = 'File Ready for Analysis';
    }
    
    // Mettre à jour la description
    const description = uploadCard.querySelector('p');
    if (description) {
        description.innerHTML = `<strong>${file.name}</strong><br><small class="text-muted">${formatFileSize(file.size)}</small>`;
    }
    
    // Afficher le statut
    showTicketStatus('success', `File "${file.name}" loaded successfully`);
}

function resetUploadUI() {
    const uploadCard = document.getElementById('ticket-upload-area');
    if (!uploadCard) return;
    
    // Réinitialiser l'icône
    const uploadIcon = uploadCard.querySelector('.upload-icon i');
    if (uploadIcon) {
        uploadIcon.className = 'fas fa-cloud-upload-alt';
        uploadIcon.style.color = 'var(--primary)';
    }
    
    // Réinitialiser le titre - chercher h4 au lieu de h3
    const title = uploadCard.querySelector('h4');
    if (title) {
        title.textContent = 'Drop Your Ticket Here';
    }
    
    // Réinitialiser la description
    const description = uploadCard.querySelector('p');
    if (description) {
        description.innerHTML = 'Drag and drop your expense ticket or click to browse';
    }
}

function showTicketStatus(type, message) {
    const statusDiv = document.getElementById('ticket-status');
    if (!statusDiv) return;
    
    const statusClass = `status-${type}`;
    
    statusDiv.innerHTML = `
        <div class="alert ${statusClass} border-0 rounded-3 mt-2">
            <i class="fas fa-info-circle me-2"></i>
            ${message}
        </div>
    `;
}

async function analyzeTicket() {
    if (!currentTicketFile || isAnalyzing) {
        console.log('Cannot analyze: no file or already analyzing');
        return;
    }
    
    console.log('Starting ticket analysis...'); 
    
    isAnalyzing = true;
    
    // Désactiver le bouton et afficher loading
    const analyzeBtn = document.getElementById('analyze-btn');
    if (analyzeBtn) {
        analyzeBtn.innerHTML = '<div class="spinner-custom"></div><span class="ms-2">Analyzing...</span>';
        analyzeBtn.disabled = true;
        analyzeBtn.style.pointerEvents = 'none';
    }
    
    // Afficher la zone de résultats avec loading - utiliser le bon conteneur
    const resultsSection = document.getElementById('analysis-results');
    const resultContent = document.getElementById('result-content');
    
    if (resultsSection && resultContent) {
        resultsSection.style.display = 'block';
        resultContent.innerHTML = `
            <div class="result-card">
                <div class="text-center py-4">
                    <div class="spinner-custom mb-3"></div>
                    <h5 class="text-primary">Analyzing Your Ticket</h5>
                    <p class="text-muted">AI is processing the document...</p>
                </div>
            </div>
        `;
    }
    
    try {
        const formData = new FormData();
        formData.append('ticket_file', currentTicketFile);
        
        const question = document.getElementById('question-input')?.value || '';
        formData.append('question', question);
        
        console.log('Sending request to API...'); 
        
        const response = await fetch('/api/analyze-ticket', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        
        console.log('API response received:', data.success); 
        
        if (data.success) {
            displayAnalysisResult(data); 
        } else {
            displayAnalysisError(data.detail || 'Analysis failed');
        }
        
    } catch (error) {
        console.error('Erreur analyse:', error);
        displayAnalysisError(error.message);
    } finally {
        // Restaurer le bouton avec un délai
        setTimeout(() => {
            if (analyzeBtn) {
                analyzeBtn.innerHTML = '<i class="fas fa-magic me-2"></i>Analyze Ticket';
                analyzeBtn.disabled = false;
                analyzeBtn.style.pointerEvents = 'auto';
            }
            isAnalyzing = false;
            console.log('Analysis completed, button restored'); 
        }, 500);
    }
}


function showResultNotification(isPass) {
    const notification = document.createElement('div');
    notification.className = `alert alert-${isPass ? 'success' : 'warning'} position-fixed`;
    notification.style.cssText = `
        top: 100px; 
        right: 20px; 
        z-index: 1050; 
        min-width: 300px;
        animation: slideInRight 0.5s ease;
    `;
    
    notification.innerHTML = `
        <div class="d-flex align-items-center justify-content-between">
            <div>
                <i class="fas fa-${isPass ? 'check' : 'exclamation-triangle'} me-2"></i>
                <strong>Analysis Complete</strong>
            </div>
            <button class="btn btn-sm btn-outline-primary" onclick="scrollToResults()">
                View <i class="fas fa-arrow-down ms-1"></i>
            </button>
        </div>
    `;
    
    document.body.appendChild(notification);
    
    // Auto-supprimer après 5 secondes
    setTimeout(() => {
        if (notification.parentNode) {
            notification.remove();
        }
    }, 5000);
}

function scrollToTop() {
    window.scrollTo({
        top: 0,
        behavior: 'smooth'
    });
}

function scrollToResults() {
    const resultsSection = document.getElementById('analysis-results');
    if (resultsSection) {
        resultsSection.scrollIntoView({ 
            behavior: 'smooth', 
            block: 'center' // Centrer au lieu de 'nearest'
        });
    }
}

function scrollToUpload() {
    const uploadArea = document.getElementById('ticket-upload-area');
    if (uploadArea) {
        uploadArea.scrollIntoView({ 
            behavior: 'smooth', 
            block: 'center' 
        });
    }
}

function displayAnalysisResult(result) {
    const resultContent = document.getElementById('result-content');
    if (!resultContent) return;
    
    const analysis = result.analysis_result;
    const ticket = result.ticket_info;
    
    const isPass = analysis.result === 'PASS';
    const statusClass = isPass ? 'success' : 'warning';
    const statusIcon = isPass ? 'fa-check-circle' : 'fa-exclamation-triangle';
    const statusColor = isPass ? 'var(--success)' : 'var(--warning)';
    
    resultContent.innerHTML = `
        <div class="result-card ${statusClass}">
            <div class="result-header">
                <div class="d-flex align-items-center justify-content-between">
                    <div class="d-flex align-items-center">
                        <div class="result-icon" style="color: ${statusColor};">
                            <i class="fas ${statusIcon}"></i>
                        </div>
                        <div>
                            <h4 class="mb-1">${analysis.result}</h4>
                            <p class="text-muted mb-0">${analysis.expense_type}</p>
                        </div>
                    </div>
                    <div class="badge bg-light text-dark px-3 py-2">
                        ${ticket.amount ? `${ticket.amount} ${ticket.currency || ''}` : 'Amount not detected'}
                    </div>
                </div>
            </div>
            
            <div class="result-content">
                ${formatAIResponse(analysis.justification)}
            </div>
            
            <div class="result-actions">
                <small class="text-muted">
                    <i class="fas fa-robot me-1"></i>
                    Analysis completed • ${new Date().toLocaleTimeString()}
                </small>
                <div class="d-flex gap-2">
                    <button class="btn btn-outline-primary btn-sm" onclick="showFeedbackModal()">
                        <i class="fas fa-star me-1"></i> Rate
                    </button>
                    <button class="btn btn-outline-secondary btn-sm" onclick="resetAnalysis()">
                        <i class="fas fa-redo me-1"></i> New Analysis
                    </button>
                </div>
            </div>
        </div>
    `;
}

function displayAnalysisError(errorMessage) {
    const resultContent = document.getElementById('result-content');
    if (!resultContent) return;
    
    resultContent.innerHTML = `
        <div class="result-card danger">
            <div class="result-header">
                <div class="d-flex align-items-center">
                    <div class="result-icon" style="color: var(--danger);">
                        <i class="fas fa-times-circle"></i>
                    </div>
                    <div>
                        <h4 class="mb-1">Analysis Failed</h4>
                        <p class="text-muted mb-0">Unable to process ticket</p>
                    </div>
                </div>
            </div>
            
            <div class="result-content">
                <div class="alert alert-danger border-0">
                    <i class="fas fa-exclamation-triangle me-2"></i>
                    ${errorMessage}
                </div>
            </div>
            
            <div class="result-actions">
                <div></div>
                <button class="btn btn-primary btn-sm" onclick="resetAnalysis()">
                    <i class="fas fa-redo me-2"></i>Try Again
                </button>
            </div>
        </div>
    `;
}

function resetAnalysis() {
    currentTicketFile = null;
    isAnalyzing = false;
    
    // Reset UI
    resetUploadUI();
    
    // Cacher les résultats
    const resultsSection = document.getElementById('analysis-results');
    if (resultsSection) {
        resultsSection.style.display = 'none';
    }
    
    // Reset bouton
    const analyzeBtn = document.getElementById('analyze-btn');
    if (analyzeBtn) {
        analyzeBtn.innerHTML = '<i class="fas fa-magic me-2"></i>Analyze Ticket';
        analyzeBtn.disabled = true;
    }
    
    // Reset question
    const questionInput = document.getElementById('question-input');
    if (questionInput) {
        questionInput.value = '';
        questionInput.style.height = 'auto';
    }
    
    // Reset status
    const ticketStatus = document.getElementById('ticket-status');
    if (ticketStatus) {
        ticketStatus.innerHTML = '';
    }
    
    // Reset file input
    const fileInput = document.getElementById('ticket-upload');
    if (fileInput) {
        fileInput.value = '';
    }
    
    // Scroll vers le haut pour voir la zone d'upload
    scrollToTop();
}


function displaySingleAnalysisResult(result) {
    // Utiliser la nouvelle fonction displayAnalysisResult
    displayAnalysisResult(result);
    
    // Ajouter à l'historique récent si la fonction existe
    if (typeof addToRecentHistory === 'function') {
        addToRecentHistory(result);
    }
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

// ===== VISUALISATION DES DOCUMENTS =====

async function viewExcelDocument() {
    console.log('Opening Excel document viewer...');
    
    try {
        const modalElement = document.getElementById('excelViewModal');
        if (!modalElement) {
            alert('Excel viewer not available');
            return;
        }
        
        const modal = new bootstrap.Modal(modalElement);
        modal.show();
        
        // Charger le contenu Excel
        await loadExcelContent();
        
    } catch (error) {
        console.error('Error opening Excel viewer:', error);
        alert('Error opening Excel document viewer');
    }
}

async function viewWordDocument() {
    console.log('Opening Word document viewer...');
    
    try {
        const modalElement = document.getElementById('wordViewModal');
        if (!modalElement) {
            alert('Word viewer not available');
            return;
        }
        
        const modal = new bootstrap.Modal(modalElement);
        modal.show();
        
        // Charger le contenu Word
        await loadWordContent();
        
    } catch (error) {
        console.error('Error opening Word viewer:', error);
        alert('Error opening Word document viewer');
    }
}

async function loadExcelContent() {
    const contentDiv = document.getElementById('excel-content');
    if (!contentDiv) return;
    
    try {
        const response = await fetch('/api/view-excel');
        const result = await response.json();
        
        if (result.success) {
            let html = `
                <div class="mb-3">
                    <div class="d-flex justify-content-between align-items-center">
                        <h6><i class="fas fa-file-excel text-success me-2"></i>${result.filename}</h6>
                        <div>
                            <span class="badge bg-primary">${result.total_rules} Rules</span>
                            <span class="badge bg-secondary">${Object.keys(result.sheets).length} Sheets</span>
                        </div>
                    </div>
                    <small class="text-muted">Last loaded: ${formatDateTime(result.last_loaded)}</small>
                </div>
            `;
            
            // Créer un accordéon pour chaque sheet
            html += '<div class="accordion" id="excelAccordion">';
            
            Object.entries(result.sheets).forEach(([sheetName, sheetData], index) => {
                const isFirst = index === 0;
                html += `
                    <div class="accordion-item">
                        <h2 class="accordion-header">
                            <button class="accordion-button ${!isFirst ? 'collapsed' : ''}" 
                                    type="button" 
                                    data-bs-toggle="collapse" 
                                    data-bs-target="#sheet-${index}">
                                <strong>${sheetName}</strong>
                                <span class="badge bg-light text-dark ms-2">${sheetData.rows.length} rules</span>
                            </button>
                        </h2>
                        <div id="sheet-${index}" 
                             class="accordion-collapse collapse ${isFirst ? 'show' : ''}" 
                             data-bs-parent="#excelAccordion">
                            <div class="accordion-body">
                                ${createExcelTable(sheetData)}
                            </div>
                        </div>
                    </div>
                `;
            });
            
            html += '</div>';
            contentDiv.innerHTML = html;
            
        } else {
            throw new Error(result.detail || 'Failed to load Excel content');
        }
        
    } catch (error) {
        console.error('Error loading Excel content:', error);
        contentDiv.innerHTML = `
            <div class="alert alert-danger">
                <i class="fas fa-exclamation-triangle me-2"></i>
                Error loading Excel document: ${error.message}
            </div>
        `;
    }
}

function createExcelTable(sheetData) {
    if (!sheetData.rows || sheetData.rows.length === 0) {
        return '<div class="alert alert-info">No data available in this sheet</div>';
    }
    
    let html = `
        <div class="table-responsive">
            <table class="table table-striped table-hover">
                <thead class="table-dark">
                    <tr>
    `;
    
    // Headers
    sheetData.columns.forEach(column => {
        html += `<th>${column}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    // Rows
    sheetData.rows.forEach((row, rowIndex) => {
        html += '<tr>';
        row.forEach((cell, cellIndex) => {
            // Formatage spécial pour les montants (dernière colonne)
            if (cellIndex === row.length - 1 && !isNaN(cell)) {
                html += `<td><strong class="text-success">${parseFloat(cell).toLocaleString()}</strong></td>`;
            } else {
                html += `<td>${escapeHtml(cell)}</td>`;
            }
        });
        html += '</tr>';
    });
    
    html += '</tbody></table></div>';
    
    // Ajouter un résumé
    html += `
        <div class="mt-3 p-2 bg-light rounded">
            <small class="text-muted">
                <i class="fas fa-info-circle me-1"></i>
                ${sheetData.rows.length} rules in this sheet
            </small>
        </div>
    `;
    
    return html;
}

async function loadWordContent() {
    const contentDiv = document.getElementById('word-content');
    if (!contentDiv) return;
    
    try {
        const response = await fetch('/api/view-word');
        const result = await response.json();
        
        if (result.success) {
            let html = `
                <div class="mb-3">
                    <div class="d-flex justify-content-between align-items-center">
                        <h6><i class="fas fa-file-word text-primary me-2"></i>${result.filename}</h6>
                        <div>
                            <span class="badge bg-primary">${result.total_sections} Sections</span>
                        </div>
                    </div>
                    <small class="text-muted">Last loaded: ${formatDateTime(result.last_loaded)}</small>
                </div>
            `;
            
            // Style PDF-like pour le contenu Word
            html += '<div class="word-document-viewer">';
            
            result.sections.forEach((section, index) => {
                html += `
                    <div class="word-section mb-4">
                        <h5 class="section-title text-primary border-bottom pb-2">
                            ${escapeHtml(section.title)}
                        </h5>
                        <div class="section-content">
                            ${formatWordContent(section.content)}
                        </div>
                    </div>
                `;
            });
            
            html += '</div>';
            contentDiv.innerHTML = html;
            
        } else {
            throw new Error(result.detail || 'Failed to load Word content');
        }
        
    } catch (error) {
        console.error('Error loading Word content:', error);
        contentDiv.innerHTML = `
            <div class="alert alert-danger">
                <i class="fas fa-exclamation-triangle me-2"></i>
                Error loading Word document: ${error.message}
            </div>
        `;
    }
}

function formatWordContent(content) {
    // Formatage pour ressembler à un document PDF
    return content
        .split('\n\n')
        .map(paragraph => {
            if (paragraph.trim()) {
                // Détecter les listes
                if (paragraph.includes('•') || paragraph.includes('-')) {
                    const listItems = paragraph.split(/[•-]/).filter(item => item.trim());
                    if (listItems.length > 1) {
                        let listHtml = '<ul class="list-unstyled ms-3">';
                        listItems.forEach(item => {
                            if (item.trim()) {
                                listHtml += `<li class="mb-1"><i class="fas fa-chevron-right text-primary me-2"></i>${escapeHtml(item.trim())}</li>`;
                            }
                        });
                        listHtml += '</ul>';
                        return listHtml;
                    }
                }
                
                // Paragraphe normal
                return `<p class="mb-3 text-justify">${escapeHtml(paragraph.trim())}</p>`;
            }
            return '';
        })
        .join('');
}

async function refreshDocuments() {
    if (!confirm('Refresh T&E documents from SharePoint? This may take a few moments.')) {
        return;
    }
    
    try {
        // Mettre à jour l'interface pour montrer le loading
        const statusText = document.getElementById('documents-status-text');
        if (statusText) {
            statusText.innerHTML = `
                <i class="fas fa-spinner fa-spin text-primary me-1"></i>
                Refreshing documents from SharePoint...
            `;
        }
        
        const response = await fetch('/api/refresh-documents', {
            method: 'POST'
        });
        
        const result = await response.json();
        
        if (result.success) {
            // Rafraîchir la page pour mettre à jour l'interface
            setTimeout(() => {
                window.location.reload();
            }, 1000);
            
            addMessageToChat('assistant', 'T&E documents refreshed successfully from SharePoint!');
        } else {
            throw new Error(result.message || 'Failed to refresh documents');
        }
        
    } catch (error) {
        console.error('Error refreshing documents:', error);
        
        // Restaurer le status d'erreur
        const statusText = document.getElementById('documents-status-text');
        if (statusText) {
            statusText.innerHTML = `
                <i class="fas fa-exclamation-triangle text-danger me-1"></i>
                Error refreshing documents: ${error.message}
            `;
        }
        
        alert('Error refreshing documents: ' + error.message);
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

// Gestion drag & drop
function setupDragAndDrop() {
    console.log('Setting up drag and drop...');
    
    const ticketUploadArea = document.getElementById('ticket-upload-area');
    if (!ticketUploadArea) {
        console.log('Ticket upload area not found');
        return;
    }
    
    // Vérifier si déjà configuré pour éviter les doublons
    if (ticketUploadArea.hasAttribute('data-drag-configured')) {
        console.log('Drag and drop already configured');
        return;
    }
    
    ticketUploadArea.setAttribute('data-drag-configured', 'true');
    
    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        ticketUploadArea.addEventListener(eventName, preventDefaults, false);
    });

    // Highlight drop area when item is dragged over it
    ['dragenter', 'dragover'].forEach(eventName => {
        ticketUploadArea.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        ticketUploadArea.addEventListener(eventName, unhighlight, false);
    });

    // Handle dropped files
    ticketUploadArea.addEventListener('drop', handleDrop, false);
    
    // Handle click to browse files
    ticketUploadArea.addEventListener('click', function(e) {
        if (!isAnalyzing && !e.target.closest('textarea') && !e.target.closest('button')) {
            document.getElementById('ticket-upload').click();
        }
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function highlight() {
        ticketUploadArea.classList.add('drag-over');
    }

    function unhighlight() {
        ticketUploadArea.classList.remove('drag-over');
    }

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleTicketUpload(files[0]);
        }
    }
    
    console.log('Drag and drop configured successfully');
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

// ---- Preview helpers ----
const el = (id) => document.getElementById(id);
const show = (node) => node.classList.remove("hidden");
const hide = (node) => node.classList.add("hidden");

function openPreview(title){
  el("preview-title").textContent = title || "Prévisualisation";
  show(el("preview-panel"));
}
function closePreview(){
  hide(el("preview-panel"));
}
function activateTab(tabId){
  ["tab-docx","tab-excel"].forEach(id => el(id).classList.remove("active"));
  el(tabId).classList.add("active");
  if(tabId === "tab-docx"){
    show(el("docx-view")); hide(el("excel-view"));
  } else {
    hide(el("docx-view")); show(el("excel-view"));
  }
}

// Render DOCX text
async function previewDocx(path){
  const r = await fetch(`/api/preview/docx?path=${encodeURIComponent(path)}`);
  if(!r.ok){ throw new Error("Erreur preview DOCX"); }
  const data = await r.json();
  el("docx-text").textContent = data.text || "(document vide)";
}

// Render Excel sheets
async function previewExcel(path){
  const r = await fetch(`/api/preview/excel?path=${encodeURIComponent(path)}&limit=200`);
  if(!r.ok){ throw new Error("Erreur preview Excel"); }
  const data = await r.json();
  const sheets = data.sheets || {};
  const sheetSelect = el("sheet-select");
  sheetSelect.innerHTML = "";

  const names = Object.keys(sheets);
  if(names.length === 0){
    el("excel-table").innerHTML = "<tbody><tr><td>Aucune donnée</td></tr></tbody>";
    return;
  }
  names.forEach(name => {
    const opt = document.createElement("option");
    opt.value = name; opt.textContent = name;
    sheetSelect.appendChild(opt);
  });

  function renderSheet(name){
    const payload = sheets[name] || {columns:[], rows:[]};
    const cols = payload.columns || [];
    const rows = payload.rows || [];
    const table = el("excel-table");
    let thead = "<thead><tr>";
    cols.forEach(c => thead += `<th style="position:sticky;top:0;background:#fafafa;">${escapeHtml(c)}</th>`);
    thead += "</tr></thead>";
    let tbody = "<tbody>";
    rows.forEach(row => {
      tbody += "<tr>";
      cols.forEach(c => tbody += `<td>${escapeHtml(row[c] ?? "")}</td>`);
      tbody += "</tr>";
    });
    tbody += "</tbody>";
    table.innerHTML = thead + tbody;
  }

  renderSheet(names[0]);
  sheetSelect.onchange = () => renderSheet(sheetSelect.value);
}

// Small HTML escaper for safe rendering
function escapeHtml(s){
  return String(s)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#39;");
}

// ---- Wire UI ----
(function initPreviewUI(){
  const btnDocx = el("btn-preview-docx");
  const btnExcel = el("btn-preview-excel");
  const closeBtn = el("preview-close");
  const tabDocx = el("tab-docx");
  const tabExcel = el("tab-excel");

  if(closeBtn) closeBtn.addEventListener("click", closePreview);

  if(tabDocx) tabDocx.addEventListener("click", () => activateTab("tab-docx"));
  if(tabExcel) tabExcel.addEventListener("click", () => activateTab("tab-excel"));

  if(btnDocx){
    btnDocx.addEventListener("click", async () => {
      const path = el("docx-path").value.trim();
      if(!path) return alert("Renseigne le chemin DOCX (SharePoint).");
      openPreview("Règles (DOCX)");
      activateTab("tab-docx");
      el("docx-text").textContent = "Chargement…";
      try { await previewDocx(path); } catch(e){ el("docx-text").textContent = "Erreur de chargement."; console.error(e); }
    });
  }

  if(btnExcel){
    btnExcel.addEventListener("click", async () => {
      const path = el("excel-path").value.trim();
      if(!path) return alert("Renseigne le chemin Excel (SharePoint).");
      openPreview("Barèmes (Excel)");
      activateTab("tab-excel");
      el("excel-table").innerHTML = "<tbody><tr><td>Chargement…</td></tr></tbody>";
      try { await previewExcel(path); } catch(e){ el("excel-table").innerHTML = "<tbody><tr><td>Erreur de chargement.</td></tr></tbody>"; console.error(e); }
    });
  }
})();

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
window.viewExcelDocument = viewExcelDocument;
window.viewWordDocument = viewWordDocument;
window.refreshDocuments = refreshDocuments;

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

window.viewExcelDocument = async function(){
  const modalEl = document.getElementById('excelViewModal');
  if(!modalEl){ console.error('excelViewModal introuvable'); return; }
  const modal = new bootstrap.Modal(modalEl);
  const mount = document.getElementById('excel-content');
  mount.innerHTML = `<div class="text-center"><div class="spinner-custom"></div><p class="mt-2">Loading Excel content...</p></div>`;
  try{
    const r = await fetch('/api/view-excel');
    const data = await r.json();
    if(!r.ok || !data.success) throw new Error(data.detail || 'Error viewing Excel');
    // rendu simple: 1 table par feuille
    const sheets = data.sheets || {};
    let html = '';
    Object.keys(sheets).forEach(name=>{
      const {columns=[], rows=[]} = sheets[name] || {};
      html += `<h6 class="mt-3">${name}</h6><div class="table-responsive"><table class="table table-sm table-hover"><thead><tr>`;
      columns.forEach(c=> html += `<th>${c}</th>`);
      html += `</tr></thead><tbody>`;
      rows.forEach(row=>{
        html += `<tr>${row.map(v=> `<td>${String(v ?? '')}</td>`).join('')}</tr>`;
      });
      html += `</tbody></table></div>`;
    });
    mount.innerHTML = html || '<div class="text-muted">No data</div>';
    modal.show();
  }catch(e){
    console.error(e);
    mount.innerHTML = `<div class="alert alert-danger">Unable to load Excel preview.</div>`;
    modal.show();
  }
};

window.viewWordDocument = async function(){
  const modalEl = document.getElementById('wordViewModal');
  if(!modalEl){ console.error('wordViewModal introuvable'); return; }
  const modal = new bootstrap.Modal(modalEl);
  const mount = document.getElementById('word-content');
  mount.innerHTML = `<div class="text-center"><div class="spinner-custom"></div><p class="mt-2">Loading Word content...</p></div>`;
  try{
    const r = await fetch('/api/view-word');
    const data = await r.json();
    if(!r.ok || !data.success) throw new Error(data.detail || 'Error viewing Word');
    const sections = data.sections || [];
    let html = `<div class="word-document-viewer">`;
    sections.forEach(s=>{
      html += `
        <div class="word-section">
          <h5 class="section-title">${s.title || ''}</h5>
          <div class="section-content">${(s.content || '').replace(/\n\n/g,'</p><p>').replace(/\n/g,'<br>')}</div>
        </div>`;
    });
    html += `</div>`;
    mount.innerHTML = html || '<div class="text-muted">No content</div>';
    modal.show();
  }catch(e){
    console.error(e);
    mount.innerHTML = `<div class="alert alert-danger">Unable to load Word preview.</div>`;
    modal.show();
  }
};

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
