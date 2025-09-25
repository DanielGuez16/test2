
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

// Fonction pour prévisualiser le ticket
async function previewTicket() {
    if (!currentTicketFile) {
        alert('Please upload a ticket first');
        return;
    }
    
    const previewBtn = document.getElementById('preview-btn');
    const originalText = previewBtn.innerHTML;
    
    try {
        // Afficher le modal avec loading
        const modal = new bootstrap.Modal(document.getElementById('ticketPreviewModal'));
        modal.show();
        
        previewBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>Extracting...';
        previewBtn.disabled = true;
        
        // Préparer les données
        const formData = new FormData();
        formData.append('ticket_file', currentTicketFile);
        
        const response = await fetch('/api/ticket-preview', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            displayTicketPreview(data.ticket_info, data.extraction_confidence);
        } else {
            displayPreviewError(data.error || 'Failed to extract ticket information');
        }
        
    } catch (error) {
        console.error('Error previewing ticket:', error);
        displayPreviewError(error.message);
    } finally {
        previewBtn.innerHTML = originalText;
        previewBtn.disabled = false;
    }
}

// COPIE EXACTE de main2.js
async function loadLogsStats() {
    try {
        const response = await fetch('/api/logs-stats');
        const result = await response.json();
        
        if (result.success) {
            const stats = result.stats;
            const statsHtml = `
                <div class="row mb-4">
                    <div class="col-md-3">
                        <div class="card text-center">
                            <div class="card-body">
                                <h5 class="card-title text-primary">${stats.total}</h5>
                                <p class="card-text">Total Logs</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card text-center">
                            <div class="card-body">
                                <h5 class="card-title text-success">${stats.users}</h5>
                                <p class="card-text">Active Users</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card text-center">
                            <div class="card-body">
                                <h5 class="card-title text-info">${stats.actions}</h5>
                                <p class="card-text">Action Types</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card text-center">
                            <div class="card-body">
                                <h5 class="card-title text-success">Online</h5>
                                <p class="card-text">System Status</p>
                            </div>
                        </div>
                    </div>
                </div>
                <hr>
            `;
            
            document.getElementById('admin-logs').innerHTML = statsHtml;
            // Charger les logs après les statistiques
            loadActivityLogs();
        }
    } catch (error) {
        console.error('Error loading stats:', error);
        document.getElementById('admin-logs').innerHTML = '<div class="alert alert-danger">Error loading statistics</div>';
    }
}

// Modifier loadAdminLogs pour utiliser la nouvelle fonction
async function loadAdminLogs() {
    loadLogsStats(); // Utilise la fonction qui charge stats + logs
}

// Aperçu ticket façon "ticket de caisse" réaliste
function displayTicketPreview(ticketInfo, confidence) {
    const content = document.getElementById('ticket-preview-content');

    // -------- Helpers sûrs (sans dépendances) --------
    const toFloat = (v) => {
        if (v == null) return null;
        const n = typeof v === 'string' ? v.replace(',', '.').replace(/[^\d.-]/g, '') : v;
        const f = parseFloat(n);
        return isNaN(f) ? null : f;
    };
    const fmtCur = (val, ccy = 'EUR', locale = 'fr-FR') => {
        const num = toFloat(val);
        if (num == null) return 'N/A';
        // Affichage style ticket: cc y après montant (ex: 12,34 EUR)
        return new Intl.NumberFormat(locale, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(num) + ' ' + (ccy || 'EUR');
    };
    const fmtDate = (d) => {
        if (!d) return '—';
        const date = new Date(d);
        return isNaN(date.getTime()) ? d : date.toLocaleDateString('fr-FR') + ' ' + date.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
    };
    const pad = (s, n) => (s || '').toString().slice(0, n).padEnd(n, ' ');
    const line = (ch='—', n=38) => ch.repeat(n);

    // -------- Données & défauts --------
    const vendor = ticketInfo.vendor || 'MERCHANT / COMMERÇANT';
    const location = ticketInfo.location || (ticketInfo.city ? ticketInfo.city : '—');
    const filename = ticketInfo.filename || '—';
    const ccy = ticketInfo.currency || 'EUR';
    const vatRate = toFloat(ticketInfo.vat_rate) ?? 20; // défaut 20% si non fourni
    const payMethod = ticketInfo.payment_method || 'CB / CREDIT CARD';
    const category = ticketInfo.category || 'GENERAL';
    const ticketDate = ticketInfo.date || new Date().toISOString();

    // Items: [{label, qty, unit_price}] attendus. Fallback si absent.
    const items = Array.isArray(ticketInfo.items) && ticketInfo.items.length
        ? ticketInfo.items.map(it => ({
            label: it.label || 'Article',
            qty: toFloat(it.qty) ?? 1,
            unit_price: toFloat(it.unit_price) ?? (toFloat(ticketInfo.amount) ?? 0)
        }))
        : [{
            label: category,
            qty: 1,
            unit_price: toFloat(ticketInfo.amount) ?? 0
        }];

    // Calculs
    const subtotal = items.reduce((s, it) => s + (it.qty * it.unit_price), 0);
    const vat = (vatRate > 0) ? subtotal * (vatRate / 100) : 0;
    // Si un total TTC explicite fourni, on le respecte, sinon on additionne
    const totalTTC = toFloat(ticketInfo.total_amount) ?? (toFloat(ticketInfo.amount) ?? (subtotal + vat));

    // Confiance
    const confidenceColor = confidence >= 0.8 ? 'success' : confidence >= 0.5 ? 'warning' : 'danger';
    const confidenceText  = confidence >= 0.8 ? 'High'    : confidence >= 0.5 ? 'Medium'  : 'Low';

    // Métadonnées transaction
    const txId = (ticketInfo.transaction_id) || ('TX' + Math.random().toString(36).slice(2,10).toUpperCase());
    const ticketNumber = (ticketInfo.ticket_number) || ('#' + Math.random().toString(36).slice(2,6).toUpperCase());
    const merchantId = ticketInfo.merchant_id || 'SIRET/TVA: —';

    // Faux code-barres (Unicode) + QR
    const barcodeData = (txId + ' ' + totalTTC).toUpperCase();
    const fakeBars = '▌▌ ▌▌▌  ▌ ▌▌ ▌▌▌ ▌  ▌▌ ▌▌  ▌▌▌ ▌ ▌  ▌'.slice(0, 38);

    // -------- Gabarit HTML (utilise tes classes .receipt-*) --------
    content.innerHTML = `
        <div class="receipt-container" style="max-width: 360px;">
            <div class="receipt-paper" style="font-family: 'Courier New', monospace; font-size: 13px;">
                
                <div class="receipt-header" style="padding: 1rem 1rem 0.75rem;">
                    <div class="company-logo"><i class="fas fa-receipt"></i></div>
                    <div class="company-info">
                        <h4 style="letter-spacing:1px;">${vendor.toUpperCase()}</h4>
                        <p class="receipt-subtitle" style="opacity:.9">${location}</p>
                    </div>
                    <div class="receipt-number">${ticketNumber}</div>
                </div>

                <div class="confidence-badge-container" style="margin-top:.75rem;">
                    <div class="confidence-badge bg-${confidenceColor}">
                        <i class="fas fa-brain me-1"></i>
                        AI Confidence: ${confidenceText} (${Math.round(confidence*100)}%)
                    </div>
                </div>

                <div class="receipt-divider thin"><span>${line('-', 38)}</span></div>

                <div class="receipt-details">

                    <div class="receipt-row" style="border-bottom:none;">
                        <span class="item-label">DATE</span>
                        <span class="item-value">${fmtDate(ticketDate)}</span>
                    </div>
                    <div class="receipt-row" style="border-bottom:none;">
                        <span class="item-label">PAYMENT</span>
                        <span class="item-value">${payMethod}</span>
                    </div>
                    <div class="receipt-row" style="border-bottom:none;">
                        <span class="item-label">MERCHANT ID</span>
                        <span class="item-value">${merchantId}</span>
                    </div>

                    <div style="margin: .5rem 0 .25rem; text-align:center;">
                        <span style="color:#666;">${line('-', 38)}</span>
                    </div>

                    <!-- En-tête articles -->
                    <div class="receipt-row" style="border-bottom:none; font-weight:bold;">
                        <span class="item-label">ITEM</span>
                        <span class="item-value">QTY  x  PRICE</span>
                    </div>

                    ${items.map(it => {
                        const lineLeft  = pad(it.label.toUpperCase(), 18);
                        const lineRight = `${(it.qty%1?it.qty:Math.round(it.qty)).toString().padStart(3,' ')} x ${fmtCur(it.unit_price, ccy)}`;
                        return `
                            <div class="receipt-row" style="border-bottom:1px dotted #eee; display:block;">
                                <pre style="margin:0; white-space:pre-wrap;">${lineLeft}    ${lineRight}</pre>
                            </div>
                        `;
                    }).join('')}

                    <div style="margin: .25rem 0 .25rem; text-align:center;">
                        <span style="color:#666;">${line('-', 38)}</span>
                    </div>

                    <!-- Totaux -->
                    <div class="receipt-row" style="border-bottom:none;">
                        <span class="item-label">SOUS-TOTAL</span>
                        <span class="item-value">${fmtCur(subtotal, ccy)}</span>
                    </div>
                    <div class="receipt-row" style="border-bottom:none;">
                        <span class="item-label">TVA ${vatRate != null ? '('+vatRate+'%)' : ''}</span>
                        <span class="item-value">${fmtCur(vat, ccy)}</span>
                    </div>

                    <div class="receipt-total" style="margin-top:.75rem;">
                        <div class="total-row">
                            <span class="total-label" style="letter-spacing:1px;">TOTAL TTC</span>
                            <span class="total-amount">${fmtCur(totalTTC, ccy)}</span>
                        </div>
                    </div>

                    <div style="margin: .5rem 0 .25rem; text-align:center;">
                        <span style="color:#666;">${line('-', 38)}</span>
                    </div>

                    <!-- Infos pied -->
                    <div class="receipt-row" style="border-bottom:none;">
                        <span class="item-label">TRANSACTION</span>
                        <span class="item-value">${txId}</span>
                    </div>
                    <div class="receipt-row" style="border-bottom:none;">
                        <span class="item-label">FICHIER</span>
                        <span class="item-value">${filename}</span>
                    </div>
                </div>

                <div class="receipt-footer">
                    <div class="qr-section" style="margin-top:.5rem;">
                        <!-- Faux code-barres -->
                        <pre class="qr-code" style="margin:0; font-size:18px; letter-spacing:1px; line-height:1;">${fakeBars}</pre>
                        <p class="qr-label" style="margin-top:.25rem;">${barcodeData}</p>
                    </div>
                    <div class="footer-info" style="text-align:center; margin-top:.5rem;">
                        <p style="justify-content:center;">
                            <i class="fas fa-info-circle me-1"></i>
                            Conservez ce ticket pour votre comptabilité
                        </p>
                        <p style="justify-content:center;">Merci de votre visite • See you soon</p>
                    </div>
                </div>
            </div>

            <!-- Perforations déjà stylées par ton CSS -->
            <div class="perforations top"></div>
            <div class="perforations bottom"></div>
        </div>
    `;
}


// Afficher les erreurs de preview
function displayPreviewError(errorMessage) {
    const content = document.getElementById('ticket-preview-content');
    content.innerHTML = `
        <div class="alert alert-danger">
            <div class="d-flex align-items-center">
                <i class="fas fa-exclamation-triangle fa-2x me-3"></i>
                <div>
                    <h6 class="alert-heading mb-1">Extraction Failed</h6>
                    <p class="mb-0">${errorMessage}</p>
                </div>
            </div>
        </div>
    `;
}

// Exposer la fonction globalement
window.previewTicket = previewTicket;


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

    const previewBtn = document.getElementById('preview-btn');
    if (previewBtn) {
        previewBtn.disabled = false;
    }
}

function updateUploadUI(file) {
    const uploadCard = document.getElementById('ticket-upload-area');
    if (!uploadCard) return;
    
    // Mettre à jour l'icône
    const uploadIcon = uploadCard.querySelector('.upload-icon-main');
    if (uploadIcon) {
        uploadIcon.className = 'fas fa-check upload-icon-main';
        uploadIcon.style.color = 'var(--success)';
        uploadIcon.style.animation = 'pulse 1.5s infinite';
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

function displayLogs(logs) {
    // Chercher le conteneur des logs ou créer une div après les stats
    let logsContainer = document.getElementById('logs-table-container');
    if (!logsContainer) {
        logsContainer = document.createElement('div');
        logsContainer.id = 'logs-table-container';
        document.getElementById('logs-container').appendChild(logsContainer);
    }
    
    if (logs.length === 0) {
        logsContainer.innerHTML = '<div class="alert alert-info">No activity logs found</div>';
        return;
    }
    
    let html = `
        <h5>Recent Activity Logs</h5>
        <div class="table-responsive">
            <table class="table table-striped table-sm">
                <thead>
                    <tr>
                        <th>Timestamp</th>
                        <th>User</th>
                        <th>Action</th>
                        <th>Details</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    logs.reverse().forEach(log => {
        html += `
            <tr>
                <td><small>${new Date(log.timestamp).toLocaleString()}</small></td>
                <td><span class="badge bg-primary">${log.username.split('@')[0]}</span></td>
                <td><strong>${log.action}</strong></td>
                <td><small>${log.details}</small></td>
            </tr>
        `;
    });
    
    html += '</tbody></table></div>';
    logsContainer.innerHTML = html;
}

async function loadAdminUsers() {
    try {
        const response = await fetch('/api/users');
        const result = await response.json();
        
        if (result.success) {
            displayUsers(result.users);
        }
    } catch (error) {
        console.error('Error loading users:', error);
        document.getElementById('users-container').innerHTML = '<div class="alert alert-danger">Error loading users</div>';
    }
}

async function loadFullHistory() {
    const historyDiv = document.getElementById('full-history');
    if (!historyDiv) return;
    
    try {
        historyDiv.innerHTML = `<div class="text-center"><div class="spinner-custom"></div><p>Loading analysis history...</p></div>`;
        
        const response = await fetch('/api/analysis-history');
        const result = await response.json();
        
        if (result.success && result.history && Array.isArray(result.history)) {
            if (result.history.length === 0) {
                historyDiv.innerHTML = `
                    <div class="text-center text-muted p-4">
                        <i class="fas fa-inbox fa-3x mb-3"></i>
                        <h5>No Analysis History</h5>
                        <p>No ticket analyses have been performed yet.</p>
                    </div>
                `;
                return;
            }
            
            historyDiv.innerHTML = `
                <div class="mb-3 d-flex justify-content-between align-items-center">
                    <h6>Analysis History (${result.history.length} analyses)</h6>
                    <div>
                        <small class="text-muted">Most recent first</small>
                    </div>
                </div>
                
                <div class="history-list">
                    ${result.history.map(item => {
                        // Vérifier la structure des données
                        const timestamp = item.timestamp || 'Unknown time';
                        const user = item.user || 'Unknown user';
                        const filename = item.ticket_filename || 'Unknown file';
                        const question = item.question || 'No specific question';
                        const result_status = item.analysis_result?.result || 'Unknown';
                        const expense_type = item.analysis_result?.expense_type || 'Unknown type';
                        const amount = item.ticket_info?.amount || 'N/A';
                        const currency = item.ticket_info?.currency || '';
                        
                        const statusBadgeClass = result_status === 'PASS' ? 'success' : 
                                               result_status === 'FAIL' ? 'danger' : 'warning';
                        
                        return `
                            <div class="history-item border rounded-3 p-3 mb-3 bg-light">
                                <div class="d-flex justify-content-between align-items-start mb-2">
                                    <div class="flex-grow-1">
                                        <h6 class="mb-1">
                                            <i class="fas fa-file-invoice me-2 text-primary"></i>
                                            ${escapeHtml(filename)}
                                        </h6>
                                        <p class="mb-2 text-dark">
                                            <strong>Question:</strong> ${escapeHtml(question)}
                                        </p>
                                        <p class="mb-1 text-muted">
                                            <strong>Type:</strong> ${escapeHtml(expense_type)}
                                        </p>
                                    </div>
                                    <div class="text-end ms-3">
                                        <span class="badge bg-${statusBadgeClass} fs-6 mb-2">
                                            ${result_status}
                                        </span>
                                        <br>
                                        <div class="badge bg-light text-dark">
                                            ${amount} ${currency}
                                        </div>
                                    </div>
                                </div>
                                
                                <div class="d-flex justify-content-between align-items-center">
                                    <small class="text-muted">
                                        <i class="fas fa-user me-1"></i>${escapeHtml(user)} •
                                        <i class="fas fa-clock me-1"></i>${formatDateTime(timestamp)}
                                    </small>
                                    <small class="text-muted">
                                        Analysis #${result.history.indexOf(item) + 1}
                                    </small>
                                </div>
                            </div>
                        `;
                    }).join('')}
                </div>
            `;
            
        } else {
            throw new Error('Invalid response format or no history data');
        }
        
    } catch (error) {
        console.error('Error loading analysis history:', error);
        historyDiv.innerHTML = `
            <div class="alert alert-danger">
                <i class="fas fa-exclamation-triangle me-2"></i>
                Error loading analysis history: ${error.message}
                <br><small>Check console for details</small>
            </div>
        `;
    }
}

async function loadAdminFeedback() {
    const feedbackDiv = document.getElementById('admin-feedback');
    if (!feedbackDiv) return;
    
    try {
        feedbackDiv.innerHTML = `<div class="text-center"><div class="spinner-custom"></div><p>Loading feedback data...</p></div>`;
        
        const response = await fetch('/api/feedback-stats');
        const result = await response.json();
        
        if (result.success && result.stats) {
            const stats = result.stats;
            
            if (stats.total_feedback === 0) {
                feedbackDiv.innerHTML = `
                    <div class="alert alert-info text-center">
                        <i class="fas fa-star fa-2x mb-3"></i>
                        <h5>No Feedback Yet</h5>
                        <p class="mb-0">No user feedback has been submitted.</p>
                    </div>
                `;
                return;
            }
            
            feedbackDiv.innerHTML = `
                <div class="row mb-4">
                    <div class="col-md-4">
                        <div class="card bg-primary text-white">
                            <div class="card-body text-center">
                                <h3>${stats.total_feedback}</h3>
                                <p class="mb-0">Total Feedbacks</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card bg-success text-white">
                            <div class="card-body text-center">
                                <h3>${stats.average_rating}<small>/5</small></h3>
                                <p class="mb-0">Average Rating</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card bg-info text-white">
                            <div class="card-body text-center">
                                <div class="d-flex justify-content-center align-items-center">
                                    ${[1,2,3,4,5].map(star => 
                                        `<i class="fas fa-star me-1 ${star <= Math.round(stats.average_rating) ? '' : 'opacity-50'}"></i>`
                                    ).join('')}
                                </div>
                                <p class="mb-0 mt-2">User Satisfaction</p>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <h6 class="mb-0">Rating Distribution</h6>
                            </div>
                            <div class="card-body">
                                ${Object.entries(stats.rating_distribution).map(([rating, count]) => {
                                    const percentage = stats.total_feedback > 0 ? Math.round((count / stats.total_feedback) * 100) : 0;
                                    return `
                                        <div class="d-flex justify-content-between align-items-center mb-2">
                                            <span>
                                                ${[...Array(parseInt(rating))].map(() => '⭐').join('')}
                                                ${rating} star${rating > 1 ? 's' : ''}
                                            </span>
                                            <div class="flex-grow-1 mx-3">
                                                <div class="progress" style="height: 8px;">
                                                    <div class="progress-bar" style="width: ${percentage}%"></div>
                                                </div>
                                            </div>
                                            <span class="badge bg-secondary">${count}</span>
                                        </div>
                                    `;
                                }).join('')}
                            </div>
                        </div>
                    </div>
                    
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <h6 class="mb-0">Common Issues</h6>
                            </div>
                            <div class="card-body">
                                ${Object.keys(stats.common_issues).length > 0 ? 
                                    Object.entries(stats.common_issues).slice(0, 5).map(([issue, count]) => `
                                        <div class="d-flex justify-content-between align-items-center mb-2">
                                            <span class="text-truncate flex-grow-1 me-2">${escapeHtml(issue)}</span>
                                            <span class="badge bg-warning">${count}</span>
                                        </div>
                                    `).join('') 
                                    : 
                                    '<div class="text-muted text-center">No issues reported</div>'
                                }
                            </div>
                        </div>
                    </div>
                </div>
            `;
            
        } else {
            throw new Error(result.error || 'Invalid response format');
        }
        
    } catch (error) {
        console.error('Error loading feedback stats:', error);
        feedbackDiv.innerHTML = `
            <div class="alert alert-danger">
                <i class="fas fa-exclamation-triangle me-2"></i>
                Error loading feedback data: ${error.message}
            </div>
        `;
    }
}

async function clearAllLogs() {
    if (!confirm('⚠️ WARNING: This will permanently delete ALL logs, analysis history, and feedback data.\n\nThis action cannot be undone. Are you sure?')) {
        return;
    }
    
    if (!confirm('This is your final warning. All data will be lost. Continue?')) {
        return;
    }
    
    try {
        const response = await fetch('/api/clear-logs', {
            method: 'POST'
        });
        
        const result = await response.json();
        
        if (result.success) {
            alert('✅ All logs have been cleared successfully');
            
            // Recharger les données admin
            setTimeout(() => {
                loadAdminLogs();
                loadAdminFeedback();
            }, 500);
            
        } else {
            throw new Error(result.message || 'Failed to clear logs');
        }
        
    } catch (error) {
        console.error('Error clearing logs:', error);
        alert('❌ Error clearing logs: ' + error.message);
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
    if (text === null || text === undefined) {
        return '';
    }
    
    const div = document.createElement('div');
    div.textContent = String(text);
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
        if (!isoString) return 'Unknown time';
        
        const date = new Date(isoString);
        
        // Vérifier si la date est valide
        if (isNaN(date.getTime())) {
            return isoString.substring(0, 16);
        }
        
        return date.toLocaleString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        });
    } catch (error) {
        console.warn('Error formatting date:', isoString, error);
        return isoString ? String(isoString).substring(0, 16) : 'Invalid date';
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

async function debugAPIs() {
    console.log('=== DEBUG APIs ===');
    
    try {
        console.log('Testing /api/logs...');
        const logsResponse = await fetch('/api/logs?limit=5');
        const logsData = await logsResponse.json();
        console.log('Logs API:', logsData);
        
        console.log('Testing /api/analysis-history...');
        const historyResponse = await fetch('/api/analysis-history');
        const historyData = await historyResponse.json();
        console.log('History API:', historyData);
        
        console.log('Testing /api/feedback-stats...');
        const feedbackResponse = await fetch('/api/feedback-stats');
        const feedbackData = await feedbackResponse.json();
        console.log('Feedback API:', feedbackData);
        
    } catch (error) {
        console.error('Error testing APIs:', error);
    }
}

window.debugAPIs = debugAPIs;
window.clearAllLogs = clearAllLogs;

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
