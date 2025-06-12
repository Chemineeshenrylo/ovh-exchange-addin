// Variables globales
let recognition = null;
let isRecording = false;
let currentTextArea = null;

// Initialisation de l'add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('AI Exchange Assistant initialisé');
        loadEmailInfo();
        // Attendre que le DOM soit chargé avant d'initialiser la reconnaissance vocale
        setTimeout(() => {
            initSpeechRecognition();
        }, 1000);
    }
});

// Initialiser la reconnaissance vocale
function initSpeechRecognition() {
    if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
        try {
            const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
            recognition = new SpeechRecognition();
            recognition.continuous = true;
            recognition.interimResults = true;
            recognition.lang = 'fr-FR';
            
            recognition.onresult = function(event) {
                let finalTranscript = '';
                for (let i = event.resultIndex; i < event.results.length; i++) {
                    if (event.results[i].isFinal) {
                        finalTranscript += event.results[i][0].transcript;
                    }
                }
                
                if (finalTranscript && currentTextArea) {
                    const textarea = document.getElementById(currentTextArea);
                    if (textarea) {
                        textarea.value += finalTranscript + ' ';
                    }
                }
            };
            
            recognition.onerror = function(event) {
                console.error('Erreur reconnaissance vocale:', event.error);
                stopVoiceInput();
                
                // Messages d'erreur spécifiques
                switch(event.error) {
                    case 'not-allowed':
                        showStatus('❌ Accès au microphone refusé. Veuillez autoriser l\'accès dans les paramètres du navigateur.', 'error');
                        break;
                    case 'no-speech':
                        showStatus('❌ Aucune parole détectée. Réessayez.', 'error');
                        break;
                    case 'audio-capture':
                        showStatus('❌ Microphone non disponible.', 'error');
                        break;
                    case 'network':
                        showStatus('❌ Erreur réseau pour la reconnaissance vocale.', 'error');
                        break;
                    default:
                        showStatus('❌ Erreur reconnaissance vocale: ' + event.error, 'error');
                }
            };
            
            recognition.onend = function() {
                stopVoiceInput();
            };
            
            console.log('Reconnaissance vocale initialisée avec succès');
        } catch (error) {
            console.error('Erreur initialisation reconnaissance vocale:', error);
            recognition = null;
        }
    } else {
        console.log('Reconnaissance vocale non supportée par ce navigateur');
    }
}

// Démarrer l'enregistrement vocal avec vérification des permissions
function startVoiceInput(textAreaId) {
    if (!recognition) {
        showStatus('❌ Reconnaissance vocale non supportée', 'error');
        return;
    }
    
    // Vérifier que les éléments DOM existent
    const voiceBtn = document.getElementById('voiceBtn');
    const stopBtn = document.getElementById('stopVoiceBtn');
    const textarea = document.getElementById(textAreaId);
    
    if (!voiceBtn || !stopBtn || !textarea) {
        console.error('Éléments DOM manquants pour la reconnaissance vocale');
        showStatus('❌ Interface non prête pour la reconnaissance vocale', 'error');
        return;
    }
    
    // Vérifier les permissions avant de démarrer
    if (navigator.permissions && navigator.permissions.query) {
        navigator.permissions.query({name: 'microphone'}).then(function(result) {
            if (result.state === 'denied') {
                showStatus('❌ Accès au microphone refusé. Veuillez l\'autoriser dans les paramètres du navigateur.', 'error');
                return;
            }
            startRecording(textAreaId, voiceBtn, stopBtn);
        }).catch(function(error) {
            console.error('Erreur vérification permissions:', error);
            // Fallback : essayer de démarrer quand même
            startRecording(textAreaId, voiceBtn, stopBtn);
        });
    } else {
        // Navigateur ne supporte pas l'API permissions
        startRecording(textAreaId, voiceBtn, stopBtn);
    }
}

// Fonction helper pour démarrer l'enregistrement
function startRecording(textAreaId, voiceBtn, stopBtn) {
    currentTextArea = textAreaId;
    isRecording = true;
    
    voiceBtn.style.display = 'none';
    stopBtn.style.display = 'block';
    voiceBtn.classList.add('recording');
    
    try {
        recognition.start();
        showStatus('🎤 Écoute en cours...', 'info');
    } catch (error) {
        console.error('Erreur démarrage reconnaissance:', error);
        showStatus('❌ Impossible de démarrer la reconnaissance vocale', 'error');
        stopVoiceInput();
    }
}

// Arrêter l'enregistrement vocal
function stopVoiceInput() {
    if (recognition && isRecording) {
        try {
            recognition.stop();
        } catch (error) {
            console.error('Erreur arrêt reconnaissance:', error);
        }
    }
    
    isRecording = false;
    currentTextArea = null;
    
    const voiceBtn = document.getElementById('voiceBtn');
    const stopBtn = document.getElementById('stopVoiceBtn');
    
    if (voiceBtn && stopBtn) {
        voiceBtn.style.display = 'block';
        stopBtn.style.display = 'none';
        voiceBtn.classList.remove('recording');
    }
    
    showStatus('✅ Enregistrement terminé', 'success');
}

// Charger les informations de l'email courant
function loadEmailInfo() {
    const item = Office.context.mailbox.item;
    
    if (item.subject) {
        document.getElementById('email-subject').textContent = item.subject;
    }
    
    if (item.from) {
        document.getElementById('email-sender').textContent = 
            item.from.displayName + ' <' + item.from.emailAddress + '>';
    }
    
    if (item.dateTimeCreated) {
        document.getElementById('email-date').textContent = 
            item.dateTimeCreated.toLocaleString('fr-FR');
    }
}

// Analyser l'email avec l'IA
function analyzeEmailWithAI() {
    const settings = getSettings();
    
    if (!settings.webhookUrl) {
        showStatus('❌ Veuillez configurer l\'URL webhook dans les paramètres', 'error');
        openSettings();
        return;
    }
    
    // Changer le bouton en mode chargement
    const analyzeBtn = document.getElementById('analyzeBtn');
    const analyzeIcon = document.getElementById('analyzeIcon');
    const analyzeText = document.getElementById('analyzeText');
    
    analyzeBtn.disabled = true;
    analyzeIcon.innerHTML = '<div class="loading"></div>';
    analyzeText.textContent = 'Analyse en cours...';
    
    showStatus('🔄 Envoi vers l\'IA...', 'info');
    
    // Récupérer les directives depuis l'interface
    const aiDirectives = document.getElementById('aiDirectives').value || 'Analyse cet email.';
    
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailData = {
                subject: Office.context.mailbox.item.subject || '',
                sender: Office.context.mailbox.item.from ? Office.context.mailbox.item.from.displayName + ' <' + Office.context.mailbox.item.from.emailAddress + '>' : '',
                recipients: Office.context.mailbox.item.to ? Office.context.mailbox.item.to.map(r => r.displayName + ' <' + r.emailAddress + '>').join(', ') : '',
                body: result.value || '',
                attachmentsCount: Office.context.mailbox.item.attachments ? Office.context.mailbox.item.attachments.length : 0,
                itemId: Office.context.mailbox.item.itemId || '',
                conversationId: Office.context.mailbox.item.conversationId || '',
                dateTimeCreated: Office.context.mailbox.item.dateTimeCreated ? Office.context.mailbox.item.dateTimeCreated.toISOString() : '',
                timestamp: new Date().toISOString(),
                action: 'analyze_email',
                aiDirectives: aiDirectives
            };
            
            // Envoyer vers le webhook
            sendToAIWebhook(emailData, settings);
            
        } else {
            resetAnalyzeButton();
            showStatus('❌ Erreur lors de la récupération du contenu', 'error');
        }
    });
}

// Envoyer vers le webhook IA
function sendToAIWebhook(emailData, settings) {
    const headers = {
        'Content-Type': 'application/json'
    };
    
    if (settings.securityToken) {
        headers['Authorization'] = `Bearer ${settings.securityToken}`;
    }
    
    fetch(settings.webhookUrl, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(emailData)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        resetAnalyzeButton();
        displayAIResponse(data);
        showStatus('✅ Réponse IA reçue!', 'success');
    })
    .catch(error => {
        resetAnalyzeButton();
        console.error('Erreur webhook IA:', error);
        showStatus(`❌ Erreur: ${error.message}`, 'error');
    });
}

// Afficher la réponse de l'IA (uniquement la réponse, sans infos supplémentaires)
function displayAIResponse(aiData) {
    const responseSection = document.getElementById('aiResponse');
    const responseContent = document.getElementById('responseContent');
    
    // Masquer les informations d'email originales
    const emailInfo = document.getElementById('email-info');
    if (emailInfo) {
        emailInfo.style.display = 'none';
    }
    
    // Extraire uniquement la réponse de l'IA
    let response = '';
    
    if (aiData.success && aiData.aiResponse) {
        response = aiData.aiResponse;
    } else {
        // Fallback pour d'autres formats
        response = aiData.aiResponse || aiData.response || aiData.message || JSON.stringify(aiData);
    }
    
    responseContent.textContent = response;
    responseSection.style.display = 'block';
    
    // Stocker la réponse pour la composition
    localStorage.setItem('lastAIResponse', response);
    localStorage.setItem('lastAIData', JSON.stringify(aiData));
}

// Composer la réponse avec copie dans le presse-papier (réponse IA uniquement)
function composeReply() {
    const aiResponse = localStorage.getItem('lastAIResponse');
    
    if (!aiResponse) {
        showStatus('❌ Aucune réponse IA disponible', 'error');
        return;
    }
    
    // Copier directement la réponse IA sans template
    copyToClipboardOffice(aiResponse);
}

// Fonction de copie spécialement conçue pour les add-ins Office
function copyToClipboardOffice(text) {
    try {
        // Créer un élément textarea temporaire
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        textArea.style.top = '-999999px';
        textArea.style.opacity = '0';
        textArea.setAttribute('readonly', '');
        document.body.appendChild(textArea);
        
        // Sélectionner et copier
        textArea.focus();
        textArea.select();
        textArea.setSelectionRange(0, 99999); // Pour les appareils mobiles
        
        const successful = document.execCommand('copy');
        document.body.removeChild(textArea);
        
        if (successful) {
            showStatus('✅ Réponse IA copiée dans le presse-papier!', 'success');
        } else {
            throw new Error('document.execCommand failed');
        }
    } catch (err) {
        console.error('Erreur copie presse-papier:', err);
        showStatus('❌ Impossible de copier automatiquement', 'error');
        // Afficher le texte pour copie manuelle
        showTextForManualCopy(text);
    }
}

// Afficher le texte pour copie manuelle en cas d'échec
function showTextForManualCopy(text) {
    // Créer une zone de texte visible pour copie manuelle
    const responseSection = document.getElementById('aiResponse');
    const existingCopyArea = document.getElementById('manualCopyArea');
    
    if (existingCopyArea) {
        existingCopyArea.remove();
    }
    
    const copyArea = document.createElement('div');
    copyArea.id = 'manualCopyArea';
    copyArea.style.marginTop = '15px';
    copyArea.style.padding = '15px';
    copyArea.style.border = '2px solid #007acc';
    copyArea.style.borderRadius = '8px';
    copyArea.style.backgroundColor = '#f8f9fa';
    
    copyArea.innerHTML = `
        <h4 style="margin: 0 0 10px 0; color: #007acc;">📋 Copie manuelle de la réponse IA :</h4>
        <textarea readonly 
                  style="width: 100%; height: 150px; margin: 10px 0; padding: 10px; 
                         border: 1px solid #ccc; border-radius: 4px; font-family: inherit;
                         resize: vertical;">${text}</textarea>
        <button onclick="selectAllText(this)" 
                style="background: #007acc; color: white; border: none; padding: 8px 16px; 
                       border-radius: 4px; cursor: pointer; margin-right: 10px;">Tout sélectionner</button>
        <p style="font-size: 12px; color: #666; margin: 10px 0 0 0;">
            Cliquez sur "Tout sélectionner" puis copiez avec Ctrl+C (ou Cmd+C sur Mac)
        </p>
    `;
    
    responseSection.appendChild(copyArea);
    
    // Auto-sélectionner le texte
    const textarea = copyArea.querySelector('textarea');
    setTimeout(() => {
        textarea.focus();
        textarea.select();
    }, 100);
    
    showStatus('📋 Zone de copie manuelle affichée ci-dessous', 'info');
}

// Fonction helper pour sélectionner tout le texte
function selectAllText(button) {
    const textarea = button.parentElement.querySelector('textarea');
    textarea.focus();
    textarea.select();
    textarea.setSelectionRange(0, 99999);
    
    // Changer temporairement le texte du bouton
    const originalText = button.textContent;
    button.textContent = '✅ Sélectionné!';
    button.style.backgroundColor = '#28a745';
    
    setTimeout(() => {
        button.textContent = originalText;
        button.style.backgroundColor = '#007acc';
    }, 2000);
}

// Alternative: Fonction pour récupérer la signature depuis les paramètres Outlook (plus avancé)
function getOutlookSignature(callback) {
    try {
        // Cette approche nécessite des permissions supplémentaires
        Office.context.mailbox.makeEwsRequestAsync(
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
            'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
            'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
            '<soap:Header>' +
            '<RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '</soap:Header>' +
            '<soap:Body>' +
            '<m:GetUserConfiguration>' +
            '<m:UserConfigurationName Name="OWA.UserOptions" />' +
            '</m:GetUserConfiguration>' +
            '</soap:Body>' +
            '</soap:Envelope>',
            function (asyncResult) {
                if (asyncResult.status === "succeeded") {
                    try {
                        // Parser la réponse XML pour extraire la signature
                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(asyncResult.value, "text/xml");
                        // Traitement de la signature...
                        callback(null, "signature_extracted");
                    } catch (error) {
                        callback(error, null);
                    }
                } else {
                    callback(asyncResult.error, null);
                }
            }
        );
    } catch (error) {
        callback(error, null);
    }
}

// Composer la réponse


// Réinitialiser le bouton d'analyse
function resetAnalyzeButton() {
    const analyzeBtn = document.getElementById('analyzeBtn');
    const analyzeIcon = document.getElementById('analyzeIcon');
    const analyzeText = document.getElementById('analyzeText');
    
    analyzeBtn.disabled = false;
    analyzeIcon.textContent = '🤖';
    analyzeText.textContent = 'Analyser avec IA';
}

// Gestion des paramètres
function openSettings() {
    loadSettings();
    document.getElementById('settingsModal').style.display = 'block';
}

function closeSettings() {
    document.getElementById('settingsModal').style.display = 'none';
}

function saveSettings() {
    const webhookUrl = document.getElementById('webhookUrl').value;
    const securityToken = document.getElementById('securityToken').value;
    const aiPrompt = document.getElementById('aiDirectives').value;
    
    if (webhookUrl && !isValidUrl(webhookUrl)) {
        showStatus('❌ URL webhook invalide', 'error');
        return;
    }
    
    const settings = {
        webhookUrl: webhookUrl,
        securityToken: securityToken,
        aiPrompt: aiPrompt,
        lastUpdated: new Date().toISOString()
    };
    
    try {
        localStorage.setItem('ovh-exchange-settings', JSON.stringify(settings));
        showStatus('✅ Paramètres sauvegardés!', 'success');
        closeSettings();
    } catch (error) {
        console.error('Erreur sauvegarde:', error);
        showStatus('❌ Erreur lors de la sauvegarde', 'error');
    }
}

function loadSettings() {
    try {
        const savedSettings = localStorage.getItem('ovh-exchange-settings');
        if (savedSettings) {
            const settings = JSON.parse(savedSettings);
            document.getElementById('webhookUrl').value = settings.webhookUrl || '';
            document.getElementById('securityToken').value = settings.securityToken || '';
            document.getElementById('aiDirectives').value = settings.aiPrompt || '';
        }
    } catch (error) {
        console.error('Erreur chargement paramètres:', error);
    }
}

function getSettings() {
    try {
        const savedSettings = localStorage.getItem('ovh-exchange-settings');
        return savedSettings ? JSON.parse(savedSettings) : {};
    } catch (error) {
        console.error('Erreur récupération paramètres:', error);
        return {};
    }
}

function isValidUrl(string) {
    try {
        new URL(string);
        return true;
    } catch (_) {
        return false;
    }
}

// Afficher un message de statut
function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
    
    setTimeout(() => {
        statusDiv.textContent = '';
        statusDiv.className = 'status';
    }, 5000);
}

// Fermer le modal en cliquant à l'extérieur
window.onclick = function(event) {
    const modal = document.getElementById('settingsModal');
    if (event.target === modal) {
        closeSettings();
    }
}

// Regénérer la réponse IA
// Regénérer la réponse IA
function regenerateResponse() {
    const settings = getSettings();
    
    if (!settings.webhookUrl) {
        showStatus('❌ Veuillez configurer l\'URL webhook dans les paramètres', 'error');
        openSettings();
        return;
    }
    
    // Changer le bouton regénérer en mode chargement
    const regenerateBtn = event.target;
    const originalText = regenerateBtn.textContent;
    regenerateBtn.disabled = true;
    regenerateBtn.textContent = '🔄 Régénération...';
    
    showStatus('🔄 Régénération en cours...', 'info');
    
    // Récupérer les directives actuelles et la réponse précédente
    const aiDirectives = document.getElementById('aiDirectives').value || 'Analyse cet email.';
    const previousResponse = localStorage.getItem('lastAIResponse') || '';
    
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailData = {
                subject: Office.context.mailbox.item.subject || '',
                sender: Office.context.mailbox.item.from ? Office.context.mailbox.item.from.displayName + ' <' + Office.context.mailbox.item.from.emailAddress + '>' : '',
                recipients: Office.context.mailbox.item.to ? Office.context.mailbox.item.to.map(r => r.displayName + ' <' + r.emailAddress + '>').join(', ') : '',
                body: result.value || '',
                attachmentsCount: Office.context.mailbox.item.attachments ? Office.context.mailbox.item.attachments.length : 0,
                itemId: Office.context.mailbox.item.itemId || '',
                conversationId: Office.context.mailbox.item.conversationId || '',
                dateTimeCreated: Office.context.mailbox.item.dateTimeCreated ? Office.context.mailbox.item.dateTimeCreated.toISOString() : '',
                timestamp: new Date().toISOString(),
                action: 'regenerate_response',
                aiDirectives: aiDirectives,
                previousResponse: previousResponse,
                regenerate: true
            };
            
            // Envoyer vers le webhook
            sendToAIWebhookRegenerate(emailData, settings, regenerateBtn, originalText);
            
        } else {
            // Restaurer le bouton en cas d'erreur
            regenerateBtn.disabled = false;
            regenerateBtn.textContent = originalText;
            showStatus('❌ Erreur lors de la récupération du contenu', 'error');
        }
    });
}

// Envoyer vers le webhook IA pour la régénération
function sendToAIWebhookRegenerate(emailData, settings, regenerateBtn, originalText) {
    const headers = {
        'Content-Type': 'application/json'
    };
    
    if (settings.securityToken) {
        headers['Authorization'] = `Bearer ${settings.securityToken}`;
    }
    
    fetch(settings.webhookUrl, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(emailData)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        // Restaurer le bouton
        regenerateBtn.disabled = false;
        regenerateBtn.textContent = originalText;
        
        // Afficher la nouvelle réponse
        displayAIResponse(data);
        showStatus('✅ Réponse régénérée!', 'success');
    })
    .catch(error => {
        // Restaurer le bouton en cas d'erreur
        regenerateBtn.disabled = false;
        regenerateBtn.textContent = originalText;
        
        console.error('Erreur webhook IA:', error);
        showStatus(`❌ Erreur: ${error.message}`, 'error');
    });

    
    // Relancer l'analyse
    analyzeEmailWithAI();
}

