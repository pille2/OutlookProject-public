// Projektnummer Manager - Taskpane JavaScript
Office.onReady((info) => {
    console.log("Office.onReady called", info);
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook host detected");
        if (document.readyState === 'loading') {
            document.addEventListener("DOMContentLoaded", initializeApp);
        } else {
            initializeApp();
        }
    } else {
        console.error("Wrong host type:", info.host);
    }
});

let currentEmail = null;
let emailData = null;
let debugLogs = [];
let errorLogs = [];
let projectList = [];
let msalInstance = null;
let accessToken = null;

// MSAL Konfiguration
const msalConfig = {
    auth: {
        clientId: "d234e330-bce6-4342-973a-36412caa3dcc",
        authority: "https://login.microsoftonline.com/638e4263-3801-4c47-9a71-d40a53295116",
        redirectUri: window.location.origin
    }
};

// SharePoint Konfiguration
const sharepointConfig = {
    siteUrl: "https://enveon.sharepoint.com",
    sitePath: "/sites/ISO17025enveon",
    listName: "3.0 Liste Projekte",
    apiUrl: "https://enveon.sharepoint.com/sites/ISO17025enveon/_api/web/lists"
};

async function initializeApp() {
    console.log("Projektnummer Manager v1.0.0 initialisiert");
    addDebugLog("App initialisiert");
    
    // MSAL initialisieren
    await initializeMSAL();
    
    // Debug Panel Setup
    setupDebugPanel();
    
    // Event Listener für Action Buttons
    document.getElementById('assignBtn').addEventListener('click', assignProjectNumber);
    document.getElementById('refreshBtn').addEventListener('click', refreshProjectList);
    document.getElementById('authBtn').addEventListener('click', authenticateWithSharePoint);
    
    // Lade E-Mail Informationen
    await loadEmailInfo();
    
    // Lade Projektliste
    await loadProjectList();
    
    // Lade Zuweisungs-Historie
    loadAssignmentHistory();
}

async function initializeMSAL() {
    try {
        if (typeof msal !== 'undefined') {
            msalInstance = new msal.PublicClientApplication(msalConfig);
            addDebugLog("MSAL initialisiert");
        } else {
            addErrorLog("MSAL nicht verfügbar - verwende statische Projektliste");
        }
    } catch (error) {
        console.error("Fehler bei MSAL-Initialisierung:", error);
        addErrorLog("MSAL-Initialisierung fehlgeschlagen: " + error.message);
    }
}

async function getSharePointToken() {
    if (!msalInstance) {
        throw new Error("MSAL nicht initialisiert");
    }
    
    const loginRequest = {
        scopes: ["https://enveon.sharepoint.com/.default"]
    };
    
    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        accessToken = loginResponse.accessToken;
        addDebugLog("SharePoint-Token erhalten");
        return accessToken;
    } catch (err) {
        console.error("Fehler beim Token-Abruf:", err);
        addErrorLog("Token-Abruf fehlgeschlagen: " + err.message);
        throw err;
    }
}

async function loadProjectsFromSharePoint() {
    try {
        if (!accessToken) {
            await getSharePointToken();
        }
        
        // Erst versuchen mit Domain-Filter
        let filteredProjects = [];
        if (emailData && emailData.sender) {
            const domain = emailData.sender.substring(emailData.sender.indexOf('@'));
            const domainFilter = `&$filter=substringof('${domain}', KundenKontaktEmail)`;
            addDebugLog(`Filtere nach Domain: ${domain}`);
            
            const filteredEndpoint = `${sharepointConfig.apiUrl}/getbytitle('${sharepointConfig.listName}')/items?$select=Projektnummer,Projektbeschreibung,KundenKontaktEmail${domainFilter}`;
            
            try {
                const filteredResponse = await fetch(filteredEndpoint, {
                    method: 'GET',
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Accept': 'application/json;odata=nometadata',
                        'Content-Type': 'application/json;odata=nometadata'
                    }
                });
                
                if (filteredResponse.ok) {
                    const filteredData = await filteredResponse.json();
                    filteredProjects = filteredData.value;
                    addDebugLog(`Gefilterte Projekte gefunden: ${filteredProjects.length}`);
                }
            } catch (filterError) {
                console.warn("Domain-Filter fehlgeschlagen:", filterError);
            }
        }
        
        // Falls keine gefilterten Projekte gefunden, alle Projekte laden
        if (filteredProjects.length === 0) {
            addDebugLog("Keine gefilterten Projekte gefunden, lade alle Projekte");
            const allEndpoint = `${sharepointConfig.apiUrl}/getbytitle('${sharepointConfig.listName}')/items?$select=Projektnummer,Projektbeschreibung,KundenKontaktEmail`;
            
            const response = await fetch(allEndpoint, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Accept': 'application/json;odata=nometadata',
                    'Content-Type': 'application/json;odata=nometadata'
                }
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const data = await response.json();
            filteredProjects = data.value;
            addDebugLog(`Alle Projekte geladen: ${filteredProjects.length}`);
        }
        
        // SharePoint-Daten in unser Format konvertieren
        projectList = filteredProjects.map(item => ({
            id: item.Projektnummer,
            name: item.Projektbeschreibung,
            description: item.KundenKontaktEmail || "",
            email: item.KundenKontaktEmail
        }));
        
        return projectList;
        
    } catch (error) {
        console.error("Fehler beim Laden der SharePoint-Projekte:", error);
        addErrorLog("SharePoint-Laden fehlgeschlagen: " + error.message);
        throw error;
    }
}

async function loadEmailInfo() {
    try {
        console.log("Lade E-Mail Informationen...");
        addDebugLog("Starte E-Mail Laden...");
        
        // Warten bis Office.js geladen ist
        if (!Office.context || !Office.context.mailbox) {
            console.error("Office.context.mailbox nicht verfügbar");
            addErrorLog("Office.context.mailbox nicht verfügbar");
            showStatus("Office.js nicht geladen", "error");
            return;
        }
        
        addDebugLog("Office.context.mailbox verfügbar");
        
        // E-Mail Item abrufen
        const item = Office.context.mailbox.item;
        
        if (!item) {
            console.error("Kein E-Mail Item gefunden");
            addErrorLog("Kein E-Mail Item gefunden");
            showStatus("Keine E-Mail ausgewählt", "error");
            return;
        }
        
        console.log("E-Mail Item gefunden:", item);
        addDebugLog("E-Mail Item gefunden: " + JSON.stringify(item, null, 2));
        
        // E-Mail Metadaten sammeln
        emailData = {
            id: item.itemId || "unknown",
            subject: item.subject || "Kein Betreff",
            sender: (item.from && item.from.emailAddress) ? item.from.emailAddress : "Unbekannt",
            senderName: (item.from && item.from.displayName) ? item.from.displayName : "Unbekannt",
            receivedTime: item.dateTimeCreated || new Date(),
            body: await getEmailBody(item)
        };
        
        console.log("E-Mail Daten:", emailData);
        
        currentEmail = emailData;
        
        // E-Mail Info anzeigen
        displayEmailInfo(emailData);
        
        // Gespeicherte Daten für diese E-Mail laden
        loadSavedEmailData(emailData.id);
        
    } catch (error) {
        console.error("Fehler beim Laden der E-Mail:", error);
        showStatus("Fehler beim Laden der E-Mail: " + error.message, "error");
        
        // Fallback: Zeige Test-Daten
        emailData = {
            id: "test-" + Date.now(),
            subject: "Test E-Mail",
            sender: "test@example.com",
            senderName: "Test Sender",
            receivedTime: new Date(),
            body: "Dies ist eine Test-E-Mail für das Projektnummer Add-in.",
        };
        
        displayEmailInfo(emailData);
    }
}

async function getEmailBody(item) {
    return new Promise((resolve, reject) => {
        try {
            if (item.body) {
                item.body.getAsync(Office.CoercionType.Text, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("E-Mail Body geladen:", result.value ? result.value.substring(0, 100) + "..." : "Leer");
                        resolve(result.value || "");
                    } else {
                        console.error("Fehler beim Laden des E-Mail Body:", result.error);
                        resolve("");
                    }
                });
            } else {
                console.log("Kein E-Mail Body verfügbar");
                resolve("");
            }
        } catch (error) {
            console.error("Fehler in getEmailBody:", error);
            resolve("");
        }
    });
}

function displayEmailInfo(data) {
    const emailInfoDiv = document.getElementById('emailInfo');
    
    emailInfoDiv.innerHTML = `
        <div><strong>Von:</strong> ${data.senderName} (${data.sender})</div>
        <div><strong>Betreff:</strong> ${data.subject}</div>
        <div><strong>Empfangen:</strong> ${new Date(data.receivedTime).toLocaleString('de-DE')}</div>
    `;
}

async function loadProjectList() {
    try {
        console.log("Lade Projektliste...");
        addDebugLog("Starte Projektliste laden...");
        
        // Versuche zuerst SharePoint zu laden
        if (msalInstance) {
            try {
                await loadProjectsFromSharePoint();
                addDebugLog(`SharePoint-Projektliste geladen: ${projectList.length} Projekte`);
            } catch (sharepointError) {
                console.warn("SharePoint-Laden fehlgeschlagen, verwende statische Liste:", sharepointError);
                addErrorLog("SharePoint-Laden fehlgeschlagen, verwende statische Liste");
                loadStaticProjectList();
            }
        } else {
            console.log("MSAL nicht verfügbar, verwende statische Projektliste");
            addDebugLog("MSAL nicht verfügbar, verwende statische Projektliste");
            loadStaticProjectList();
        }
        
        updateProjectSelect();
        
        // Auth-Button anzeigen/verstecken
        const authBtn = document.getElementById('authBtn');
        if (msalInstance && !accessToken) {
            authBtn.style.display = 'inline-block';
        } else {
            authBtn.style.display = 'none';
        }
        
    } catch (error) {
        console.error("Fehler beim Laden der Projektliste:", error);
        addErrorLog("Fehler beim Laden der Projektliste: " + error.message);
        showStatus("Fehler beim Laden der Projektliste", "error");
        
        // Fallback auf statische Liste
        loadStaticProjectList();
        updateProjectSelect();
    }
}

function loadStaticProjectList() {
    // Statische Projektliste als Fallback
    projectList = [
        { id: "PROJ-2024-001", name: "Kundenprojekt Alpha", description: "Projekt Alpha Beschreibung" },
        { id: "PROJ-2024-002", name: "Kundenprojekt Beta", description: "Projekt Beta Beschreibung" },
        { id: "PROJ-2024-003", name: "Kundenprojekt Gamma", description: "Projekt Gamma Beschreibung" },
        { id: "PROJ-2024-004", name: "Kundenprojekt Delta", description: "Projekt Delta Beschreibung" },
        { id: "PROJ-2024-005", name: "Kundenprojekt Epsilon", description: "Projekt Epsilon Beschreibung" }
    ];
    addDebugLog(`Statische Projektliste geladen: ${projectList.length} Projekte`);
}

function updateProjectSelect() {
    const select = document.getElementById('projectNumber');
    
    // Aktueller Wert speichern
    const currentValue = select.value;
    
    // Select leeren (außer der ersten Option)
    select.innerHTML = '<option value="">Projektnummer wählen...</option>';
    
    // Projekte hinzufügen
    projectList.forEach(project => {
        const option = document.createElement('option');
        option.value = project.id;
        option.textContent = `${project.id} - ${project.name}`;
        if (project.email) {
            option.title = `Kontakt: ${project.email}`;
        }
        select.appendChild(option);
    });
    
    // Vorherigen Wert wiederherstellen
    if (currentValue) {
        select.value = currentValue;
    }
}

async function refreshProjectList() {
    showStatus("Aktualisiere Projektliste...", "info");
    await loadProjectList();
    showStatus("Projektliste aktualisiert", "success");
}

async function authenticateWithSharePoint() {
    try {
        showStatus("Verbinde mit SharePoint...", "info");
        await getSharePointToken();
        await loadProjectList();
        showStatus("Erfolgreich mit SharePoint verbunden!", "success");
        
        // Auth-Button ausblenden
        document.getElementById('authBtn').style.display = 'none';
        
    } catch (error) {
        console.error("SharePoint-Authentifizierung fehlgeschlagen:", error);
        showStatus("SharePoint-Authentifizierung fehlgeschlagen: " + error.message, "error");
    }
}

function loadSavedEmailData(emailId) {
    const savedData = localStorage.getItem(`email_${emailId}`);
    if (savedData) {
        const data = JSON.parse(savedData);
        document.getElementById('comment').value = data.comment || '';
        
        if (data.projectNumber) {
            document.getElementById('projectNumber').value = data.projectNumber;
        }
    }
}

async function assignProjectNumber() {
    if (!emailData) {
        showStatus("Keine E-Mail-Daten verfügbar", "error");
        return;
    }
    
    const projectNumber = document.getElementById('projectNumber').value;
    if (!projectNumber) {
        showStatus("Bitte eine Projektnummer auswählen", "error");
        return;
    }
    
    try {
        showStatus("Projektnummer wird zugewiesen...", "info");
        
        // 1. Kategorie setzen
        await setEmailCategory(projectNumber);
        
        // 2. Power Automate Webhook triggern
        await triggerPowerAutomate(projectNumber);
        
        // 3. Erfolg protokollieren
        const logEntry = {
            emailId: emailData.id,
            subject: emailData.subject,
            sender: emailData.sender,
            projectNumber: projectNumber,
            comment: document.getElementById('comment').value,
            processedAt: new Date().toISOString(),
            categorySet: true,
            powerAutomateTriggered: true
        };
        
        // LocalStorage speichern
        localStorage.setItem(`email_${emailData.id}_processed`, JSON.stringify(logEntry));
        
        // Historie aktualisieren
        const assignmentHistory = JSON.parse(localStorage.getItem('assignmentHistory') || '[]');
        assignmentHistory.unshift(logEntry);
        
        // Nur die letzten 50 Einträge behalten
        if (assignmentHistory.length > 50) {
            assignmentHistory.splice(50);
        }
        
        localStorage.setItem('assignmentHistory', JSON.stringify(assignmentHistory));
        
        addDebugLog(`Projektnummer zugewiesen: ${projectNumber} für E-Mail ${emailData.subject}`);
        
        showStatus(`Projektnummer ${projectNumber} erfolgreich zugewiesen!`, "success");
        
        // Formular zurücksetzen
        document.getElementById('comment').value = '';
        document.getElementById('projectNumber').value = '';
        
        // Historie aktualisieren
        loadAssignmentHistory();
        
    } catch (error) {
        console.error("Fehler beim Zuweisen der Projektnummer:", error);
        showStatus("Fehler beim Zuweisen der Projektnummer: " + error.message, "error");
    }
}

async function setEmailCategory(projectNumber) {
    return new Promise((resolve, reject) => {
        try {
            const item = Office.context.mailbox.item;
            
            if (item && item.categories) {
                const categoryName = `Projekt-${projectNumber}`;
                
                item.categories.addAsync([categoryName], (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        console.log(`Kategorie ${categoryName} gesetzt`);
                        addDebugLog(`Kategorie ${categoryName} erfolgreich gesetzt`);
                        resolve();
                    } else {
                        console.error("Fehler beim Setzen der Kategorie:", result.error);
                        addErrorLog("Fehler beim Setzen der Kategorie: " + result.error.message);
                        reject(new Error(result.error.message));
                    }
                });
            } else {
                console.warn("Kategorie-Funktion nicht verfügbar");
                addDebugLog("Kategorie-Funktion nicht verfügbar - überspringe");
                resolve(); // Nicht kritisch, daher resolve statt reject
            }
        } catch (error) {
            console.error("Fehler in setEmailCategory:", error);
            addErrorLog("Fehler in setEmailCategory: " + error.message);
            reject(error);
        }
    });
}

async function triggerPowerAutomate(projectNumber) {
    try {
        // TODO: Hier wird später der Power Automate Webhook implementiert
        // Für jetzt simulieren wir den Aufruf
        console.log(`Power Automate Webhook würde getriggert für Projekt: ${projectNumber}`);
        addDebugLog(`Power Automate Webhook simuliert für Projekt: ${projectNumber}`);
        
        // Simulierte Webhook-URL (später durch echte URL ersetzen)
        const webhookUrl = 'https://your-power-automate-webhook-url.com';
        
        const webhookData = {
            email: {
                id: emailData.id,
                subject: emailData.subject,
                sender_email: emailData.sender,
                sender_name: emailData.senderName,
                project_number: projectNumber,
                comment: document.getElementById('comment').value,
                processed_at: new Date().toISOString(),
                source: "Outlook Add-in Projektnummer Manager"
            }
        };
        
        // Webhook senden (aktuell deaktiviert, da URL nicht konfiguriert)
        /*
        const response = await fetch(webhookUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(webhookData)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        */
        
        console.log("Power Automate Webhook erfolgreich gesendet");
        addDebugLog("Power Automate Webhook erfolgreich gesendet");
        
    } catch (error) {
        console.error("Fehler beim Triggern von Power Automate:", error);
        addErrorLog("Fehler beim Triggern von Power Automate: " + error.message);
        throw error;
    }
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.innerHTML = `<div class="status ${type}">${message}</div>`;
    
    // Status nach 5 Sekunden ausblenden
    setTimeout(() => {
        statusDiv.innerHTML = '';
    }, 5000);
}

function loadAssignmentHistory() {
    const assignmentHistoryList = document.getElementById('assignmentHistoryList');
    const assignmentHistory = JSON.parse(localStorage.getItem('assignmentHistory') || '[]');
    
    if (assignmentHistory.length > 0) {
        const historyHtml = assignmentHistory.slice(0, 10).map(entry => {
            const date = new Date(entry.processedAt).toLocaleString('de-DE');
            return `
                <div class="history-item">
                    <div class="history-date">${date}</div>
                    <div class="history-project">${entry.projectNumber}</div>
                    <div class="history-comment">${entry.comment || 'Kein Kommentar'}</div>
                </div>
            `;
        }).join('');
        
        assignmentHistoryList.innerHTML = historyHtml;
    } else {
        assignmentHistoryList.innerHTML = '<div class="history-note">Hinweis: Die Historie ist cookie-basiert und daher nur temporär</div>';
    }
}

// Debug Functions
function setupDebugPanel() {
    const toggleBtn = document.getElementById('toggleDebug');
    const debugContent = document.getElementById('debugContent');
    const debugPanel = document.getElementById('debugPanel');
    
    toggleBtn.addEventListener('click', function() {
        if (debugPanel.style.display === 'none') {
            debugPanel.style.display = 'block';
            debugContent.style.display = 'block';
            toggleBtn.textContent = 'Ausblenden';
            updateDebugInfo();
        } else {
            debugPanel.style.display = 'none';
            debugContent.style.display = 'none';
            toggleBtn.textContent = 'Einblenden';
        }
    });
    
    // Console override für Debug-Logs
    const originalLog = console.log;
    const originalError = console.error;
    
    console.log = function(...args) {
        originalLog.apply(console, args);
        addDebugLog(args.join(' '));
    };
    
    console.error = function(...args) {
        originalError.apply(console, args);
        addErrorLog(args.join(' '));
    };
}

function addDebugLog(message) {
    const timestamp = new Date().toLocaleTimeString();
    debugLogs.push(`[${timestamp}] ${message}`);
    
    // Nur die letzten 20 Logs behalten
    if (debugLogs.length > 20) {
        debugLogs.shift();
    }
    
    updateDebugInfo();
}

function addErrorLog(message) {
    const timestamp = new Date().toLocaleTimeString();
    errorLogs.push(`[${timestamp}] ${message}`);
    
    // Nur die letzten 10 Fehler behalten
    if (errorLogs.length > 10) {
        errorLogs.shift();
    }
    
    updateDebugInfo();
}

function updateDebugInfo() {
    // Office.js Status
    const officeStatus = document.getElementById('officeStatus');
    if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
        officeStatus.innerHTML = '<span style="color: green;">✓ Office.js geladen</span>';
    } else {
        officeStatus.innerHTML = '<span style="color: red;">✗ Office.js nicht verfügbar</span>';
    }
    
    // E-Mail Item Status
    const emailItemStatus = document.getElementById('emailItemStatus');
    if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        emailItemStatus.innerHTML = '<span style="color: green;">✓ E-Mail Item verfügbar</span>';
    } else {
        emailItemStatus.innerHTML = '<span style="color: red;">✗ Kein E-Mail Item</span>';
    }
    
    // E-Mail Daten
    const emailDataDebug = document.getElementById('emailDataDebug');
    if (emailData) {
        emailDataDebug.textContent = JSON.stringify(emailData, null, 2);
    } else {
        emailDataDebug.textContent = 'Keine E-Mail Daten verfügbar';
    }
    
    // SharePoint Status
    const sharepointStatus = document.getElementById('sharepointStatus');
    if (msalInstance && accessToken) {
        sharepointStatus.innerHTML = `<span style="color: green;">✓ SharePoint verbunden - ${projectList.length} Projekte</span>`;
    } else if (msalInstance) {
        sharepointStatus.innerHTML = '<span style="color: orange;">⚠ MSAL verfügbar, aber nicht authentifiziert</span>';
    } else {
        sharepointStatus.innerHTML = '<span style="color: red;">✗ MSAL nicht verfügbar - statische Liste</span>';
    }
    
    // Console Logs
    const consoleLogs = document.getElementById('consoleLogs');
    consoleLogs.innerHTML = debugLogs.join('<br>');
    
    // Error Logs
    const errorLogsDiv = document.getElementById('errorLogs');
    errorLogsDiv.innerHTML = errorLogs.length > 0 ? errorLogs.join('<br>') : 'Keine Fehler';
    
    // Assignment History
    const assignmentHistoryDiv = document.getElementById('assignmentHistory');
    const assignmentHistory = JSON.parse(localStorage.getItem('assignmentHistory') || '[]');
    if (assignmentHistory.length > 0) {
        const historyHtml = assignmentHistory.slice(0, 10).map(entry => 
            `<div style="margin-bottom: 5px; font-size: 10px;">
                <strong>${entry.subject}</strong><br>
                Projekt: ${entry.projectNumber} | Zeit: ${new Date(entry.processedAt).toLocaleString('de-DE')}
            </div>`
        ).join('');
        assignmentHistoryDiv.innerHTML = historyHtml;
    } else {
        assignmentHistoryDiv.innerHTML = 'Keine Zuweisungen';
    }
    
    // LocalStorage Keys
    const localStorageKeysDiv = document.getElementById('localStorageKeys');
    const keys = Object.keys(localStorage).filter(key => key.startsWith('email_') || key === 'assignmentHistory');
    localStorageKeysDiv.innerHTML = keys.length > 0 ? 
        keys.map(key => `<div style="font-size: 10px;">${key}</div>`).join('') : 
        'Keine E-Mail-Daten gespeichert';
}