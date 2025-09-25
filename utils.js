// Variabili globali per il sistema
let iscrizioniData = [];
let movimentiData = [];
let results = [];

// Contatori per numerazione ricevute
let numeroRicevutaPerCF = {};
let importiAnnualiPerCF = {};

// Sistema Progress Bar Avanzato
let progressState = {
    isActive: false,
    currentStep: '',
    currentProgress: 0,
    totalSteps: 0,
    currentStepIndex: 0,
    startTime: null,
    stepStartTime: null
};

// Inizializza progress bar per un processo
function initProgressBar(steps, processName = '') {
    progressState = {
        isActive: true,
        currentStep: processName,
        currentProgress: 0,
        totalSteps: steps.length,
        currentStepIndex: 0,
        startTime: Date.now(),
        stepStartTime: Date.now(),
        steps: steps
    };
    
    const progressBar = document.getElementById('progressBar');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const progressPercent = document.getElementById('progressPercent');
    
    if (progressBar) {
        progressBar.style.display = 'block';
        progressBar.innerHTML = `
            <div class="progress-header">
                <div class="progress-title">${processName}</div>
                <div class="progress-percent" id="progressPercent">0%</div>
            </div>
            <div class="progress-fill" id="progressFill" style="width: 0%;"></div>
            <div class="progress-text" id="progressText">Inizializzazione...</div>
            <div class="progress-stats" id="progressStats">
                <span id="progressStep">Step 0/${steps.length}</span>
                <span id="progressTime">Tempo: 0s</span>
                <span id="progressETA">ETA: --</span>
            </div>
        `;
    }
    
    console.log(`📊 Progress Bar inizializzata: ${processName} (${steps.length} steps)`);
}

// Aggiorna step corrente
function updateProgressStep(stepIndex, stepName, progress = 0) {
    if (!progressState.isActive) return;
    
    progressState.currentStepIndex = stepIndex;
    progressState.currentStep = stepName;
    progressState.stepStartTime = Date.now();
    
    // Calcola progress globale
    const stepProgress = stepIndex / progressState.totalSteps * 100;
    const innerProgress = (progress / 100) * (100 / progressState.totalSteps);
    const totalProgress = Math.min(stepProgress + innerProgress, 100);
    
    updateProgressDisplay(totalProgress, stepName, stepIndex);
    
    console.log(`📈 Step ${stepIndex + 1}/${progressState.totalSteps}: ${stepName} (${progress}%)`);
}

// Aggiorna progress all'interno dello step corrente
function updateProgressInStep(progress, detail = '') {
    if (!progressState.isActive) return;
    
    // Calcola progress totale
    const stepProgress = progressState.currentStepIndex / progressState.totalSteps * 100;
    const innerProgress = (progress / 100) * (100 / progressState.totalSteps);
    const totalProgress = Math.min(stepProgress + innerProgress, 100);
    
    const displayText = detail ? `${progressState.currentStep}: ${detail}` : progressState.currentStep;
    updateProgressDisplay(totalProgress, displayText, progressState.currentStepIndex, progress);
}

// Aggiorna la visualizzazione della progress bar
function updateProgressDisplay(totalProgress, stepText, stepIndex, stepProgress = null) {
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const progressPercent = document.getElementById('progressPercent');
    const progressStep = document.getElementById('progressStep');
    const progressTime = document.getElementById('progressTime');
    const progressETA = document.getElementById('progressETA');
    
    if (progressFill) {
        progressFill.style.width = `${totalProgress}%`;
        
        // Colore progressivo
        if (totalProgress < 30) {
            progressFill.style.background = '#dc3545'; // Rosso
        } else if (totalProgress < 70) {
            progressFill.style.background = '#ffc107'; // Giallo
        } else {
            progressFill.style.background = '#28a745'; // Verde
        }
    }
    
    if (progressPercent) {
        progressPercent.textContent = `${Math.round(totalProgress)}%`;
    }
    
    if (progressText) {
        const displayText = stepProgress !== null ? 
            `${stepText} (${stepProgress}%)` : stepText;
        progressText.textContent = displayText;
    }
    
    // Aggiorna statistiche
    const elapsed = (Date.now() - progressState.startTime) / 1000;
    const eta = totalProgress > 5 ? 
        Math.round((elapsed / totalProgress) * (100 - totalProgress)) : null;
    
    if (progressStep) {
        progressStep.textContent = `Step ${stepIndex + 1}/${progressState.totalSteps}`;
    }
    
    if (progressTime) {
        progressTime.textContent = `Tempo: ${Math.round(elapsed)}s`;
    }
    
    if (progressETA && eta) {
        progressETA.textContent = `ETA: ${eta}s`;
    }
}

// Completa progress bar
function completeProgressBar(message = 'Completato!') {
    if (!progressState.isActive) return;
    
    const totalTime = (Date.now() - progressState.startTime) / 1000;
    
    updateProgressDisplay(100, message, progressState.totalSteps - 1);
    
    // Mostra completamento per 2 secondi
    setTimeout(() => {
        const progressBar = document.getElementById('progressBar');
        if (progressBar) {
            progressBar.style.display = 'none';
        }
        progressState.isActive = false;
    }, 2000);
    
    console.log(`✅ Progress completata: ${message} (${Math.round(totalTime)}s)`);
}

// Sistema di numerazione unificato
function assignProgressiveNumbers() {
    console.log('🔢 Assegnazione numeri progressivi cronologici...');
    
    if (!window.results || window.results.length === 0) {
        console.log('Nessun risultato da numerare');
        return;
    }

    // Ordina i results per data cronologica
    const resultsOrdinati = [...window.results].sort((a, b) => {
        if (a.anno !== b.anno) {
            return a.anno - b.anno;
        }
        return a.mese - b.mese;
    });

    // Resetta la numerazione per CF
    const numeroProgressivoPerCF = {};

    // Assegna numeri progressivi nell'ordine cronologico
    resultsOrdinati.forEach((person) => {
        const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
        
        if (!numeroProgressivoPerCF[cfKey]) {
            numeroProgressivoPerCF[cfKey] = 1;
        } else {
            numeroProgressivoPerCF[cfKey]++;
        }

        person.numeroProgressivo = numeroProgressivoPerCF[cfKey];
        console.log(`${person.nome} ${person.cognome} - ${person.mese}/${person.anno} → Numero: ${person.numeroProgressivo}`);
    });

    window.results = resultsOrdinati;
    console.log('✅ Numerazione progressiva completata');
}

// Gestione numerazione (mantenuta per compatibilità)
function getNextReceiptNumber(cfKey) {
    if (!numeroRicevutaPerCF[cfKey]) {
        numeroRicevutaPerCF[cfKey] = 1;
    } else {
        numeroRicevutaPerCF[cfKey]++;
    }
    
    saveToLocalStorage();
    return numeroRicevutaPerCF[cfKey];
}

function getCurrentReceiptNumber(cfKey) {
    return numeroRicevutaPerCF[cfKey] || 1;
}

function resetAllCounters() {
    const conferma = confirm('Sei sicuro di voler resettare tutti i contatori di numerazione ricevute?');
    if (conferma) {
        numeroRicevutaPerCF = {};
        importiAnnualiPerCF = {};
        localStorage.removeItem('numeroRicevutaPerCF');
        localStorage.removeItem('importiAnnualiPerCF');
        alert('Contatori resettati con successo!');
        console.log('🔄 Tutti i contatori sono stati resettati');
    }
}

function saveToLocalStorage() {
    try {
        localStorage.setItem('numeroRicevutaPerCF', JSON.stringify(numeroRicevutaPerCF));
        localStorage.setItem('importiAnnualiPerCF', JSON.stringify(importiAnnualiPerCF));
    } catch (error) {
        console.error('Errore nel salvare in localStorage:', error);
    }
}

function loadFromLocalStorage() {
    try {
        const savedNumbers = localStorage.getItem('numeroRicevutaPerCF');
        const savedAmounts = localStorage.getItem('importiAnnualiPerCF');
        
        if (savedNumbers) {
            numeroRicevutaPerCF = JSON.parse(savedNumbers);
        }
        if (savedAmounts) {
            importiAnnualiPerCF = JSON.parse(savedAmounts);
        }
        
        console.log('📥 Dati caricati dal localStorage');
    } catch (error) {
        console.error('Errore nel caricare da localStorage:', error);
        numeroRicevutaPerCF = {};
        importiAnnualiPerCF = {};
    }
}

// Controllo importi annuali
function checkAndUpdateAnnualAmount(cfKey, amount) {
    const currentYear = new Date().getFullYear();
    
    if (!importiAnnualiPerCF[cfKey]) {
        importiAnnualiPerCF[cfKey] = {};
    }
    if (!importiAnnualiPerCF[cfKey][currentYear]) {
        importiAnnualiPerCF[cfKey][currentYear] = 0;
    }
    
    importiAnnualiPerCF[cfKey][currentYear] += amount;
    
    if (importiAnnualiPerCF[cfKey][currentYear] > 5000) {
        console.warn(`⚠️ ATTENZIONE: ${cfKey} ha superato 5.000€ nel ${currentYear}!`);
        return {
            warning: true,
            total: importiAnnualiPerCF[cfKey][currentYear],
            year: currentYear
        };
    }
    
    saveToLocalStorage();
    return {
        warning: false,
        total: importiAnnualiPerCF[cfKey][currentYear],
        year: currentYear
    };
}

// Utility functions
function normalizeString(str) {
    if (!str) return '';
    return str.toString()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]/g, '');
}

function calculateSimilarity(str1, str2) {
    if (!str1 || !str2) return 0;
    
    const s1 = normalizeString(str1);
    const s2 = normalizeString(str2);
    
    if (s1 === s2) return 1;
    
    const matrix = [];
    const len1 = s1.length;
    const len2 = s2.length;
    
    for (let i = 0; i <= len2; i++) {
        matrix[i] = [i];
    }
    
    for (let j = 0; j <= len1; j++) {
        matrix[0][j] = j;
    }
    
    for (let i = 1; i <= len2; i++) {
        for (let j = 1; j <= len1; j++) {
            if (s2.charAt(i - 1) === s1.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j] + 1
                );
            }
        }
    }
    
    const maxLen = Math.max(len1, len2);
    return maxLen === 0 ? 1 : (maxLen - matrix[len2][len1]) / maxLen;
}

function calculateRimborsoSpese(importo) {
    if (importo >= 500) return importo * 0.40;
    if (importo >= 450) return 200;
    if (importo >= 350) return 150;
    if (importo >= 250) return 100;
    if (importo >= 150) return 60;
    if (importo >= 80) return 40;
    if (importo >= 51) return 30;
    if (importo >= 1) return 20;
    return 0;
}

function findColumnValue(row, columnNames) {
    if (!row || typeof row !== 'object') {
        console.warn('findColumnValue ricevuto row non valido:', row);
        return null;
    }
    
    for (let colName of columnNames) {
        if (row[colName] !== undefined && row[colName] !== null) {
            return row[colName];
        }
    }
    return null;
}

function getImportoFromMovimento(movimento) {
    const importoColumns = ['IMPORTO', 'Importo', 'ADDEBITI', 'Addebiti', 'ADDEBITO', 'Addebito'];
    
    for (let col of importoColumns) {
        if (movimento[col] !== undefined && movimento[col] !== null) {
            let value = movimento[col];
            if (typeof value === 'string') {
                value = value.replace(/[^\d,-]/g, '').replace(',', '.');
            }
            return Math.abs(parseFloat(value)) || 0;
        }
    }
    return 0;
}

function getDataFromMovimento(movimento) {
    const dateColumns = ['DATA', 'Data', 'DATA VALUTA', 'Data Valuta', 'DATA OPERAZIONE', 'Data Operazione'];
    
    for (let col of dateColumns) {
        if (movimento[col]) {
            let dateValue = movimento[col];
            
            if (dateValue instanceof Date) return dateValue;
            
            if (typeof dateValue === 'number') {
                const excelEpoch = new Date(1900, 0, 1);
                const days = dateValue - 2;
                return new Date(excelEpoch.getTime() + days * 24 * 60 * 60 * 1000);
            }
            
            if (typeof dateValue === 'string') {
                const italianMatch = dateValue.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                if (italianMatch) {
                    return new Date(italianMatch[3], italianMatch[2] - 1, italianMatch[1]);
                }
                
                const parsed = new Date(dateValue);
                if (!isNaN(parsed.getTime())) {
                    return parsed;
                }
            }
        }
    }
    
    return new Date();
}

// Gestione tab
function showTab(tabName) {
    document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    const selectedTab = document.querySelector(`[onclick="showTab('${tabName}')"]`);
    const selectedContent = document.getElementById(`${tabName}Tab`);
    
    if (selectedTab) selectedTab.classList.add('active');
    if (selectedContent) selectedContent.classList.add('active');
}

// Progress bar legacy (mantenuta per compatibilità)
function updateProgressBar(percentage) {
    updateProgressInStep(percentage);
}

// Verifica librerie
function checkLibrariesLoaded() {
    const libraries = {
        'XLSX': typeof XLSX !== 'undefined',
        'jsPDF': typeof window.jspdf !== 'undefined',
        'html2canvas': typeof html2canvas !== 'undefined',
        'JSZip': typeof JSZip !== 'undefined'
    };
    
    const missing = Object.entries(libraries)
        .filter(([name, loaded]) => !loaded)
        .map(([name]) => name);
    
    if (missing.length > 0) {
        console.warn('📚 Librerie mancanti:', missing);
        return false;
    }
    
    console.log('✅ Tutte le librerie sono caricate');
    return true;
}

// Debug utilities
function debugObject(obj, name = 'Object') {
    console.group(`🔍 Debug ${name}`);
    console.log('Type:', typeof obj);
    console.log('Value:', obj);
    if (typeof obj === 'object' && obj !== null) {
        console.log('Keys:', Object.keys(obj));
        console.log('Length:', Array.isArray(obj) ? obj.length : 'N/A');
    }
    console.groupEnd();
}

function validateIscrizioniData() {
    if (!iscrizioniData || !Array.isArray(iscrizioniData)) {
        console.error('❌ iscrizioniData non valida');
        return false;
    }
    
    console.log(`✅ iscrizioniData valida: ${iscrizioniData.length} record`);
    return true;
}

function validateMovimentiData() {
    if (!movimentiData || !Array.isArray(movimentiData)) {
        console.error('❌ movimentiData non valida');
        return false;
    }
    
    console.log(`✅ movimentiData valida: ${movimentiData.length} record`);
    return true;
}

// Inizializzazione
loadFromLocalStorage();

// Esposizione globale delle funzioni e variabili
window.iscrizioniData = iscrizioniData;
window.movimentiData = movimentiData;
window.results = results;
window.numeroRicevutaPerCF = numeroRicevutaPerCF;
window.importiAnnualiPerCF = importiAnnualiPerCF;

// Progress bar system
window.initProgressBar = initProgressBar;
window.updateProgressStep = updateProgressStep;
window.updateProgressInStep = updateProgressInStep;
window.updateProgressDisplay = updateProgressDisplay;
window.completeProgressBar = completeProgressBar;

// Numerazione
window.assignProgressiveNumbers = assignProgressiveNumbers;
window.getNextReceiptNumber = getNextReceiptNumber;
window.getCurrentReceiptNumber = getCurrentReceiptNumber;
window.resetAllCounters = resetAllCounters;

// Utilities
window.normalizeString = normalizeString;
window.calculateSimilarity = calculateSimilarity;
window.calculateRimborsoSpese = calculateRimborsoSpese;
window.findColumnValue = findColumnValue;
window.getImportoFromMovimento = getImportoFromMovimento;
window.getDataFromMovimento = getDataFromMovimento;
window.showTab = showTab;
window.updateProgressBar = updateProgressBar;
window.checkLibrariesLoaded = checkLibrariesLoaded;
window.debugObject = debugObject;
window.validateIscrizioniData = validateIscrizioniData;
window.validateMovimentiData = validateMovimentiData;
window.checkAndUpdateAnnualAmount = checkAndUpdateAnnualAmount;

console.log('✅ utils.js caricato con sistema Progress Bar avanzato');
console.log('📊 Funzioni progress disponibili:', {
    initProgressBar: typeof window.initProgressBar,
    updateProgressStep: typeof window.updateProgressStep,
    updateProgressInStep: typeof window.updateProgressInStep,
    completeProgressBar: typeof window.completeProgressBar
});
