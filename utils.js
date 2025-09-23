// Variabili globali
let iscrizioniData = [];
let movimentiData = [];
let results = [];

// Memoria locale per numeri progressivi e importi annuali
let numeroRicevutaPerCF = JSON.parse(localStorage.getItem('numeroRicevutaPerCF')) || {};
let importiAnnualiPerCF = JSON.parse(localStorage.getItem('importiAnnualiPerCF')) || {};

// Salvataggio localStorage
function saveToLocalStorage() {
    localStorage.setItem('numeroRicevutaPerCF', JSON.stringify(numeroRicevutaPerCF));
    localStorage.setItem('importiAnnualiPerCF', JSON.stringify(importiAnnualiPerCF));
}

// Gestione numerazione ricevute
function getNextReceiptNumber(cf) {
    if (!numeroRicevutaPerCF[cf]) {
        numeroRicevutaPerCF[cf] = 0;
    }
    numeroRicevutaPerCF[cf]++;
    saveToLocalStorage();
    return numeroRicevutaPerCF[cf];
}

function getCurrentReceiptNumber(cf) {
    return numeroRicevutaPerCF[cf] || 0;
}

function resetAllCounters() {
    if (confirm('Sei sicuro di voler resettare tutti i numeri progressivi delle ricevute e gli importi annuali?')) {
        numeroRicevutaPerCF = {};
        importiAnnualiPerCF = {};
        localStorage.removeItem('numeroRicevutaPerCF');
        localStorage.removeItem('importiAnnualiPerCF');
        alert('Numeri progressivi e importi annuali resettati!');
    }
}

// Controllo limite annuale €2.500
function checkAndUpdateAnnualAmount(cf, compensoNetto) {
    if (!importiAnnualiPerCF[cf]) {
        importiAnnualiPerCF[cf] = 0;
    }
    
    const totaleAttuale = importiAnnualiPerCF[cf];
    const nuovoTotale = totaleAttuale + compensoNetto;
    
    importiAnnualiPerCF[cf] = nuovoTotale;
    saveToLocalStorage();
    
    if (totaleAttuale <= 2500 && nuovoTotale > 2500) {
        return {
            superaLimite: true,
            totaleAttuale: totaleAttuale,
            nuovoTotale: nuovoTotale
        };
    }
    
    return {
        superaLimite: false,
        totaleAttuale: totaleAttuale,
        nuovoTotale: nuovoTotale
    };
}

// Normalizzazione stringhe
function normalizeString(str) {
    if (!str) return '';
    return str.toString()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]/g, '');
}

// Calcolo similarità (Levenshtein)
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

// Calcolo rimborsi spese - NUOVA SCALA AGGIORNATA
function calculateRimborsoSpese(importo) {
    if (importo >= 500) return importo * 0.40; // 40% per importi ≥ 500€
    if (importo >= 450) return 200;
    if (importo >= 350) return 150;
    if (importo >= 250) return 100;
    if (importo >= 150) return 60;
    if (importo >= 80) return 40;
    if (importo >= 51) return 30;   // NUOVO: 30€ tra 51-80€
    if (importo >= 1) return 20;    // NUOVO: 20€ sotto 51€ (minimo 1€)
    return 0;
}

// Gestione tab
function showTab(tabName) {
    document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    event.target.classList.add('active');
    document.getElementById(tabName + 'Tab').classList.add('active');
}

// Progress bar
function updateProgressBar(progress) {
    const progressFill = document.getElementById('progressFill');
    if (progressFill) {
        progressFill.style.width = progress + '%';
        progressFill.textContent = Math.round(progress) + '%';
    }
}

// Verifica caricamento librerie
function checkLibrariesLoaded() {
    const libs = {
        XLSX: typeof XLSX !== 'undefined',
        jsPDF: typeof window.jspdf !== 'undefined',
        html2canvas: typeof html2canvas !== 'undefined',
        JSZip: typeof JSZip !== 'undefined'
    };
    
    console.log('Stato librerie:', libs);
    
    const missing = Object.keys(libs).filter(lib => !libs[lib]);
    if (missing.length > 0) {
        console.error('Librerie mancanti:', missing);
        return false;
    }
    
    console.log('✓ Tutte le librerie sono caricate correttamente');
    return true;
}

// Utility per trovare valori nelle colonne
function findColumnValue(row, columnNames) {
    for (let colName of columnNames) {
        if (row[colName] !== undefined && row[colName] !== null) {
            return row[colName];
        }
    }
    return null;
}

// Estrazione importo da movimento
function getImportoFromMovimento(movimento) {
    // Priorità agli ADDEBITI invece che agli accrediti
    const importoColumns = ['ADDEBITI', 'Addebiti', 'ADDEBITO', 'Addebito', 'IMPORTO', 'Importo', 'ACCREDITI', 'Accrediti'];
    
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

// Estrazione data da movimento
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
