// Variabili globali
let iscrizioniData = [];
let movimentiData = [];
let results = [];

// Memoria locale per numeri progressivi e importi annuali
let numeroRicevutaPerCF = JSON.parse(localStorage.getItem('numeroRicevutaPerCF')) || {};
let importiAnnualiPerCF = JSON.parse(localStorage.getItem('importiAnnualiPerCF')) || {};

// Funzioni di utilità
function normalizeString(str) {
    if (!str) return '';
    return str.toString()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]/g, '');
}

function findColumnValue(row, possibleColumns) {
    for (let col of possibleColumns) {
        if (row[col] !== undefined && row[col] !== null && row[col] !== '') {
            return row[col];
        }
    }
    for (let key of Object.keys(row)) {
        if (possibleColumns.some(col => key.toUpperCase().includes(col.toUpperCase()))) {
            return row[key];
        }
    }
    return null;
}

function saveToLocalStorage() {
    localStorage.setItem('numeroRicevutaPerCF', JSON.stringify(numeroRicevutaPerCF));
    localStorage.setItem('importiAnnualiPerCF', JSON.stringify(importiAnnualiPerCF));
}

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

function getImportoFromMovimento(movimento) {
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

function updateProgressBar(progress) {
    const progressFill = document.getElementById('progressFill');
    if (progressFill) {
        progressFill.style.width = progress + '%';
    }
}

function showTab(tabName) {
    document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    event.target.classList.add('active');
    document.getElementById(tabName + 'Tab').classList.add('active');
}

// Debug librerie
function checkLibraries() {
    const libraries = {
        'XLSX': typeof XLSX !== 'undefined',
        'jsPDF': typeof window.jspdf !== 'undefined',
        'html2canvas': typeof html2canvas !== 'undefined',
        'JSZip': typeof JSZip !== 'undefined'
    };
    
    console.log('Stato librerie:', libraries);
    return libraries;
}
