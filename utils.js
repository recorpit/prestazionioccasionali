// Variabili globali
let iscrizioniData = [];
let movimentiData = [];
let results = [];

// Memoria locale per numeri progressivi e importi annuali
let numeroRicevutaPerCF = JSON.parse(localStorage.getItem('numeroRicevutaPerCF')) || {};
let importiAnnualiPerCF = JSON.parse(localStorage.getItem('importiAnnualiPerCF')) || {};

// Salvataggio localStorage
function saveToLocalStorage() {
    try {
        localStorage.setItem('numeroRicevutaPerCF', JSON.stringify(numeroRicevutaPerCF));
        localStorage.setItem('importiAnnualiPerCF', JSON.stringify(importiAnnualiPerCF));
        console.log('Dati salvati in localStorage');
    } catch (error) {
        console.error('Errore salvataggio localStorage:', error);
    }
}

// SISTEMA DI NUMERAZIONE UNIFICATO - FUNZIONE PRINCIPALE
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

    console.log('Results ordinati per numerazione:', 
        resultsOrdinati.map(r => `${r.mese}/${r.anno} - ${r.nome} ${r.cognome}`));

    // Resetta la numerazione per CF
    const numeroProgressivoPerCF = {};

    // Assegna numeri progressivi nell'ordine cronologico
    resultsOrdinati.forEach((person) => {
        const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
        
        // Incrementa il numero per questo CF
        if (!numeroProgressivoPerCF[cfKey]) {
            numeroProgressivoPerCF[cfKey] = 1;
        } else {
            numeroProgressivoPerCF[cfKey]++;
        }

        // Assegna il numero progressivo alla ricevuta
        person.numeroProgressivo = numeroProgressivoPerCF[cfKey];
        
        console.log(`${person.nome} ${person.cognome} - ${person.mese}/${person.anno} → Numero: ${person.numeroProgressivo}`);
    });

    // Aggiorna i results originali con la numerazione corretta
    window.results = resultsOrdinati;
    
    console.log('✅ Numerazione progressiva completata');
}

// Gestione numerazione ricevute - MANTIENE COMPATIBILITÀ
function getNextReceiptNumber(cf) {
    if (!cf) {
        console.error('getNextReceiptNumber chiamato senza CF');
        return 1;
    }
    
    if (!numeroRicevutaPerCF[cf]) {
        numeroRicevutaPerCF[cf] = 0;
    }
    numeroRicevutaPerCF[cf]++;
    saveToLocalStorage();
    
    console.log(`Nuovo numero ricevuta per ${cf}: ${numeroRicevutaPerCF[cf]}`);
    return numeroRicevutaPerCF[cf];
}

function getCurrentReceiptNumber(cf) {
    if (!cf) {
        console.error('getCurrentReceiptNumber chiamato senza CF');
        return 0;
    }
    return numeroRicevutaPerCF[cf] || 0;
}

function resetAllCounters() {
    const conferma = confirm(
        'Sei sicuro di voler resettare tutti i numeri progressivi delle ricevute e gli importi annuali?\n\n' +
        'Questa operazione cancellerà:\n' +
        '• Tutti i numeri progressivi delle ricevute\n' +
        '• Tutti gli importi annuali registrati\n\n' +
        'Questa operazione NON può essere annullata!'
    );
    
    if (conferma) {
        const secondaConferma = confirm('ULTIMA CONFERMA: Procedere con il reset completo?');
        if (secondaConferma) {
            numeroRicevutaPerCF = {};
            importiAnnualiPerCF = {};
            
            try {
                localStorage.removeItem('numeroRicevutaPerCF');
                localStorage.removeItem('importiAnnualiPerCF');
                console.log('Reset completato - localStorage pulito');
                alert('Reset completato! Tutti i contatori sono stati azzerati.');
            } catch (error) {
                console.error('Errore durante il reset:', error);
                alert('Errore durante il reset. Controlla la console.');
            }
        }
    }
}

// Controllo limite annuale €2.500
function checkAndUpdateAnnualAmount(cf, compensoNetto) {
    if (!cf || typeof compensoNetto !== 'number') {
        console.error('checkAndUpdateAnnualAmount: parametri non validi', { cf, compensoNetto });
        return {
            superaLimite: false,
            totaleAttuale: 0,
            nuovoTotale: 0
        };
    }
    
    if (!importiAnnualiPerCF[cf]) {
        importiAnnualiPerCF[cf] = 0;
    }
    
    const totaleAttuale = importiAnnualiPerCF[cf];
    const nuovoTotale = totaleAttuale + compensoNetto;
    
    importiAnnualiPerCF[cf] = nuovoTotale;
    saveToLocalStorage();
    
    console.log(`Controllo limite per ${cf}: ${totaleAttuale} + ${compensoNetto} = ${nuovoTotale}`);
    
    const superaLimite = totaleAttuale <= 2500 && nuovoTotale > 2500;
    
    return {
        superaLimite: superaLimite,
        totaleAttuale: totaleAttuale,
        nuovoTotale: nuovoTotale
    };
}

// Normalizzazione stringhe per matching
function normalizeString(str) {
    if (!str) return '';
    return str.toString()
        .toLowerCase()
        .trim()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // Rimuove accenti
        .replace(/[^a-z0-9\s]/g, '') // Solo lettere, numeri e spazi
        .replace(/\s+/g, ''); // Rimuove tutti gli spazi
}

// Calcolo similarità Levenshtein
function calculateSimilarity(str1, str2) {
    if (!str1 || !str2) return 0;
    
    const s1 = normalizeString(str1);
    const s2 = normalizeString(str2);
    
    if (s1 === s2) return 1;
    if (s1.length === 0) return s2.length === 0 ? 1 : 0;
    if (s2.length === 0) return 0;
    
    const matrix = [];
    const len1 = s1.length;
    const len2 = s2.length;
    
    // Inizializza matrice
    for (let i = 0; i <= len2; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= len1; j++) {
        matrix[0][j] = j;
    }
    
    // Calcola distanza
    for (let i = 1; i <= len2; i++) {
        for (let j = 1; j <= len1; j++) {
            if (s2.charAt(i - 1) === s1.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1, // sostituzione
                    matrix[i][j - 1] + 1,     // inserimento
                    matrix[i - 1][j] + 1      // cancellazione
                );
            }
        }
    }
    
    const distance = matrix[len2][len1];
    const maxLen = Math.max(len1, len2);
    const similarity = maxLen === 0 ? 1 : (maxLen - distance) / maxLen;
    
    return Math.max(0, similarity); // Assicura che sia >= 0
}

// Calcolo rimborsi spese - SCALA AGGIORNATA
function calculateRimborsoSpese(importo) {
    if (typeof importo !== 'number' || importo < 0) {
        console.error('calculateRimborsoSpese: importo non valido', importo);
        return 0;
    }
    
    if (importo >= 500) {
        const rimborso = importo * 0.40; // 40% per importi ≥ 500€
        console.log(`Rimborso 40% per ${importo}: ${rimborso}`);
        return rimborso;
    }
    if (importo >= 450) return 200;
    if (importo >= 350) return 150;
    if (importo >= 250) return 100;
    if (importo >= 150) return 60;
    if (importo >= 80) return 40;
    if (importo >= 51) return 30;   // 30€ tra 51-80€
    if (importo >= 1) return 20;    // 20€ sotto 51€ (minimo 1€)
    
    return 0;
}

// Gestione tab interfaccia
function showTab(tabName) {
    try {
        // Rimuovi active da tutti i tab
        document.querySelectorAll('.tab').forEach(tab => {
            tab.classList.remove('active');
        });
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        
        // Attiva il tab selezionato
        if (event && event.target) {
            event.target.classList.add('active');
        }
        
        const targetContent = document.getElementById(tabName + 'Tab');
        if (targetContent) {
            targetContent.classList.add('active');
            console.log(`Tab attivato: ${tabName}`);
        } else {
            console.error(`Tab content non trovato: ${tabName}Tab`);
        }
    } catch (error) {
        console.error('Errore nella gestione tab:', error);
    }
}

// Progress bar
function updateProgressBar(progress) {
    try {
        const progressFill = document.getElementById('progressFill');
        if (progressFill) {
            // Assicura che progress sia tra 0 e 100
            const normalizedProgress = Math.min(100, Math.max(0, progress));
            progressFill.style.width = normalizedProgress + '%';
            
            // Mostra percentuale se l'elemento lo supporta
            if (normalizedProgress > 10) {
                progressFill.textContent = Math.round(normalizedProgress) + '%';
            } else {
                progressFill.textContent = '';
            }
            
            console.log(`Progress bar aggiornata: ${normalizedProgress}%`);
        } else {
            console.warn('Elemento progressFill non trovato');
        }
    } catch (error) {
        console.error('Errore aggiornamento progress bar:', error);
    }
}

// Verifica caricamento librerie
function checkLibrariesLoaded() {
    const libraries = {
        XLSX: typeof XLSX !== 'undefined',
        jsPDF: typeof window.jspdf !== 'undefined',
        html2canvas: typeof html2canvas !== 'undefined',
        JSZip: typeof JSZip !== 'undefined'
    };
    
    console.log('Stato librerie esterne:', libraries);
    
    const missing = Object.keys(libraries).filter(lib => !libraries[lib]);
    
    if (missing.length > 0) {
        console.error('Librerie mancanti:', missing);
        const errorMsg = `Librerie non caricate: ${missing.join(', ')}\n\nRicarica la pagina. Se il problema persiste, controlla la connessione internet.`;
        alert(errorMsg);
        return false;
    }
    
    console.log('Tutte le librerie esterne sono state caricate correttamente');
    return true;
}

// Utility per trovare valori nelle colonne Excel
function findColumnValue(row, columnNames) {
    if (!row || typeof row !== 'object') {
        console.warn('findColumnValue: row non valido', row);
        return null;
    }
    
    if (!Array.isArray(columnNames)) {
        console.warn('findColumnValue: columnNames deve essere un array', columnNames);
        return null;
    }
    
    for (let colName of columnNames) {
        if (row[colName] !== undefined && row[colName] !== null && row[colName] !== '') {
            return row[colName];
        }
    }
    
    return null;
}

// Estrazione importo da movimento bancario
function getImportoFromMovimento(movimento) {
    if (!movimento || typeof movimento !== 'object') {
        console.warn('getImportoFromMovimento: movimento non valido', movimento);
        return 0;
    }
    
    // Priorità agli ADDEBITI per le ricevute
    const importoColumns = [
        'ADDEBITI', 'Addebiti', 'ADDEBITO', 'Addebito',
        'G', // Colonna G spesso usata per addebiti
        'IMPORTO', 'Importo',
        'ACCREDITI', 'Accrediti', 'ACCREDITO', 'Accredito',
        'F' // Colonna F spesso usata per accrediti
    ];
    
    for (let col of importoColumns) {
        if (movimento[col] !== undefined && movimento[col] !== null) {
            let value = movimento[col];
            
            // Se è stringa, pulisci e converti
            if (typeof value === 'string') {
                value = value.replace(/[^\d,.-]/g, '').replace(',', '.');
            }
            
            const parsedValue = Math.abs(parseFloat(value)) || 0;
            
            if (parsedValue > 0) {
                console.log(`Importo trovato in colonna ${col}:`, parsedValue);
                return parsedValue;
            }
        }
    }
    
    console.warn('Nessun importo valido trovato nel movimento:', movimento);
    return 0;
}

// Estrazione data da movimento bancario
function getDataFromMovimento(movimento) {
    if (!movimento || typeof movimento !== 'object') {
        console.warn('getDataFromMovimento: movimento non valido', movimento);
        return new Date();
    }
    
    const dateColumns = [
        'DATA', 'Data', 'data',
        'DATA VALUTA', 'Data Valuta', 'data_valuta',
        'DATA OPERAZIONE', 'Data Operazione', 'data_operazione',
        'DATA_CONTABILE', 'Data_Contabile'
    ];
    
    for (let col of dateColumns) {
        if (movimento[col]) {
            let dateValue = movimento[col];
            
            // Se è già un oggetto Date
            if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
                return dateValue;
            }
            
            // Se è un numero (formato Excel seriale)
            if (typeof dateValue === 'number' && dateValue > 0) {
                try {
                    // Formula per convertire data seriale Excel
                    const excelEpoch = new Date(1900, 0, 1);
                    const days = dateValue - 2; // Correzione per leap year bug Excel
                    const converted = new Date(excelEpoch.getTime() + days * 24 * 60 * 60 * 1000);
                    
                    if (!isNaN(converted.getTime()) && converted.getFullYear() > 1900) {
                        console.log(`Data convertita da Excel: ${dateValue} -> ${converted.toLocaleDateString('it-IT')}`);
                        return converted;
                    }
                } catch (error) {
                    console.warn('Errore conversione data Excel:', error);
                }
            }
            
            // Se è una stringa
            if (typeof dateValue === 'string' && dateValue.trim()) {
                try {
                    // Formato italiano DD/MM/YYYY
                    const italianMatch = dateValue.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                    if (italianMatch) {
                        const [, day, month, year] = italianMatch;
                        const italianDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                        if (!isNaN(italianDate.getTime())) {
                            return italianDate;
                        }
                    }
                    
                    // Formato ISO o altri formati standard
                    const standardDate = new Date(dateValue);
                    if (!isNaN(standardDate.getTime()) && standardDate.getFullYear() > 1900) {
                        return standardDate;
                    }
                } catch (error) {
                    console.warn('Errore parsing data stringa:', error);
                }
            }
        }
    }
    
    console.warn('Data non trovata o non valida, usando data corrente:', movimento);
    return new Date();
}

// Utility per debug - mostra struttura oggetto
function debugObject(obj, name = 'Object') {
    if (typeof obj === 'object' && obj !== null) {
        console.log(`${name} keys:`, Object.keys(obj));
        console.log(`${name} sample:`, obj);
    } else {
        console.log(`${name}:`, obj);
    }
}

// Verifica integrità dati iscrizioni
function validateIscrizioniData() {
    if (!Array.isArray(iscrizioniData) || iscrizioniData.length === 0) {
        console.error('Dati iscrizioni non validi o vuoti');
        return false;
    }
    
    let validRecords = 0;
    iscrizioniData.forEach((record, index) => {
        if (record.B && record.C && record.D) { // CF, Cognome, Nome
            validRecords++;
        } else {
            console.warn(`Record iscrizione ${index} incompleto:`, record);
        }
    });
    
    console.log(`Iscrizioni valide: ${validRecords}/${iscrizioniData.length}`);
    return validRecords > 0;
}

// Verifica integrità dati movimenti
function validateMovimentiData() {
    if (!Array.isArray(movimentiData) || movimentiData.length === 0) {
        console.error('Dati movimenti non validi o vuoti');
        return false;
    }
    
    let validMovements = 0;
    movimentiData.forEach((movimento, index) => {
        const controparte = findColumnValue(movimento, ['CONTROPARTE', 'Controparte', 'controparte', 'C']);
        const importo = getImportoFromMovimento(movimento);
        
        if (controparte && importo > 0) {
            validMovements++;
        } else {
            console.warn(`Movimento ${index} incompleto o senza importo:`, movimento);
        }
    });
    
    console.log(`Movimenti validi: ${validMovements}/${movimentiData.length}`);
    return validMovements > 0;
}

// Esposizione funzioni al contesto globale
window.assignProgressiveNumbers = assignProgressiveNumbers; // NUOVO SISTEMA
window.saveToLocalStorage = saveToLocalStorage;
window.getNextReceiptNumber = getNextReceiptNumber;
window.getCurrentReceiptNumber = getCurrentReceiptNumber;
window.resetAllCounters = resetAllCounters;
window.checkAndUpdateAnnualAmount = checkAndUpdateAnnualAmount;
window.normalizeString = normalizeString;
window.calculateSimilarity = calculateSimilarity;
window.calculateRimborsoSpese = calculateRimborsoSpese;
window.showTab = showTab;
window.updateProgressBar = updateProgressBar;
window.checkLibrariesLoaded = checkLibrariesLoaded;
window.findColumnValue = findColumnValue;
window.getImportoFromMovimento = getImportoFromMovimento;
window.getDataFromMovimento = getDataFromMovimento;
window.debugObject = debugObject;
window.validateIscrizioniData = validateIscrizioniData;
window.validateMovimentiData = validateMovimentiData;

// Variabili globali esposte
window.iscrizioniData = iscrizioniData;
window.movimentiData = movimentiData;
window.results = results;
window.numeroRicevutaPerCF = numeroRicevutaPerCF;
window.importiAnnualiPerCF = importiAnnualiPerCF;

// Debug - verifica che le funzioni siano esposte
console.log('utils.js caricato - Funzioni esposte:', {
    assignProgressiveNumbers: typeof window.assignProgressiveNumbers,
    normalizeString: typeof window.normalizeString,
    getCurrentReceiptNumber: typeof window.getCurrentReceiptNumber,
    updateProgressBar: typeof window.updateProgressBar,
    calculateRimborsoSpese: typeof window.calculateRimborsoSpese,
    checkLibrariesLoaded: typeof window.checkLibrariesLoaded
});

console.log('utils.js - Variabili globali inizializzate:', {
    iscrizioniData: window.iscrizioniData.length,
    movimentiData: window.movimentiData.length,
    results: window.results.length,
    numeroRicevutaPerCF: Object.keys(window.numeroRicevutaPerCF).length,
    importiAnnualiPerCF: Object.keys(window.importiAnnualiPerCF).length
});
