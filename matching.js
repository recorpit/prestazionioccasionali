// Variabili globali per gestione matching
let allMovimentiProcessati = [];
let movimentiMatchati = [];
let movimentiNonMatchati = [];
let movimentiIgnorati = [];
let accreditiDaControllare = [];
let results = [];

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

// Calcolo rimborsi spese - NUOVA SCALA
function calculateRimborsoSpese(importo) {
    if (importo >= 500) return importo * 0.40; // 40% per importi ≥ 500€
    if (importo >= 450) return 200;
    if (importo >= 350) return 150;
    if (importo >= 250) return 100;
    if (importo >= 150) return 60;
    if (importo >= 80) return 40;
    if (importo >= 51) return 30;   // 30€ tra 51-80€
    if (importo >= 1) return 20;    // 20€ sotto 51€
    return 0;
}

// Utility per trovare valori nelle colonne
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

// Estrazione importo addebiti
function getAddebitiFromMovimento(movimento) {
    const addebitiColumns = ['ADDEBITI', 'Addebiti', 'ADDEBITO', 'Addebito', 'G'];
    
    for (let col of addebitiColumns) {
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

// Estrazione importo accrediti
function getAccreditiFromMovimento(movimento) {
    const accreditiColumns = ['ACCREDITI', 'Accrediti', 'ACCREDITO', 'Accredito', 'F'];
    
    for (let col of accreditiColumns) {
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

// Trova persona nelle iscrizioni per controparte
function findPersonaByControparte(controparte, iscrizioni) {
    if (!controparte) return null;
    
    const controparteNorm = normalizeString(controparte);
    
    return iscrizioni.find(isc => {
        const cognome = isc.C;
        const nome = isc.D;
        
        if (!nome || !cognome) return false;
        
        const nomeNorm = normalizeString(nome);
        const cognomeNorm = normalizeString(cognome);
        
        // Verifica se la controparte contiene sia nome che cognome
        return (controparteNorm.includes(nomeNorm) && controparteNorm.includes(cognomeNorm)) ||
               controparteNorm === (nomeNorm + cognomeNorm) ||
               controparteNorm === (cognomeNorm + nomeNorm);
    });
}

// Caricamento file iscrizioni
function loadIscrizioni() {
    const file = document.getElementById('iscrizioniFile').files[0];
    if (!file) {
        alert('Seleziona un file Excel con le iscrizioni');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // Cerca specificamente il foglio "ISCRIZIONI"
            let targetSheet = null;
            const sheetNames = workbook.SheetNames;
            
            for (const name of sheetNames) {
                if (name.toLowerCase().includes('iscrizioni')) {
                    targetSheet = workbook.Sheets[name];
                    break;
                }
            }
            
            if (!targetSheet) {
                alert(`Foglio "ISCRIZIONI" non trovato!\nFogli disponibili: ${sheetNames.join(', ')}\nAssicurati che il foglio si chiami esattamente "ISCRIZIONI"`);
                return;
            }
            
            const rawData = XLSX.utils.sheet_to_json(targetSheet, { header: 1 });
            
            // Converti array in oggetti con le chiavi delle colonne
            iscrizioniData = [];
            for (let i = 0; i < rawData.length; i++) {
                const row = rawData[i];
                if (row.length > 0 && row[1]) { // Verifica che ci sia almeno il CF (colonna B)
                    iscrizioniData.push({
                        B: row[1], // Codice Fiscale
                        C: row[2], // Cognome
                        D: row[3], // Nome
                        E: row[4], // Data di nascita
                        F: row[5], // Cittadinanza
                        G: row[6], // Nome d'arte
                        H: row[7], // Via Indirizzo 1
                        I: row[8], // Città
                        J: row[9], // Provincia
                        K: row[10], // CAP
                        L: row[11], // Paese
                        M: row[12], // P.IVA
                        N: row[13], // Mansione
                        O: row[14], // IBAN
                        P: row[15], // BIC
                        Q: row[16], // E-mail
                        R: row[17], // Cellulare
                        S: row[18], // Comune
                        T: row[19]  // Nascita
                    });
                }
            }
            
            const resultDiv = document.getElementById('iscrizioniResult');
            resultDiv.innerHTML = `
                <div class="success-box">
                    ✅ File caricato con successo!<br>
                    <strong>${iscrizioniData.length}</strong> artisti trovati
                </div>
            `;
            
            // Abilita step 2
            document.getElementById('step2').classList.remove('inactive');
            document.getElementById('movimentiFile').disabled = false;
            document.getElementById('loadMovimentiBtn').disabled = false;
            
        } catch (error) {
            document.getElementById('iscrizioniResult').innerHTML = `
                <div class="error-box">
                    ❌ Errore nel caricamento: ${error.message}
                </div>
            `;
        }
    };
    reader.readAsArrayBuffer(file);
}

// Caricamento file movimenti
function loadMovimenti() {
    const file = document.getElementById('movimentiFile').files[0];
    if (!file) {
        alert('Seleziona un file Excel con i movimenti bancari');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            movimentiData = XLSX.utils.sheet_to_json(firstSheet);
            
            const resultDiv = document.getElementById('movimentiResult');
            resultDiv.innerHTML = `
                <div class="success-box">
                    ✅ File caricato con successo!<br>
                    <strong>${movimentiData.length}</strong> movimenti trovati
                </div>
            `;
            
            // Abilita step 3
            document.getElementById('step3').classList.remove('inactive');
            document.getElementById('matchBtn').disabled = false;
            
        } catch (error) {
            document.getElementById('movimentiResult').innerHTML = `
                <div class="error-box">
                    ❌ Errore nel caricamento: ${error.message}
                </div>
            `;
        }
    };
    reader.readAsArrayBuffer(file);
}

// Matching principale - NUOVA LOGICA SEMPLIFICATA
function performMatching() {
    if (iscrizioniData.length === 0 || movimentiData.length === 0) {
        alert('Carica prima entrambi i file Excel!');
        return;
    }
    
    // Reset array
    movimentiMatchati = [];
    movimentiNonMatchati = [];
    movimentiIgnorati = [];
    accreditiDaControllare = [];
    
    console.log('=== INIZIO MATCHING ===');
    console.log('Movimenti totali:', movimentiData.length);
    console.log('Iscrizioni totali:', iscrizioniData.length);
    
    // STEP 1: Separa ADDEBITI e ACCREDITI
    let addebiti = [];
    let accrediti = [];
    
    movimentiData.forEach((movimento, index) => {
        const importoAddebito = getAddebitiFromMovimento(movimento);
        const importoAccredito = getAccreditiFromMovimento(movimento);
        const controparte = findColumnValue(movimento, ['CONTROPARTE', 'Controparte', 'controparte', 'C']);
        const data = getDataFromMovimento(movimento);
        
        if (importoAddebito > 0) {
            addebiti.push({
                movimento: movimento,
                controparte: controparte,
                importo: importoAddebito,
                data: data,
                index: index,
                tipo: 'ADDEBITO'
            });
        }
        
        if (importoAccredito > 0) {
            accrediti.push({
                movimento: movimento,
                controparte: controparte,
                importo: importoAccredito,
                data: data,
                index: index,
                tipo: 'ACCREDITO'
            });
        }
    });
    
    console.log('Addebiti trovati:', addebiti.length);
    console.log('Accrediti trovati:', accrediti.length);
    
    // STEP 2: Processa ADDEBITI - Solo se presente nelle iscrizioni
    addebiti.forEach(addebito => {
        const persona = findPersonaByControparte(addebito.controparte, iscrizioniData);
        
        if (persona) {
            // MATCHATO - Persona presente nelle iscrizioni
            movimentiMatchati.push({
                movimento: addebito.movimento,
                persona: persona,
                controparte: addebito.controparte,
                importoTotale: addebito.importo,
                data: addebito.data,
                tipo: 'ADDEBITO',
                index: addebito.index
            });
        } else {
            // NON MATCHATO ma addebito - Va in archivio
            movimentiNonMatchati.push({
                movimento: addebito.movimento,
                controparte: addebito.controparte,
                importo: addebito.importo,
                data: addebito.data,
                tipo: 'ADDEBITO',
                index: addebito.index,
                motivo: 'Persona non presente nelle iscrizioni'
            });
        }
    });
    
    // STEP 3: Processa ACCREDITI - Solo se presente nelle iscrizioni (sospetti)
    accrediti.forEach(accredito => {
        const persona = findPersonaByControparte(accredito.controparte, iscrizioniData);
        
        if (persona) {
            // ACCREDITO DA PERSONA NELLE ISCRIZIONI - Sospetto rimborso
            accreditiDaControllare.push({
                movimento: accredito.movimento,
                persona: persona,
                controparte: accredito.controparte,
                importo: accredito.importo,
                data: accredito.data,
                tipo: 'ACCREDITO',
                index: accredito.index,
                motivo: 'Possibile rimborso da artista'
            });
        }
        // Gli accrediti da persone NON nelle iscrizioni vengono completamente ignorati
    });
    
    console.log('=== RISULTATI MATCHING ===');
    console.log('Movimenti matchati (addebiti da iscritti):', movimentiMatchati.length);
    console.log('Movimenti non matchati (addebiti da non iscritti):', movimentiNonMatchati.length);
    console.log('Accrediti da controllare (da iscritti):', accreditiDaControllare.length);
    
    // Mostra i risultati
    showMatchingResults();
}

// Visualizzazione risultati semplificata
function showMatchingResults() {
    const resultDiv = document.getElementById('matchingResult');
    let html = `
        <div class="info-box">
            <strong>📊 RISULTATI MATCHING:</strong><br>
            ✅ Addebiti processabili (da persone nelle iscrizioni): <strong>${movimentiMatchati.length}</strong><br>
            📁 Addebiti non matchati (archivio): <strong>${movimentiNonMatchati.length}</strong><br>
            ⚠️ Accrediti da controllare (possibili rimborsi): <strong>${accreditiDaControllare.length}</strong>
        </div>
    `;
    
    // SEZIONE 1: MOVIMENTI PROCESSABILI (Addebiti matchati)
    if (movimentiMatchati.length > 0) {
        html += `
            <h3 style="color: #155724; background: #d4edda; padding: 10px; border-radius: 5px; margin-top: 20px;">
                ✅ Movimenti Processabili (${movimentiMatchati.length})
            </h3>
            <p>Questi addebiti genereranno ricevute:</p>
            <table>
                <thead>
                    <tr>
                        <th>Controparte</th>
                        <th>Importo Totale<br>(con rimborsi)</th>
                        <th>Importo Netto</th>
                        <th>Rimborso Spese</th>
                        <th>Data</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        let totaleImportiProcessabili = 0;
        let totaleRimborsiProcessabili = 0;
        
        movimentiMatchati.forEach(match => {
            const rimborso = calculateRimborsoSpese(match.importoTotale);
            const netto = match.importoTotale - rimborso;
            
            totaleImportiProcessabili += match.importoTotale;
            totaleRimborsiProcessabili += rimborso;
            
            html += `
                <tr style="background-color: #d4edda;">
                    <td><strong>${match.controparte}</strong></td>
                    <td><strong>€ ${match.importoTotale.toFixed(2)}</strong></td>
                    <td>€ ${netto.toFixed(2)}</td>
                    <td>€ ${rimborso.toFixed(2)}</td>
                    <td>${match.data.toLocaleDateString('it-IT')}</td>
                </tr>
            `;
        });
        
        html += `
                </tbody>
                <tfoot style="background-color: #155724; color: white;">
                    <tr>
                        <td><strong>TOTALI</strong></td>
                        <td><strong>€ ${totaleImportiProcessabili.toFixed(2)}</strong></td>
                        <td><strong>€ ${(totaleImportiProcessabili - totaleRimborsiProcessabili).toFixed(2)}</strong></td>
                        <td><strong>€ ${totaleRimborsiProcessabili.toFixed(2)}</strong></td>
                        <td>-</td>
                    </tr>
                </tfoot>
            </table>
        `;
    }
    
    // SEZIONE 2: DA CONTROLLARE (Accrediti + Addebiti dello stesso CF)
    if (accreditiDaControllare.length > 0) {
        html += `
            <h3 style="color: #856404; background: #fff3cd; padding: 10px; border-radius: 5px; margin-top: 20px;">
                ⚠️ Da Controllare Manualmente (${accreditiDaControllare.length})
            </h3>
            <p><strong>ATTENZIONE:</strong> Accrediti da persone nelle iscrizioni - Possibili rimborsi di bonifici errati!</p>
            <table>
                <thead>
                    <tr>
                        <th>Nome Completo</th>
                        <th>Controparte</th>
                        <th>Importo Ricevuto</th>
                        <th>Data</th>
                        <th>Addebiti Corrispondenti</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        accreditiDaControllare.forEach(accredito => {
            // Cerca addebiti corrispondenti stesso CF
            const addebitiCorrispondenti = movimentiMatchati.filter(match => 
                match.persona.B === accredito.persona.B
            );
            
            let addebitiInfo = 'Nessuno';
            if (addebitiCorrispondenti.length > 0) {
                addebitiInfo = addebitiCorrispondenti.map(add => 
                    `€${add.importoTotale.toFixed(2)} (${add.data.toLocaleDateString('it-IT')})`
                ).join('<br>');
            }
            
            html += `
                <tr style="background-color: #fff3cd;">
                    <td><strong>${accredito.persona.D} ${accredito.persona.C}</strong><br>
                        <small>CF: ${accredito.persona.B}</small></td>
                    <td>${accredito.controparte}</td>
                    <td><strong style="color: #856404;">€ ${accredito.importo.toFixed(2)}</strong></td>
                    <td>${accredito.data.toLocaleDateString('it-IT')}</td>
                    <td><small>${addebitiInfo}</small></td>
                </tr>
            `;
        });
        
        html += `
                </tbody>
            </table>
            <div style="background: #ffeaa7; padding: 10px; border-radius: 5px; margin-top: 10px;">
                <strong>💡 Azione richiesta:</strong> Verifica se gli addebiti corrispondenti sono legittimi o dovrebbero essere rimossi per evitare duplicazioni.
            </div>
        `;
    }
    
    // SEZIONE 3: ARCHIVIO NON MATCHATI
    if (movimentiNonMatchati.length > 0) {
        html += `
            <h3 style="color: #721c24; background: #f8d7da; padding: 10px; border-radius: 5px; margin-top: 20px;">
                📁 Archivio Non Matchati (${movimentiNonMatchati.length})
            </h3>
            <p>Addebiti verso persone NON presenti nelle iscrizioni:</p>
            
            <div style="margin-bottom: 15px;">
                <button onclick="showArchiveControls(true)" style="background: #28a745; margin-right: 10px;">
                    Mostra Archivio Completo
                </button>
                <button onclick="createReceiptFromArchive()" style="background: #17a2b8; margin-right: 10px;">
                    Crea Ricevuta da Archivio
                </button>
                <button onclick="deleteFromArchive()" style="background: #dc3545;">
                    Elimina Selezionati
                </button>
            </div>
            
            <div id="archiveSection" style="display: none;">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 30px;">☑️</th>
                            <th>Controparte</th>
                            <th>Importo</th>
                            <th>Data</th>
                            <th>Azioni</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        movimentiNonMatchati.forEach((movimento, index) => {
            html += `
                <tr id="archive_${index}">
                    <td style="text-align: center;">
                        <input type="checkbox" id="arch_${index}" style="transform: scale(1.2);">
                    </td>
                    <td>${movimento.controparte}</td>
                    <td><strong>€ ${movimento.importo.toFixed(2)}</strong></td>
                    <td>${movimento.data.toLocaleDateString('it-IT')}</td>
                    <td>
                        <button onclick="createSingleReceipt(${index})" style="background: #007bff; font-size: 12px; padding: 5px;">
                            Crea Ricevuta
                        </button>
                    </td>
                </tr>
            `;
        });
        
        html += `
                    </tbody>
                </table>
            </div>
            
            <div style="background: #d1ecf1; padding: 10px; border-radius: 5px; margin-top: 10px;">
                <strong>ℹ️ Informazioni:</strong> Questi movimenti non genereranno ricevute automaticamente poiché le controparti non sono presenti nel file iscrizioni.
            </div>
        `;
    }
    
    // PULSANTE PROCEDI
    if (movimentiMatchati.length > 0) {
        html += `
            <div style="text-align: center; margin: 30px 0; padding: 20px; background: #e8f5e9; border-radius: 10px;">
                <h4 style="color: #155724;">🚀 Pronto per Generare le Ricevute!</h4>
                <p>Trovati <strong>${movimentiMatchati.length}</strong> movimenti processabili.</p>
                <button onclick="proceedToGeneration()" 
                        style="background: #28a745; font-size: 18px; padding: 15px 30px; border-radius: 8px;">
                    Procedi alla Generazione Ricevute
                </button>
            </div>
        `;
    }
    
    resultDiv.innerHTML = html;
}

// Controlli archivio
function showArchiveControls(show) {
    const archiveSection = document.getElementById('archiveSection');
    if (archiveSection) {
        archiveSection.style.display = show ? 'block' : 'none';
    }
}

function createReceiptFromArchive() {
    alert('Funzionalità in sviluppo: Creazione ricevuta da archivio');
}

function deleteFromArchive() {
    const checkboxes = document.querySelectorAll('input[id^="arch_"]:checked');
    if (checkboxes.length === 0) {
        alert('Seleziona almeno un movimento da eliminare');
        return;
    }
    
    if (confirm(`Eliminare ${checkboxes.length} movimenti dall'archivio?`)) {
        checkboxes.forEach(checkbox => {
            const index = checkbox.id.replace('arch_', '');
            const row = document.getElementById(`archive_${index}`);
            if (row) row.remove();
        });
        alert(`${checkboxes.length} movimenti eliminati dall'archivio`);
    }
}

function createSingleReceipt(index) {
    const movimento = movimentiNonMatchati[index];
    if (movimento) {
        alert(`Creazione ricevuta singola per: ${movimento.controparte}\nImporto: €${movimento.importo.toFixed(2)}\n(Funzionalità in sviluppo)`);
    }
}

// Procedi alla generazione - NUOVA LOGICA SEMPLIFICATA
function proceedToGeneration() {
    if (movimentiMatchati.length === 0) {
        alert('Nessun movimento processabile trovato!');
        return;
    }
    
    // Raggruppa per persona E per mese
    const gruppiPerPersonaMese = {};
    
    movimentiMatchati.forEach(match => {
        const cf = match.persona.B;
        const mese = match.data.getMonth() + 1;
        const anno = match.data.getFullYear();
        const chiave = `${cf}_${anno}_${mese}`;
        
        if (!gruppiPerPersonaMese[chiave]) {
            gruppiPerPersonaMese[chiave] = {
                persona: match.persona,
                movimenti: [],
                importoTotale: 0,
                mese: mese,
                anno: anno,
                cf: cf
            };
        }
        
        gruppiPerPersonaMese[chiave].movimenti.push(match);
        gruppiPerPersonaMese[chiave].importoTotale += match.importoTotale;
    });
    
    // Genera results per ricevute
    results = [];
    
    Object.values(gruppiPerPersonaMese).forEach(gruppo => {
        const rimborsoSpese = calculateRimborsoSpese(gruppo.importoTotale);
        const compensoNetto = gruppo.importoTotale - rimborsoSpese;
        const compensoLordo = compensoNetto / 0.8;
        
        // Crea oggetto per ricevuta
        results.push({
            nome: gruppo.persona.D,
            cognome: gruppo.persona.C,
            codiceFiscale: gruppo.persona.B,
            partitaIva: gruppo.persona.M,
            iban: gruppo.persona.O,
            indirizzo: gruppo.persona.H,
            cap: gruppo.persona.K,
            citta: gruppo.persona.I,
            provincia: gruppo.persona.J,
            compenso: compensoLordo,
            rimborsoSpese: rimborsoSpese,
            movimentoBancario: gruppo.importoTotale,
            mese: gruppo.mese,
            anno: gruppo.anno,
            movimenti: gruppo.movimenti.map(m => m.movimento),
            // Dettagli aggiuntivi per la visualizzazione
            dettaglioMovimenti: gruppo.movimenti.map(m => ({
                controparte: m.controparte,
                importo: m.importoTotale,
                data: m.data
            }))
        });
    });
    
    console.log('Results generati:', results.length);
    
    // Abilita generazione ricevute
    document.getElementById('generateBtn').disabled = false;
    
    alert(`✅ Elaborazione completata!\n\nSaranno generate ${results.length} ricevute (raggruppate per mese).\n\nOra puoi cliccare su "Genera Ricevute" per creare i documenti HTML.`);
}

// Esposizione IMMEDIATA funzioni al contesto globale
window.loadIscrizioni = loadIscrizioni;
window.loadMovimenti = loadMovimenti;
window.performMatching = performMatching;
window.showArchiveControls = showArchiveControls;
window.createReceiptFromArchive = createReceiptFromArchive;
window.deleteFromArchive = deleteFromArchive;
window.createSingleReceipt = createSingleReceipt;
window.proceedToGeneration = proceedToGeneration;

// Debug IMMEDIATO - verifica che le funzioni siano esposte
console.log('🔍 matching.js CARICATO - Verifico esposizione funzioni...');
console.log('loadIscrizioni:', typeof window.loadIscrizioni);
console.log('loadMovimenti:', typeof window.loadMovimenti);
console.log('performMatching:', typeof window.performMatching);

if (typeof window.loadIscrizioni !== 'function') {
    console.error('❌ ERRORE: loadIscrizioni non è esposta correttamente!');
} else {
    console.log('✅ loadIscrizioni esposta correttamente');
}

if (typeof window.performMatching !== 'function') {
    console.error('❌ ERRORE: performMatching non è esposta correttamente!');
} else {
    console.log('✅ performMatching esposta correttamente');
}
