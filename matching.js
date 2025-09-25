// Variabili locali per gestione matching
let allMovimentiProcessati = [];
let movimentiMatchati = [];
let movimentiNonMatchati = [];
let movimentiIgnorati = [];
let accreditiDaControllare = [];

// Normalizzazione stringhe
function normalizeString(str) {
    if (!str) return '';
    return str.toString()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]/g, '');
}

// Estrai nome base per raggruppamento archivio
function extractBaseName(controparte) {
    if (!controparte) return '';
    
    // Estrae solo la parte alfabetica iniziale
    const match = controparte.match(/^([A-Za-z]+)/);
    return match ? match[1].toUpperCase() : controparte.toUpperCase();
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

// Trova corrispondenze simili per controllo manuale
function findSimilarMatches(controparte, iscrizioni, threshold = 0.7) {
    const suggestions = [];
    
    iscrizioni.forEach(isc => {
        const nome = isc.D;
        const cognome = isc.C;
        
        if (!nome || !cognome) return;
        
        const nomeCompleto = `${nome} ${cognome}`;
        const similarity = calculateSimilarity(controparte, nomeCompleto);
        
        if (similarity >= threshold && similarity < 1) { // Escludi match perfetti (già processati)
            suggestions.push({
                iscrizione: isc,
                nomeCompleto: nomeCompleto,
                similarity: similarity
            });
        }
    });
    
    // Ordina per similarità decrescente
    return suggestions.sort((a, b) => b.similarity - a.similarity).slice(0, 5); // Max 5 suggerimenti
}

// Dialog per gestione manuale nomi simili
async function showUnmatchedDialog(notMatched) {
    return new Promise((resolve) => {
        let manualMatches = [];
        let currentIndex = 0;
        
        function showNextUnmatched() {
            if (currentIndex >= notMatched.length) {
                resolve(manualMatches);
                return;
            }
            
            const movimento = notMatched[currentIndex].movimento;
            const controparte = findColumnValue(movimento, ['CONTROPARTE', 'Controparte', 'controparte', 'C']);
            const importo = getAddebitiFromMovimento(movimento);
            const data = getDataFromMovimento(movimento);
            
            // Trova suggerimenti simili
            const suggestions = findSimilarMatches(controparte, iscrizioniData);
            
            let message = `CONTROLLO MANUALE NOMI SIMILI\n`;
            message += `Movimento ${currentIndex + 1} di ${notMatched.length}\n\n`;
            message += `Controparte: "${controparte}"\n`;
            message += `Importo: €${importo.toFixed(2)}\n`;
            message += `Data: ${data.toLocaleDateString('it-IT')}\n\n`;
            
            if (suggestions.length > 0) {
                message += `POSSIBILI CORRISPONDENZE SIMILI:\n\n`;
                suggestions.forEach((sugg, idx) => {
                    const similarity = (sugg.similarity * 100).toFixed(1);
                    message += `${idx + 1}. ${sugg.nomeCompleto} (${similarity}% simile)\n`;
                    message += `   CF: ${sugg.iscrizione.B}\n`;
                    message += `   Indirizzo: ${sugg.iscrizione.H || 'N/D'}\n\n`;
                });
                message += `SCEGLI UN'OPZIONE:\n`;
                message += `1-${suggestions.length}: Associa al numero corrispondente\n`;
                message += `S: Salta (metti in archivio)\n`;
                message += `T: Termina controllo\n\n`;
                message += `Scelta:`;
                
                const choice = prompt(message);
                
                if (choice === null || choice.toUpperCase() === 'T') {
                    resolve(manualMatches);
                    return;
                } else if (choice.toUpperCase() === 'S') {
                    currentIndex++;
                    showNextUnmatched();
                    return;
                } else {
                    const choiceNum = parseInt(choice);
                    if (choiceNum >= 1 && choiceNum <= suggestions.length) {
                        manualMatches.push({
                            movimento: movimento,
                            iscrizione: suggestions[choiceNum - 1].iscrizione,
                            importoMovimento: importo,
                            data: data,
                            tipo: 'MATCH_MANUALE',
                            similarity: suggestions[choiceNum - 1].similarity
                        });
                        console.log(`Match manuale: "${controparte}" → "${suggestions[choiceNum - 1].nomeCompleto}" (${(suggestions[choiceNum - 1].similarity * 100).toFixed(1)}%)`);
                    }
                }
            } else {
                message += `Nessuna corrispondenza simile trovata.\n\n`;
                message += `S: Salta (metti in archivio)\n`;
                message += `T: Termina controllo`;
                
                const choice = prompt(message);
                
                if (choice === null || choice.toUpperCase() === 'T') {
                    resolve(manualMatches);
                    return;
                }
            }
            
            currentIndex++;
            showNextUnmatched();
        }
        
        if (notMatched.length > 0) {
            const conferma = confirm(
                `CONTROLLO NOMI SIMILI\n\n` +
                `Trovati ${notMatched.length} movimenti con possibili nomi simili agli iscritti.\n\n` +
                `Vuoi controllare manualmente le corrispondenze?\n` +
                `(Es: "Sar Zen" potrebbe corrispondere a "Sara Zen")\n\n` +
                `Clicca OK per iniziare il controllo manuale\n` +
                `Clicca Annulla per saltare`
            );
            
            if (conferma) {
                showNextUnmatched();
            } else {
                resolve([]);
            }
        } else {
            resolve([]);
        }
    });
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

// Trova persona nelle iscrizioni per controparte CON FILTRO PRESTAZIONE OCCASIONALE
function findPersonaByControparte(controparte, movimento, iscrizioni) {
    if (!controparte) return null;
    
    // FILTRO OBBLIGATORIO: Controlla colonna D per "prestazione occasionale"
    const descrizione = findColumnValue(movimento, ['DESCRIZIONE', 'Descrizione', 'descrizione', 'D']) || '';
    const descrizioneNorm = normalizeString(descrizione);
    
    if (!descrizioneNorm.includes('prestazioneoccasionale') && 
        !descrizioneNorm.includes('prestazione') && 
        !descrizioneNorm.includes('occasionale')) {
        console.log(`Movimento ignorato - Descrizione non contiene "prestazione occasionale":`, descrizione);
        return null;
    }
    
    console.log(`Movimento valido - Descrizione contiene prestazione occasionale:`, descrizione);
    
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
                    ✓ File caricato con successo!<br>
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
                    ✗ Errore nel caricamento: ${error.message}
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
                    ✓ File caricato con successo!<br>
                    <strong>${movimentiData.length}</strong> movimenti trovati
                </div>
            `;
            
            // Abilita step 3
            document.getElementById('step3').classList.remove('inactive');
            document.getElementById('matchBtn').disabled = false;
            
        } catch (error) {
            document.getElementById('movimentiResult').innerHTML = `
                <div class="error-box">
                    ✗ Errore nel caricamento: ${error.message}
                </div>
            `;
        }
    };
    reader.readAsArrayBuffer(file);
}

// Matching principale - CON PRE-FILTRO NOMI SIMILI
async function performMatching() {
    if (iscrizioniData.length === 0 || movimentiData.length === 0) {
        alert('Carica prima entrambi i file Excel!');
        return;
    }
    
    // Reset array
    movimentiMatchati = [];
    movimentiNonMatchati = [];
    movimentiIgnorati = [];
    accreditiDaControllare = [];
    
    console.log('=== INIZIO MATCHING CON PRE-FILTRO NOMI SIMILI ===');
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
    
    // STEP 2: Processa ADDEBITI - Solo se presente nelle iscrizioni E ha "prestazione occasionale"
    let addebitiNonMatchati = [];
    
    addebiti.forEach(addebito => {
        const persona = findPersonaByControparte(addebito.controparte, addebito.movimento, iscrizioniData);
        
        if (persona) {
            // MATCHATO - Persona presente nelle iscrizioni E ha prestazione occasionale
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
            // Controlla se è una persona nelle iscrizioni ma senza "prestazione occasionale"
            const personaSenzaPrestazione = iscrizioniData.find(isc => {
                const cognome = isc.C;
                const nome = isc.D;
                if (!nome || !cognome) return false;
                const controparteNorm = normalizeString(addebito.controparte);
                const nomeNorm = normalizeString(nome);
                const cognomeNorm = normalizeString(cognome);
                return (controparteNorm.includes(nomeNorm) && controparteNorm.includes(cognomeNorm));
            });
            
            if (personaSenzaPrestazione) {
                console.log(`Movimento ignorato - Persona nelle iscrizioni ma senza "prestazione occasionale":`, addebito.controparte);
                // Non aggiungere né ai matchati né ai non matchati - completamente ignorato
            } else {
                // NON MATCHATO - Potrebbe avere nomi simili
                addebitiNonMatchati.push({
                    movimento: addebito.movimento,
                    controparte: addebito.controparte,
                    importo: addebito.importo,
                    data: addebito.data,
                    tipo: 'ADDEBITO',
                    index: addebito.index,
                    motivo: 'Persona non presente nelle iscrizioni'
                });
            }
        }
    });
    
    console.log('Movimenti matchati automaticamente:', movimentiMatchati.length);
    console.log('Addebiti non matchati:', addebitiNonMatchati.length);
    
    // STEP 3: PRE-FILTRA I NON MATCHATI - Solo quelli con similarità agli iscritti
    let addebitiConNomiSimili = [];
    let addebitiSenzaNomiSimili = [];
    
    if (addebitiNonMatchati.length > 0) {
        console.log('Pre-filtraggio addebiti non matchati per nomi simili...');
        
        addebitiNonMatchati.forEach(addebito => {
            // Cerca se ha almeno un nome simile negli iscritti
            const hasSimilarNames = findSimilarMatches(addebito.controparte, iscrizioniData, 0.7);
            
            if (hasSimilarNames.length > 0) {
                // Ha nomi simili - va al controllo manuale
                addebitiConNomiSimili.push(addebito);
                console.log(`Controllo manuale necessario per "${addebito.controparte}" - ${hasSimilarNames.length} nomi simili trovati`);
            } else {
                // Nessun nome simile - va direttamente in archivio
                addebitiSenzaNomiSimili.push(addebito);
                console.log(`Archivio diretto per "${addebito.controparte}" - nessuna similarità trovata`);
            }
        });
        
        console.log(`Pre-filtraggio completato:`);
        console.log(`- Da controllare manualmente: ${addebitiConNomiSimili.length}`);
        console.log(`- Archivio diretto: ${addebitiSenzaNomiSimili.length}`);
    }
    
    // STEP 4: CONTROLLO MANUALE NOMI SIMILI - Solo sui pre-filtrati
    let matchManuali = [];
    if (addebitiConNomiSimili.length > 0) {
        console.log('Avvio controllo manuale solo sui nomi potenzialmente simili...');
        matchManuali = await showUnmatchedDialog(addebitiConNomiSimili);
        console.log('Match manuali ottenuti:', matchManuali.length);
        
        // Aggiungi i match manuali a quelli automatici
        matchManuali.forEach(match => {
            movimentiMatchati.push({
                movimento: match.movimento,
                persona: match.iscrizione,
                controparte: findColumnValue(match.movimento, ['CONTROPARTE', 'Controparte', 'controparte', 'C']),
                importoTotale: match.importoMovimento,
                data: match.data,
                tipo: 'MATCH_MANUALE',
                similarity: match.similarity,
                index: match.movimento.index || 0
            });
        });
        
        // Rimuovi dai "con nomi simili" quelli che sono stati matchati manualmente
        const matchedMovements = new Set(matchManuali.map(m => m.movimento));
        addebitiConNomiSimili = addebitiConNomiSimili.filter(addebito => !matchedMovements.has(addebito.movimento));
    }
    
    // STEP 5: Componi l'archivio finale
    // Gli addebiti rimasti (senza nomi simili + con nomi simili non matchati) vanno in archivio
    movimentiNonMatchati = [...addebitiSenzaNomiSimili, ...addebitiConNomiSimili];
    
    // STEP 6: Processa ACCREDITI - Solo se presente nelle iscrizioni (sospetti)
    accrediti.forEach(accredito => {
        const persona = findPersonaByControparte(accredito.controparte, accredito.movimento, iscrizioniData);
        
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
    
    console.log('=== RISULTATI MATCHING CON PRE-FILTRO NOMI SIMILI ===');
    console.log('Movimenti matchati totali (auto + manuali):', movimentiMatchati.length);
    console.log('- Match automatici:', movimentiMatchati.filter(m => m.tipo === 'ADDEBITO').length);
    console.log('- Match manuali:', movimentiMatchati.filter(m => m.tipo === 'MATCH_MANUALE').length);
    console.log('Movimenti non matchati (archivio):', movimentiNonMatchati.length);
    console.log('- Senza nomi simili:', addebitiSenzaNomiSimili.length);
    console.log('- Con nomi simili non associati:', addebitiConNomiSimili.length);
    console.log('Accrediti da controllare:', accreditiDaControllare.length);
    
    // Mostra i risultati
    showMatchingResults();
}

// Visualizzazione risultati con archivio raggruppato per nome base
function showMatchingResults() {
    const resultDiv = document.getElementById('matchingResult');
    
    const matchAutomatici = movimentiMatchati.filter(m => m.tipo === 'ADDEBITO').length;
    const matchManuali = movimentiMatchati.filter(m => m.tipo === 'MATCH_MANUALE').length;
    
    let html = `
        <div class="info-box">
            <strong>📊 RISULTATI MATCHING CON PRE-FILTRO NOMI SIMILI:</strong><br>
            ✅ Match automatici: <strong>${matchAutomatici}</strong><br>
            🔍 Match manuali (nomi simili): <strong>${matchManuali}</strong><br>
            📝 Totale processabili: <strong>${movimentiMatchati.length}</strong><br>
            📁 Addebiti non matchati (archivio): <strong>${movimentiNonMatchati.length}</strong><br>
            ⚠️ Accrediti da controllare: <strong>${accreditiDaControllare.length}</strong>
        </div>
    `;
    
    // SEZIONE 1: MOVIMENTI PROCESSABILI (Addebiti matchati + manuali)
    if (movimentiMatchati.length > 0) {
        html += `
            <h3 style="color: #155724; background: #d4edda; padding: 10px; border-radius: 5px; margin-top: 20px;">
                ✅ Movimenti Processabili (${movimentiMatchati.length})
            </h3>
            <p>Questi addebiti genereranno ricevute (inclusi i match manuali per nomi simili):</p>
            <table>
                <thead>
                    <tr>
                        <th>Controparte</th>
                        <th>Match</th>
                        <th>Importo Totale</th>
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
            
            const matchType = match.tipo === 'MATCH_MANUALE' ? 
                `🔍 Manuale (${(match.similarity * 100).toFixed(1)}%)` : 
                '✅ Automatico';
                
            const rowColor = match.tipo === 'MATCH_MANUALE' ? '#fff3cd' : '#d4edda';
            
            html += `
                <tr style="background-color: ${rowColor};">
                    <td><strong>${match.controparte}</strong></td>
                    <td>${matchType}</td>
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
                        <td colspan="2"><strong>TOTALI</strong></td>
                        <td><strong>€ ${totaleImportiProcessabili.toFixed(2)}</strong></td>
                        <td><strong>€ ${(totaleImportiProcessabili - totaleRimborsiProcessabili).toFixed(2)}</strong></td>
                        <td><strong>€ ${totaleRimborsiProcessabili.toFixed(2)}</strong></td>
                        <td>-</td>
                    </tr>
                </tfoot>
            </table>
        `;
        
        if (matchManuali > 0) {
            html += `
                <div style="background: #fff3cd; padding: 10px; border-radius: 5px; margin-top: 10px;">
                    <strong>🔍 Controllo nomi simili completato:</strong> ${matchManuali} corrispondenze trovate manualmente.<br>
                    Il controllo è stato fatto solo sui movimenti con possibili nomi simili agli iscritti.
                </div>
            `;
        }
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
        `;
    }
    
    // SEZIONE 3: ARCHIVIO RAGGRUPPATO PER NOME BASE
    if (movimentiNonMatchati.length > 0) {
        // RAGGRUPPA per nome base (parte alfabetica)
        const raggruppamentiArchivio = {};
        let totaleArchivio = 0;
        
        movimentiNonMatchati.forEach(movimento => {
            const nomeBase = extractBaseName(movimento.controparte);
            const key = normalizeString(nomeBase);
            
            if (!raggruppamentiArchivio[key]) {
                raggruppamentiArchivio[key] = {
                    nomeBase: nomeBase, // Nome base (es: MICROSOFT)
                    controparti: new Set(), // Set di controparti complete
                    importoTotale: 0,
                    movimenti: [],
                    numeroMovimenti: 0
                };
            }
            
            raggruppamentiArchivio[key].controparti.add(movimento.controparte);
            raggruppamentiArchivio[key].importoTotale += movimento.importo;
            raggruppamentiArchivio[key].numeroMovimenti++;
            raggruppamentiArchivio[key].movimenti.push({
                controparte: movimento.controparte, // Nome completo
                importo: movimento.importo,
                data: movimento.data,
                index: movimento.index
            });
            
            totaleArchivio += movimento.importo;
        });
        
        // Converti Set in Array e ordina per importo decrescente
        const archivioOrdinato = Object.values(raggruppamentiArchivio)
            .map(gruppo => ({
                ...gruppo,
                controparti: Array.from(gruppo.controparti).sort()
            }))
            .sort((a, b) => b.importoTotale - a.importoTotale);
        
        html += `
            <h3 style="color: #721c24; background: #f8d7da; padding: 10px; border-radius: 5px; margin-top: 20px;">
                📁 Archivio Non Matchati - ${archivioOrdinato.length} nomi base diversi
            </h3>
            <p>Addebiti raggruppati per nome base (es: MICROSOFT*, AMAZON*, ecc.)</p>
            
            <table>
                <thead>
                    <tr>
                        <th>Nome Base</th>
                        <th>Varianti</th>
                        <th>N° Movimenti</th>
                        <th>Importo Totale</th>
                        <th>Azioni</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        archivioOrdinato.forEach((gruppo, index) => {
            const importoMedio = gruppo.importoTotale / gruppo.numeroMovimenti;
            const variantiMostrate = gruppo.controparti.slice(0, 3).join(', ');
            const altreVarianti = gruppo.controparti.length > 3 ? `... (+${gruppo.controparti.length - 3})` : '';
            
            html += `
                <tr>
                    <td>
                        <strong style="font-size: 16px; color: #721c24;">${gruppo.nomeBase}</strong>
                        <br><small style="color: #666;">Prima: ${gruppo.movimenti[0].data.toLocaleDateString('it-IT')} - 
                        Ultima: ${gruppo.movimenti[gruppo.movimenti.length - 1].data.toLocaleDateString('it-IT')}</small>
                    </td>
                    <td style="font-size: 12px; max-width: 200px;">
                        ${variantiMostrate}${altreVarianti}
                    </td>
                    <td style="text-align: center;">
                        <strong>${gruppo.numeroMovimenti}</strong>
                        <br><small>€${importoMedio.toFixed(2)} medio</small>
                    </td>
                    <td style="text-align: right;">
                        <strong style="font-size: 16px; color: #721c24;">€ ${gruppo.importoTotale.toFixed(2)}</strong>
                    </td>
                    <td>
                        <button onclick="showGroupDetails(${index})" 
                                style="background: #007bff; color: white; border: none; font-size: 11px; padding: 4px 8px; margin: 2px; border-radius: 3px; cursor: pointer;">
                            Dettagli
                        </button>
                    </td>
                </tr>
                <tr id="details_${index}" style="display: none; background-color: #f8f9fa;">
                    <td colspan="5">
                        <div style="padding: 15px;">
                            <h5 style="margin: 0 0 10px 0; color: #721c24;">Dettaglio movimenti per ${gruppo.nomeBase}:</h5>
                            <div style="max-height: 200px; overflow-y: auto; font-size: 13px;">
                                ${gruppo.movimenti.map(mov => 
                                    `<div style="padding: 3px 0; border-bottom: 1px solid #dee2e6;">
                                        <strong>${mov.controparte}</strong> - €${mov.importo.toFixed(2)} - ${mov.data.toLocaleDateString('it-IT')}
                                    </div>`
                                ).join('')}
                            </div>
                            <div style="margin-top: 10px; padding: 8px; background: #fff3cd; border-radius: 4px; font-size: 12px;">
                                <strong>Riepilogo ${gruppo.nomeBase}:</strong> ${gruppo.controparti.length} varianti diverse, 
                                ${gruppo.numeroMovimenti} movimenti, €${gruppo.importoTotale.toFixed(2)} totali
                            </div>
                        </div>
                    </td>
                </tr>
            `;
        });
        
        html += `
                </tbody>
                <tfoot style="background-color: #721c24; color: white;">
                    <tr>
                        <td><strong>TOTALI ARCHIVIO</strong></td>
                        <td style="text-align: center;"><strong>${movimentiNonMatchati.length} originali</strong></td>
                        <td style="text-align: center;"><strong>${movimentiNonMatchati.length}</strong></td>
                        <td style="text-align: right;"><strong>€ ${totaleArchivio.toFixed(2)}</strong></td>
                        <td>-</td>
                    </tr>
                </tfoot>
            </table>
        `;
        
        // Conserva i dati raggruppati per le funzioni JavaScript
        window.archivioRaggruppato = archivioOrdinato;
    }
    
    // PULSANTE PROCEDI
    if (movimentiMatchati.length > 0) {
        html += `
            <div style="text-align: center; margin: 30px 0; padding: 20px; background: #e8f5e9; border-radius: 10px;">
                <h4 style="color: #155724;">🚀 Pronto per Generare le Ricevute!</h4>
                <p>Trovati <strong>${movimentiMatchati.length}</strong> movimenti processabili (${matchAutomatici} automatici + ${matchManuali} manuali).</p>
                <button onclick="proceedToGeneration()" 
                        style="background: #28a745; color: white; border: none; font-size: 18px; padding: 15px 30px; border-radius: 8px; cursor: pointer;">
                    Procedi alla Generazione Ricevute
                </button>
            </div>
        `;
    }
    
    resultDiv.innerHTML = html;
}

// Funzioni di gestione archivio
function showGroupDetails(index) {
    const detailRow = document.getElementById(`details_${index}`);
    if (detailRow) {
        if (detailRow.style.display === 'none') {
            detailRow.style.display = '';
        } else {
            detailRow.style.display = 'none';
        }
    }
}

// Procedi alla generazione
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
    results.length = 0;
    
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
    
    console.log('Results generati con pre-filtro nomi simili e archivio raggruppato:', results.length);
    
    // Abilita generazione ricevute
    document.getElementById('generateBtn').disabled = false;
    
    const matchAutomatici = movimentiMatchati.filter(m => m.tipo === 'ADDEBITO').length;
    const matchManuali = movimentiMatchati.filter(m => m.tipo === 'MATCH_MANUALE').length;
    
    alert(`✅ Elaborazione completata con archivio intelligente!\n\nSaranno generate ${results.length} ricevute (raggruppate per mese).\n\nMatch trovati:\n• ${matchAutomatici} automatici\n• ${matchManuali} manuali (solo per nomi simili)\n\nL'archivio raggruppa i movimenti per nome base (es: MICROSOFT*)\n\nOra puoi cliccare su "Genera Ricevute" per creare i documenti HTML.`);
}

// Esposizione funzioni al contesto globale
window.loadIscrizioni = loadIscrizioni;
window.loadMovimenti = loadMovimenti;
window.performMatching = performMatching;
window.showGroupDetails = showGroupDetails;
window.proceedToGeneration = proceedToGeneration;
window.extractBaseName = extractBaseName;

console.log('🔍 matching.js CARICATO - Sistema con archivio raggruppato per nome base implementato');
console.log('Funzioni esposte:', {
    loadIscrizioni: typeof window.loadIscrizioni,
    performMatching: typeof window.performMatching,
    extractBaseName: typeof window.extractBaseName
});
