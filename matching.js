// Estrazione importo accrediti da movimento
function getAccreditiFromMovimento(movimento) {
    const accreditiColumns = ['ACCREDITI', 'Accrediti', 'ACCREDITO', 'Accredito'];
    
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

// Mostra avviso per accrediti da persone nelle iscrizioni
function showAccreditiWarning(accreditiFromIscritti) {
    let message = '⚠️ ATTENZIONE - POSSIBILI RIMBORSI RILEVATI!\n\n';
    message += 'Trovati bonifici IN ENTRATA da persone presenti nelle iscrizioni:\n\n';
    
    accreditiFromIscritti.forEach((accredito, index) => {
        const controparte = findColumnValue(accredito.movimento, ['CONTROPARTE', 'Controparte', 'controparte']) || 'N/D';
        const data = accredito.data.toLocaleDateString('it-IT');
        message += `${index + 1}. ${accredito.iscrizione.D} ${accredito.iscrizione.C}\n`;
        message += `   Controparte: ${controparte}\n`;
        message += `   Importo ricevuto: € ${accredito.importoMovimento.toFixed(2)}\n`;
        message += `   Data: ${data}\n\n`;
    });
    
    message += 'Questi potrebbero essere rimborsi di bonifici errati.\n';
    message += 'VERIFICA che non siano stati generati pagamenti duplicati per queste persone!\n\n';
    message += 'Controlla manualmente se ci sono addebiti corrispondenti da escludere.';
    
    alert(message);
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
            
            // Converti array in oggetti con le chiavi delle colonne - STRUTTURA CORRETTA
            iscrizioniData = [];
            for (let i = 0; i < rawData.length; i++) {
                const row = rawData[i];
                if (row.length > 0 && row[1]) { // Verifica che ci sia almeno il CF (colonna B)
                    iscrizioniData.push({
                        // Struttura corretta del file ISCRIZIONI
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

// Trova match per controparte con filtri per escludere buste paga/stipendi
function findBestMatch(movimento, iscrizioni) {
    const controparteColumns = ['CONTROPARTE', 'Controparte', 'controparte'];
    let controparte = null;
    
    for (let col of controparteColumns) {
        if (movimento[col]) {
            controparte = movimento[col];
            break;
        }
    }
    
    if (!controparte) return null;
    
    // Filtri per escludere buste paga, stipendi e similari
    const descrizioneCompleta = (controparte + ' ' + (movimento.Descrizione || movimento.DESCRIZIONE || '')).toLowerCase();
    const terminiEsclusi = [
        'stipendio', 'stipendi', 'busta paga', 'bustapaga', 'salario', 'retribuzione',
        'cedolino', 'paga', 'salary', 'payroll', 'wage', 'wages',
        'tfr', 'liquidazione', 'indennita', 'indennità', 'contributi',
        'previdenza', 'pensione', 'inps', 'inail', 'irpef',
        'trattenute', 'ritenute', 'detrazioni', 'assegni familiari',
        'straordinari', 'ferie', 'permessi', 'malattia',
        'mensilita', 'mensilità', 'tredicesima', 'quattordicesima'
    ];
    
    for (let termine of terminiEsclusi) {
        if (descrizioneCompleta.includes(termine)) {
            console.log(`Movimento escluso (contiene "${termine}"):`, controparte);
            return null;
        }
    }
    
    const controparteNorm = normalizeString(controparte);
    
    return iscrizioni.find(isc => {
        // Usa le colonne corrette: C = Cognome, D = Nome
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

// Calcolo similarità (Levenshtein) - duplicata qui per evitare dipendenze
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

// Calcolo rimborsi spese - NUOVA SCALA (duplicata qui per evitare dipendenze)
function calculateRimborsoSpese(importo) {
    if (importo >= 500) return importo * 0.40; // 40% per importi ≥ 500€
    if (importo >= 450) return 200;
    if (importo >= 350) return 150;
    if (importo >= 250) return 100;
    if (importo >= 150) return 60;
    if (importo >= 80) return 40;
    return 0;
}

// Trova corrispondenze simili per matching fuzzy - MOLTO PIÙ SELETTIVO
function findSimilarMatches(controparte, iscrizioni, threshold = 0.90) { // Soglia molto alta: 90%
    const suggestions = [];
    
    iscrizioni.forEach(isc => {
        const nome = isc.D;
        const cognome = isc.C;
        
        if (!nome || !cognome) return;
        
        const nomeCompleto = `${nome} ${cognome}`;
        const similarity = calculateSimilarity(controparte, nomeCompleto);
        
        // Solo match molto simili (90%+ di similarità)
        if (similarity >= threshold && similarity < 1) {
            suggestions.push({
                iscrizione: isc,
                nomeCompleto: nomeCompleto,
                similarity: similarity
            });
        }
    });
    
    // Ordina per similarità decrescente e prende solo il migliore
    return suggestions.sort((a, b) => b.similarity - a.similarity).slice(0, 1); // Solo 1 suggerimento
}

// Dialog interattivo per matching fuzzy
async function showUnmatchedDialog(notMatched) {
    return new Promise((resolve) => {
        let manualMatches = [];
        let currentIndex = 0;
        
        function processNext() {
            if (currentIndex >= notMatched.length) {
                resolve(manualMatches);
                return;
            }
            
            const movimento = notMatched[currentIndex];
            const controparte = findColumnValue(movimento.movimento, ['CONTROPARTE', 'Controparte', 'controparte']);
            const importo = getImportoFromMovimento(movimento.movimento);
            
            // Trova suggerimenti simili
            const suggestions = findSimilarMatches(controparte, iscrizioniData);
            
            if (suggestions.length === 0) {
                currentIndex++;
                processNext();
                return;
            }
            
            let message = `MOVIMENTO NON MATCHATO #${currentIndex + 1} di ${notMatched.length}\n\n`;
            message += `Controparte: "${controparte}"\n`;
            message += `Importo: € ${importo.toFixed(2)}\n\n`;
            message += `Possibili corrispondenze simili:\n\n`;
            
            suggestions.forEach((sugg, idx) => {
                const similarity = (sugg.similarity * 100).toFixed(1);
                message += `${idx + 1}. ${sugg.nomeCompleto} (${similarity}% simile)\n`;
                message += `   CF: ${sugg.iscrizione.B}\n\n`;
            });
            
            message += `Scegli un'opzione:\n`;
            message += `1-${suggestions.length}: Matcha con il numero corrispondente\n`;
            message += `S: Salta questo movimento\n`;
            message += `A: Annulla tutto`;
            
            const choice = prompt(message);
            
            if (!choice || choice.toUpperCase() === 'A') {
                resolve([]);
                return;
            } else if (choice.toUpperCase() === 'S') {
                // Salta questo movimento
            } else {
                const choiceNum = parseInt(choice);
                if (choiceNum >= 1 && choiceNum <= suggestions.length) {
                    manualMatches.push({
                        movimento: movimento.movimento,
                        iscrizione: suggestions[choiceNum - 1].iscrizione,
                        importoMovimento: importo,
                        data: getDataFromMovimento(movimento.movimento)
                    });
                }
            }
            
            currentIndex++;
            setTimeout(processNext, 100);
        }
        
        processNext();
    });
}

// Matching principale
function performMatching() {
    if (iscrizioniData.length === 0 || movimentiData.length === 0) {
        alert('Carica prima entrambi i file Excel!');
        return;
    }
    
    let matchedMovimenti = [];
    let notMatched = [];
    
    console.log('Inizio matching con', movimentiData.length, 'movimenti e', iscrizioniData.length, 'iscrizioni');
    
    movimentiData.forEach((movimento, index) => {
        const match = findBestMatch(movimento, iscrizioniData);
        const importo = getImportoFromMovimento(movimento);
        const data = getDataFromMovimento(movimento);
        
        if (match) {
            matchedMovimenti.push({
                movimento: movimento,
                iscrizione: match,
                importoMovimento: importo,
                data: data
            });
        } else {
            notMatched.push({ movimento: movimento });
        }
    });
    
    // Mostra risultati
    showMatchingResults(matchedMovimenti, notMatched);
    
    // Se ci sono non matchati, offri matching intelligente
    if (notMatched.length > 0) {
        const useIntelligent = confirm(
            `Trovati ${notMatched.length} movimenti non matchati.\n\n` +
            `Vuoi attivare il matching intelligente per cercare corrispondenze simili?`
        );
        
        if (useIntelligent) {
            showUnmatchedDialog(notMatched).then(manualMatches => {
                const finalMatched = [...matchedMovimenti, ...manualMatches];
                const finalUnmatched = notMatched.filter(item => 
                    !manualMatches.some(match => match.movimento === item.movimento)
                );
                
                // Processa i risultati per creare le ricevute
                processMatchingResults(finalMatched);
                
                // Mostra i risultati aggiornati DOPO aver processato
                showMatchingResults(finalMatched, finalUnmatched);
            });
        } else {
            // Processa solo i matchati ma mostra anche i non matchati
            processMatchingResults(matchedMovimenti);
            // Mostra tutti i risultati inclusi i non matchati
            showMatchingResults(matchedMovimenti, notMatched);
        }
    } else {
        processMatchingResults(matchedMovimenti);
        // Anche qui mostra i risultati (in questo caso nessun non matchato)
        showMatchingResults(matchedMovimenti, []);
    }
}

// Processamento risultati del matching
function processMatchingResults(matchedMovimenti) {
    // Raggruppa per persona E per mese
    const gruppiPerPersonaMese = {};
    
    matchedMovimenti.forEach(item => {
        const cf = item.iscrizione.B || `${item.iscrizione.D}_${item.iscrizione.C}`;
        const mese = item.data.getMonth() + 1;
        const anno = item.data.getFullYear();
        const chiaveMese = `${cf}_${anno}_${mese}`;
        
        if (!gruppiPerPersonaMese[chiaveMese]) {
            gruppiPerPersonaMese[chiaveMese] = {
                iscrizione: item.iscrizione,
                movimenti: [],
                totaleMovimenti: 0,
                mese: mese,
                anno: anno,
                cf: cf
            };
        }
        
        gruppiPerPersonaMese[chiaveMese].movimenti.push(item.movimento);
        gruppiPerPersonaMese[chiaveMese].totaleMovimenti += item.importoMovimento;
    });
    
    // Gestione compensazioni tra mesi
    const gruppiPerCF = {};
    Object.values(gruppiPerPersonaMese).forEach(gruppo => {
        if (!gruppiPerCF[gruppo.cf]) {
            gruppiPerCF[gruppo.cf] = [];
        }
        gruppiPerCF[gruppo.cf].push(gruppo);
    });
    
    results = [];
    
    Object.values(gruppiPerCF).forEach(gruppiPersona => {
        gruppiPersona.sort((a, b) => {
            if (a.anno !== b.anno) return a.anno - b.anno;
            return a.mese - b.mese;
        });
        
        let creditoResiduo = 0;
        
        gruppiPersona.forEach(gruppo => {
            let totaleNetto = gruppo.totaleMovimenti + creditoResiduo;
            
            if (totaleNetto > 0) {
                // Calcola rimborso spese con nuova scala
                const rimborsoSpese = calculateRimborsoSpese(totaleNetto);
                const compensoNetto = totaleNetto - rimborsoSpese;
                const compensoLordo = compensoNetto / 0.8;
                
                results.push({
                    nome: gruppo.iscrizione.D, // Colonna D = Nome
                    cognome: gruppo.iscrizione.C, // Colonna C = Cognome
                    codiceFiscale: gruppo.iscrizione.B, // Colonna B = CF
                    partitaIva: gruppo.iscrizione.M, // Colonna M = P.IVA
                    iban: gruppo.iscrizione.O, // Colonna O = IBAN
                    indirizzo: gruppo.iscrizione.H, // Colonna H = Via Indirizzo 1
                    cap: gruppo.iscrizione.K, // Colonna K = CAP
                    citta: gruppo.iscrizione.I, // Colonna I = Città
                    provincia: gruppo.iscrizione.J, // Colonna J = Provincia
                    compenso: compensoLordo,
                    rimborsoSpese: rimborsoSpese,
                    movimentoBancario: totaleNetto,
                    mese: gruppo.mese,
                    anno: gruppo.anno,
                    movimenti: gruppo.movimenti
                });
                
                creditoResiduo = 0;
            } else if (totaleNetto < 0) {
                creditoResiduo = totaleNetto;
            }
        });
    });
    
    // Mostra risultati
    showMatchingResults(results, matchedMovimenti.length);
    
    // Abilita generazione ricevute
    if (results.length > 0) {
        document.getElementById('generateBtn').disabled = false;
        
        // Mostra riepilogo finale delle ricevute da generare
        showFinalSummary();
        
        console.log('Pulsante Genera Ricevute abilitato per', results.length, 'ricevute');
    } else {
        console.log('Nessuna ricevuta da generare, pulsante rimane disabilitato');
        alert('Nessuna ricevuta da generare. Verifica che ci siano movimenti validi da persone nelle iscrizioni.');
    }
}

// Mostra riepilogo finale delle ricevute da generare
function showFinalSummary() {
    const resultDiv = document.getElementById('matchingResult');
    let summaryHtml = `
        <div style="margin-top: 30px; padding: 20px; background: #e8f5e9; border: 1px solid #4CAF50; border-radius: 5px;">
            <h4 style="color: #155724; margin-top: 0;">✅ RICEVUTE PRONTE PER LA GENERAZIONE</h4>
            <p><strong>${results.length}</strong> ricevute verranno generate (divise per mese):</p>
            
            <table style="margin-top: 15px;">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Cognome</th>
                        <th>CF</th>
                        <th>Periodo</th>
                        <th>Movimenti</th>
                        <th>Totale</th>
                        <th>Rimborso</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    results.forEach(item => {
        summaryHtml += `
            <tr>
                <td>${item.nome || ''}</td>
                <td>${item.cognome || ''}</td>
                <td>${item.codiceFiscale || ''}</td>
                <td>${item.mese}/${item.anno}</td>
                <td>${item.movimenti.length}</td>
                <td>€ ${item.movimentoBancario.toFixed(2)}</td>
                <td>€ ${item.rimborsoSpese.toFixed(2)}</td>
            </tr>
        `;
    });
    
    summaryHtml += `
                </tbody>
            </table>
            <p style="margin: 15px 0 0 0; font-weight: bold; color: #155724;">
                Ora puoi cliccare su "Genera Ricevute" per creare i documenti!
            </p>
        </div>
    `;
    
    resultDiv.innerHTML += summaryHtml;
}

// Variabili globali per gestione esclusioni
let movimentiEsclusi = [];
let allMatchedMovimenti = []; // Conserva tutti i match per il controllo manuale

// Mostra risultati matching con controlli manuali
function showMatchingResults(matched, unmatched, accrediti = []) {
    // Salva tutti i match per il controllo manuale
    allMatchedMovimenti = [...matched];
    
    const resultDiv = document.getElementById('matchingResult');
    let html = `
        <div class="info-box">
            <strong>Risultati del matching:</strong><br>
            ✓ Movimenti matchati: ${matched.length}<br>
            ${unmatched.length > 0 ? `✗ Movimenti non matchati: ${unmatched.length}<br>` : ''}
            ${accrediti.length > 0 ? `⚠️ Accrediti da iscritti rilevati: ${accrediti.length}<br>` : ''}
            ${movimentiEsclusi.length > 0 ? `🚫 Movimenti esclusi manualmente: ${movimentiEsclusi.length}<br>` : ''}
        </div>
    `;
    
    if (matched.length > 0) {
        html += `
            <h4>Movimenti Matchati (Controllo Manuale):</h4>
            <div style="margin-bottom: 15px;">
                <button onclick="selectAllMovimenti(true)" style="background: #28a745; margin-right: 10px;">Seleziona Tutti</button>
                <button onclick="selectAllMovimenti(false)" style="background: #dc3545; margin-right: 10px;">Deseleziona Tutti</button>
                <button onclick="toggleEsclusioni()" style="background: #fd7e14;">Applica Esclusioni</button>
            </div>
            <table id="movimentiTable">
                <thead>
                    <tr>
                        <th style="width: 30px;">✓</th>
                        <th>Nome</th>
                        <th>Cognome</th>
                        <th>CF</th>
                        <th>Controparte</th>
                        <th>Importo</th>
                        <th>Data</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        matched.forEach((match, index) => {
            // Verifica che movimento e iscrizione siano validi
            if (!match || !match.movimento || !match.iscrizione) {
                console.warn('Match non valido saltato:', match);
                return;
            }
            
            const controparte = findColumnValue(match.movimento, ['CONTROPARTE', 'Controparte', 'controparte']) || 'N/D';
            const data = match.data ? match.data.toLocaleDateString('it-IT') : 'N/D';
            const movimentoId = `movimento_${index}`;
            const isEscluso = movimentiEsclusi.some(escl => escl.index === index);
            
            html += `
                <tr id="row_${index}" style="${isEscluso ? 'background-color: #ffebee; opacity: 0.7;' : ''}">
                    <td style="text-align: center;">
                        <input type="checkbox" id="${movimentoId}" ${!isEscluso ? 'checked' : ''} 
                               onchange="toggleMovimento(${index}, this.checked)" 
                               style="transform: scale(1.2);">
                    </td>
                    <td>${match.iscrizione.D || ''}</td>
                    <td>${match.iscrizione.C || ''}</td>
                    <td>${match.iscrizione.B || ''}</td>
                    <td>${controparte}</td>
                    <td>€ ${match.importoMovimento.toFixed(2)}</td>
                    <td>${data}</td>
                    <td>
                        <span class="match-status ${isEscluso ? 'match-not-found' : 'match-found'}">
                            ${isEscluso ? 'ESCLUSO' : 'INCLUSO'}
                        </span>
                    </td>
                </tr>
            `;
        });
        
        html += '</tbody></table>';
        
        // Riepilogo movimenti selezionati
        const movimentiInclusi = matched.filter((_, index) => !movimentiEsclusi.some(escl => escl.index === index));
        const totaleInclusi = movimentiInclusi.reduce((sum, match) => sum + match.importoMovimento, 0);
        const totaleEsclusi = movimentiEsclusi.reduce((sum, escl) => sum + escl.importo, 0);
        
        html += `
            <div style="background: #e8f5e9; border: 1px solid #4CAF50; border-radius: 5px; padding: 15px; margin: 15px 0;">
                <h5 style="margin-top: 0; color: #155724;">Riepilogo Selezione:</h5>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                    <div>
                        <strong>Movimenti da processare:</strong> ${movimentiInclusi.length}<br>
                        <strong>Importo totale:</strong> € ${totaleInclusi.toFixed(2)}
                    </div>
                    <div>
                        <strong>Movimenti esclusi:</strong> ${movimentiEsclusi.length}<br>
                        <strong>Importo escluso:</strong> € ${totaleEsclusi.toFixed(2)}
                    </div>
                </div>
            </div>
        `;
    }
    
    if (accrediti.length > 0) {
        html += `
            <h4 style="color: #ff9800;">⚠️ Accrediti da Persone nelle Iscrizioni (${accrediti.length}):</h4>
            <div style="background: #fff3cd; border: 1px solid #ff9800; border-radius: 5px; padding: 15px; margin: 10px 0;">
                <p style="margin: 0 0 10px 0; font-weight: bold; color: #856404;">
                    ATTENZIONE: Trovati bonifici IN ENTRATA da persone che sono anche nelle iscrizioni!
                </p>
                <p style="margin: 0 0 15px 0; font-style: italic; color: #856404;">
                    Potrebbero essere rimborsi di bonifici errati. Verifica se esistono addebiti corrispondenti da escludere.
                </p>
                
                <button onclick="autoEscludiDuplicati(${JSON.stringify(accrediti).replace(/"/g, '&quot;')})" 
                        style="background: #ff9800; margin-bottom: 15px;">
                    Escludi Automaticamente Duplicati Sospetti
                </button>
                
                <table>
                    <thead>
                        <tr style="background-color: #ff9800;">
                            <th style="color: white;">Nome</th>
                            <th style="color: white;">Cognome</th>
                            <th style="color: white;">Controparte</th>
                            <th style="color: white;">Importo Ricevuto</th>
                            <th style="color: white;">Data</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        accrediti.forEach((accredito, index) => {
            const controparte = findColumnValue(accredito.movimento, ['CONTROPARTE', 'Controparte', 'controparte']) || 'N/D';
            const data = accredito.data.toLocaleDateString('it-IT');
            html += `
                <tr style="background-color: #fff3cd;">
                    <td><strong>${accredito.iscrizione.D || ''}</strong></td>
                    <td><strong>${accredito.iscrizione.C || ''}</strong></td>
                    <td>${controparte}</td>
                    <td><strong style="color: #ff9800;">€ ${accredito.importoMovimento.toFixed(2)}</strong></td>
                    <td>${data}</td>
                </tr>
            `;
        });
        
        html += `
                    </tbody>
                </table>
            </div>
        `;
    }
    
    if (unmatched.length > 0) {
        html += `
            <h4 style="color: #dc3545;">⚠️ Movimenti Non Matchati (${unmatched.length}):</h4>
            <div style="background: #f8d7da; border: 1px solid #dc3545; border-radius: 5px; padding: 15px; margin: 10px 0;">
                <table>
                    <thead>
                        <tr style="background-color: #dc3545;">
                            <th style="color: white;">Controparte</th>
                            <th style="color: white;">Importo</th>
                            <th style="color: white;">Status</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        unmatched.forEach((movimento, index) => {
            const controparte = findColumnValue(movimento.movimento, ['CONTROPARTE', 'Controparte', 'controparte']) || 'N/D';
            const importo = getImportoFromMovimento(movimento.movimento);
            
            html += `
                <tr style="background-color: #f8d7da;">
                    <td><strong>${controparte}</strong></td>
                    <td><strong>€ ${importo.toFixed(2)}</strong></td>
                    <td><span class="match-status match-not-found">NON PROCESSATO</span></td>
                </tr>
            `;
        });
        
        html += '</tbody></table></div>';
    }
    
    resultDiv.innerHTML = html;
}

// Funzioni per il controllo manuale
function selectAllMovimenti(select) {
    allMatchedMovimenti.forEach((_, index) => {
        const checkbox = document.getElementById(`movimento_${index}`);
        if (checkbox) {
            checkbox.checked = select;
            toggleMovimento(index, select);
        }
    });
}

function toggleMovimento(index, includi) {
    const movimento = allMatchedMovimenti[index];
    if (!movimento) return;
    
    if (includi) {
        // Rimuovi dalle esclusioni se presente
        movimentiEsclusi = movimentiEsclusi.filter(escl => escl.index !== index);
    } else {
        // Aggiungi alle esclusioni se non presente
        if (!movimentiEsclusi.some(escl => escl.index === index)) {
            movimentiEsclusi.push({
                index: index,
                movimento: movimento,
                importo: movimento.importoMovimento,
                motivo: 'Esclusione manuale'
            });
        }
    }
    
    // Aggiorna visivamente la riga
    const row = document.getElementById(`row_${index}`);
    if (row) {
        if (includi) {
            row.style.backgroundColor = '';
            row.style.opacity = '1';
            const statusCell = row.querySelector('.match-status');
            statusCell.textContent = 'INCLUSO';
            statusCell.className = 'match-status match-found';
        } else {
            row.style.backgroundColor = '#ffebee';
            row.style.opacity = '0.7';
            const statusCell = row.querySelector('.match-status');
            statusCell.textContent = 'ESCLUSO';
            statusCell.className = 'match-status match-not-found';
        }
    }
    
    // Aggiorna il riepilogo
    updateRiepilogoSelezione();
}

function updateRiepilogoSelezione() {
    const movimentiInclusi = allMatchedMovimenti.filter((_, index) => 
        !movimentiEsclusi.some(escl => escl.index === index)
    );
    const totaleInclusi = movimentiInclusi.reduce((sum, match) => sum + match.importoMovimento, 0);
    const totaleEsclusi = movimentiEsclusi.reduce((sum, escl) => sum + escl.importo, 0);
    
    // Aggiorna il riepilogo nella pagina se esiste
    const riepilogoDiv = document.querySelector('[style*="background: #e8f5e9"]');
    if (riepilogoDiv) {
        const content = riepilogoDiv.querySelector('div[style*="grid-template-columns"]');
        if (content) {
            content.innerHTML = `
                <div>
                    <strong>Movimenti da processare:</strong> ${movimentiInclusi.length}<br>
                    <strong>Importo totale:</strong> € ${totaleInclusi.toFixed(2)}
                </div>
                <div>
                    <strong>Movimenti esclusi:</strong> ${movimentiEsclusi.length}<br>
                    <strong>Importo escluso:</strong> € ${totaleEsclusi.toFixed(2)}
                </div>
            `;
        }
    }
}

function toggleEsclusioni() {
    if (movimentiEsclusi.length === 0) {
        alert('Nessun movimento da escludere selezionato.');
        return;
    }
    
    const movimentiInclusi = allMatchedMovimenti.filter((_, index) => 
        !movimentiEsclusi.some(escl => escl.index === index)
    );
    
    const conferma = confirm(
        `Confermi le esclusioni?\n\n` +
        `Movimenti da processare: ${movimentiInclusi.length}\n` +
        `Movimenti esclusi: ${movimentiEsclusi.length}\n\n` +
        `Solo i movimenti selezionati genereranno ricevute.`
    );
    
    if (conferma) {
        // Processa solo i movimenti inclusi
        console.log('Processing movimenti inclusi:', movimentiInclusi.length);
        processMatchingResults(movimentiInclusi);
        alert(`Esclusioni applicate! ${movimentiInclusi.length} movimenti verranno processati.`);
        
        // Forza l'abilitazione del pulsante se ci sono risultati
        if (results.length > 0) {
            document.getElementById('generateBtn').disabled = false;
            console.log('Pulsante Generate Ricevute forzatamente abilitato');
        }
    }
}

function autoEscludiDuplicati(accrediti) {
    let esclusiAutomatici = 0;
    
    accrediti.forEach(accredito => {
        // Cerca addebiti con stesso CF e importo simile
        allMatchedMovimenti.forEach((match, index) => {
            if (match.iscrizione.B === accredito.iscrizione.B) {
                const differenzaImporto = Math.abs(match.importoMovimento - accredito.importoMovimento);
                const tolleranza = accredito.importoMovimento * 0.1; // 10% di tolleranza
                
                if (differenzaImporto <= tolleranza) {
                    // Esclude automaticamente questo movimento
                    if (!movimentiEsclusi.some(escl => escl.index === index)) {
                        movimentiEsclusi.push({
                            index: index,
                            movimento: match,
                            importo: match.importoMovimento,
                            motivo: 'Duplicato automatico - trovato accredito corrispondente'
                        });
                        
                        const checkbox = document.getElementById(`movimento_${index}`);
                        if (checkbox) checkbox.checked = false;
                        toggleMovimento(index, false);
                        esclusiAutomatici++;
                    }
                }
            }
        });
    });
    
    if (esclusiAutomatici > 0) {
        alert(`Esclusi automaticamente ${esclusiAutomatici} movimenti duplicati.`);
        updateRiepilogoSelezione();
    } else {
        alert('Nessun duplicato automatico rilevato.');
    }
}

// Esposizione funzioni al contesto globale
window.loadIscrizioni = loadIscrizioni;
window.loadMovimenti = loadMovimenti;
window.performMatching = performMatching;
window.selectAllMovimenti = selectAllMovimenti;
window.toggleMovimento = toggleMovimento;
window.toggleEsclusioni = toggleEsclusioni;
window.autoEscludiDuplicati = autoEscludiDuplicati;
