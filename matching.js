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

// Trova match per controparte
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

// Trova corrispondenze simili per matching fuzzy
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
    return suggestions.sort((a, b) => b.similarity - a.similarity).slice(0, 3); // Max 3 suggerimenti
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
    
    console.log('Match trovati:', matchedMovimenti.length, 'Non matchati:', notMatched.length);
    
    // Se ci sono non matchati, offri matching intelligente
    if (notMatched.length > 0) {
        const useIntelligent = confirm(
            `Trovati ${notMatched.length} movimenti non matchati.\n\n` +
            `Vuoi attivare il matching intelligente per cercare corrispondenze simili?`
        );
        
        if (useIntelligent) {
            showUnmatchedDialog(notMatched).then(manualMatches => {
                const finalMatched = [...matchedMovimenti, ...manualMatches];
                processMatchingResults(finalMatched);
            });
        } else {
            processMatchingResults(matchedMovimenti);
        }
    } else {
        processMatchingResults(matchedMovimenti);
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
    }
}

// Mostra risultati matching
function showMatchingResults(results, numMatched) {
    const resultDiv = document.getElementById('matchingResult');
    let html = `
        <div class="info-box">
            <strong>Risultati del matching:</strong><br>
            ✓ Ricevute da generare: ${results.length}<br>
            ✓ Movimenti matchati: ${numMatched}<br>
        </div>
    `;
    
    if (results.length > 0) {
        html += `
            <h4>Ricevute da generare (divise per mese):</h4>
            <table>
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Cognome</th>
                        <th>CF</th>
                        <th>Mese/Anno</th>
                        <th>N° Mov.</th>
                        <th>Totale</th>
                        <th>Rimborso</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        results.forEach(item => {
            html += `
                <tr>
                    <td>${item.nome || ''}</td>
                    <td>${item.cognome || ''}</td>
                    <td>${item.codiceFiscale || ''}</td>
                    <td>${item.mese}/${item.anno}</td>
                    <td>${item.movimenti.length}</td>
                    <td>€ ${item.movimentoBancario.toFixed(2)}</td>
                    <td>€ ${item.rimborsoSpese.toFixed(2)}</td>
                    <td><span class="match-status match-found">OK</span></td>
                </tr>
            `;
        });
        
        html += '</tbody></table>';
    }
    
    resultDiv.innerHTML = html;
}
