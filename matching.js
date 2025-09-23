// Funzioni per il caricamento dei file
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
            if (workbook.Sheets['ISCRIZIONI']) {
                targetSheet = workbook.Sheets['ISCRIZIONI'];
            } else if (workbook.Sheets['Iscrizioni']) {
                targetSheet = workbook.Sheets['Iscrizioni'];
            } else if (workbook.Sheets['iscrizioni']) {
                targetSheet = workbook.Sheets['iscrizioni'];
            } else {
                alert(`Foglio "ISCRIZIONI" non trovato!\nFogli disponibili: ${workbook.SheetNames.join(', ')}\nAssicurati che il foglio si chiami esattamente "ISCRIZIONI"`);
                return;
            }
            
            iscrizioniData = XLSX.utils.sheet_to_json(targetSheet, { header: 1 });
            
            // Converti array in oggetti con le chiavi delle colonne
            const processedData = [];
            for (let i = 0; i < iscrizioniData.length; i++) {
                const row = iscrizioniData[i];
                if (row.length > 0) {
                    processedData.push({
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
            iscrizioniData = processedData.filter(row => row.B); // Filtra righe con CF
            
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

// Funzioni di matching
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
        const nome = isc.D; // Colonna D = Nome
        const cognome = isc.C; // Colonna C = Cognome
        
        if (!nome || !cognome) return false;
        
        const nomeNorm = normalizeString(nome);
        const cognomeNorm = normalizeString(cognome);
        
        return (controparteNorm.includes(nomeNorm) && controparteNorm.includes(cognomeNorm)) ||
               controparteNorm === (nomeNorm + cognomeNorm) ||
               controparteNorm === (cognomeNorm + nomeNorm);
    });
}

// Funzioni di matching fuzzy
function calculateSimilarity(str1, str2) {
    if (!str1 || !str2) return 0;
    
    const s1 = normalizeString(str1);
    const s2 = normalizeString(str2);
    
    if (s1 === s2) return 1;
    
    // Calcolo della similarità di Levenshtein
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

function findSimilarMatches(controparte, iscrizioni, threshold = 0.7) {
    const suggestions = [];
    
    iscrizioni.forEach(isc => {
        const nome = isc.D;
        const cognome = isc.C;
        
        if (!nome || !cognome) return;
        
        const nomeCompleto = `${nome} ${cognome}`;
        const similarity = calculateSimilarity(controparte, nomeCompleto);
        
        if (similarity >= threshold && similarity < 1) {
            suggestions.push({
                iscrizione: isc,
                nomeCompleto: nomeCompleto,
                similarity: similarity
            });
        }
    });
    
    return suggestions.sort((a, b) => b.similarity - a.similarity).slice(0, 3);
}

async function showUnmatchedDialog(notMatched) {
    return new Promise((resolve) => {
        let manualMatches = [];
        let currentIndex = 0;
        
        function showNextUnmatched() {
            if (currentIndex >= notMatched.length) {
                resolve(manualMatches);
                return;
            }
            
            const movimento = notMatched[currentIndex];
            const controparte = findColumnValue(movimento.movimento, ['CONTROPARTE', 'Controparte', 'controparte']);
            const importo = getImportoFromMovimento(movimento.movimento);
            
            const suggestions = findSimilarMatches(controparte, iscrizioniData);
            
            let message = `MOVIMENTO NON MATCHATO #${currentIndex + 1} di ${notMatched.length}\n\n`;
            message += `Controparte: "${controparte}"\n`;
            message += `Importo: € ${importo.toFixed(2)}\n\n`;
            
            if (suggestions.length > 0) {
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
                
                if (choice === null || choice.toUpperCase() === 'A') {
                    resolve([]);
                    return;
                } else if (choice.toUpperCase() === 'S') {
                    currentIndex++;
                    showNextUnmatched();
                    return;
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
            } else {
                message += `Nessuna corrispondenza simile trovata.\n\n`;
                message += `S: Salta questo movimento\n`;
                message += `A: Annulla tutto`;
                
                const choice = prompt(message);
                
                if (choice === null || choice.toUpperCase() === 'A') {
                    resolve([]);
                    return;
                }
            }
            
            currentIndex++;
            showNextUnmatched();
        }
        
        showNextUnmatched();
    });
}

// Funzione principale di matching
async function performMatching() {
    let matchedMovimenti = [];
    let notMatched = [];
    
    console.log('Inizio matching con', movimentiData.length, 'movimenti e', iscrizioniData.length, 'iscrizioni');
    
    // Primo giro di matching automatico
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
    
    console.log('Match trovati:', matchedMovimenti.length);
    console.log('Non matchati:', notMatched.length);
    
    // Matching fuzzy per i non matchati
    if (notMatched.length > 0) {
        const proceedWithFuzzy = confirm(
            `Trovati ${notMatched.length} movimenti non matchati.\n\n` +
            `Vuoi procedere con il matching intelligente per trovare corrispondenze simili?\n\n` +
            `(Ti verranno mostrati i suggerimenti uno per uno)`
        );
        
        if (proceedWithFuzzy) {
            const manualMatches = await showUnmatchedDialog(notMatched);
            
            if (manualMatches.length > 0) {
                matchedMovimenti.push(...manualMatches);
                
                manualMatches.forEach(manualMatch => {
                    const index = notMatched.findIndex(nm => nm.movimento === manualMatch.movimento);
                    if (index !== -1) {
                        notMatched.splice(index, 1);
                    }
                });
                
                alert(`Match aggiuntivi trovati: ${manualMatches.length}\nMovimenti ancora non matchati: ${notMatched.length}`);
            }
        }
    }
    
    // Processa i risultati finali
    processMatchingResults(matchedMovimenti, notMatched);
}

function processMatchingResults(matchedMovimenti, notMatched) {
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
                // Nuova scala rimborsi spese
                let rimborsoSpese = 0;
                if (totaleNetto >= 500) {
                    rimborsoSpese = totaleNetto * 0.40; // 40% per importi ≥ 500€
                } else if (totaleNetto >= 450) {
                    rimborsoSpese = 200;
                } else if (totaleNetto >= 350) {
                    rimborsoSpese = 150;
                } else if (totaleNetto >= 250) {
                    rimborsoSpese = 100;
                } else if (totaleNetto >= 150) {
                    rimborsoSpese = 60;
                } else if (totaleNetto >= 80) {
                    rimborsoSpese = 40;
                }
                
                const compensoNetto = totaleNetto - rimborsoSpese;
                const compensoLordo = compensoNetto / 0.8;
                
                results.push({
                    nome: gruppo.iscrizione.D,
                    cognome: gruppo.iscrizione.C,
                    codiceFiscale: gruppo.iscrizione.B,
                    partitaIva: gruppo.iscrizione.M,
                    iban: gruppo.iscrizione.O,
                    indirizzo: gruppo.iscrizione.H,
                    cap: gruppo.iscrizione.K,
                    citta: gruppo.iscrizione.I,
                    provincia: gruppo.iscrizione.J,
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
    
    // Mostra i risultati nell'interfaccia
    displayMatchingResults(matchedMovimenti, notMatched);
}

function displayMatchingResults(matchedMovimenti, notMatched) {
    const resultDiv = document.getElementById('matchingResult');
    let html = `
        <div class="info-box">
            <strong>Risultati del matching:</strong><br>
            ✓ Ricevute da generare: ${results.length} (divise per mese)<br>
            ✓ Movimenti matchati: ${matchedMovimenti.length}<br>
            ✗ Non matchati: ${notMatched.length}<br>
            Totale movimenti: ${movimentiData.length}
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
        document.getElementById('generateBtn').disabled = false;
    }
    
    if (notMatched.length > 0) {
        html += `
            <h4>Movimenti ancora non matchati:</h4>
            <table>
                <thead>
                    <tr>
                        <th>Controparte</th>
                        <th>Importo</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        notMatched.forEach(item => {
            const controparte = item.movimento.CONTROPARTE || item.movimento.Controparte || 'N/D';
            const importo = getImportoFromMovimento(item.movimento);
            
            html += `
                <tr>
                    <td>${controparte}</td>
                    <td>€ ${importo.toFixed(2)}</td>
                </tr>
            `;
        });
        
        html += '</tbody></table>';
    }
    
    resultDiv.innerHTML = html;
}
