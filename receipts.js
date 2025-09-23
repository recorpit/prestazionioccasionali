// Generazione ricevute
function generateReceipts() {
    console.log('generateReceipts chiamata, results:', results);
    console.log('results.length:', results ? results.length : 'results undefined');
    
    if (!results || results.length === 0) {
        console.log('Nessun result trovato');
        alert('Nessun match trovato per generare ricevute');
        return;
    }
    
    console.log('Inizio generazione ricevute per', results.length, 'elementi');
    
    const container = document.getElementById('receiptsContainer');
    const previewContainer = document.getElementById('previewArea');
    
    if (!container) {
        console.error('receiptsContainer non trovato');
        return;
    }
    
    if (!previewContainer) {
        console.error('previewArea non trovato');
        return;
    }
    
    console.log('Container trovati, inizio elaborazione...');
    
    const container = document.getElementById('receiptsContainer');
    const previewContainer = document.getElementById('previewArea');
    
    container.innerHTML = '';
    previewContainer.innerHTML = '<h4>Anteprima Ricevute Generate:</h4>';
    
    let alertiSuperamento = [];
    let totalePrestazioni = 0;
    let totaleRimborsi = 0;
    
    results.forEach((person, index) => {
        // Controllo limite €2.500
        const compensoNetto = person.compenso * 0.8;
        const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
        const risultatoControllo = checkAndUpdateAnnualAmount(cfKey, compensoNetto);
        
        // Accumula i totali
        totalePrestazioni += person.compenso;
        totaleRimborsi += person.rimborsoSpese;
        
        if (risultatoControllo.superaLimite) {
            alertiSuperamento.push({
                nome: person.nome,
                cognome: person.cognome,
                cf: person.codiceFiscale,
                totalePrec: risultatoControllo.totaleAttuale,
                nuovoTotale: risultatoControllo.nuovoTotale,
                compensoRicevuta: compensoNetto
            });
        }
        
        const receipt = generateReceipt(person);
        receiptsContainer.innerHTML += receipt;
        
        // Anteprima
        const numeroRicevuta = getCurrentReceiptNumber(cfKey);
        const previewItem = document.createElement('div');
        previewItem.style.cssText = 'border: 1px solid #ddd; padding: 10px; margin: 10px 0; background: #f9f9f9;';
        
        let warningHtml = '';
        if (risultatoControllo.superaLimite) {
            warningHtml = '<br><span style="color: red; font-weight: bold;">⚠️ ATTENZIONE: Supera limite €2500!</span>';
        }
        
        const mesi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                     'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
        
        previewItem.innerHTML = `
            <strong>#${numeroRicevuta} - ${person.nome} ${person.cognome}</strong><br>
            CF: ${person.codiceFiscale}<br>
            Periodo: ${mesi[person.mese]} ${person.anno}<br>
            ${person.movimenti.length > 1 ? `Somma di ${person.movimenti.length} movimenti<br>` : ''}
            Totale movimenti: € ${person.movimentoBancario.toFixed(2)}<br>
            ${person.rimborsoSpese > 0 ? `Rimborso spese: € ${person.rimborsoSpese.toFixed(2)}<br>` : 'Nessun rimborso spese<br>'}
            Compenso lordo: € ${person.compenso.toFixed(2)}<br>
            <strong>Netto a pagare: € ${(person.compenso * 0.8 + person.rimborsoSpese).toFixed(2)}</strong><br>
            Totale annuale: € ${risultatoControllo.nuovoTotale.toFixed(2)}${warningHtml}
        `;
        previewContainer.appendChild(previewItem);
    });
    
    // Aggiungi riepilogo totali alla fine dell'anteprima
    const totalsItem = document.createElement('div');
    totalsItem.style.cssText = 'border: 2px solid #007bff; padding: 15px; margin: 20px 0; background: #e7f3ff; font-weight: bold;';
    totalsItem.innerHTML = `
        <h4 style="color: #007bff; margin-top: 0;">RIEPILOGO TOTALI</h4>
        <div>Numero ricevute generate: <strong>${results.length}</strong></div>
        <div>Totale prestazioni (lordo): <strong>€ ${totalePrestazioni.toFixed(2)}</strong></div>
        <div>Totale rimborsi spese: <strong>€ ${totaleRimborsi.toFixed(2)}</strong></div>
        <div style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #007bff;">
            <strong>TOTALE COMPLESSIVO: € ${(totalePrestazioni * 0.8 + totaleRimborsi).toFixed(2)}</strong>
        </div>
    `;
    previewContainer.appendChild(totalsItem);
    
    // Mostra alert superamento
    if (alertiSuperamento.length > 0) {
        let messaggioAlert = '⚠️ ATTENZIONE - SUPERAMENTO LIMITE €2500 NETTI ANNUALI!\n\n';
        messaggioAlert += 'I seguenti artisti superano il limite con questa ricevuta:\n\n';
        
        alertiSuperamento.forEach(alert => {
            messaggioAlert += `${alert.nome} ${alert.cognome} (CF: ${alert.cf})\n`;
            messaggioAlert += `- Totale precedente: € ${alert.totalePrec.toFixed(2)}\n`;
            messaggioAlert += `- Compenso questa ricevuta: € ${alert.compensoRicevuta.toFixed(2)}\n`;
            messaggioAlert += `- NUOVO TOTALE: € ${alert.nuovoTotale.toFixed(2)}\n\n`;
        });
        
        messaggioAlert += 'Questi artisti potrebbero dover aprire Partita IVA!';
        alert(messaggioAlert);
    }
    
    // Abilita download e anteprima PDF
    document.getElementById('downloadBtn').disabled = false;
    document.getElementById('downloadByMonthBtn').disabled = false;
    document.getElementById('pdfPreviewBtn').disabled = false;
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('exportByMonthBtn').disabled = false;
    
    // Mostra tab anteprima
    document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    const tabs = document.querySelectorAll('.tab');
    const previewTab = tabs[1]; // Il secondo tab è l'anteprima
    previewTab.classList.add('active');
    document.getElementById('previewTab').classList.add('active');
    
    document.getElementById('previewArea').scrollIntoView({ behavior: 'smooth' });
}

// Generazione HTML ricevuta singola
function generateReceipt(person) {
    const today = new Date();
    const dateStr = today.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
    
    const ritenuta = person.compenso * 0.20;
    const compensoNetto = person.compenso - ritenuta;
    const nettoPagare = compensoNetto + person.rimborsoSpese;
    const needsStamp = person.compenso > 77.47;
    
    // Ottieni numero progressivo
    const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
    const numeroRicevuta = getNextReceiptNumber(cfKey);
    
    const mesi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    
    // Template HTML ricevuta - formato identico all'originale
    if (person.rimborsoSpese === 0) {
        // Ricevuta SENZA rimborso spese
        return `
            <div class="ricevuta">
                <div class="ricevuta-header">
                    <div style="text-align: left; margin-bottom: 30px;">
                        <strong>${person.nome} ${person.cognome}</strong><br>
                        ${person.indirizzo}<br>
                        ${person.cap} – ${person.citta} – ${person.provincia}<br>
                        ${person.codiceFiscale}
                    </div>
                </div>
                
                <hr style="border: 1px solid #000; margin: 30px 0;">
                
                <div class="ricevuta-info">
                    <div>
                        <strong>SPETT.LE</strong><br>
                        OKL SRL<br>
                        VIA MONTE PASUBIO<br>
                        36010 – ZANE' – (VI)<br>
                        P.I. 04433920248
                    </div>
                    <div style="text-align: right;">
                        <strong>RICEVUTA NUM: ${numeroRicevuta}</strong><br>
                        <strong>DATA: ${dateStr}</strong>
                    </div>
                </div>
                
                <div style="margin: 30px 0;">
                    <strong>DESCRIZIONE ATTIVITÀ:</strong> COMPENSO PER PRESTAZIONE ARTISTICA DELLO SPETTACOLO
                </div>
                
                <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
                    <thead>
                        <tr style="background-color: #003d7a;">
                            <th style="color: white; padding: 10px; text-align: left; border: 1px solid #000;">DESCRIZIONE COMPENSO</th>
                            <th style="color: white; padding: 10px; text-align: center; border: 1px solid #000;">IMPORTO</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td style="border: 1px solid #000; padding: 10px;">COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE</td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${person.compenso.toFixed(2)} €</td>
                        </tr>
                        <tr style="background-color: #f0f0f0;">
                            <td style="border: 1px solid #000; padding: 10px; text-align: center;"><strong>COMPENSO LORDO</strong></td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;"><strong>${person.compenso.toFixed(2)} €</strong></td>
                        </tr>
                        <tr>
                            <td style="border: 1px solid #000; padding: 10px;">RITENUTA D'ACCONTO IRPEF 20% - Art. 25 DPR 633/72</td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${ritenuta.toFixed(2)} €</td>
                        </tr>
                        <tr style="background-color: #f0f0f0;">
                            <td style="border: 1px solid #000; padding: 10px; text-align: center;"><strong>NETTO A PAGARE</strong></td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;"><strong>${compensoNetto.toFixed(2)} €</strong></td>
                        </tr>
                    </tbody>
                </table>
                
                <div style="margin-top: 30px; font-size: 12px; line-height: 1.6;">
                    <p><em>${needsStamp ? 'Imposta di bollo da 2,00 euro assolta sull\'originale per importi maggiori di 77,47 euro.' : ''}</em></p>
                    <p><em>Operazione esclusa da IVA ai sensi dell'art. 5 del D.P.R. 633/72.</em></p>
                    
                    <p style="margin-top: 20px;">
                        • Il sottoscritto dichiara che, nell'anno solare in corso, <strong>alla data odierna</strong>:
                    </p>
                    <p style="margin-left: 20px;">
                        non ha conseguito redditi derivanti dall'esercizio di attività di lavoro autonomo occasionale pari o eccedenti<br>
                        i 5.000 euro e <strong>si obbliga a comunicare l'eventuale superamento</strong> del limite annuo, anche successivamente<br>
                        alla data odierna.
                    </p>
                    
                    <div style="margin-top: 50px; text-align: right;">
                        <div style="border-bottom: 1px solid #000; width: 200px; margin-left: auto; margin-bottom: 5px;"></div>
                        <div class="signature">
                            ${person.nome} ${person.cognome}
                        </div>
                        <div style="font-size: 12px;">(Firma)</div>
                    </div>
                </div>
            </div>
        `;
    } else {
        // Ricevuta CON rimborso spese
        return `
            <div class="ricevuta">
                <div class="ricevuta-header">
                    <div style="text-align: left; margin-bottom: 30px;">
                        <strong>${person.nome} ${person.cognome}</strong><br>
                        ${person.indirizzo}<br>
                        ${person.cap} – ${person.citta} – ${person.provincia}<br>
                        ${person.codiceFiscale}
                    </div>
                </div>
                
                <hr style="border: 1px solid #000; margin: 30px 0;">
                
                <div class="ricevuta-info">
                    <div>
                        <strong>SPETT.LE</strong><br>
                        OKL SRL<br>
                        VIA MONTE PASUBIO<br>
                        36010 – ZANE' – (VI)<br>
                        P.I. 04433920248
                    </div>
                    <div style="text-align: right;">
                        <strong>RICEVUTA NUM: ${numeroRicevuta}</strong><br>
                        <strong>DATA: ${dateStr}</strong>
                    </div>
                </div>
                
                <div style="margin: 30px 0;">
                    <strong>DESCRIZIONE ATTIVITÀ:</strong> COMPENSO PER PRESTAZIONE ARTISTICA DELLO SPETTACOLO
                </div>
                
                <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
                    <thead>
                        <tr style="background-color: #003d7a;">
                            <th style="color: white; padding: 10px; text-align: left; border: 1px solid #000;">DESCRIZIONE COMPENSO</th>
                            <th style="color: white; padding: 10px; text-align: center; border: 1px solid #000;">IMPORTO</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td style="border: 1px solid #000; padding: 10px;">COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE</td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${person.compenso.toFixed(2)} €</td>
                        </tr>
                        <tr style="background-color: #f0f0f0;">
                            <td style="border: 1px solid #000; padding: 10px; text-align: center;"><strong>COMPENSO LORDO</strong></td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;"><strong>${person.compenso.toFixed(2)} €</strong></td>
                        </tr>
                        <tr>
                            <td style="border: 1px solid #000; padding: 10px;">RITENUTA D'ACCONTO IRPEF 20% - Art. 25 DPR 633/72</td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${ritenuta.toFixed(2)} €</td>
                        </tr>
                        <tr style="background-color: #f0f0f0;">
                            <td style="border: 1px solid #000; padding: 10px;">COMPENSO NETTO</td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${compensoNetto.toFixed(2)} €</td>
                        </tr>
                        <tr>
                            <td style="border: 1px solid #000; padding: 10px;">RIMBORSO SPESE</td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${person.rimborsoSpese.toFixed(2)} €</td>
                        </tr>
                        <tr style="background-color: #e0e0e0;">
                            <td style="border: 1px solid #000; padding: 10px; text-align: center;"><strong>NETTO A PAGARE</strong></td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;"><strong>${nettoPagare.toFixed(2)} €</strong></td>
                        </tr>
                    </tbody>
                </table>
                
                <div style="margin-top: 30px; font-size: 12px; line-height: 1.6;">
                    <p><em>${needsStamp ? 'Imposta di bollo da 2,00 euro assolta sull\'originale per importi maggiori di 77,47 euro.' : ''}</em></p>
                    <p><em>Operazione esclusa da IVA ai sensi dell'art. 5 del D.P.R. 633/72.</em></p>
                    
                    <p style="margin-top: 20px;">
                        • Il sottoscritto dichiara che, nell'anno solare in corso, <strong>alla data odierna</strong>:
                    </p>
                    <p style="margin-left: 20px;">
                        non ha conseguito redditi derivanti dall'esercizio di attività di lavoro autonomo occasionale pari o eccedenti<br>
                        i 5.000 euro e <strong>si obbliga a comunicare l'eventuale superamento</strong> del limite annuo, anche successivamente<br>
                        alla data odierna.
                    </p>
                    
                    <div style="margin-top: 50px; text-align: right;">
                        <div style="border-bottom: 1px solid #000; width: 200px; margin-left: auto; margin-bottom: 5px;"></div>
                        <div class="signature">
                            ${person.nome} ${person.cognome}
                        </div>
                        <div style="font-size: 12px;">(Firma)</div>
                    </div>
                </div>
            </div>
        `;
    }
}

// Esposizione funzioni al contesto globale
window.generateReceipts = generateReceipts;
window.generateReceipt = generateReceipt;
