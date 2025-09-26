// Generazione ricevute - CON SISTEMA NUMERAZIONE UNIFICATO + FIX RIMBORSI
function generateReceipts() {
    console.log('generateReceipts chiamata, results:', results);
    console.log('results.length:', results ? results.length : 'results undefined');
    
    if (!results || results.length === 0) {
        console.log('Nessun result trovato');
        alert('Nessuna ricevuta da generare. Esegui prima il matching e clicca "Procedi alla Generazione".');
        return;
    }
    
    console.log('Inizio generazione ricevute per', results.length, 'elementi');
    
    // PRIMO PASSO: Assegna numeri progressivi cronologici
    assignProgressiveNumbers();
    
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
    
    container.innerHTML = '';
    previewContainer.innerHTML = '<h3>📋 Anteprima Ricevute Generate</h3>';
    
    let alertiSuperamento = [];
    let totalePrestazioni = 0;
    let totaleRimborsi = 0;
    
    // I results sono già ordinati dalla funzione assignProgressiveNumbers()
    console.log('Results elaborazione:', 
        results.map(r => `${r.mese}/${r.anno} - ${r.nome} ${r.cognome} - Numero: ${r.numeroProgressivo}`));
    
    results.forEach((person, index) => {
        // FIX ERRORE: Assicura che rimborsoSpese sia sempre un numero
        if (person.rimborsoSpese === undefined || person.rimborsoSpese === null || isNaN(person.rimborsoSpese)) {
            console.warn(`Fix rimborsoSpese per ${person.nome} ${person.cognome}: ${person.rimborsoSpese} → 0`);
            person.rimborsoSpese = 0;
        }
        
        // Controllo limite €2.500
        const compensoNetto = person.compenso * 0.8;
        const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
        const risultatoControllo = checkAndUpdateAnnualAmount(cfKey, compensoNetto);
        
        // Accumula i totali - CON CONTROLLO SICUREZZA
        totalePrestazioni += person.compenso || 0;
        totaleRimborsi += person.rimborsoSpese || 0;
        
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
        
        // Genera ricevuta HTML usando il numero progressivo unificato
        const receipt = createReceiptHTML(person, person.numeroProgressivo, getReceiptDate(person.anno, person.mese));
        container.innerHTML += receipt;
        
        // Anteprima con dettagli movimenti
        const previewItem = document.createElement('div');
        previewItem.style.cssText = 'border: 1px solid #dee2e6; padding: 15px; margin: 15px 0; background: #f8f9fa; border-radius: 8px;';
        
        let warningHtml = '';
        if (risultatoControllo.superaLimite) {
            warningHtml = '<br><span style="color: #dc3545; font-weight: bold;">⚠️ ATTENZIONE: Supera limite €2500!</span>';
        }
        
        const mesi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                     'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
        
        // Dettaglio movimenti se presente
        let movimentiDettaglio = '';
        if (person.dettaglioMovimenti && person.dettaglioMovimenti.length > 0) {
            movimentiDettaglio = '<br><strong>Dettaglio movimenti:</strong><br>';
            person.dettaglioMovimenti.forEach(mov => {
                movimentiDettaglio += `• ${mov.controparte}: €${mov.importo.toFixed(2)} (${mov.data.toLocaleDateString('it-IT')})<br>`;
            });
        }
        
        // RIMBORSI SICURI - RIGA 103 ORIGINALE CORRETTA
        const rimborsiSafe = person.rimborsoSpese || 0;
        const rimborsiDisplay = rimborsiSafe > 0 ? 
            `<strong>Rimborso spese:</strong> €${rimborsiSafe.toFixed(2)}<br>` : 
            '<em>Nessun rimborso spese</em><br>';
        
        previewItem.innerHTML = `
            <div style="display: grid; grid-template-columns: 1fr 200px; gap: 15px; align-items: start;">
                <div>
                    <h4 style="margin: 0 0 10px 0; color: #495057;">
                        #${person.numeroProgressivo} - ${person.nome} ${person.cognome}
                    </h4>
                    <div style="font-size: 14px; line-height: 1.6;">
                        <strong>CF:</strong> ${person.codiceFiscale}<br>
                        <strong>Periodo:</strong> ${mesi[person.mese]} ${person.anno}<br>
                        ${person.movimenti && person.movimenti.length > 1 ? `<strong>Movimenti aggregati:</strong> ${person.movimenti.length}<br>` : ''}
                        ${movimentiDettaglio}
                        <strong>Movimento bancario totale:</strong> €${(person.movimentoBancario || 0).toFixed(2)}<br>
                        ${rimborsiDisplay}
                        <strong>Compenso lordo:</strong> €${(person.compenso || 0).toFixed(2)}<br>
                        <strong>Totale annuale:</strong> €${risultatoControllo.nuovoTotale.toFixed(2)}${warningHtml}
                    </div>
                </div>
                <div style="background: #e8f5e9; padding: 10px; border-radius: 5px; text-align: center;">
                    <strong style="color: #155724;">Netto a Pagare</strong><br>
                    <span style="font-size: 18px; font-weight: bold; color: #155724;">
                        €${((person.compenso || 0) * 0.8 + (person.rimborsoSpese || 0)).toFixed(2)}
                    </span>
                </div>
            </div>
        `;
        previewContainer.appendChild(previewItem);
    });
    
    // Aggiungi riepilogo totali alla fine dell'anteprima - CON CONTROLLI SICUREZZA
    const totalsItem = document.createElement('div');
    totalsItem.style.cssText = 'border: 2px solid #007bff; padding: 20px; margin: 25px 0; background: #e3f2fd; border-radius: 10px;';
    totalsItem.innerHTML = `
        <h4 style="color: #1976d2; margin-top: 0; text-align: center;">📊 RIEPILOGO TOTALI RICEVUTE</h4>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; text-align: center;">
            <div style="background: white; padding: 15px; border-radius: 8px;">
                <strong style="color: #495057;">Ricevute Generate</strong><br>
                <span style="font-size: 24px; font-weight: bold; color: #007bff;">${results.length}</span>
            </div>
            <div style="background: white; padding: 15px; border-radius: 8px;">
                <strong style="color: #495057;">Prestazioni Lordo</strong><br>
                <span style="font-size: 20px; font-weight: bold; color: #28a745;">€${totalePrestazioni.toFixed(2)}</span>
            </div>
            <div style="background: white; padding: 15px; border-radius: 8px;">
                <strong style="color: #495057;">Rimborsi Spese</strong><br>
                <span style="font-size: 20px; font-weight: bold; color: #ffc107;">€${totaleRimborsi.toFixed(2)}</span>
            </div>
            <div style="background: white; padding: 15px; border-radius: 8px;">
                <strong style="color: #495057;">Totale Netto</strong><br>
                <span style="font-size: 20px; font-weight: bold; color: #dc3545;">€${(totalePrestazioni * 0.8 + totaleRimborsi).toFixed(2)}</span>
            </div>
        </div>
        <div style="margin-top: 15px; padding: 15px; background: #d4edda; border-radius: 8px; text-align: center;">
            <strong style="color: #155724;">💰 Risparmio Fiscale (20%): €${(totaleRimborsi * 0.2).toFixed(2)}</strong>
        </div>
    `;
    previewContainer.appendChild(totalsItem);
    
    // Mostra alert superamento se necessario
    if (alertiSuperamento.length > 0) {
        let messaggioAlert = '⚠️ ATTENZIONE - SUPERAMENTO LIMITE €2500 NETTI ANNUALI!\n\n';
        messaggioAlert += 'I seguenti artisti superano il limite con questa ricevuta:\n\n';
        
        alertiSuperamento.forEach(alert => {
            messaggioAlert += `${alert.nome} ${alert.cognome} (CF: ${alert.cf})\n`;
            messaggioAlert += `- Totale precedente: €${alert.totalePrec.toFixed(2)}\n`;
            messaggioAlert += `- Compenso questa ricevuta: €${alert.compensoRicevuta.toFixed(2)}\n`;
            messaggioAlert += `- NUOVO TOTALE: €${alert.nuovoTotale.toFixed(2)}\n\n`;
        });
        
        messaggioAlert += 'Questi artisti potrebbero dover aprire Partita IVA!';
        alert(messaggioAlert);
    }
    
    // Abilita tutti i pulsanti di export
    const buttonsToEnable = [
        'downloadBtn', 'downloadByMonthBtn', 'pdfPreviewBtn', 
        'exportBtn', 'exportByMonthBtn'
    ];
    
    buttonsToEnable.forEach(buttonId => {
        const button = document.getElementById(buttonId);
        if (button) {
            button.disabled = false;
            console.log(`Pulsante ${buttonId} abilitato`);
        }
    });
    
    // Mostra tab anteprima
    document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    const tabs = document.querySelectorAll('.tab');
    const previewTab = tabs[1]; // Il secondo tab è l'anteprima
    if (previewTab) {
        previewTab.classList.add('active');
        document.getElementById('previewTab').classList.add('active');
    }
    
    document.getElementById('previewArea').scrollIntoView({ behavior: 'smooth' });
    
    console.log(`✅ Ricevute generate con successo: ${results.length}`);
    alert(`✅ Generazione completata!\n\n${results.length} ricevute HTML create.\nRisparmio fiscale totale: €${(totaleRimborsi * 0.2).toFixed(2)}\n\nI pulsanti di export sono ora attivi.`);
}

// Utility per calcolare data ricevuta (ultimo giorno del mese)
function getReceiptDate(anno, mese) {
    const lastDayOfMonth = new Date(anno, mese, 0);
    return lastDayOfMonth.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

// Generazione HTML ricevuta singola - USA NUMERAZIONE PASSATA COME PARAMETRO + FIX RIMBORSI
function createReceiptHTML(person, numeroRicevuta, dataRicevuta) {
    console.log(`Creazione HTML per ${person.nome} ${person.cognome} - Numero: ${numeroRicevuta} - Data: ${dataRicevuta}`);
    
    // FIX SICUREZZA: Assicura che tutti i valori numerici siano validi
    const compenso = parseFloat(person.compenso) || 0;
    const rimborsiSpese = parseFloat(person.rimborsoSpese) || 0;
    
    const ritenuta = compenso * 0.20;
    const compensoNetto = compenso - ritenuta;
    const nettoPagare = compensoNetto + rimborsiSpese;
    const needsStamp = compenso > 77.47;
    
    console.log(`${person.nome} ${person.cognome} - Compenso: €${compenso}, Rimborsi: €${rimborsiSpese}, Netto: €${nettoPagare.toFixed(2)}`);
    
    // Template HTML ricevuta - formato identico all'originale
    if (rimborsiSpese === 0) {
        // Ricevuta SENZA rimborso spese
        return `
            <div class="ricevuta" id="receipt-${results.indexOf(person)}">
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
                        VIA MONTE PASUBIO 222/1<br>
                        36010 – ZANE' – (VI)<br>
                        P.I. 04433920248
                    </div>
                    <div style="text-align: right;">
                        <strong>RICEVUTA NUM: ${numeroRicevuta}</strong><br>
                        <strong>DATA: ${dataRicevuta}</strong>
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
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${compenso.toFixed(2)} €</td>
                        </tr>
                        <tr style="background-color: #f0f0f0;">
                            <td style="border: 1px solid #000; padding: 10px; text-align: center;"><strong>COMPENSO LORDO</strong></td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;"><strong>${compenso.toFixed(2)} €</strong></td>
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
            <div class="ricevuta" id="receipt-${results.indexOf(person)}">
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
                        VIA MONTE PASUBIO 222/1<br>
                        36010 – ZANE' – (VI)<br>
                        P.I. 04433920248
                    </div>
                    <div style="text-align: right;">
                        <strong>RICEVUTA NUM: ${numeroRicevuta}</strong><br>
                        <strong>DATA: ${dataRicevuta}</strong>
                    </div>
                </div>
                
                <div style="margin: 30px 0;">
                    <strong>DESCRIZIONE ATTIVITÀ:</strong> COMPENSO PER PRESTAZIONE ARTISTICA DELLO SPETTACOLO
                </div>
                
                <table class="ricevuta-table" style="width: 100%; border-collapse: collapse; margin: 20px 0;">
                    <thead>
                        <tr style="background-color: #003d7a;">
                            <th style="color: white; padding: 10px; text-align: left; border: 1px solid #000;">DESCRIZIONE COMPENSO</th>
                            <th style="color: white; padding: 10px; text-align: center; border: 1px solid #000;">IMPORTO</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td style="border: 1px solid #000; padding: 10px;">COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE</td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${compenso.toFixed(2)} €</td>
                        </tr>
                        <tr style="background-color: #f0f0f0;">
                            <td style="border: 1px solid #000; padding: 10px; text-align: center;"><strong>COMPENSO LORDO</strong></td>
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;"><strong>${compenso.toFixed(2)} €</strong></td>
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
                            <td style="border: 1px solid #000; padding: 10px; text-align: right;">${rimborsiSpese.toFixed(2)} €</td>
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

// Funzione compatibilità con il vecchio sistema
function generateReceipt(person) {
    const numeroRicevuta = person.numeroProgressivo || getCurrentReceiptNumber(person.codiceFiscale || `${person.nome}_${person.cognome}`);
    const dataRicevuta = getReceiptDate(person.anno, person.mese);
    return createReceiptHTML(person, numeroRicevuta, dataRicevuta);
}

// Esposizione funzioni al contesto globale
window.generateReceipts = generateReceipts;
window.generateReceipt = generateReceipt; // Compatibilità
window.createReceiptHTML = createReceiptHTML;
window.getReceiptDate = getReceiptDate;

console.log('✅ receipts.js caricato - Sistema numerazione unificato + FIX errore rimborsoSpese implementato');
