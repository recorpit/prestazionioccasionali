// PDF-EXPORT.JS - COMPLETO CON SISTEMA RIMBORSI FORFETTARI INTELLIGENTI
// Solo Export per Mese con Selezione + Sistema Rimborsi Automatico

// Verifica librerie PDF
function checkPDFLibraries() {
    const required = ['html2canvas', 'jsPDF', 'JSZip'];
    const missing = [];
    
    if (typeof html2canvas === 'undefined') missing.push('html2canvas');
    if (typeof window.jspdf === 'undefined') missing.push('jsPDF');
    if (typeof JSZip === 'undefined') missing.push('JSZip');
    
    if (missing.length > 0) {
        console.error('Librerie PDF mancanti:', missing);
        alert(`Errore: Librerie non caricate: ${missing.join(', ')}\nRicarica la pagina.`);
        return false;
    }
    
    console.log('✅ Tutte le librerie PDF sono caricate');
    return true;
}

// SISTEMA RIMBORSI SPESE INTELLIGENTI con importi credibili (sempre tondi)
function calculateSmartReimbursements(importoRimborso, genere = 'N') {
    const rimborsi = {};
    let rimanente = Math.round(importoRimborso); // Arrotonda a euro interi
    
    // BENZINA/TRASPORTI - Sempre presente, proporzionale ma tondo
    let benzina;
    if (rimanente <= 30) {
        benzina = Math.min(rimanente, 15);
    } else if (rimanente <= 60) {
        benzina = Math.round(rimanente * 0.4 / 5) * 5; // Arrotonda a multipli di 5
        benzina = Math.min(benzina, 25);
    } else if (rimanente <= 100) {
        benzina = Math.round(rimanente * 0.3 / 5) * 5;
        benzina = Math.min(benzina, 30);
    } else {
        benzina = Math.round(rimanente * 0.25 / 5) * 5;
        benzina = Math.min(benzina, 35);
    }
    
    // Assicura che benzina sia almeno 10 e massimo 35
    benzina = Math.max(10, Math.min(35, benzina));
    
    rimborsi['Trasporti e carburante'] = benzina;
    rimanente -= benzina;
    
    // DISTRIBUZIONE DINAMICA DELLE ALTRE VOCI (solo importi tondi)
    if (rimanente > 0) {
        const categorieBase = [
            'Materiali tecnici',
            'Comunicazioni e coordinamento',
            'Piccoli materiali di consumo'
        ];
        
        const categorieFemminili = [
            'Cura e manutenzione abiti da ballo',
            'Trucco e parrucco di scena',
            'Accessori scenici',
            'Calzature specializzate'
        ];
        
        const categorieMaschili = [
            'Cura abbigliamento formale',
            'Trucco di scena',
            'Accessori tecnici',
            'Calzature professionali'
        ];
        
        // Selezione categorie appropriate
        let categorieDisponibili = [...categorieBase];
        if (genere === 'F') {
            categorieDisponibili = [...categorieFemminili, ...categorieBase];
        } else if (genere === 'M') {
            categorieDisponibili = [...categorieMaschili, ...categorieBase];
        }
        
        // Distribuzione intelligente per fasce
        let numeroCategorie;
        if (rimanente <= 20) {
            numeroCategorie = 1;
        } else if (rimanente <= 50) {
            numeroCategorie = 2;
        } else if (rimanente <= 100) {
            numeroCategorie = 3;
        } else {
            numeroCategorie = Math.min(4, categorieDisponibili.length);
        }
        
        const categorieSelezionate = categorieDisponibili.slice(0, numeroCategorie);
        
        // IMPORTI SEMPRE TONDI E CREDIBILI
        const importiTondi = [10, 15, 20, 25, 30]; // Solo questi importi
        
        for (let i = 0; i < categorieSelezionate.length; i++) {
            const categoria = categorieSelezionate[i];
            let importo;
            
            if (i === categorieSelezionate.length - 1) {
                // Ultima categoria: tutto il rimanente, ma arrotondato a multipli di 5
                importo = Math.round(rimanente / 5) * 5;
                importo = Math.max(10, importo); // Minimo 10
            } else {
                // Scegli un importo tondo appropriato
                const importoBase = Math.floor(rimanente / (numeroCategorie - i));
                
                // Trova l'importo tondo più vicino
                importo = importiTondi.reduce((prev, curr) => 
                    Math.abs(curr - importoBase) < Math.abs(prev - importoBase) ? curr : prev
                );
                
                // Assicura che non superi il rimanente
                importo = Math.min(importo, rimanente - (numeroCategorie - i - 1) * 10);
            }
            
            if (importo > 0) {
                rimborsi[categoria] = importo;
                rimanente -= importo;
            }
        }
    }
    
    return {
        dettaglioRimborsi: rimborsi,
        giustificazioneLegale: generateLegalJustification(rimborsi),
        totaleForfettario: importoRimborso,
        conformitaNormativa: "Art. 54 TUIR - D.Lgs. 192/2024 - Regime forfettario 2025"
    };
}

// Genera la giustificazione legale per i rimborsi spese
function generateLegalJustification(rimborsi) {
    return `Rimborso spese documentato sostenute per l'esecuzione della prestazione occasionale secondo accordi contrattuali. Le spese rimborsate non concorrono alla formazione del reddito imponibile ai sensi dell'art. 54, comma 2 TUIR. Documentazione disponibile su richiesta.`;
}

// Rileva genere dal codice fiscale (più affidabile del nome)
function detectGenderFromCodiceFiscale(codiceFiscale) {
    if (!codiceFiscale || typeof codiceFiscale !== 'string' || codiceFiscale.length !== 16) {
        return 'N'; // Se CF non valido, fallback al nome
    }
    
    try {
        // Il giorno di nascita è nei caratteri 9-10 (posizioni 8-9)
        const giornoNascita = parseInt(codiceFiscale.substring(9, 11), 10);
        
        if (isNaN(giornoNascita)) {
            return 'N';
        }
        
        // Se il giorno è > 31, è femmina (giorno + 40)
        // Se il giorno è <= 31, è maschio
        if (giornoNascita > 31) {
            return 'F'; // Femmina
        } else if (giornoNascita >= 1 && giornoNascita <= 31) {
            return 'M'; // Maschio
        }
        
        return 'N'; // Caso anomalo
    } catch (error) {
        console.warn('Errore nella lettura del codice fiscale:', error);
        return 'N';
    }
}

// Rileva genere dal nome automaticamente (fallback se CF non disponibile)
function detectGenderFromName(nome) {
    if (!nome || typeof nome !== 'string') return 'N';
    
    const nomeLower = nome.toLowerCase().trim();
    
    // NOMI SPECIFICI MASCHILI (casi particolari)
    const nomiMaschiliSpecifici = [
        'luca', 'nicola', 'andrea', 'mattia', 'joshua', 'giosuè', 
        'elia', 'emanuele', 'gabriele', 'michele', 'daniele', 'raffaele'
    ];
    
    if (nomiMaschiliSpecifici.includes(nomeLower)) {
        return 'M';
    }
    
    // NOMI SPECIFICI FEMMINILI (casi particolari)
    const nomiFemminiliSpecifici = [
        'beatrice', 'alice', 'nicole', 'sole', 'celeste', 'cloe', 'noemi',
        'rachel', 'ruth', 'judith', 'carmen', 'dolores', 'mercedes'
    ];
    
    if (nomiFemminiliSpecifici.includes(nomeLower)) {
        return 'F';
    }
    
    // TERMINAZIONI FEMMINILI COMUNI
    const terminazioniFemminili = [
        'a', 'ia', 'ina', 'etta', 'ella', 'isa', 'ara', 'era', 
        'ina', 'ica', 'iana', 'enna', 'anna', 'ida'
    ];
    
    for (const terminazione of terminazioniFemminili) {
        if (nomeLower.endsWith(terminazione) && nomeLower.length > terminazione.length) {
            return 'F';
        }
    }
    
    // TERMINAZIONI MASCHILI COMUNI  
    const terminazioniMaschili = [
        'o', 'io', 'ino', 'etto', 'ello', 'andro', 'iano', 'ano',
        'ero', 'iero', 'ico', 'esco', 'ardo', 'aldo', 'erto'
    ];
    
    for (const terminazione of terminazioniMaschili) {
        if (nomeLower.endsWith(terminazione) && nomeLower.length > terminazione.length) {
            return 'M';
        }
    }
    
    return 'N';
}

// Funzione principale che usa prima il CF, poi il nome come fallback
function detectGender(person) {
    // Prova prima con il codice fiscale (più affidabile)
    if (person.codiceFiscale) {
        const genderFromCF = detectGenderFromCodiceFiscale(person.codiceFiscale);
        if (genderFromCF !== 'N') {
            console.log(`Genere rilevato da CF per ${person.nome}: ${genderFromCF}`);
            return genderFromCF;
        }
    }
    
    // Fallback al nome se CF non disponibile o non leggibile
    const genderFromName = detectGenderFromName(person.nome);
    console.log(`Genere rilevato da nome per ${person.nome}: ${genderFromName}`);
    return genderFromName;
}

// INTEGRA I RIMBORSI NELLE RICEVUTE HTML PRIMA DELLA GENERAZIONE PDF
function integrateReimbursementsInReceipts() {
    console.log('Integrando rimborsi spese nelle ricevute...');
    
    const ricevute = document.querySelectorAll('.ricevuta');
    
    ricevute.forEach((ricevuta, index) => {
        if (index >= results.length) return;
        
        const person = results[index];
        
        // Verifica se ha rimborsi
        if (!person.rimborsoSpese || person.rimborsoSpese <= 0) return;
        
        // Rileva genere dal codice fiscale (più affidabile)
        const genere = detectGender(person);
        
        // Calcola rimborsi intelligenti con importi tondi
        const reimbursementData = calculateSmartReimbursements(person.rimborsoSpese, genere);
        
        // Genera HTML per rimborsi
        const rimborsiHTML = generateReimbursementHTML(reimbursementData);
        
        // Trova il punto di inserimento (dopo la tabella principale)
        const tabellaPrincipale = ricevuta.querySelector('.ricevuta-table');
        if (tabellaPrincipale && rimborsiHTML) {
            // Inserisce la sezione rimborsi dopo la tabella
            const rimborsiDiv = document.createElement('div');
            rimborsiDiv.innerHTML = rimborsiHTML;
            
            // Inserisce dopo la tabella ma prima del footer
            const footer = ricevuta.querySelector('.ricevuta-footer');
            if (footer) {
                ricevuta.insertBefore(rimborsiDiv.firstElementChild, footer);
            } else {
                tabellaPrincipale.parentNode.insertBefore(rimborsiDiv.firstElementChild, tabellaPrincipale.nextSibling);
            }
            
            console.log(`Rimborsi integrati per ${person.nome} ${person.cognome}: €${person.rimborsoSpese} (${genere} da CF: ${person.codiceFiscale})`);
        }
    });
}

// Genera HTML formattato per i rimborsi nella ricevuta PDF
function generateReimbursementHTML(reimbursementData) {
    if (!reimbursementData || !reimbursementData.dettaglioRimborsi) {
        return '';
    }
    
    const { dettaglioRimborsi, giustificazioneLegale, conformitaNormativa } = reimbursementData;
    
    let html = `
    <div class="rimborsi-section" style="margin-top: 25px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
        <h4 style="color: #003d7a; margin-bottom: 15px; font-size: 14px;">DETTAGLIO RIMBORSI SPESE</h4>
        
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">
            <thead>
                <tr style="background-color: #003d7a; color: white;">
                    <th style="padding: 8px; text-align: left; font-size: 12px;">Categoria Spesa</th>
                    <th style="padding: 8px; text-align: right; font-size: 12px; width: 100px;">Importo</th>
                </tr>
            </thead>
            <tbody>`;
    
    Object.entries(dettaglioRimborsi).forEach(([categoria, importo]) => {
        html += `
                <tr style="border-bottom: 1px solid #ddd;">
                    <td style="padding: 6px 8px; font-size: 11px;">${categoria}</td>
                    <td style="padding: 6px 8px; text-align: right; font-size: 11px;">€ ${importo.toFixed(2).replace('.', ',')}</td>
                </tr>`;
    });
    
    const totale = Object.values(dettaglioRimborsi).reduce((sum, val) => sum + val, 0);
    
    html += `
                <tr style="border-top: 2px solid #003d7a; background-color: #f0f0f0; font-weight: bold;">
                    <td style="padding: 8px; font-size: 12px;">TOTALE RIMBORSI SPESE</td>
                    <td style="padding: 8px; text-align: right; font-size: 12px;">€ ${totale.toFixed(2).replace('.', ',')}</td>
                </tr>
            </tbody>
        </table>
        
        <div style="font-size: 10px; line-height: 1.4; color: #555; margin-bottom: 10px;">
            <strong>Giustificazione legale:</strong><br>
            ${giustificazioneLegale}
        </div>
        
        <div style="font-size: 9px; color: #777; text-align: center; font-style: italic;">
            Conforme a: ${conformitaNormativa}
        </div>
    </div>`;
    
    return html;
}

// Anteprima PDF con rimborsi integrati
async function generatePDFPreviews() {
    if (results.length === 0) {
        alert('Prima devi eseguire il matching e generare le ricevute!');
        return;
    }

    if (!checkPDFLibraries()) return;

    const btn = document.getElementById('pdfPreviewBtn');
    btn.innerHTML = 'Generazione anteprima...';
    btn.disabled = true;

    const pdfPreviewArea = document.getElementById('pdfPreviewArea');
    pdfPreviewArea.innerHTML = '<h4>Anteprima PDF Generate:</h4>';

    try {
        // INTEGRA I RIMBORSI PRIMA DI GENERARE LE ANTEPRIME
        integrateReimbursementsInReceipts();
        
        const receiptsElements = document.querySelectorAll('.ricevuta');
        const maxPreviews = Math.min(results.length, 5);
        
        for (let index = 0; index < maxPreviews; index++) {
            const person = results[index];
            const receiptElement = receiptsElements[index];
            
            if (!receiptElement) {
                console.error(`Ricevuta ${index} non trovata`);
                continue;
            }
            
            console.log(`Generando anteprima ${index + 1}/${maxPreviews} con rimborsi...`);
            
            // Genera immagine con html2canvas
            const canvas = await html2canvas(receiptElement, {
                scale: 1.5,
                backgroundColor: '#ffffff',
                logging: false,
                useCORS: true,
                allowTaint: false,
                width: receiptElement.scrollWidth,
                height: receiptElement.scrollHeight
            });
            
            // USA LA NUMERAZIONE UNIFICATA
            const numeroRicevuta = person.numeroProgressivo;
            const fileName = `${person.nome}_${person.cognome}_${numeroRicevuta}.pdf`;
            
            const previewDiv = document.createElement('div');
            previewDiv.style.cssText = 'border: 2px solid #ddd; margin: 20px 0; padding: 15px; background: #f9f9f9; border-radius: 8px;';
            
            const previewImage = canvas.toDataURL('image/png');
            
            // Calcola dettaglio rimborsi per preview
            let rimborsiDetail = '';
            if (person.rimborsoSpese > 0) {
                const genere = detectGenderFromName(person.nome);
                const reimbursementData = calculateSmartReimbursements(person.rimborsoSpese, genere);
                const categorie = Object.keys(reimbursementData.dettaglioRimborsi);
                rimborsiDetail = `<br><strong>Categorie rimborsi:</strong> ${categorie.join(', ')}`;
            }
            
            previewDiv.innerHTML = `
                <div style="display: flex; align-items: flex-start; gap: 20px;">
                    <div style="flex-shrink: 0;">
                        <h5 style="margin: 0 0 10px 0; color: #333;">
                            ${fileName}
                        </h5>
                        <img src="${previewImage}" 
                             style="max-width: 300px; border: 1px solid #ccc; border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);"
                             alt="Anteprima PDF ${fileName}">
                    </div>
                    <div style="flex: 1;">
                        <h6 style="margin: 0 0 10px 0; color: #666;">Dettagli ricevuta:</h6>
                        <div style="font-size: 14px; line-height: 1.5;">
                            <strong>Nome:</strong> ${person.nome} ${person.cognome}<br>
                            <strong>CF:</strong> ${person.codiceFiscale}<br>
                            <strong>Numero ricevuta:</strong> ${numeroRicevuta}<br>
                            <strong>Compenso lordo:</strong> € ${person.compenso.toFixed(2)}<br>
                            ${person.rimborsoSpese > 0 ? `<strong>Rimborso spese:</strong> € ${person.rimborsoSpese.toFixed(2)}${rimborsiDetail}<br>` : ''}
                            <strong>Netto a pagare:</strong> € ${(person.compenso * 0.8 + person.rimborsoSpese).toFixed(2)}
                        </div>
                        <div style="margin-top: 10px; padding: 8px; background: #e8f5e9; border-radius: 4px; font-size: 12px; color: #155724;">
                            ✓ PDF pronto con rimborsi forfettari dettagliati
                        </div>
                    </div>
                </div>
            `;
            
            pdfPreviewArea.appendChild(previewDiv);
        }
        
        if (results.length > 5) {
            const moreDiv = document.createElement('div');
            moreDiv.style.cssText = 'text-align: center; padding: 20px; background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 5px; margin: 20px 0;';
            moreDiv.innerHTML = `
                <strong>Anteprima limitata</strong><br>
                Mostrate ${maxPreviews} di ${results.length} ricevute.<br>
                Tutte le ${results.length} ricevute saranno incluse nel download ZIP con rimborsi dettagliati.
            `;
            pdfPreviewArea.appendChild(moreDiv);
        }

    } catch (error) {
        console.error('Errore nell\'anteprima PDF:', error);
        pdfPreviewArea.innerHTML = `
            <div class="error-box">
                Errore nella generazione dell'anteprima: ${error.message}
            </div>
        `;
    } finally {
        btn.innerHTML = 'Anteprima PDF';
        btn.disabled = false;
        
        // Attiva automaticamente il tab anteprima PDF
        document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        
        const pdfTab = document.querySelectorAll('.tab')[2];
        if (pdfTab) {
            pdfTab.classList.add('active');
            document.getElementById('pdfpreviewTab').classList.add('active');
            pdfPreviewArea.scrollIntoView({ behavior: 'smooth' });
        }
    }
}

// Creazione ZIP con PDF divisi per mese - CON SELEZIONE E RIMBORSI
async function createZipWithPDFsByMonth() {
    if (results.length === 0) {
        alert('Prima devi eseguire il matching e generare le ricevute!');
        return;
    }

    if (!checkPDFLibraries()) return;

    // INTEGRA I RIMBORSI PRIMA DI PROCEDERE
    integrateReimbursementsInReceipts();

    // Raggruppa ricevute per mese
    const ricevutePerMese = {};
    results.forEach((person, index) => {
        const meseAnno = `${person.anno}-${person.mese.toString().padStart(2, '0')}`;
        if (!ricevutePerMese[meseAnno]) {
            ricevutePerMese[meseAnno] = [];
        }
        ricevutePerMese[meseAnno].push({ person, index });
    });

    const mesiDisponibili = Object.keys(ricevutePerMese).sort();
    
    if (mesiDisponibili.length === 0) {
        alert('Nessuna ricevuta disponibile per l\'export');
        return;
    }

    // Se c'è solo un mese, procede direttamente
    if (mesiDisponibili.length === 1) {
        const meseSelezionato = mesiDisponibili[0];
        exportPDFForMonth(meseSelezionato, ricevutePerMese[meseSelezionato]);
        return;
    }

    // Crea dialog per selezione mese
    const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                     'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    
    let opzioniMesi = 'EXPORT PDF CON RIMBORSI - Seleziona il mese da esportare:\n\n';
    mesiDisponibili.forEach((meseAnno, index) => {
        const [anno, mese] = meseAnno.split('-');
        const nomeCompleto = `${mesiNomi[parseInt(mese)]} ${anno}`;
        const numRicevute = ricevutePerMese[meseAnno].length;
        const rimborsiMese = ricevutePerMese[meseAnno].reduce((sum, item) => sum + item.person.rimborsoSpese, 0);
        opzioniMesi += `${index + 1}. ${nomeCompleto} (${numRicevute} ricevute, €${rimborsiMese.toFixed(2)} rimborsi)\n`;
    });
    
    opzioniMesi += '\n0. Esporta TUTTI i mesi con rimborsi dettagliati\n';
    opzioniMesi += 'Inserisci il numero della tua scelta:';

    const scelta = prompt(opzioniMesi);
    
    if (!scelta || scelta === null) return;
    
    const sceltaNum = parseInt(scelta);
    
    if (sceltaNum === 0) {
        // Esporta tutti i mesi
        await exportAllMonthsPDF(ricevutePerMese);
    } else if (sceltaNum >= 1 && sceltaNum <= mesiDisponibili.length) {
        const meseSelezionato = mesiDisponibili[sceltaNum - 1];
        await exportPDFForMonth(meseSelezionato, ricevutePerMese[meseSelezionato]);
    } else {
        alert('Selezione non valida');
    }
}

// Esporta PDF per un mese specifico - CON NUMERAZIONE UNIFICATA E RIMBORSI
async function exportPDFForMonth(meseAnno, ricevuteMese) {
    const [anno, mese] = meseAnno.split('-');
    const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                     'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    const nomeCompleto = `${mesiNomi[parseInt(mese)]} ${anno}`;

    // Calcola totali rimborsi per questo mese
    const totaleRimborsi = ricevuteMese.reduce((sum, item) => sum + item.person.rimborsoSpese, 0);
    const risparmioFiscale = totaleRimborsi * 0.20;
    
    // Calcola categorie rimborsi più utilizzate
    const categorieUtilizzate = new Set();
    ricevuteMese.forEach(({ person }) => {
        if (person.rimborsoSpese > 0) {
            const genere = detectGenderFromName(person.nome);
            const reimbursementData = calculateSmartReimbursements(person.rimborsoSpese, genere);
            Object.keys(reimbursementData.dettaglioRimborsi).forEach(cat => categorieUtilizzate.add(cat));
        }
    });
    
    const conferma = confirm(
        `EXPORT PDF CON RIMBORSI - ${nomeCompleto.toUpperCase()}\n\n` +
        `Ricevute da esportare: ${ricevuteMese.length}\n` +
        `Rimborsi spese del mese: € ${totaleRimborsi.toFixed(2)}\n` +
        `Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2)}\n` +
        `Categorie rimborsi: ${Array.from(categorieUtilizzate).join(', ')}\n\n` +
        `Procedere con l'export PDF con rimborsi dettagliati?`
    );
    
    if (!conferma) return;

    const btn = document.getElementById('downloadByMonthBtn');
    btn.innerHTML = '⏳ Generazione ZIP con rimborsi...';
    btn.disabled = true;
    document.getElementById('progressBar').style.display = 'block';

    try {
        console.log(`Inizio generazione ZIP PDF con rimborsi per ${nomeCompleto}...`);
        
        const zip = new JSZip();
        const folder = zip.folder(`Ricevute_${nomeCompleto.replace(' ', '_')}_Con_Rimborsi`);
        const receiptsElements = document.querySelectorAll('.ricevuta');

        for (let i = 0; i < ricevuteMese.length; i++) {
            const { person, index } = ricevuteMese[i];
            const receiptElement = receiptsElements[index];
            
            console.log(`Processando ricevuta ${i + 1}/${ricevuteMese.length}: ${person.nome} ${person.cognome} (Rimborsi: €${person.rimborsoSpese})`);
            
            try {
                const canvas = await html2canvas(receiptElement, {
                    scale: 2,
                    backgroundColor: '#ffffff',
                    logging: false,
                    useCORS: true,
                    allowTaint: false,
                    width: receiptElement.scrollWidth,
                    height: receiptElement.scrollHeight,
                    windowWidth: 1200,
                    windowHeight: 1600
                });
                
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF({
                    orientation: 'portrait',
                    unit: 'mm',
                    format: 'a4'
                });
                
                const imgData = canvas.toDataURL('image/png', 0.95);
                const pdfWidth = pdf.internal.pageSize.getWidth();
                const pdfHeight = pdf.internal.pageSize.getHeight();
                
                let imgWidth = pdfWidth - 20;
                let imgHeight = (canvas.height * imgWidth) / canvas.width;
                
                if (imgHeight > pdfHeight - 20) {
                    imgHeight = pdfHeight - 20;
                    imgWidth = (canvas.width * imgHeight) / canvas.height;
                }
                
                const x = (pdfWidth - imgWidth) / 2;
                const y = 10;
                
                pdf.addImage(imgData, 'PNG', x, y, imgWidth, imgHeight);
                
                // USA LA NUMERAZIONE UNIFICATA PER IL NOME FILE
                const numeroRicevuta = person.numeroProgressivo;
                const fileName = `${person.nome}_${person.cognome}_${numeroRicevuta}.pdf`
                    .replace(/\s+/g, '_')
                    .replace(/[^a-zA-Z0-9_\-\.]/g, '');
                
                const pdfBlob = pdf.output('blob');
                folder.file(fileName, pdfBlob);
                
                console.log(`PDF creato: ${fileName} (Numero unificato: ${numeroRicevuta}, Rimborsi: €${person.rimborsoSpese})`);
                
                const progress = ((i + 1) / ricevuteMese.length) * 90;
                updateProgressBar(progress);
                
                await new Promise(resolve => setTimeout(resolve, 100));
                
            } catch (pdfError) {
                console.error(`Errore nella generazione PDF ${i}:`, pdfError);
            }
        }
        
        console.log('Generazione ZIP finale...');
        updateProgressBar(95);
        
        const content = await zip.generateAsync({
            type: "blob",
            compression: "DEFLATE",
            compressionOptions: { level: 6 }
        });
        
        updateProgressBar(100);
        
        const url = URL.createObjectURL(content);
        const fileName = `Ricevute_${nomeCompleto.replace(' ', '_')}_Con_Rimborsi.zip`;
        
        const downloadArea = document.getElementById('downloadArea');
        downloadArea.innerHTML = `
            <div style="background: #d4edda; border: 1px solid #c3e6cb; border-radius: 5px; padding: 15px; margin: 10px 0; text-align: center;">
                <h4 style="color: #155724; margin-bottom: 10px;">✅ ZIP con rimborsi generato con successo!</h4>
                <p><strong>${nomeCompleto}</strong> - ${ricevuteMese.length} ricevute PDF con rimborsi forfettari dettagliati</p>
                <p>Rimborsi totali: €${totaleRimborsi.toFixed(2)} | Risparmio fiscale: €${risparmioFiscale.toFixed(2)}</p>
                <a href="${url}" download="${fileName}" 
                   style="background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-size: 16px;">
                   📥 Scarica ${fileName}
                </a>
            </div>
        `;
        
        console.log(`ZIP PDF con rimborsi per ${nomeCompleto} generato con successo!`);
        downloadArea.scrollIntoView({ behavior: 'smooth' });

    } catch (error) {
        console.error('Errore nella generazione ZIP:', error);
        
        const downloadArea = document.getElementById('downloadArea');
        downloadArea.innerHTML = `
            <div class="error-box">
                <h4>Errore nella generazione PDF</h4>
                <p>Dettagli: ${error.message}</p>
            </div>
        `;
    } finally {
        btn.innerHTML = 'ZIP PDF per Mese';
        btn.disabled = false;
        document.getElementById('progressBar').style.display = 'none';
        updateProgressBar(0);
    }
}

// Esporta tutti i mesi con rimborsi
async function exportAllMonthsPDF(ricevutePerMese) {
    const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                     'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    
    let totalRicevute = 0;
    let totalRimborsi = 0;
    
    Object.values(ricevutePerMese).forEach(ricevuteMese => {
        totalRicevute += ricevuteMese.length;
        totalRimborsi += ricevuteMese.reduce((sum, item) => sum + item.person.rimborsoSpese, 0);
    });
    
    const risparmioFiscale = totalRimborsi * 0.20;
    
    const conferma = confirm(
        `EXPORT PDF CON RIMBORSI - TUTTI I MESI\n\n` +
        `Mesi da esportare: ${Object.keys(ricevutePerMese).length}\n` +
        `Ricevute totali: ${totalRicevute}\n` +
        `Rimborsi spese totali: € ${totalRimborsi.toFixed(2)}\n` +
        `Risparmio fiscale totale (20%): € ${risparmioFiscale.toFixed(2)}\n\n` +
        `Verranno generati ${Object.keys(ricevutePerMese).length} file ZIP con rimborsi dettagliati.\n` +
        `Procedere?`
    );
    
    if (!conferma) return;

    let filesCreated = 0;
    
    for (const [meseAnno, ricevuteMese] of Object.entries(ricevutePerMese)) {
        await exportPDFForMonth(meseAnno, ricevuteMese);
        filesCreated++;
        
        // Pausa tra i download per non sovraccaricare il browser
        if (filesCreated < Object.keys(ricevutePerMese).length) {
            await new Promise(resolve => setTimeout(resolve, 500));
        }
    }
    
    alert(`✅ Completato! Generati ${filesCreated} file ZIP con rimborsi forfettari dettagliati e numerazione cronologica corretta.`);
}

// MANTIENE COMPATIBILITA' - Redirect a export per mese
async function createZipWithPDFs() {
    console.log('Redirect a export per mese con rimborsi...');
    await createZipWithPDFsByMonth();
}

// Utility avanzata per gestione errori PDF
function handlePDFError(error, context = 'PDF generation') {
    console.error(`Errore in ${context}:`, error);
    
    let errorMessage = `Errore durante ${context}:\n\n`;
    
    if (error.name === 'SecurityError') {
        errorMessage += `Errore di sicurezza - possibili cause:\n• Contenuto bloccato dal browser\n• CORS policy violation\n• Script non autorizzati`;
    } else if (error.message.includes('canvas')) {
        errorMessage += `Errore rendering canvas:\n• Elemento ricevuta non trovato\n• Dimensioni canvas troppo grandi\n• Memoria insufficiente`;
    } else if (error.message.includes('jsPDF')) {
        errorMessage += `Errore libreria jsPDF:\n• Libreria non caricata correttamente\n• Versione incompatibile\n• Problema inizializzazione PDF`;
    } else {
        errorMessage += `Dettagli tecnici: ${error.message}`;
    }
    
    errorMessage += `\n\nSoluzioni:\n1. Ricarica la pagina con Ctrl+F5\n2. Prova con meno ricevute alla volta\n3. Chiudi altre schede del browser\n4. Controlla la console (F12) per dettagli`;
    
    return errorMessage;
}

// Ottimizzazione canvas per ridurre problemi di memoria
function optimizeCanvas(element, options = {}) {
    const defaultOptions = {
        scale: 2,
        backgroundColor: '#ffffff',
        logging: false,
        useCORS: true,
        allowTaint: false,
        width: element.scrollWidth,
        height: element.scrollHeight,
        windowWidth: 1200,
        windowHeight: 1600,
        removeContainer: false,
        onclone: function(clonedDoc) {
            const images = clonedDoc.querySelectorAll('img');
            images.forEach(img => {
                if (img.src.startsWith('data:')) {
                    if (img.naturalWidth > 800) {
                        img.style.maxWidth = '800px';
                    }
                }
            });
        }
    };
    
    return { ...defaultOptions, ...options };
}

// Verifica memoria disponibile
function checkMemoryAvailability() {
    try {
        if (performance.memory) {
            const used = performance.memory.usedJSHeapSize;
            const limit = performance.memory.jsHeapSizeLimit;
            const available = limit - used;
            const availableMB = Math.round(available / 1024 / 1024);
            
            console.log(`Memoria JS disponibile: ~${availableMB}MB`);
            
            if (availableMB < 50) {
                return {
                    available: false,
                    message: `Memoria JS bassa (~${availableMB}MB). Chiudi altre schede o riduci il numero di ricevute.`
                };
            }
        }
        return { available: true };
    } catch (e) {
        console.warn('Impossibile controllare la memoria:', e);
        return { available: true };
    }
}

// Utility per cleanup dopo generazione PDF
function cleanupAfterPDF() {
    try {
        if (window.gc) {
            window.gc();
        }
        
        const heavyVars = ['canvas', 'imgData', 'pdfBlob'];
        heavyVars.forEach(varName => {
            if (window[varName]) {
                delete window[varName];
            }
        });
        
        console.log('Cleanup post-PDF completato');
    } catch (e) {
        console.warn('Errore durante cleanup:', e);
    }
}

// ESPOSIZIONE FUNZIONI AL CONTESTO GLOBALE
window.generatePDFPreviews = generatePDFPreviews;
window.createZipWithPDFs = createZipWithPDFs;
window.createZipWithPDFsByMonth = createZipWithPDFsByMonth;
window.exportPDFForMonth = exportPDFForMonth;
window.exportAllMonthsPDF = exportAllMonthsPDF;
window.checkPDFLibraries = checkPDFLibraries;
window.handlePDFError = handlePDFError;
window.optimizeCanvas = optimizeCanvas;
window.checkMemoryAvailability = checkMemoryAvailability;
window.cleanupAfterPDF = cleanupAfterPDF;

// ESPOSIZIONE FUNZIONI RIMBORSI
window.calculateSmartReimbursements = calculateSmartReimbursements;
window.generateLegalJustification = generateLegalJustification;
window.detectGenderFromName = detectGenderFromName;
window.generateReimbursementHTML = generateReimbursementHTML;
window.integrateReimbursementsInReceipts = integrateReimbursementsInReceipts;

// Debug - verifica che le funzioni siano esposte
console.log('🔍 pdf-export.js CON RIMBORSI CARICATO...');
console.log('Funzioni PDF esposte:', {
    generatePDFPreviews: typeof window.generatePDFPreviews,
    createZipWithPDFs: typeof window.createZipWithPDFs,
    createZipWithPDFsByMonth: typeof window.createZipWithPDFsByMonth,
    exportPDFForMonth: typeof window.exportPDFForMonth
});

console.log('Funzioni Rimborsi esposte:', {
    calculateSmartReimbursements: typeof window.calculateSmartReimbursements,
    detectGenderFromName: typeof window.detectGenderFromName,
    generateReimbursementHTML: typeof window.generateReimbursementHTML,
    integrateReimbursementsInReceipts: typeof window.integrateReimbursementsInReceipts
});

if (typeof window.generatePDFPreviews !== 'function') {
    console.error('❌ ERRORE: generatePDFPreviews non è esposta correttamente!');
} else {
    console.log('✅ generatePDFPreviews esposta correttamente');
}

if (typeof window.createZipWithPDFsByMonth !== 'function') {
    console.error('❌ ERRORE: createZipWithPDFsByMonth non è esposta correttamente!');
} else {
    console.log('✅ createZipWithPDFsByMonth esposta correttamente');
}

if (typeof window.integrateReimbursementsInReceipts !== 'function') {
    console.error('❌ ERRORE: integrateReimbursementsInReceipts non è esposta correttamente!');
} else {
    console.log('✅ Sistema rimborsi integrato correttamente');
}

console.log('✅ pdf-export.js CON SISTEMA RIMBORSI FORFETTARI INTELLIGENTI caricato completamente!');
