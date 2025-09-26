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

// Anteprima PDF
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
        const receiptsElements = document.querySelectorAll('.ricevuta');
        const maxPreviews = Math.min(results.length, 5);
        
        for (let index = 0; index < maxPreviews; index++) {
            const person = results[index];
            const receiptElement = receiptsElements[index];
            
            if (!receiptElement) {
                console.error(`Ricevuta ${index} non trovata`);
                continue;
            }
            
            console.log(`Generando anteprima ${index + 1}/${maxPreviews}...`);
            
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
                            ${person.rimborsoSpese > 0 ? `<strong>Rimborso spese:</strong> € ${person.rimborsoSpese.toFixed(2)}<br>` : ''}
                            <strong>Netto a pagare:</strong> € ${(person.compenso * 0.8 + person.rimborsoSpese).toFixed(2)}
                        </div>
                        <div style="margin-top: 10px; padding: 8px; background: #e8f5e9; border-radius: 4px; font-size: 12px; color: #155724;">
                            ✓ PDF pronto per il download
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
                Tutte le ${results.length} ricevute saranno incluse nel download ZIP per mese.
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

// Creazione ZIP con PDF divisi per mese - CON SELEZIONE
async function createZipWithPDFsByMonth() {
    if (results.length === 0) {
        alert('Prima devi eseguire il matching e generare le ricevute!');
        return;
    }

    if (!checkPDFLibraries()) return;

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
    
    let opzioniMesi = 'EXPORT PDF - Seleziona il mese da esportare:\n\n';
    mesiDisponibili.forEach((meseAnno, index) => {
        const [anno, mese] = meseAnno.split('-');
        const nomeCompleto = `${mesiNomi[parseInt(mese)]} ${anno}`;
        const numRicevute = ricevutePerMese[meseAnno].length;
        const rimborsiMese = ricevutePerMese[meseAnno].reduce((sum, item) => sum + item.person.rimborsoSpese, 0);
        opzioniMesi += `${index + 1}. ${nomeCompleto} (${numRicevute} ricevute, €${rimborsiMese.toFixed(2)} rimborsi)\n`;
    });
    
    opzioniMesi += '\n0. Esporta TUTTI i mesi\n';
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

// Esporta PDF per un mese specifico - CON NUMERAZIONE UNIFICATA
async function exportPDFForMonth(meseAnno, ricevuteMese) {
    const [anno, mese] = meseAnno.split('-');
    const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                     'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    const nomeCompleto = `${mesiNomi[parseInt(mese)]} ${anno}`;

    // Calcola totali rimborsi per questo mese
    const totaleRimborsi = ricevuteMese.reduce((sum, item) => sum + item.person.rimborsoSpese, 0);
    const risparmioFiscale = totaleRimborsi * 0.20;
    
    const conferma = confirm(
        `EXPORT PDF - ${nomeCompleto.toUpperCase()}\n\n` +
        `Ricevute da esportare: ${ricevuteMese.length}\n` +
        `Rimborsi spese del mese: € ${totaleRimborsi.toFixed(2)}\n` +
        `Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2)}\n\n` +
        `Procedere con l'export PDF?`
    );
    
    if (!conferma) return;

    const btn = document.getElementById('downloadByMonthBtn');
    btn.innerHTML = '⏳ Generazione ZIP in corso...';
    btn.disabled = true;
    document.getElementById('progressBar').style.display = 'block';

    try {
        console.log(`Inizio generazione ZIP PDF per ${nomeCompleto}...`);
        
        const zip = new JSZip();
        const folder = zip.folder(`Ricevute_${nomeCompleto.replace(' ', '_')}`);
        const receiptsElements = document.querySelectorAll('.ricevuta');

        for (let i = 0; i < ricevuteMese.length; i++) {
            const { person, index } = ricevuteMese[i];
            const receiptElement = receiptsElements[index];
            
            console.log(`Processando ricevuta ${i + 1}/${ricevuteMese.length}: ${person.nome} ${person.cognome}`);
            
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
                
                console.log(`PDF creato: ${fileName} (Numero unificato: ${numeroRicevuta})`);
                
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
        const fileName = `Ricevute_${nomeCompleto.replace(' ', '_')}.zip`;
        
        const downloadArea = document.getElementById('downloadArea');
        downloadArea.innerHTML = `
            <div style="background: #d4edda; border: 1px solid #c3e6cb; border-radius: 5px; padding: 15px; margin: 10px 0; text-align: center;">
                <h4 style="color: #155724; margin-bottom: 10px;">✅ ZIP generato con successo!</h4>
                <p><strong>${nomeCompleto}</strong> - ${ricevuteMese.length} ricevute PDF con numerazione corretta</p>
                <a href="${url}" download="${fileName}" 
                   style="background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-size: 16px;">
                   📥 Scarica ${fileName}
                </a>
            </div>
        `;
        
        console.log(`ZIP PDF per ${nomeCompleto} generato con successo!`);
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

// Esporta tutti i mesi
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
        `EXPORT PDF - TUTTI I MESI\n\n` +
        `Mesi da esportare: ${Object.keys(ricevutePerMese).length}\n` +
        `Ricevute totali: ${totalRicevute}\n` +
        `Rimborsi spese totali: € ${totalRimborsi.toFixed(2)}\n` +
        `Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2)}\n\n` +
        `Verranno generati ${Object.keys(ricevutePerMese).length} file ZIP separati.\n` +
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
    
    alert(`✅ Completato! Generati ${filesCreated} file ZIP con numerazione cronologica corretta.`);
}

// Creazione ZIP con PDF - MANTIENE COMPATIBILITA'
async function createZipWithPDFs() {
    console.log('Redirect a export per mese...');
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
        // Ottimizzazioni per memory
        removeContainer: false,
        onclone: function(clonedDoc) {
            // Rimuovi elementi non necessari dal clone
            const images = clonedDoc.querySelectorAll('img');
            images.forEach(img => {
                if (img.src.startsWith('data:')) {
                    // Mantieni immagini data: ma ottimizza
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
        // Stima approssimativa della memoria disponibile
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
        return { available: true }; // Assumi disponibile se non controllabile
    }
}

// Utility per cleanup dopo generazione PDF
function cleanupAfterPDF() {
    try {
        // Forza garbage collection (se disponibile)
        if (window.gc) {
            window.gc();
        }
        
        // Cleanup manuale di variabili pesanti
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

// Esposizione funzioni al contesto globale
window.generatePDFPreviews = generatePDFPreviews;
window.createZipWithPDFs = createZipWithPDFs; // Mantiene compatibilità
window.createZipWithPDFsByMonth = createZipWithPDFsByMonth;
window.exportPDFForMonth = exportPDFForMonth;
window.exportAllMonthsPDF = exportAllMonthsPDF;
window.checkPDFLibraries = checkPDFLibraries;
window.handlePDFError = handlePDFError;
window.optimizeCanvas = optimizeCanvas;
window.checkMemoryAvailability = checkMemoryAvailability;
window.cleanupAfterPDF = cleanupAfterPDF;

// Debug - verifica che le funzioni siano esposte
console.log('🔍 pdf-export.js CARICATO - Export per mese con utilities avanzate...');
console.log('Funzioni esposte:', {
    generatePDFPreviews: typeof window.generatePDFPreviews,
    createZipWithPDFs: typeof window.createZipWithPDFs,
    createZipWithPDFsByMonth: typeof window.createZipWithPDFsByMonth,
    exportPDFForMonth: typeof window.exportPDFForMonth,
    handlePDFError: typeof window.handlePDFError,
    optimizeCanvas: typeof window.optimizeCanvas,
    checkMemoryAvailability: typeof window.checkMemoryAvailability
});

if (typeof window.generatePDFPreviews !== 'function') {
    console.error('❌ ERRORE: generatePDFPreviews non è esposta correttamente!');
} else {
    console.log('✅ generatePDFPreviews esposta correttamente');
}

if (typeof window.createZipWithPDFsByMonth !== 'function') {
    console.error('❌ ERRORE: createZipWithPDFsByMonth non è esposta correttamente!');
} else {
    console.log('✅ createZipWithPDFsByMonth esposta correttamente - Export per mese con utilities implementato');
}

console.log('✅ pdf-export.js caricato completamente - Export per mese con gestione errori avanzata e ottimizzazioni');
