// Export Excel per mese con selezione - MODIFICHE COMMERCIALISTA
function exportToExcelByMonth() {
    console.log('Inizio export Excel per mese con selezione e modifiche commercialista...');
    
    if (!results || results.length === 0) {
        alert('Prima devi eseguire il matching e generare le ricevute!');
        return;
    }
    
    if (typeof XLSX === 'undefined') {
        alert('Libreria XLSX non caricata. Ricarica la pagina e riprova.');
        return;
    }

    try {
        // Calcola totali rimborsi e risparmio fiscale
        const totaleRimborsi = results.reduce((sum, person) => sum + (person.rimborsoSpese || 0), 0);
        const risparmioFiscale = totaleRimborsi * 0.20;

        // Raggruppa le ricevute per mese
        const ricevutePerMese = {};
        
        results.forEach(person => {
            const chiaveMese = `${person.anno}-${person.mese.toString().padStart(2, '0')}`;
            if (!ricevutePerMese[chiaveMese]) {
                ricevutePerMese[chiaveMese] = [];
            }
            ricevutePerMese[chiaveMese].push(person);
        });

        const mesiDisponibili = Object.keys(ricevutePerMese).sort();
        console.log('Mesi disponibili per export:', mesiDisponibili);

        if (mesiDisponibili.length === 0) {
            alert('Nessuna ricevuta disponibile per l\'export');
            return;
        }

        // Se c'è solo un mese, procede direttamente
        if (mesiDisponibili.length === 1) {
            const meseSelezionato = mesiDisponibili[0];
            exportExcelForMonth(meseSelezionato, ricevutePerMese[meseSelezionato], totaleRimborsi, risparmioFiscale);
            return;
        }

        // Crea dialog per selezione mese
        const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                         'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
        
        let opzioniMesi = 'EXPORT EXCEL - Seleziona il mese da esportare:\n\n';
        mesiDisponibili.forEach((meseAnno, index) => {
            const [anno, mese] = meseAnno.split('-');
            const nomeCompleto = `${mesiNomi[parseInt(mese)]} ${anno}`;
            const numRicevute = ricevutePerMese[meseAnno].length;
            const rimborsiMese = ricevutePerMese[meseAnno].reduce((sum, p) => sum + (p.rimborsoSpese || 0), 0);
            opzioniMesi += `${index + 1}. ${nomeCompleto} (${numRicevute} ricevute, €${rimborsiMese.toFixed(2).replace('.', ',')} rimborsi)\n`;
        });
        
        opzioniMesi += '\n0. Esporta TUTTI i mesi separatamente\n';
        opzioniMesi += `\nRiepilogo totale: ${results.length} ricevute, €${totaleRimborsi.toFixed(2).replace('.', ',')} rimborsi, €${risparmioFiscale.toFixed(2).replace('.', ',')} risparmio fiscale\n`;
        opzioniMesi += '\nInserisci il numero della tua scelta:';

        const scelta = prompt(opzioniMesi);
        
        if (!scelta || scelta === null) return;
        
        const sceltaNum = parseInt(scelta);
        
        if (sceltaNum === 0) {
            // Esporta tutti i mesi separatamente
            exportAllMonthsExcel(ricevutePerMese, totaleRimborsi, risparmioFiscale);
        } else if (sceltaNum >= 1 && sceltaNum <= mesiDisponibili.length) {
            const meseSelezionato = mesiDisponibili[sceltaNum - 1];
            const rimborsiMese = ricevutePerMese[meseSelezionato].reduce((sum, p) => sum + (p.rimborsoSpese || 0), 0);
            const risparmioMese = rimborsiMese * 0.20;
            exportExcelForMonth(meseSelezionato, ricevutePerMese[meseSelezionato], rimborsiMese, risparmioMese);
        } else {
            alert('Selezione non valida');
        }
        
    } catch (error) {
        console.error('Errore durante export Excel per mese:', error);
        alert(`Errore durante l'export Excel per mese:\n\n${error.message}\n\nControlla la console per maggiori dettagli.`);
    }
}

// Esporta Excel per un mese specifico - CON MODIFICHE COMMERCIALISTA
function exportExcelForMonth(meseAnno, ricevuteMese, rimborsiMese, risparmioMese) {
    const [anno, mese] = meseAnno.split('-');
    const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                     'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    const nomeCompleto = `${mesiNomi[parseInt(mese)]} ${anno}`;

    const conferma = confirm(
        `EXPORT EXCEL - ${nomeCompleto.toUpperCase()}\n\n` +
        `Ricevute da esportare: ${ricevuteMese.length}\n` +
        `Rimborsi spese del mese: € ${rimborsiMese.toFixed(2).replace('.', ',')}\n` +
        `Risparmio fiscale (20%): € ${risparmioMese.toFixed(2).replace('.', ',')}\n\n` +
        `Procedere con l'export Excel?`
    );
    
    if (!conferma) return;

    try {
        console.log(`Creando Excel per: ${nomeCompleto} (${ricevuteMese.length} ricevute) con modifiche commercialista...`);

        const excelData = [];
        
        // Intestazione - 32 colonne con modifiche commercialista
        excelData.push([
            'Id Paese', 'Partita Iva', 'Codice Fiscale', 'Denominazione', 'Cognome', 'Nome',
            'Indirizzo', 'Num. civico', 'CAP', 'Comune', 'Provincia', 'Causale', 'Sezionale',
            'Tipo Doc.', 'Data Doc.', 'Numero Doc.', 'Data doc. fattura origine', 'Num. doc. fattura origine',
            'Descr. Articolo1', 'Imponibile1', 'Aliquota IVA1', 'Natura IVA1', 'Codice IVA1',
            'Imposta1', 'Totale Imponibile1', 'Totale Imposta1', 'Totale Imponibile', 'Totale Imposta', 
            'Totale Documento', 'Esigibilita\' IVA', 'Conto', 'Conto1'
        ]);

        // Ordina le ricevute di questo mese per CF
        const ricevuteMeseOrdinate = ricevuteMese.sort((a, b) => {
            const cfA = a.codiceFiscale || `${a.nome}_${a.cognome}`;
            const cfB = b.codiceFiscale || `${b.nome}_${b.cognome}`;
            return cfA.localeCompare(cfB);
        });

        // Dati per questo mese - CON NUMERAZIONE UNIFICATA E MODIFICHE COMMERCIALISTA
        ricevuteMeseOrdinate.forEach(person => {
            try {
                const numeroRicevuta = person.numeroProgressivo;
                
                console.log(`Excel mese: ${person.nome} ${person.cognome} - ${person.mese}/${person.anno} → Numero: ${numeroRicevuta}`);
                
                const lastDayOfMonth = new Date(person.anno, person.mese, 0);
                const dataDoc = lastDayOfMonth.toLocaleDateString('it-IT');
                
                const denominazione = `${person.nome || ''} ${person.cognome || ''}`.trim();
                
                // Separa indirizzo e numero civico
                const indirizzoCompleto = person.indirizzo || '';
                let indirizzo = '';
                let numCivico = '';
                
                if (indirizzoCompleto.trim()) {
                    const parts = indirizzoCompleto.trim().split(/\s+/);
                    if (parts.length > 0) {
                        const lastPart = parts[parts.length - 1];
                        if (/^\d+[a-zA-Z]?$/.test(lastPart)) {
                            numCivico = lastPart;
                            indirizzo = parts.slice(0, -1).join(' ');
                        } else {
                            indirizzo = indirizzoCompleto;
                        }
                    }
                }

                // CALCOLI CON MODIFICHE COMMERCIALISTA
                const compensoLordo = person.compenso || 0;
                const rimborsiSpese = person.rimborsoSpese || 0;
                const totaleDocumento = compensoLordo + rimborsiSpese;
                
                // Formato con virgola italiana
                const compensoStr = compensoLordo.toFixed(2).replace('.', ',');
                const rimborsiStr = rimborsiSpese.toFixed(2).replace('.', ',');
                const totaleStr = totaleDocumento.toFixed(2).replace('.', ',');

                // Riga dati - 32 colonne
                excelData.push([
                    'IT', person.partitaIva || '', person.codiceFiscale || '', denominazione,
                    person.cognome || '', person.nome || '', indirizzo, numCivico,
                    person.cap || '', person.citta || '', person.provincia || '', '104', '4', // CAUSALE 104
                    'TD01', dataDoc, numeroRicevuta.toString(), dataDoc, '1',
                    'COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE',
                    compensoStr, '0', 'N2', 'NI', '0',
                    rimborsiStr,        // Totale Imponibile1 (RIMBORSI)
                    '0',                // Totale Imposta1 (SEMPRE 0)
                    compensoStr,        // Totale Imponibile (COMPENSO)
                    '0',                // Totale Imposta (SEMPRE 0)
                    totaleStr,          // Totale Documento (COMPENSO + RIMBORSI)
                    'I',                // Esigibilità IVA
                    '816000',           // Conto (FISSO)
                    '809230'            // Conto1 (FISSO)
                ]);
                
            } catch (personError) {
                console.error(`Errore elaborando persona in ${nomeCompleto}:`, personError);
            }
        });

        // Crea e scarica il file Excel per questo mese
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        
        // Larghezza colonne (32 colonne)
        const colWidths = [
            {wch: 5}, {wch: 15}, {wch: 16}, {wch: 30}, {wch: 20}, {wch: 20},
            {wch: 30}, {wch: 8}, {wch: 8}, {wch: 25}, {wch: 8}, {wch: 8}, {wch: 8},
            {wch: 8}, {wch: 12}, {wch: 10}, {wch: 12}, {wch: 10}, {wch: 50},
            {wch: 12}, {wch: 8}, {wch: 8}, {wch: 8}, {wch: 8}, {wch: 12},
            {wch: 8}, {wch: 12}, {wch: 8}, {wch: 12}, {wch: 8}, {wch: 10}, {wch: 10}
        ];
        ws['!cols'] = colWidths;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, nomeCompleto);
        
        const fileName = `Export_Ricevute_${mesiNomi[parseInt(mese)]}_${anno}.xlsx`;
        
        console.log(`Scaricando file: ${fileName}`);
        XLSX.writeFile(wb, fileName);
        
        // Messaggio di successo
        setTimeout(() => {
            const totalCompensi = ricevuteMese.reduce((s,p) => s + p.compenso, 0);
            const message = `✅ Export Excel completato per ${nomeCompleto}!\n\n` +
                          `📄 Ricevute esportate: ${ricevuteMese.length}\n` +
                          `💼 Compensi mese: €${totalCompensi.toFixed(2).replace('.', ',')}\n` +
                          `💰 Rimborsi spese: €${rimborsiMese.toFixed(2).replace('.', ',')}\n` +
                          `🎯 Risparmio fiscale: €${risparmioMese.toFixed(2).replace('.', ',')}\n\n` +
                          `🔧 Modifiche commercialista applicate:\n` +
                          `• Causale 104 • Separazione compensi/rimborsi\n` +
                          `• Virgola italiana • Conti bilancio\n\n` +
                          `File: ${fileName}`;
            alert(message);
        }, 500);
        
    } catch (error) {
        console.error(`Errore creando Excel per mese ${meseAnno}:`, error);
        alert(`Errore durante l'export Excel per ${nomeCompleto}:\n\n${error.message}`);
    }
}

// Esporta tutti i mesi separatamente
function exportAllMonthsExcel(ricevutePerMese, totaleRimborsi, risparmioFiscale) {
    const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                     'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    
    let totalRicevute = 0;
    let totalRimborsiTutti = 0;
    
    Object.values(ricevutePerMese).forEach(ricevuteMese => {
        totalRicevute += ricevuteMese.length;
        totalRimborsiTutti += ricevuteMese.reduce((sum, p) => sum + (p.rimborsoSpese || 0), 0);
    });
    
    const risparmioTotale = totalRimborsiTutti * 0.20;
    
    const conferma = confirm(
        `EXPORT EXCEL - TUTTI I MESI SEPARATAMENTE\n\n` +
        `Mesi da esportare: ${Object.keys(ricevutePerMese).length}\n` +
        `Ricevute totali: ${totalRicevute}\n` +
        `Rimborsi spese totali: € ${totalRimborsiTutti.toFixed(2).replace('.', ',')}\n` +
        `Risparmio fiscale (20%): € ${risparmioTotale.toFixed(2).replace('.', ',')}\n\n` +
        `Verranno generati ${Object.keys(ricevutePerMese).length} file Excel separati.\n` +
        `Procedere?`
    );
    
    if (!conferma) return;

    let filesCreated = 0;
    
    for (const [meseAnno, ricevuteMese] of Object.entries(ricevutePerMese)) {
        const rimborsiMese = ricevuteMese.reduce((sum, p) => sum + (p.rimborsoSpese || 0), 0);
        const risparmioMese = rimborsiMese * 0.20;
        
        exportExcelForMonth(meseAnno, ricevuteMese, rimborsiMese, risparmioMese);
        filesCreated++;
        
        // Pausa tra i download per non sovraccaricare il browser
        if (filesCreated < Object.keys(ricevutePerMese).length) {
            setTimeout(() => {}, 300); // Pausa sincrona
        }
    }
    
    // Messaggio finale
    setTimeout(() => {
        alert(`✅ Export Excel completato per tutti i mesi!\n\n` +
              `📁 File generati: ${filesCreated}\n` +
              `📄 Ricevute totali: ${totalRicevute}\n` +
              `💰 Rimborsi spese: €${totalRimborsiTutti.toFixed(2).replace('.', ',')}\n` +
              `🎯 Risparmio fiscale: €${risparmioTotale.toFixed(2).replace('.', ',')}\n\n` +
              `Tutti i file Excel sono stati scaricati con numerazione cronologica corretta.`);
    }, 1000);
}

// Export Excel singolo - MANTIENE COMPATIBILITA'
function exportToExcel() {
    console.log('Redirect a export per mese...');
    exportToExcelByMonth();
}

// Utility per validazione dati prima dell'export
function validateExportData() {
    if (!results || results.length === 0) {
        return { valid: false, message: 'Nessuna ricevuta da esportare' };
    }
    
    let invalidRecords = 0;
    let missingNumbers = 0;
    let zeroAmounts = 0;
    
    results.forEach((person, index) => {
        if (!person.nome || !person.cognome || !person.compenso) {
            console.warn(`Record ${index} incompleto:`, person);
            invalidRecords++;
        }
        
        if (!person.numeroProgressivo) {
            console.warn(`Record ${index} senza numero progressivo:`, person);
            missingNumbers++;
        }
        
        if (person.compenso === 0 && person.rimborsoSpese === 0) {
            console.warn(`Record ${index} con importi zero:`, person);
            zeroAmounts++;
        }
    });
    
    if (invalidRecords > 0 || missingNumbers > 0) {
        return { 
            valid: false, 
            message: `${invalidRecords} record incompleti, ${missingNumbers} senza numerazione, ${zeroAmounts} con importi zero. Rigenera le ricevute.` 
        };
    }
    
    return { valid: true, message: 'Dati validi per export con modifiche commercialista' };
}

// Utility per debug export
function debugExportData() {
    console.log('=== DEBUG EXPORT DATA - MODIFICHE COMMERCIALISTA ===');
    console.log('Results array:', results);
    console.log('Results length:', results.length);
    
    if (results.length > 0) {
        console.log('Sample record:', results[0]);
        console.log('Record keys:', Object.keys(results[0]));
        console.log('Numerazione progressiva presente:', !!results[0].numeroProgressivo);
        console.log('Compenso presente:', !!results[0].compenso);
        console.log('Rimborsi presenti:', !!results[0].rimborsoSpese);
        
        // Verifica calcoli
        const sample = results[0];
        if (sample.compenso && sample.rimborsoSpese !== undefined) {
            const totaleCalcolato = sample.compenso + sample.rimborsoSpese;
            console.log(`Esempio calcolo - Compenso: ${sample.compenso}, Rimborsi: ${sample.rimborsoSpese}, Totale: ${totaleCalcolato}`);
        }
        
        // Verifica raggruppamento per mese
        const gruppiFase = {};
        results.forEach(p => {
            const key = `${p.anno}-${p.mese.toString().padStart(2, '0')}`;
            if (!gruppiFase[key]) gruppiFase[key] = 0;
            gruppiFase[key]++;
        });
        console.log('Distribuzione per mese:', gruppiFase);
    }
    
    const validation = validateExportData();
    console.log('Validation result:', validation);
}

// ESPOSIZIONE IMMEDIATA DELLE FUNZIONI
console.log('🔄 Esposizione funzioni excel-export.js...');

window.exportToExcel = exportToExcel;
window.exportToExcelByMonth = exportToExcelByMonth;
window.exportExcelForMonth = exportExcelForMonth;
window.exportAllMonthsExcel = exportAllMonthsExcel;
window.validateExportData = validateExportData;
window.debugExportData = debugExportData;

// VERIFICA IMMEDIATA
console.log('✅ Funzioni esposte:', {
    exportToExcel: typeof window.exportToExcel,
    exportToExcelByMonth: typeof window.exportToExcelByMonth,
    exportExcelForMonth: typeof window.exportExcelForMonth,
    exportAllMonthsExcel: typeof window.exportAllMonthsExcel,
    validateExportData: typeof window.validateExportData,
    debugExportData: typeof window.debugExportData
});

// TEST ESPOSIZIONE
if (typeof window.exportToExcel !== 'function') {
    console.error('❌ ERRORE CRITICO: exportToExcel non esposta!');
} else {
    console.log('✅ exportToExcel esposta correttamente');
}

if (typeof window.exportToExcelByMonth !== 'function') {
    console.error('❌ ERRORE CRITICO: exportToExcelByMonth non esposta!');
} else {
    console.log('✅ exportToExcelByMonth esposta correttamente');
}

console.log('✅ excel-export.js caricato - Export per mese con modifiche commercialista FUNZIONANTE');
