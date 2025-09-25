// Export Excel singolo - CON MODIFICHE DEL COMMERCIALISTA
function exportToExcel() {
    console.log('Inizio export Excel singolo con modifiche commercialista...');
    
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
        
        // Mostra popup informativo
        const conferma = confirm(
            `EXPORT EXCEL - RIEPILOGO RIMBORSI SPESE\n\n` +
            `• Ricevute da esportare: ${results.length}\n` +
            `• Totale rimborsi spese: € ${totaleRimborsi.toFixed(2).replace('.', ',')}\n` +
            `• Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2).replace('.', ',')}\n\n` +
            `Procedere con l'export Excel?`
        );
        
        if (!conferma) return;

        console.log(`Esportando ${results.length} ricevute in Excel con numerazione unificata e modifiche commercialista...`);

        // Prepara dati Excel con intestazione a 32 colonne (30 + 2 conti)
        const excelData = [];
        
        // Intestazione - 32 colonne totali
        excelData.push([
            'Id Paese',              // 1
            'Partita Iva',           // 2
            'Codice Fiscale',        // 3
            'Denominazione',         // 4
            'Cognome',               // 5
            'Nome',                  // 6
            'Indirizzo',             // 7
            'Num. civico',           // 8
            'CAP',                   // 9
            'Comune',                // 10
            'Provincia',             // 11
            'Causale',               // 12
            'Sezionale',             // 13
            'Tipo Doc.',             // 14
            'Data Doc.',             // 15
            'Numero Doc.',           // 16
            'Data doc. fattura origine', // 17
            'Num. doc. fattura origine', // 18
            'Descr. Articolo1',      // 19
            'Imponibile1',           // 20
            'Aliquota IVA1',         // 21
            'Natura IVA1',           // 22
            'Codice IVA1',           // 23
            'Imposta1',              // 24
            'Totale Imponibile1',    // 25 (AA) - RIMBORSI
            'Totale Imposta1',       // 26 (AB) - SEMPRE 0
            'Totale Imponibile',     // 27 (AC) - COMPENSO
            'Totale Imposta',        // 28 (AD) - SEMPRE 0
            'Totale Documento',      // 29 (AE) - COMPENSO + RIMBORSI
            'Esigibilita\' IVA',     // 30 (AF)
            'Conto',                 // 31 (AG) - 816000
            'Conto1'                 // 32 (AH) - 809230
        ]);

        // Elabora ogni ricevuta USANDO LA NUMERAZIONE UNIFICATA
        results.forEach((person, index) => {
            try {
                // USA IL NUMERO PROGRESSIVO UNIFICATO
                const numeroRicevuta = person.numeroProgressivo;
                
                console.log(`Excel: ${person.nome} ${person.cognome} - ${person.mese}/${person.anno} → Numero: ${numeroRicevuta}`);
                
                // Data ricevuta = ultimo giorno del mese del pagamento
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

                // CALCOLI SECONDO LE SPECIFICHE DEL COMMERCIALISTA
                const compensoLordo = person.compenso || 0;
                const rimborsiSpese = person.rimborsoSpese || 0;
                const totaleDocumento = compensoLordo + rimborsiSpese;
                
                // Formattazione numeri con VIRGOLA (non punto)
                const compensoStr = compensoLordo.toFixed(2).replace('.', ',');
                const rimborsiStr = rimborsiSpese.toFixed(2).replace('.', ',');
                const totaleStr = totaleDocumento.toFixed(2).replace('.', ',');

                // Riga dati - 32 colonne
                const rowData = [
                    'IT',                           // 1 - Id Paese
                    person.partitaIva || '',        // 2 - Partita IVA
                    person.codiceFiscale || '',     // 3 - Codice Fiscale
                    denominazione,                  // 4 - Denominazione
                    person.cognome || '',           // 5 - Cognome
                    person.nome || '',              // 6 - Nome
                    indirizzo,                      // 7 - Indirizzo
                    numCivico,                      // 8 - Numero civico
                    person.cap || '',               // 9 - CAP
                    person.citta || '',             // 10 - Comune
                    person.provincia || '',         // 11 - Provincia
                    '104',                          // 12 - Causale (CAMBIATA DA 135 A 104)
                    '4',                            // 13 - Sezionale
                    'TD01',                         // 14 - Tipo Doc.
                    dataDoc,                        // 15 - Data Documento
                    numeroRicevuta.toString(),      // 16 - NUMERO PROGRESSIVO UNIFICATO
                    dataDoc,                        // 17 - Data doc. fattura origine
                    '1',                            // 18 - Num. doc. fattura origine
                    'COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE', // 19 - Descrizione
                    compensoStr,                    // 20 - Imponibile1 (compenso con virgola)
                    '0',                            // 21 - Aliquota IVA1
                    'N2',                           // 22 - Natura IVA1
                    'NI',                           // 23 - Codice IVA1
                    '0',                            // 24 - Imposta1
                    rimborsiStr,                    // 25 (AA) - Totale Imponibile1 (RIMBORSI)
                    '0',                            // 26 (AB) - Totale Imposta1 (SEMPRE 0)
                    compensoStr,                    // 27 (AC) - Totale Imponibile (COMPENSO)
                    '0',                            // 28 (AD) - Totale Imposta (SEMPRE 0)
                    totaleStr,                      // 29 (AE) - Totale Documento (COMPENSO + RIMBORSI)
                    'I',                            // 30 (AF) - Esigibilità IVA
                    '816000',                       // 31 (AG) - Conto (FISSO)
                    '809230'                        // 32 (AH) - Conto1 (FISSO)
                ];
                
                excelData.push(rowData);
                
                console.log(`Excel row: ${denominazione} - Compenso: ${compensoStr}, Rimborsi: ${rimborsiStr}, Totale: ${totaleStr}`);
                
            } catch (error) {
                console.error(`Errore elaborando ricevuta ${index}:`, error);
            }
        });

        // Crea e scarica il file Excel
        console.log('Creazione file Excel con 32 colonne...');
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        
        // Imposta larghezza colonne per leggibilità (32 colonne)
        const colWidths = [
            {wch: 5},   // Id Paese
            {wch: 15},  // Partita IVA
            {wch: 16},  // Codice Fiscale
            {wch: 30},  // Denominazione
            {wch: 20},  // Cognome
            {wch: 20},  // Nome
            {wch: 30},  // Indirizzo
            {wch: 8},   // Num civico
            {wch: 8},   // CAP
            {wch: 25},  // Comune
            {wch: 8},   // Provincia
            {wch: 8},   // Causale
            {wch: 8},   // Sezionale
            {wch: 8},   // Tipo Doc
            {wch: 12},  // Data Doc
            {wch: 10},  // Numero Doc
            {wch: 12},  // Data origine
            {wch: 10},  // Num origine
            {wch: 50},  // Descrizione
            {wch: 12},  // Imponibile1
            {wch: 8},   // Aliquota IVA1
            {wch: 8},   // Natura IVA1
            {wch: 8},   // Codice IVA1
            {wch: 8},   // Imposta1
            {wch: 12},  // Totale Imponibile1 (rimborsi)
            {wch: 8},   // Totale Imposta1
            {wch: 12},  // Totale Imponibile (compenso)
            {wch: 8},   // Totale Imposta
            {wch: 12},  // Totale Documento
            {wch: 8},   // Esigibilità IVA
            {wch: 10},  // Conto
            {wch: 10}   // Conto1
        ];
        ws['!cols'] = colWidths;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Export Ricevute");
        
        const currentDate = new Date().toISOString().split('T')[0];
        const fileName = `Export_Ricevute_${currentDate}.xlsx`;
        
        console.log(`Scaricando file: ${fileName}`);
        XLSX.writeFile(wb, fileName);
        
        // Messaggio di successo
        setTimeout(() => {
            alert(`✅ Export Excel completato con modifiche del commercialista!\n\nFile: ${fileName}\n\n📊 Statistiche:\n• Ricevute: ${results.length}\n• Compensi: €${results.reduce((s,p) => s + p.compenso, 0).toFixed(2).replace('.', ',')}\n• Rimborsi spese: €${totaleRimborsi.toFixed(2).replace('.', ',')}\n• Risparmio fiscale: €${risparmioFiscale.toFixed(2).replace('.', ',')}\n\n🔧 Modifiche applicate:\n• Causale 104 (era 135)\n• Separazione compensi/rimborsi\n• Formato virgola italiana\n• Conti bilancio aggiunti`);
        }, 500);
        
    } catch (error) {
        console.error('Errore durante export Excel:', error);
        alert(`Errore durante l'export Excel:\n\n${error.message}\n\nControlla la console per maggiori dettagli.`);
    }
}

// Export Excel diviso per mese - CON MODIFICHE DEL COMMERCIALISTA
function exportToExcelByMonth() {
    console.log('Inizio export Excel per mese con modifiche commercialista...');
    
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
        
        // Mostra popup informativo
        const conferma = confirm(
            `EXPORT EXCEL PER MESE - RIEPILOGO RIMBORSI\n\n` +
            `• Ricevute da esportare: ${results.length}\n` +
            `• Totale rimborsi spese: € ${totaleRimborsi.toFixed(2).replace('.', ',')}\n` +
            `• Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2).replace('.', ',')}\n\n` +
            `Verranno creati file Excel separati per ogni mese.\nProcedere?`
        );
        
        if (!conferma) return;

        console.log('Raggruppamento ricevute per mese con numerazione unificata e modifiche commercialista...');

        // Raggruppa le ricevute per mese
        const ricevutePerMese = {};
        
        results.forEach(person => {
            const chiaveMese = `${person.anno}-${person.mese.toString().padStart(2, '0')}`;
            if (!ricevutePerMese[chiaveMese]) {
                ricevutePerMese[chiaveMese] = [];
            }
            ricevutePerMese[chiaveMese].push(person);
        });

        const mesiOrdinati = Object.keys(ricevutePerMese).sort();
        console.log('Mesi da esportare:', mesiOrdinati);

        // Nomi mesi in italiano
        const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                         'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];

        let filesCreated = 0;
        let totalExportedReceipts = 0;
        
        // Crea un file Excel per ogni mese
        mesiOrdinati.forEach(meseAnno => {
            try {
                const ricevuteMese = ricevutePerMese[meseAnno];
                const [anno, mese] = meseAnno.split('-');
                const nomeCompletoMese = `${mesiNomi[parseInt(mese)]} ${anno}`;

                console.log(`Creando Excel per: ${nomeCompletoMese} (${ricevuteMese.length} ricevute)`);

                const excelData = [];
                
                // Intestazione - 32 colonne identica al singolo
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

                // Dati per questo mese - CON MODIFICHE COMMERCIALISTA
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
                        
                        // Formato con virgola
                        const compensoStr = compensoLordo.toFixed(2).replace('.', ',');
                        const rimborsiStr = rimborsiSpese.toFixed(2).replace('.', ',');
                        const totaleStr = totaleDocumento.toFixed(2).replace('.', ',');

                        // Riga dati - 32 colonne
                        excelData.push([
                            'IT', person.partitaIva || '', person.codiceFiscale || '', denominazione,
                            person.cognome || '', person.nome || '', indirizzo, numCivico,
                            person.cap || '', person.citta || '', person.provincia || '', '104', '4',
                            'TD01', dataDoc, numeroRicevuta.toString(), dataDoc, '1',
                            'COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE',
                            compensoStr, '0', 'N2', 'NI', '0',
                            rimborsiStr,        // Totale Imponibile1 (rimborsi)
                            '0',                // Totale Imposta1
                            compensoStr,        // Totale Imponibile (compenso)
                            '0',                // Totale Imposta
                            totaleStr,          // Totale Documento (compenso + rimborsi)
                            'I',                // Esigibilità IVA
                            '816000',           // Conto
                            '809230'            // Conto1
                        ]);
                        
                        totalExportedReceipts++;
                        
                    } catch (personError) {
                        console.error(`Errore elaborando persona in ${nomeCompletoMese}:`, personError);
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
                XLSX.utils.book_append_sheet(wb, ws, nomeCompletoMese);
                
                const fileName = `Export_Ricevute_${mesiNomi[parseInt(mese)]}_${anno}.xlsx`;
                
                console.log(`Scaricando file: ${fileName}`);
                XLSX.writeFile(wb, fileName);
                filesCreated++;
                
                if (filesCreated < mesiOrdinati.length) {
                    setTimeout(() => {}, 200);
                }
                
            } catch (monthError) {
                console.error(`Errore creando Excel per mese ${meseAnno}:`, monthError);
            }
        });

        // Messaggio finale di successo
        setTimeout(() => {
            const totalCompensi = results.reduce((s,p) => s + p.compenso, 0);
            const message = `✅ Export Excel per mese completato con modifiche commercialista!\n\n` +
                          `📁 File generati: ${filesCreated}\n` +
                          `📄 Ricevute esportate: ${totalExportedReceipts}\n` +
                          `💼 Compensi totali: €${totalCompensi.toFixed(2).replace('.', ',')}\n` +
                          `💰 Rimborsi spese totali: €${totaleRimborsi.toFixed(2).replace('.', ',')}\n` +
                          `🎯 Risparmio fiscale: €${risparmioFiscale.toFixed(2).replace('.', ',')}\n\n` +
                          `🔧 Modifiche commercialista applicate:\n` +
                          `• Causale 104 • Separazione compensi/rimborsi\n` +
                          `• Virgola italiana • Conti bilancio\n` +
                          `I file sono stati scaricati automaticamente.`;
            alert(message);
        }, 1000);
        
    } catch (error) {
        console.error('Errore durante export Excel per mese:', error);
        alert(`Errore durante l'export Excel per mese:\n\n${error.message}\n\nControlla la console per maggiori dettagli.`);
    }
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
    }
    
    const validation = validateExportData();
    console.log('Validation result:', validation);
}

// Esposizione funzioni al contesto globale
window.exportToExcel = exportToExcel;
window.exportToExcelByMonth = exportToExcelByMonth;
window.validateExportData = validateExportData;
window.debugExportData = debugExportData;

// Debug - verifica che le funzioni siano esposte
console.log('excel-export.js caricato - Modifiche commercialista implementate:', {
    exportToExcel: typeof window.exportToExcel,
    exportToExcelByMonth: typeof window.exportToExcelByMonth,
    validateExportData: typeof window.validateExportData,
    debugExportData: typeof window.debugExportData
});

console.log('✅ excel-export.js caricato - Modifiche del commercialista applicate: Causale 104, separazione compensi/rimborsi, virgola italiana, conti bilancio');
