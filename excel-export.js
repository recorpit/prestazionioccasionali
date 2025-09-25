// Export Excel singolo - COMPLETO
function exportToExcel() {
    console.log('Inizio export Excel singolo...');
    
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
            `• Totale rimborsi spese: € ${totaleRimborsi.toFixed(2)}\n` +
            `• Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2)}\n\n` +
            `Procedere con l'export Excel?`
        );
        
        if (!conferma) return;

        console.log(`Esportando ${results.length} ricevute in Excel...`);

        // ORDINAMENTO CRUCIALE: Ordina i results per data prima dell'export
        const resultsOrdinati = [...results].sort((a, b) => {
            if (a.anno !== b.anno) {
                return a.anno - b.anno;
            }
            return a.mese - b.mese;
        });
        
        console.log('Export Excel - Results ordinati per data:', 
            resultsOrdinati.map(r => `${r.mese}/${r.anno} - ${r.nome} ${r.cognome}`));

        // Prepara dati Excel con intestazione a 28 colonne
        const excelData = [];
        
        // Intestazione - 28 colonne come richiesto
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
            'Totale Imponibile',     // 25
            'Totale Imposta',        // 26
            'Totale Documento',      // 27
            'Esigibilita\' IVA'      // 28
        ]);

        // CORREZIONE NUMERAZIONE PROGRESSIVA PER EXCEL
        // Crea una mappa temporanea per numerazione corretta nell'ordine cronologico
        const numeroTemporaneoPerCF = {};
        
        // Elabora ogni ricevuta NELL'ORDINE CORRETTO per ricostruire numerazione
        resultsOrdinati.forEach((person, index) => {
            try {
                const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
                const receiptNumber = getCurrentReceiptNumber(cfKey);
                
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
                        // Se l'ultima parte sembra un numero civico (numero + eventuale lettera)
                        if (/^\d+[a-zA-Z]?$/.test(lastPart)) {
                            numCivico = lastPart;
                            indirizzo = parts.slice(0, -1).join(' ');
                        } else {
                            indirizzo = indirizzoCompleto;
                        }
                    }
                }

                // Riga dati - 28 colonne
                const rowData = [
                    'IT',                           // 1 - Id Paese
                    person.partitaIva || '',        // 2 - Partita IVA (dalla colonna M iscrizioni)
                    person.codiceFiscale || '',     // 3 - Codice Fiscale
                    denominazione,                  // 4 - Denominazione
                    person.cognome || '',           // 5 - Cognome
                    person.nome || '',              // 6 - Nome
                    indirizzo,                      // 7 - Indirizzo
                    numCivico,                      // 8 - Numero civico
                    person.cap || '',               // 9 - CAP
                    person.citta || '',             // 10 - Comune
                    person.provincia || '',         // 11 - Provincia
                    '135',                          // 12 - Causale (135 per prestazioni occasionali)
                    '4',                            // 13 - Sezionale
                    'TD01',                         // 14 - Tipo Doc. (TD01 corretto per fatture)
                    dataDoc,                        // 15 - Data Documento
                    receiptNumber.toString(),       // 16 - Numero Documento
                    dataDoc,                        // 17 - Data doc. fattura origine
                    '1',                            // 18 - Num. doc. fattura origine
                    'COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE', // 19 - Descrizione
                    person.compenso.toFixed(2),     // 20 - Imponibile1
                    '0',                            // 21 - Aliquota IVA1 (0% per occasionali)
                    'N2',                           // 22 - Natura IVA1 (N2 = non soggetta)
                    'NI',                           // 23 - Codice IVA1 (NI = non imponibile)
                    '0',                            // 24 - Imposta1 (0 per occasionali)
                    person.compenso.toFixed(2),     // 25 - Totale Imponibile
                    '0',                            // 26 - Totale Imposta
                    person.compenso.toFixed(2),     // 27 - Totale Documento
                    'I'                             // 28 - Esigibilità IVA (I = immediata)
                ];
                
                excelData.push(rowData);
                console.log(`Riga Excel creata per: ${denominazione}`);
                
            } catch (error) {
                console.error(`Errore elaborando ricevuta ${index}:`, error);
                // Continua con la prossima ricevuta invece di fermarsi
            }
        });

        // Crea e scarica il file Excel
        console.log('Creazione file Excel...');
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        
        // Imposta larghezza colonne per leggibilità
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
            {wch: 12},  // Totale Imponibile
            {wch: 12},  // Totale Imposta
            {wch: 12},  // Totale Documento
            {wch: 8}    // Esigibilità IVA
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
            alert(`✅ Export Excel completato!\n\nFile: ${fileName}\n\n📊 Statistiche:\n• Ricevute: ${results.length}\n• Rimborsi spese: €${totaleRimborsi.toFixed(2)}\n• Risparmio fiscale: €${risparmioFiscale.toFixed(2)}`);
        }, 500);
        
    } catch (error) {
        console.error('Errore durante export Excel:', error);
        alert(`Errore durante l'export Excel:\n\n${error.message}\n\nControlla la console per maggiori dettagli.`);
    }
}

// Export Excel diviso per mese - COMPLETO
function exportToExcelByMonth() {
    console.log('Inizio export Excel per mese...');
    
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
            `• Totale rimborsi spese: € ${totaleRimborsi.toFixed(2)}\n` +
            `• Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2)}\n\n` +
            `Verranno creati file Excel separati per ogni mese.\nProcedere?`
        );
        
        if (!conferma) return;

        // Raggruppa le ricevute per mese CON ORDINAMENTO
        const ricevutePerMese = {};
        
        // Prima ordina tutti i results per data
        const resultsOrdinatiPerMese = [...results].sort((a, b) => {
            if (a.anno !== b.anno) {
                return a.anno - b.anno;
            }
            return a.mese - b.mese;
        });
        
        resultsOrdinatiPerMese.forEach(person => {
            // Crea chiave mese nel formato YYYY-MM
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
                
                // NUMERAZIONE TEMPORANEA PER QUESTO MESE (nell'ordine cronologico)
                const numeroTemporaneoPerCF = {};
                
                // Ordina le ricevute di questo mese per CF per mantenere coerenza numerica
                const ricevuteMeseOrdinate = ricevuteMese.sort((a, b) => {
                    const cfA = a.codiceFiscale || `${a.nome}_${a.cognome}`;
                    const cfB = b.codiceFiscale || `${b.nome}_${b.cognome}`;
                    return cfA.localeCompare(cfB);
                });
                
                // Intestazione - 28 colonne identica al singolo
                excelData.push([
                    'Id Paese', 'Partita Iva', 'Codice Fiscale', 'Denominazione', 'Cognome', 'Nome',
                    'Indirizzo', 'Num. civico', 'CAP', 'Comune', 'Provincia', 'Causale', 'Sezionale',
                    'Tipo Doc.', 'Data Doc.', 'Numero Doc.', 'Data doc. fattura origine', 'Num. doc. fattura origine',
                    'Descr. Articolo1', 'Imponibile1', 'Aliquota IVA1', 'Natura IVA1', 'Codice IVA1',
                    'Imposta1', 'Totale Imponibile', 'Totale Imposta', 'Totale Documento', 'Esigibilita\' IVA'
                ]);

                // Dati per questo mese
                ricevuteMese.forEach(person => {
                    try {
                        const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
                        const receiptNumber = getCurrentReceiptNumber(cfKey);
                        
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

                        // Riga dati - 28 colonne
                        excelData.push([
                            'IT', person.partitaIva || '', person.codiceFiscale || '', denominazione,
                            person.cognome || '', person.nome || '', indirizzo, numCivico,
                            person.cap || '', person.citta || '', person.provincia || '', '135', '4',
                            'TD01', dataDoc, receiptNumber.toString(), dataDoc, '1',
                            'COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE',
                            person.compenso.toFixed(2), '0', 'N2', 'NI', '0',
                            person.compenso.toFixed(2), '0', person.compenso.toFixed(2), 'I'
                        ]);
                        
                        totalExportedReceipts++;
                        
                    } catch (personError) {
                        console.error(`Errore elaborando persona in ${nomeCompletoMese}:`, personError);
                    }
                });

                // Crea e scarica il file Excel per questo mese
                const ws = XLSX.utils.aoa_to_sheet(excelData);
                
                // Larghezza colonne
                const colWidths = [
                    {wch: 5}, {wch: 15}, {wch: 16}, {wch: 30}, {wch: 20}, {wch: 20},
                    {wch: 30}, {wch: 8}, {wch: 8}, {wch: 25}, {wch: 8}, {wch: 8}, {wch: 8},
                    {wch: 8}, {wch: 12}, {wch: 10}, {wch: 12}, {wch: 10}, {wch: 50},
                    {wch: 12}, {wch: 8}, {wch: 8}, {wch: 8}, {wch: 8}, {wch: 12},
                    {wch: 12}, {wch: 12}, {wch: 8}
                ];
                ws['!cols'] = colWidths;

                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, nomeCompletoMese);
                
                const fileName = `Export_Ricevute_${mesiNomi[parseInt(mese)]}_${anno}.xlsx`;
                
                console.log(`Scaricando file: ${fileName}`);
                XLSX.writeFile(wb, fileName);
                filesCreated++;
                
                // Pausa piccola tra i download per evitare problemi browser
                if (filesCreated < mesiOrdinati.length) {
                    setTimeout(() => {}, 200);
                }
                
            } catch (monthError) {
                console.error(`Errore creando Excel per mese ${meseAnno}:`, monthError);
            }
        });

        // Messaggio finale di successo
        setTimeout(() => {
            const message = `✅ Export Excel per mese completato!\n\n` +
                          `📁 File generati: ${filesCreated}\n` +
                          `📄 Ricevute esportate: ${totalExportedReceipts}\n` +
                          `💰 Rimborsi spese totali: €${totaleRimborsi.toFixed(2)}\n` +
                          `🎯 Risparmio fiscale: €${risparmioFiscale.toFixed(2)}\n\n` +
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
    results.forEach((person, index) => {
        if (!person.nome || !person.cognome || !person.compenso) {
            console.warn(`Record ${index} incompleto:`, person);
            invalidRecords++;
        }
    });
    
    if (invalidRecords > 0) {
        return { 
            valid: false, 
            message: `${invalidRecords} record incompleti trovati. Controlla i dati.` 
        };
    }
    
    return { valid: true, message: 'Dati validi per export' };
}

// Utility per debug export
function debugExportData() {
    console.log('=== DEBUG EXPORT DATA ===');
    console.log('Results array:', results);
    console.log('Results length:', results.length);
    
    if (results.length > 0) {
        console.log('Sample record:', results[0]);
        console.log('Record keys:', Object.keys(results[0]));
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
console.log('excel-export.js caricato - Funzioni esposte:', {
    exportToExcel: typeof window.exportToExcel,
    exportToExcelByMonth: typeof window.exportToExcelByMonth,
    validateExportData: typeof window.validateExportData,
    debugExportData: typeof window.debugExportData
});

// Log caricamento modulo
console.log('✅ excel-export.js caricato completamente');
