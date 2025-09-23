// Export Excel singolo
function exportToExcel() {
    if (results.length === 0) {
        alert('Prima devi eseguire il matching e generare le ricevute!');
        return;
    }

    // Calcola totali rimborsi e risparmio fiscale
    const totaleRimborsi = results.reduce((sum, person) => sum + person.rimborsoSpese, 0);
    const risparmioFiscale = totaleRimborsi * 0.20;
    
    // Mostra popup informativo
    const conferma = confirm(
        `RIEPILOGO RIMBORSI SPESE\n\n` +
        `• Totale rimborsi spese: € ${totaleRimborsi.toFixed(2)}\n` +
        `• Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2)}\n\n` +
        `Procedere con l'export Excel?`
    );
    
    if (!conferma) return;

    const excelData = [];
    
    // Intestazione pulita (senza commenti)
    excelData.push([
        'Id Paese',
        'Partita Iva', 
        'Codice Fiscale',
        'Denominazione',
        'Cognome',
        'Nome',
        'Indirizzo',
        'Num. civico',
        'CAP',
        'Comune',
        'Provincia',
        'Causale',
        'Sezionale',
        'Tipo Doc.',
        'Data Doc.',
        'Numero Doc.',
        'Data doc. fattura origine',
        'Num. doc. fattura origine',
        'Descr. Articolo1',
        'Imponibile1',
        'Aliquota IVA1',
        'Natura IVA1',
        'Codice IVA1',
        'Imposta1',
        'Totale Imponibile',
        'Totale Imposta',
        'Totale Documento',
        'Esigibilita\' IVA'
    ]);

    // Dati
    results.forEach(person => {
        const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
        const receiptNumber = getCurrentReceiptNumber(cfKey);
        const dataDoc = new Date().toLocaleDateString('it-IT');
        const denominazione = `${person.nome} ${person.cognome}`;
        
        // Separa indirizzo e numero civico
        const indirizzoCompleto = person.indirizzo || '';
        const indirizzoParti = indirizzoCompleto.split(' ');
        let indirizzo = '';
        let numCivico = '';
        
        if (indirizzoParti.length > 0) {
            const ultimaParte = indirizzoParti[indirizzoParti.length - 1];
            if (/^\d+[a-zA-Z]?$/.test(ultimaParte)) {
                numCivico = ultimaParte;
                indirizzo = indirizzoParti.slice(0, -1).join(' ');
            } else {
                indirizzo = indirizzoCompleto;
            }
        }

        excelData.push([
            'IT',
            person.partitaIva || '',
            person.codiceFiscale,
            denominazione,
            person.cognome,
            person.nome,
            indirizzo,
            numCivico,
            person.cap || '',
            person.citta || '',
            person.provincia || '',
            '135',
            '4',
            'TD01', // Corretto da TF01 a TD01
            dataDoc,
            receiptNumber,
            dataDoc,
            '1',
            'COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE',
            person.compenso.toFixed(2),
            '0',
            'N2',
            'NI',
            '0',
            person.compenso.toFixed(2),
            '0',
            person.compenso.toFixed(2),
            'I'
        ]);
    });

    // Crea e scarica il file Excel
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Export Ricevute");
    
    const currentDate = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Export_Ricevute_${currentDate}.xlsx`);
}

// Export Excel diviso per mese
function exportToExcelByMonth() {
    if (results.length === 0) {
        alert('Prima devi eseguire il matching e generare le ricevute!');
        return;
    }

    // Calcola totali rimborsi e risparmio fiscale
    const totaleRimborsi = results.reduce((sum, person) => sum + person.rimborsoSpese, 0);
    const risparmioFiscale = totaleRimborsi * 0.20;
    
    // Mostra popup informativo
    const conferma = confirm(
        `RIEPILOGO RIMBORSI SPESE\n\n` +
        `• Totale rimborsi spese: € ${totaleRimborsi.toFixed(2)}\n` +
        `• Risparmio fiscale (20%): € ${risparmioFiscale.toFixed(2)}\n\n` +
        `Procedere con l'export Excel per mese?`
    );
    
    if (!conferma) return;

    // Raggruppa le ricevute per mese
    const ricevutePerMese = {};
    
    results.forEach(person => {
        const chiaveMese = `${person.anno}-${person.mese.toString().padStart(2, '0')}`;
        if (!ricevutePerMese[chiaveMese]) {
            ricevutePerMese[chiaveMese] = [];
        }
        ricevutePerMese[chiaveMese].push(person);
    });

    // Crea un file Excel per ogni mese
    Object.keys(ricevutePerMese).forEach(meseAnno => {
        const ricevuteMese = ricevutePerMese[meseAnno];
        const [anno, mese] = meseAnno.split('-');
        const mesiNomi = ['', 'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                         'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];

        const excelData = [];
        
        // Intestazione pulita
        excelData.push([
            'Id Paese',
            'Partita Iva', 
            'Codice Fiscale',
            'Denominazione',
            'Cognome',
            'Nome',
            'Indirizzo',
            'Num. civico',
            'CAP',
            'Comune',
            'Provincia',
            'Causale',
            'Sezionale',
            'Tipo Doc.',
            'Data Doc.',
            'Numero Doc.',
            'Data doc. fattura origine',
            'Num. doc. fattura origine',
            'Descr. Articolo1',
            'Imponibile1',
            'Aliquota IVA1',
            'Natura IVA1',
            'Codice IVA1',
            'Imposta1',
            'Totale Imponibile',
            'Totale Imposta',
            'Totale Documento',
            'Esigibilita\' IVA'
        ]);

        // Dati per questo mese
        ricevuteMese.forEach(person => {
            const cfKey = person.codiceFiscale || `${person.nome}_${person.cognome}`;
            const receiptNumber = getCurrentReceiptNumber(cfKey);
            const dataDoc = new Date().toLocaleDateString('it-IT');
            const denominazione = `${person.nome} ${person.cognome}`;
            
            // Separa indirizzo e numero civico
            const indirizzoCompleto = person.indirizzo || '';
            const indirizzoParti = indirizzoCompleto.split(' ');
            let indirizzo = '';
            let numCivico = '';
            
            if (indirizzoParti.length > 0) {
                const ultimaParte = indirizzoParti[indirizzoParti.length - 1];
                if (/^\d+[a-zA-Z]?$/.test(ultimaParte)) {
                    numCivico = ultimaParte;
                    indirizzo = indirizzoParti.slice(0, -1).join(' ');
                } else {
                    indirizzo = indirizzoCompleto;
                }
            }

            excelData.push([
                'IT',
                person.partitaIva || '',
                person.codiceFiscale,
                denominazione,
                person.cognome,
                person.nome,
                indirizzo,
                numCivico,
                person.cap || '',
                person.citta || '',
                person.provincia || '',
                '135',
                '4',
                'TD01', // Corretto da TF01 a TD01
                dataDoc,
                receiptNumber,
                dataDoc,
                '1',
                'COMPENSO PER PRESTAZIONE DI LAVORO AUTONOMO OCCASIONALE',
                person.compenso.toFixed(2),
                '0',
                'N2',
                'NI',
                '0',
                person.compenso.toFixed(2),
                '0',
                person.compenso.toFixed(2),
                'I'
            ]);
        });

        // Crea e scarica il file Excel per questo mese
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, `${mesiNomi[parseInt(mese)]} ${anno}`);
        
        const fileName = `Export_Ricevute_${mesiNomi[parseInt(mese)]}_${anno}.xlsx`;
        XLSX.writeFile(wb, fileName);
    });

    const numMesi = Object.keys(ricevutePerMese).length;
    alert(`Generati ${numMesi} file Excel (uno per ogni mese)!`);
}
