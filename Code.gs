// Restituisce sempre il foglio dati corretto cercandolo per NOME ("Dati"),
// invece di affidarsi alla sua posizione tra le schede del file (che può
// cambiare, es. se altri fogli come "Mappa Nazionalità" o "Report" vengono
// spostati prima di esso). Se per qualche motivo il foglio "Dati" non viene
// trovato con questo nome esatto, usa comunque il primo foglio come ripiego.
function getFoglioDati() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Archivio mediazione') || ss.getSheets()[0];
}

function doPost(e) {
  const sheet = getFoglioDati();
  const data = JSON.parse(e.postData.contents);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const id = String(data._id || Date.now().toString());
  data._id = id;
  
  // Cerca riga esistente con stesso ID
  const idCol = sheet.getRange(2, 1, Math.max(sheet.getLastRow()-1, 1), 1).getValues();
  for (let i = 0; i < idCol.length; i++) {
    if (String(idCol[i][0]) === id) {
      // Aggiorna riga esistente
      const row = headers.map(h => data[h] !== undefined ? String(data[h]) : '');
      sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({status:'updated', id:id}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // Nuova riga
  const row = headers.map(h => h === '_id' ? id : (data[h] !== undefined ? String(data[h]) : ''));
  sheet.appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({status:'created', id:id}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // Modalità diagnostica: apri l'URL con ?debug=1 in fondo per verificare che il
  // deployment sia aggiornato, senza restituire nessun dato personale dei beneficiari.
  if (e && e.parameter && e.parameter.debug === '1') {
    const sheet = getFoglioDati();
    return ContentService.createTextOutput(JSON.stringify({
      versione: 'getFoglioDati-2026-08-11',
      foglioUsato: sheet.getName(),
      righeTotali: Math.max(sheet.getLastRow() - 1, 0)
    })).setMimeType(ContentService.MimeType.JSON);
  }
  const sheet = getFoglioDati();
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return ContentService.createTextOutput('[]')
    .setMimeType(ContentService.MimeType.JSON);
  const headers = rows[0];
  const fusoOrario = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => {
      const val = row[i];
      // Se Google Sheets ha convertito automaticamente il testo in un vero valore Data
      // (succede spesso con stringhe tipo "2024-09-12"), lo riformatto qui esplicitamente
      // nel fuso orario del foglio, PRIMA che JSON.stringify lo converta da solo in UTC
      // spostando la data di un giorno indietro (bug che causava lo "slittamento" della data).
      obj[h] = (val instanceof Date) ? Utilities.formatDate(val, fusoOrario, 'yyyy-MM-dd') : val;
    });
    return obj;
  });
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function debugHeaders() {
  const sheet = getFoglioDati();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Numero colonne: ' + headers.length);
  Logger.log('Prime 5: ' + headers.slice(0,5).join(' | '));
  Logger.log('Prima cella: "' + headers[0] + '" lunghezza: ' + headers[0].length);
}

// Trova i minori con un anno di nascita palesemente errato: il report ora usa
// automaticamente l'età attuale quando un minore non era ancora nato alla data
// del colloquio (es. schede inserite retroattivamente per bambini nati nel
// frattempo), quindi questa funzione segnala solo i casi in cui l'anno di
// nascita è nel FUTURO anche rispetto a oggi — quasi certamente un errore di
// battitura (es. 2025 invece di 2015). Esegui questa funzione, guarda i
// risultati in Log (Visualizza > Log oppure Ctrl+Cronologia esecuzioni) o nel
// popup, e correggi la cella indicata direttamente nel foglio dati.
function trovaMinoriEtaNonDisponibile() {
  const sheet = getFoglioDati();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getDataRange().getValues();

  function idx(nome) { return headers.indexOf(nome); }
  const idxId = idx('_id');
  const annoCorrente = new Date().getFullYear();

  const problemi = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row[0]) continue; // riga vuota

    for (let i = 1; i <= 6; i++) {
      const idxAnno = idx('minore_' + i + '_anno');
      if (idxAnno < 0) continue;
      const anno = parseInt(row[idxAnno]);
      if (!anno) continue;
      const eta = annoCorrente - anno;
      if (eta < 0) {
        const id = idxId >= 0 ? row[idxId] : '(senza _id)';
        problemi.push(
          'Riga foglio N° ' + (r + 1) +
          ' | _id: ' + id +
          ' | minore_' + i + '_anno = ' + anno +
          ' | anno corrente: ' + annoCorrente +
          ' | età calcolata: ' + eta + ' (anno di nascita nel futuro: probabile errore di battitura)'
        );
      }
    }
  }

  if (problemi.length === 0) {
    SpreadsheetApp.getUi().alert('Nessuna anomalia trovata: tutte le età dei minori risultano valide.');
    return;
  }

  Logger.log(problemi.join('\n'));
  SpreadsheetApp.getUi().alert(
    'Trovate ' + problemi.length + ' anomalie:\n\n' +
    problemi.join('\n\n') +
    '\n\nCorreggi l\'anno di nascita nella cella indicata (colonna minore_N_anno) sulla riga del foglio dati.'
  );
}


function verificaColonneNuove() {
  const sheet = getFoglioDati();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  const attese = ['istruzione_origine', 'istruzione_italia', 'n_minori'];
  for (let i = 1; i <= 6; i++) {
    attese.push('minore_' + i + '_anno');
    attese.push('minore_' + i + '_sesso');
  }

  const mancanti = attese.filter(nome => headers.indexOf(nome) === -1);
  const presenti = attese.filter(nome => headers.indexOf(nome) !== -1);

  Logger.log('COLONNE PRESENTI (' + presenti.length + '/' + attese.length + '): ' + presenti.join(', '));
  Logger.log('COLONNE MANCANTI: ' + (mancanti.length ? mancanti.join(', ') : 'nessuna'));

  const msg = mancanti.length
    ? 'Mancano queste colonne in riga 1 del foglio dati:\n\n' + mancanti.join('\n') +
      '\n\nAggiungile come nuove intestazioni (esattamente con questo nome) prima di generare il report.'
    : 'Tutte le colonne nuove sono presenti. Se il report è comunque vuoto, il problema è che non ci sono ancora righe compilate con questi dati.';

  SpreadsheetApp.getUi().alert(msg);
}

// Etichette leggibili per i livelli ISCED 0-8 (i valori salvati nella scheda sono i codici numerici)
const ISCED_LABELS = {
  '0': 'ISCED 0 - Educazione prima infanzia',
  '1': 'ISCED 1 - Istruzione primaria',
  '2': 'ISCED 2 - Istruzione secondaria inferiore',
  '3': 'ISCED 3 - Istruzione secondaria superiore',
  '4': 'ISCED 4 - Post-secondaria non terziaria',
  '5': 'ISCED 5 - Terziaria breve (tecnica)',
  '6': 'ISCED 6 - Laurea / Bachelor',
  '7': 'ISCED 7 - Laurea magistrale / Master',
  '8': 'ISCED 8 - Dottorato'
};

function generaReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const foglioD = getFoglioDati();
  
  let report = ss.getSheetByName('Report');
  if (!report) report = ss.insertSheet('Report');
  else report.clearContents();
  
  const headers = foglioD.getRange(1, 1, 1, foglioD.getLastColumn()).getValues()[0];
  const rows = foglioD.getDataRange().getValues().slice(1).filter(r => r[0] !== '');
  
  function idx(nome) { return headers.indexOf(nome); }
  
  function conta(colonna, pulisci) {
    const map = {};
    rows.forEach(r => {
      let val = String(r[idx(colonna)] || '').trim();
      if (!val) return;
      if (pulisci) val = val.split('—')[0].split('-')[0].trim();
      val.split(',').forEach(v => {
        v = v.trim();
        if (v) map[v] = (map[v] || 0) + 1;
      });
    });
    return Object.entries(map).sort((a,b) => b[1]-a[1]);
  }

  // Conta la nazionalità/residenza dell'adulto in ogni riga, e aggiunge al totale
  // anche i minori presenti in quel nucleo (stessa nazionalità/residenza della famiglia).
  function contaConMinori(colonna, pulisci) {
    const map = {};
    rows.forEach(r => {
      let val = String(r[idx(colonna)] || '').trim();
      if (!val) return;
      if (pulisci) val = val.split('—')[0].split('-')[0].trim();
      let numMinoriRiga = 0;
      for (let i = 1; i <= 6; i++) {
        const idxAnno = idx('minore_' + i + '_anno');
        if (idxAnno >= 0 && parseInt(r[idxAnno])) numMinoriRiga++;
      }
      const peso = 1 + numMinoriRiga;
      val.split(',').forEach(v => {
        v = v.trim();
        if (v) map[v] = (map[v] || 0) + peso;
      });
    });
    return Object.entries(map).sort((a,b) => b[1]-a[1]);
  }

  // Come conta(), ma traduce i codici salvati (es. ISCED 0-8) in etichette leggibili
  function contaConMappa(colonna, mappa) {
    const map = {};
    rows.forEach(r => {
      const val = String(r[idx(colonna)] || '').trim();
      if (!val) return;
      const label = mappa[val] || val;
      map[label] = (map[label] || 0) + 1;
    });
    return Object.entries(map).sort((a,b) => b[1]-a[1]);
  }

  // Età calcolata dinamicamente (anno corrente - anno di nascita) dalla data di nascita completa
  // Adulti: conteggio per anno di nascita (non più a fasce), per poter incrociare
  // con gli anni in cui sono avvenuti gli interventi.
  function contaAnnoNascita() {
    const map = {};
    rows.forEach(r => {
      const dn = r[idx('data_nascita')];
      let anno = 'Non disponibile';
      if (dn) {
        const nascita = new Date(dn);
        if (!isNaN(nascita)) anno = String(nascita.getFullYear());
      }
      map[anno] = (map[anno] || 0) + 1;
    });
    const entries = Object.entries(map);
    entries.sort((a, b) => {
      if (a[0] === 'Non disponibile') return 1;
      if (b[0] === 'Non disponibile') return -1;
      return Number(a[0]) - Number(b[0]);
    });
    return entries;
  }

  // Conteggio degli interventi (int_1_data ... int_8_data) per anno, da confrontare
  // con la distribuzione per anno di nascita qui sopra.
  function contaInterventiPerAnno() {
    const map = {};
    rows.forEach(r => {
      for (let i = 1; i <= 8; i++) {
        const idxData = idx('int_' + i + '_data');
        if (idxData < 0) continue;
        const val = r[idxData];
        if (!val) continue;
        const d = new Date(val);
        if (isNaN(d)) continue;
        const anno = String(d.getFullYear());
        map[anno] = (map[anno] || 0) + 1;
      }
    });
    return Object.entries(map).sort((a, b) => Number(a[0]) - Number(b[0]));
  }

  // Dati sui minori: per ciascuno dei 6 slot per nucleo legge anno di nascita e sesso,
  // calcola l'età come (anno del colloquio della riga - anno di nascita) e aggrega per età e per sesso.
  // Minori: età singola e sesso, non più a fasce.
  // Se manca la data colloquio su una riga si usa l'anno corrente come fallback.
  // Se il minore non era ancora nato alla data del colloquio (es. schede inserite
  // retroattivamente per bambini nati nel frattempo), si usa l'età attuale (anno
  // corrente - anno di nascita) invece di segnare "Non disponibile".
  function datiMinori() {
    const idxColloquio = idx('data_colloquio');
    const annoCorrente = new Date().getFullYear();
    const eta = {};
    const sesso = {};
    let totale = 0;
    rows.forEach(r => {
      const dc = idxColloquio >= 0 ? r[idxColloquio] : '';
      const dataColloquio = dc ? new Date(dc) : null;
      const annoRif = (dataColloquio && !isNaN(dataColloquio)) ? dataColloquio.getFullYear() : annoCorrente;
      for (let i = 1; i <= 6; i++) {
        const idxAnno = idx('minore_' + i + '_anno');
        const idxSesso = idx('minore_' + i + '_sesso');
        const annoRaw = idxAnno >= 0 ? r[idxAnno] : '';
        const anno = parseInt(annoRaw);
        if (!anno) continue;
        totale++;
        let e = annoRif - anno;
        if (e < 0) e = annoCorrente - anno; // non ancora nato al colloquio: uso l'età attuale
        const etaLabel = e >= 0 ? (e + ' anni') : 'Non disponibile';
        eta[etaLabel] = (eta[etaLabel] || 0) + 1;
        let s = idxSesso >= 0 ? String(r[idxSesso] || '').trim() : '';
        if (!s) s = 'Non rilevato';
        sesso[s] = (sesso[s] || 0) + 1;
      }
    });
    return { totale, eta, sesso };
  }
  
  // Conta semplicemente quante schede hanno "Cittadino italiano" spuntato nel campo permesso
  function contaCittadinanzaItaliana() {
    let totale = 0;
    rows.forEach(r => {
      const val = String(r[idx('permesso')] || '').trim();
      if (!val) return;
      const valori = val.split(',').map(v => v.trim());
      if (valori.includes('Cittadino italiano')) totale++;
    });
    return [['Cittadino italiano (cittadinanza acquisita)', totale], ['Totale schede', rows.length]];
  }

  const sezioni = [
    ['GENERE', conta('genere')],
    ['ANNO DI NASCITA (ADULTI)', contaAnnoNascita()],
    ['INTERVENTI PER ANNO', contaInterventiPerAnno()],
    ['CITTADINANZA ITALIANA ACQUISITA', contaCittadinanzaItaliana()],
    ['NAZIONALITÀ (adulti + minori del nucleo)', contaConMinori('nazionalita')],
    ['TITOLO DI STUDIO - PAESE D\'ORIGINE', contaConMappa('istruzione_origine', ISCED_LABELS)],
    ['TITOLO DI STUDIO - ITALIA', contaConMappa('istruzione_italia', ISCED_LABELS)],
    ['COMUNI DI RESIDENZA (adulti + minori del nucleo)', contaConMinori('residenza', true)],
    ['ANNO DI ARRIVO IN ITALIA', conta('anno_arrivo')],
    ['AREE DI BISOGNO / MOTIVO DEL COLLOQUIO', conta('area')],
    ['SERVIZI / ENTI DI INVIO', conta('servizi_invio')],
    ['TIPO DI PERMESSO', conta('permesso')],
    ['TIPO DI SPOSTAMENTO', conta('tipo_spostamento')],
    ['DURATA SPOSTAMENTO PREVISTA', conta('durata_spostamento')],
    ['SETTORE LAVORATIVO', conta('lavoro')],
  ];
  
  const output = [];
  sezioni.forEach(([titolo, dati]) => {
    output.push([titolo, '']);
    output.push(['Valore', 'N°']);
    if (dati.length === 0) {
      output.push(['(nessun dato)', '']);
    } else {
      dati.forEach(([k, v]) => output.push([k, v]));
    }
    output.push(['', '']);
  });

  // Sezione minori (anno di nascita, sesso, età calcolata dinamicamente)
  const minori = datiMinori();
  output.push(['MINORI NEL NUCLEO FAMILIARE - ETÀ', '']);
  output.push(['Età', 'Totale']);
  Object.entries(minori.eta).sort((a, b) => {
    const na = parseInt(a[0]), nb = parseInt(b[0]);
    if (isNaN(na)) return 1;
    if (isNaN(nb)) return -1;
    return na - nb;
  }).forEach(([k, v]) => output.push([k, v]));
  output.push(['TOTALE MINORI', minori.totale]);
  output.push(['', '']);

  output.push(['MINORI NEL NUCLEO FAMILIARE - SESSO', '']);
  output.push(['Sesso', 'Totale']);
  Object.entries(minori.sesso).sort((a,b) => b[1]-a[1]).forEach(([k,v]) => output.push([k, v]));
  output.push(['', '']);

  output.push(['TOTALE PERSONE IN ARCHIVIO', rows.length]);
  output.push(['TOTALE MINORI NEL NUCLEO', minori.totale]);
  output.push(['TOTALE PERSONE COMPLESSIVAMENTE SEGUITE', rows.length + minori.totale]);
  
  report.getRange(1, 1, output.length, 2).setValues(output);

  // Aggiorna anche la mappa del mondo con la distribuzione delle nazionalità,
  // riusando gli stessi dati già calcolati per la sezione "NAZIONALITÀ" del report qui sopra.
  generaMappaNazionalita(contaConMinori('nazionalita'));

  Logger.log('Report completato: ' + output.length + ' righe');
}

// Scrive/aggiorna la tabella dati (Paese | Persone) sul foglio dedicato "Mappa Nazionalità".
// Viene richiamata automaticamente da generaReport(), quindi la tabella si aggiorna da sola
// ogni volta che generi il report. Il grafico va inserito UNA SOLA VOLTA manualmente
// (Inserisci → Grafico → Grafico geografico, selezionando l'intervallo A1:B su questo foglio):
// una volta collegato all'intervallo, il grafico nativo di Sheets si aggiorna da solo ogni
// volta che questi dati cambiano, senza bisogno che lo script lo ricrei.
// NOTA: la mappa riconosce i nomi dei paesi scritti in italiano (es. "Marocco", "Nigeria",
// "Albania") nella maggior parte dei casi; nomi imprecisi o abbreviati potrebbero non
// comparire sulla mappa pur restando conteggiati nella tabella.
function generaMappaNazionalita(datiNazionalita) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mappaSheet = ss.getSheetByName('Mappa Nazionalità');
  if (!mappaSheet) mappaSheet = ss.insertSheet('Mappa Nazionalità');

  // Pulisce solo i valori delle celle, senza toccare eventuali grafici già presenti sul foglio
  mappaSheet.getRange(1, 1, Math.max(mappaSheet.getLastRow(), 1), 2).clearContent();

  mappaSheet.getRange(1, 1, 1, 2).setValues([['Paese', 'Persone']]);
  if (datiNazionalita.length > 0) {
    mappaSheet.getRange(2, 1, datiNazionalita.length, 2).setValues(datiNazionalita);
  }
}

function convertiNazioni() {
  const sheet = getFoglioDati();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colNaz = headers.indexOf('nazionalita');
  if (colNaz < 0) { Logger.log('Colonna nazionalita non trovata'); return; }
  
  const codici = {
    'AFG':'Afghanistan','ALB':'Albania','DZA':'Algeria','AGO':'Angola',
    'BGD':'Bangladesh','BEN':'Benin','BFA':'Burkina Faso','BDI':'Burundi',
    'CMR':'Camerun','CPV':'Capo Verde','CAF':'Rep. Centrafricana',
    'TCD':'Ciad','CIV':'Costa d\'Avorio','COD':'Rep. Dem. Congo',
    'COG':'Congo','EGY':'Egitto','ERI':'Eritrea','ETH':'Etiopia',
    'GAB':'Gabon','GMB':'Gambia','GHA':'Ghana','GIN':'Guinea',
    'GNB':'Guinea-Bissau','GNQ':'Guinea Equatoriale','IND':'India',
    'IRQ':'Iraq','IRN':'Iran','JOR':'Giordania','KEN':'Kenya',
    'LBR':'Liberia','LBY':'Libia','MDG':'Madagascar','MWI':'Malawi',
    'MLI':'Mali','MRT':'Mauritania','MAR':'Marocco','MOZ':'Mozambico',
    'MMR':'Myanmar','NGA':'Nigeria','NER':'Niger','PAK':'Pakistan',
    'PSE':'Palestina','PHL':'Filippine','RWA':'Ruanda','SEN':'Senegal',
    'SLE':'Sierra Leone','SOM':'Somalia','SDN':'Sudan','SSD':'Sudan del Sud',
    'SYR':'Siria','TZA':'Tanzania','TGO':'Togo','TUN':'Tunisia',
    'UGA':'Uganda','UKR':'Ucraina','YEM':'Yemen','ZMB':'Zambia',
    'ZWE':'Zimbabwe','ROM':'Romania','BGR':'Bulgaria','MDA':'Moldova',
    'SRB':'Serbia','BLR':'Bielorussia','CHN':'Cina','VNM':'Vietnam',
    'LKA':'Sri Lanka','NPL':'Nepal','ECU':'Ecuador','PER':'Perù',
    'COL':'Colombia','BRA':'Brasile','DOM':'Rep. Dominicana','CUB':'Cuba',
    'MEX':'Messico','TUR':'Turchia','KGZ':'Kirghizistan','UZB':'Uzbekistan',
    'TKM':'Turkmenistan','KAZ':'Kazakhstan','AZE':'Azerbaigian',
    'GEO':'Georgia','ARM':'Armenia'
  };
  
  const rows = sheet.getDataRange().getValues();
  const nuoviValori = [];
  let modificate = 0;
  
  for (let i = 1; i < rows.length; i++) {
    const codice = String(rows[i][colNaz] || '').trim().toUpperCase();
    if (codici[codice]) {
      nuoviValori.push({row: i + 1, val: codici[codice]});
      modificate++;
    }
  }
  
  // Scrivi tutto in batch
  nuoviValori.forEach(item => {
    sheet.getRange(item.row, colNaz + 1).setValue(item.val);
  });
  
  SpreadsheetApp.flush();
  Logger.log('Convertite: ' + modificate);
  SpreadsheetApp.getUi().alert('✓ Completato! ' + modificate + ' nazionalità convertite.');
}

// DIAGNOSTICA: conta quante schede hanno almeno un intervento con data compilata,
// e quante celle int_N_data risultano effettivamente valorizzate nel foglio.
// Aiuta a distinguere tra "i dati mancano davvero" e "le colonne non vengono trovate".
function diagnosticaInterventi() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archivio mediazione');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getDataRange().getValues().slice(1).filter(r => r[0] !== '');
  function idx(nome) { return headers.indexOf(nome); }

  let colonneNonTrovate = [];
  for (let i = 1; i <= 8; i++) {
    if (idx('int_' + i + '_data') < 0) colonneNonTrovate.push('int_' + i + '_data');
  }

  let schedeConAlmenoUnIntervento = 0;
  let celleValorizzate = 0;
  let totaleSchede = rows.length;

  rows.forEach(r => {
    let haInterventi = false;
    for (let i = 1; i <= 8; i++) {
      const idxData = idx('int_' + i + '_data');
      if (idxData < 0) continue;
      const val = r[idxData];
      if (val) {
        celleValorizzate++;
        haInterventi = true;
      }
    }
    if (haInterventi) schedeConAlmenoUnIntervento++;
  });

  const msg = 'Totale schede: ' + totaleSchede +
    '\nSchede con almeno un intervento compilato: ' + schedeConAlmenoUnIntervento +
    '\nCelle int_N_data valorizzate in totale: ' + celleValorizzate +
    '\nColonne int_N_data NON trovate nell\'intestazione: ' + (colonneNonTrovate.length ? colonneNonTrovate.join(', ') : 'nessuna (tutte trovate)');

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// DIAGNOSTICA: individua righe con _id mancante, vuoto, o duplicato nel foglio dati.
// Un _id duplicato o "sporco" (spazi, formattazione diversa) è la causa più comune
// del problema "modifico una scheda ma la modifica non si salva sulla riga giusta":
// il salvataggio cerca una riga con lo stesso _id e, se non la trova con corrispondenza
// esatta, ne crea una nuova invece di aggiornare quella esistente.
function trovaIdDuplicatiOMancanti() {
  const sheet = getFoglioDati(); // stesso foglio usato da doPost/doGet
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idColIndex = headers.indexOf('_id');
  if (idColIndex !== 0) {
    SpreadsheetApp.getUi().alert('Attenzione: la colonna "_id" non è la prima colonna (A) del foglio, oppure non è stata trovata. Questo può causare problemi di salvataggio. Colonna trovata: ' + (idColIndex + 1));
    return;
  }

  const rows = sheet.getDataRange().getValues().slice(1);
  const conteggio = {};
  const vuoti = [];
  rows.forEach((r, i) => {
    const id = String(r[0]).trim();
    const rigaFoglio = i + 2; // riga reale nel foglio (1 = intestazione)
    if (!id) { vuoti.push(rigaFoglio); return; }
    if (!conteggio[id]) conteggio[id] = [];
    conteggio[id].push(rigaFoglio);
  });

  const duplicati = Object.entries(conteggio).filter(([, righe]) => righe.length > 1);

  let msg = 'Righe totali: ' + rows.length + '\n';
  msg += 'Righe con _id vuoto: ' + vuoti.length + (vuoti.length ? ' (righe foglio: ' + vuoti.join(', ') + ')' : '') + '\n';
  msg += 'ID duplicati trovati: ' + duplicati.length;
  if (duplicati.length) {
    msg += '\n' + duplicati.map(([id, righe]) => 'ID ' + id + ' → righe ' + righe.join(', ')).join('\n');
  }

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// FIX: assegna un _id univoco a tutte le righe che ne sono prive (colonna A vuota).
// Una scheda senza _id viene sempre trattata come "nuova" ad ogni salvataggio,
// quindi ogni modifica crea una riga duplicata invece di aggiornare quella esistente
// (è la causa del problema "modifico la data ma dopo il reload torna quella vecchia").
// Esegui questa funzione UNA VOLTA per sistemare le righe già presenti nel foglio.
function assegnaIdMancanti() {
  const sheet = getFoglioDati();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getActiveSpreadsheet().toast('Nessuna riga dati trovata.'); return; }

  const idRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const valori = idRange.getValues();
  let corrette = 0;

  valori.forEach((r, i) => {
    if (!String(r[0]).trim()) {
      // Timestamp + indice riga per garantire l'unicità anche se assegnati nello stesso istante
      valori[i][0] = String(Date.now()) + '_' + (i + 2);
      corrette++;
    }
  });

  idRange.setValues(valori);
  const msg = 'Righe corrette: ' + corrette;
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Assegna ID mancanti', 8);
}

// DIAGNOSTICA: controlla, tra le schede già inserite, quali hanno la data
// colloquio o una data di intervento (1-8) che cade di sabato o domenica.
// Utile per individuare inserimenti retroattivi da correggere.
function trovaDateWeekendEsistenti() {
  const sheet = getFoglioDati();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getDataRange().getValues().slice(1).filter(r => r[0] !== '');
  function idx(nome) { return headers.indexOf(nome); }
  const idxId = idx('_id');
  const idxCognome = idx('cognome');
  const idxNome = idx('nome');

  // Restituisce il giorno della settimana (0=domenica..6=sabato) di una data,
  // gestendo sia stringhe "YYYY-MM-DD" sia veri valori Data del foglio.
  function giornoSettimana(val) {
    if (!val) return null;
    let d;
    if (val instanceof Date) {
      d = val;
    } else {
      const parti = String(val).split('-');
      if (parti.length !== 3) return null;
      d = new Date(parseInt(parti[0]), parseInt(parti[1]) - 1, parseInt(parti[2]));
    }
    if (isNaN(d)) return null;
    return d.getDay();
  }

  function eWeekend(val) {
    const g = giornoSettimana(val);
    return g === 0 || g === 6;
  }

  const problemi = [];

  rows.forEach(row => {
    const id = idxId >= 0 ? row[idxId] : '(senza _id)';
    const nomeCompleto = ((idxCognome >= 0 ? row[idxCognome] : '') + ' ' + (idxNome >= 0 ? row[idxNome] : '')).trim();
    const etichetta = nomeCompleto || ('_id: ' + id);

    const idxDataColloquio = idx('data_colloquio');
    if (idxDataColloquio >= 0 && eWeekend(row[idxDataColloquio])) {
      problemi.push(etichetta + ' → Data colloquio: ' + row[idxDataColloquio]);
    }

    for (let i = 1; i <= 8; i++) {
      const idxData = idx('int_' + i + '_data');
      if (idxData < 0) continue;
      if (eWeekend(row[idxData])) {
        problemi.push(etichetta + ' → Intervento ' + i + ': ' + row[idxData]);
      }
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let esitoSheet = ss.getSheetByName('Controllo Weekend');
  if (!esitoSheet) esitoSheet = ss.insertSheet('Controllo Weekend');
  esitoSheet.clearContents();
  esitoSheet.getRange(1, 1).setValue('Controllo eseguito il: ' + new Date().toLocaleString('it-IT'));

  if (problemi.length === 0) {
    esitoSheet.getRange(2, 1).setValue('Nessuna data di sabato o domenica trovata tra le schede inserite.');
  } else {
    esitoSheet.getRange(2, 1).setValue('Trovate ' + problemi.length + ' date di sabato/domenica:');
    esitoSheet.getRange(3, 1, problemi.length, 1).setValues(problemi.map(p => [p]));
  }

  return 'Controllo completato: vedi il foglio "Controllo Weekend" per il dettaglio.';
}
