/**
 * Sistema gestione Consigli di Classe - ITIS G. Riva
 * Backend Apps Script
 *
 * Lo script crea automaticamente i fogli di servizio (UTENTI, PARTECIPAZIONI, OTP)
 * dentro SPREADSHEET_ID al primo avvio. Lanciare diagnostica() una tantum per
 * verificare che la lettura dei fogli sorgente sia corretta.
 */

// ============================ CONFIGURAZIONE ============================

const ID_ATTIVITA = "1fiPLKYT6rjyRUZPPbU56kRyu9uiHlNxQLExRz7dDAh8";
const SPREADSHEET_ID  = "1fPnwpMwAES-LFQYklYwe5YryAslJsi3teW_j3qfdFoQ"; // Foglio DB privato
const ID_RUBRICA      = "1VZhqw7q1Ss95_XpIJIr0cD9-EVLIb-W060Z3SwKlyUU";
const ID_CATTEDRE     = "1Ac7Dzy2OZ_qf3zVJD08gMmkYl9ewk_6HzJ58ADqSzm8";
const ID_CALENDARIO   = "1-XLzS6yBUkau5o5YnMxEFBoLbaP5-8rO6Akj3vCfBpw";
const ID_COORDINATORI = "1XPni3WHn6dEcl4oejAV-pOqqX_RE45WgGred07LekNk";
const ID_ORARIO_CDC = "1-L5zUNvpFa9MluOnGpcXfnVryTESqzqzQ9Zbl5-jB0Y";

const ADMIN_EMAIL     = "vatf02006@istruzione.edu.it";
const DOMINIO_DOCENTI = "@itisriva.edu.it";
const DEFAULT_PWD     = "BardoMagno26";

const ORE_FULL_TIME   = 18;     // Cattedra piena
const ORE_CDC_DOVUTE  = 40;     // Ore CdC dovute al full time
const OTP_VALID_MIN   = 15;     // Validità OTP in minuti

const GIORNI_SETTIMANA = ["Lunedì","Martedì","Mercoledì","Giovedì","Venerdì","Sabato","Domenica"];
const GIORNI_BREVI     = ["LUN","MAR","MER","GIO","VEN","SAB","DOM"];
const SLOT_ORARI_CDC = ["08h10", "09h10", "10h10", "11h10", "12h10", "13h10", "14h10", "15h10"];


// Mappatura giorno lettera -> numero ISO (lunedì=1)
const MAPPA_GIORNO = {
  "L":1,"M":2,"M":2,"G":4,"V":5,"S":6,"D":7
};

// ============================ ENTRY POINT WEBAPP ============================

function doGet(e) {
  var output = HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Gestione CdC - ITIS G. Riva')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  // Questa riga è quella che abilita la modalità "App" su Android e iOS
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  
  return output;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================ INGESTION DA CALENDARIO ATTIVITÀ ============================

function caricaOrario() {
  loaderDash(true);
  google.script.run
    .withSuccessHandler(function(res) {
      loaderDash(false);
      if (res.error) return alert(res.error);
      renderOrario(res);
    })
    .withFailureHandler(function(err) {
      loaderDash(false);
      alert('Errore: ' + err.message);
    })
    .getMioOrario(stateUser.email);
}

function getMioOrario(email) {
  email = String(email || "").trim().toLowerCase();
  const utente = _getUtente(email);
  if (!utente) return { error: "Utente non trovato." };
  
  const nomeCompleto = (utente.COGNOME + " " + utente.NOME).trim();
  
  try {
    const orario = _getOrarioDocente(nomeCompleto);
    return { success: true, nome: nomeCompleto, orario: orario };
  } catch(e) {
    return { error: "Errore nel recupero orario: " + e.toString() };
  }
}

function _getOrarioDocente(nomeDocente) {
  var nomeUpper = nomeDocente.toUpperCase().trim();
  var ss = SpreadsheetApp.openById(ID_ORARIO_CDC);
  var dati = ss.getSheets()[0].getDataRange().getValues();
  
  var giorni = ["lunedì","martedì","mercoledì","giovedì","venerdì"];
  var griglia = {};
  giorni.forEach(function(g) { griglia[g] = {}; });

  for (var i = 1; i < dati.length; i++) {
    var r = dati[i];
    var giorno = r[10] ? r[10].toString().toLowerCase().trim() : "";
    if (giorni.indexOf(giorno) === -1) continue;

    var docentiInCella = r[5] ? r[5].toString().toUpperCase().split(",")
                               .map(function(s) { return s.trim(); }) : [];

    // Cerca il docente nella cella (match flessibile)
    var trovato = false;
    var partiNome = nomeUpper.split(" ");
    var cognome = partiNome[0];
    
    trovato = docentiInCella.some(function(d) {
      return d === nomeUpper || d === cognome || nomeUpper.includes(d) || d.includes(cognome);
    });
    
    if (!trovato) continue;

    var oraString = r[11] ? r[11].toString().trim() : "";
    var oraInizioIdx = SLOT_ORARI_CDC.indexOf(oraString);
    if (oraInizioIdx === -1) continue;

    var durataRaw = String(r[2] || "1").toLowerCase().trim();
    var durata = 1;
    if (durataRaw.indexOf("h") !== -1) {
      durata = parseInt(durataRaw.split("h")[0], 10) || 1;
    } else {
      durata = parseInt(durataRaw, 10) || 1;
    }

    var classe = r[7] ? String(r[7]).trim() : "";
    var aula = r[8] ? String(r[8]).trim() : "";

    var compresenti = docentiInCella.filter(function(d) {
      return d !== "" && d !== cognome && !nomeUpper.includes(d) && !d.includes(cognome);
    }).join(", ");

    // Per ogni ora di durata, riempi la griglia
    for (var d = 0; d < durata; d++) {
      var oraNum = oraInizioIdx + 1 + d;
      if (oraNum <= 8) {
        griglia[giorno][oraNum] = {
          classe: classe,
          aula: aula,
          compresenti: compresenti
        };
      }
    }
  }

  return griglia;
}

function renderOrario(data) {
  document.getElementById('orario-nome-docente').textContent = 'Docente: ' + data.nome;
  
  const giorni = ["lunedì","martedì","mercoledì","giovedì","venerdì"];
  const oreLabels = ["1ª","2ª","3ª","4ª","5ª","6ª","7ª","8ª"];
  const orario = data.orario;
  
  let html = '<table class="table table-bordered table-sm text-center">';
  html += '<thead class="table-primary"><tr><th>Ora</th>';
  giorni.forEach(function(g) {
    html += '<th>' + g.charAt(0).toUpperCase() + g.slice(1) + '</th>';
  });
  html += '</tr></thead><tbody>';
  
  for (let ora = 1; ora <= 8; ora++) {
    html += '<tr><td class="fw-bold">' + ora + 'ª</td>';
    giorni.forEach(function(g) {
      const cella = (orario[g] && orario[g][ora]) ? orario[g][ora] : null;
      if (cella) {
        html += '<td class="bg-light">';
        html += '<strong>' + escapeHtml(cella.classe) + '</strong><br>';
        if (cella.aula) html += '<small>Aula ' + escapeHtml(cella.aula) + '</small><br>';
        if (cella.compresenti) html += '<small class="text-muted">con ' + escapeHtml(cella.compresenti) + '</small>';
        html += '</td>';
      } else {
        html += '<td class="text-muted">—</td>';
      }
    });
    html += '</tr>';
  }
  
  html += '</tbody></table>';
  document.getElementById('orario-container').innerHTML = html;
}

function _caricaAttivita() {
  return _cachedRead("ATTIVITA", 60, function() { return _caricaAttivitaRaw(); });
}

function _caricaAttivitaRaw() {
  Logger.log("Inizio lettura calendario attività..."); // DEBUG
  const ss = SpreadsheetApp.openById(ID_ATTIVITA);
  const sheets = ss.getSheets();
  
  let shDati = null;
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getLastColumn() > 10) {
      shDati = sheets[i];
      break;
    }
  }
  if (!shDati) shDati = sheets[sheets.length - 1];
  
  const data = shDati.getDataRange().getValues();
  if (data.length < 2) return { attivita: [], mesiDisponibili: [] };
  
  const headerRow = data[0];
  
  // 1. Trova TUTTI i mesi e le loro colonne di inizio
  const blocchiMese = [];
  const mesiNomi = ["SETTEMBRE","OTTOBRE","NOVEMBRE","DICEMBRE","GENNAIO","FEBBRAIO",
                    "MARZO","APRILE","MAGGIO","GIUGNO","LUGLIO","AGOSTO"];
  
  for (let col = 0; col < headerRow.length; col++) {
    const cell = String(headerRow[col] || "").trim().toUpperCase();
    if (mesiNomi.includes(cell)) {
      blocchiMese.push({ mese: cell, colInizio: col });
      Logger.log(`Trovato mese: ${cell} a colonna ${col}`); // DEBUG
    }
  }
  
  // 2. Prepara strutture
  const attivita = [];
  const mesiDisponibili = [];
  const MESE_NUMERO = {
    "GENNAIO":1,"FEBBRAIO":2,"MARZO":3,"APRILE":4,"MAGGIO":5,"GIUGNO":6,
    "LUGLIO":7,"AGOSTO":8,"SETTEMBRE":9,"OTTOBRE":10,"NOVEMBRE":11,"DICEMBRE":12
  };
  const GIORNI_SETTIMANA = ["Lunedì","Martedì","Mercoledì","Giovedì","Venerdì","Sabato","Domenica"];
  const GIORNI_BREVI     = ["LUN","MAR","MER","GIO","VEN","SAB","DOM"];
  
  // 3. Per ogni mese
  blocchiMese.forEach((blocco, index) => {
    const nomeMese = blocco.mese;
    // Determina anno: Sett-Dic 2025, Gen-Ago 2026
    const anno = ["SETTEMBRE","OTTOBRE","NOVEMBRE","DICEMBRE"].includes(nomeMese) ? 2025 : 2026;
    const meseNum = MESE_NUMERO[nomeMese];
    
    Logger.log(`Processando ${nomeMese} ${anno}...`); // DEBUG
    
    mesiDisponibili.push({
      nome: nomeMese.charAt(0) + nomeMese.slice(1).toLowerCase(),
      mese: meseNum,
      anno: anno,
      key: anno + "-" + ("0"+meseNum).slice(-2)
    });
    
    // LE 3 COLONNE DEL MESE: [ Numero | Lettera | Attività ]
    const colNumero   = blocco.colInizio;
    const colLettera  = blocco.colInizio + 1;
    const colAttivita = blocco.colInizio + 2;
    
    Logger.log(`  Colonne: Num=${colNumero}, Lett=${colLettera}, Att=${colAttivita}`); // DEBUG
    
    // Scorri le righe (dalla 1 in poi, la 0 è l'intestazione)
    for (let riga = 1; riga < data.length; riga++) {
      const cellaNumero = data[riga][colNumero];
      const cellaAttivita = data[riga][colAttivita];
      
      // Se il numero è vuoto o non è un numero valido, salta
      if (cellaNumero === "" || cellaNumero === null || cellaNumero === undefined) continue;
      const num = parseInt(cellaNumero, 10);
      if (isNaN(num) || num < 1 || num > 31) continue;
      
      // Controlla se c'è un'attività scritta
      const attivitaTesto = String(cellaAttivita || "").trim();
      if (!attivitaTesto) continue; // Giorno senza attività, non ci interessa
      
      // Crea l'oggetto data
      const dataObj = new Date(anno, meseNum - 1, num);
      if (dataObj.getMonth() + 1 !== meseNum) {
        Logger.log(`  Data non valida: ${num}/${meseNum}/${anno}`); // DEBUG
        continue; // Data inesistente (es. 30 Febbraio)
      }
      
      const giornoSett = dataObj.getDay();
      const giornoIso = giornoSett === 0 ? 7 : giornoSett;
      
      attivita.push({
        data: Utilities.formatDate(dataObj, Session.getScriptTimeZone() || "Europe/Rome", "yyyy-MM-dd"),
        dataIT: `${("0"+num).slice(-2)}/${("0"+meseNum).slice(-2)}/${anno}`,
        giornoSettimana: GIORNI_SETTIMANA[giornoIso - 1],
        giornoBreve: GIORNI_BREVI[giornoIso - 1],
        giornoIso: giornoIso,
        numeroGiorno: num,
        mese: meseNum,
        anno: anno,
        meseNome: nomeMese.charAt(0) + nomeMese.slice(1).toLowerCase(),
        attivita: attivitaTesto
      });
      Logger.log(`    OK: ${num}/${meseNum} -> ${attivitaTesto}`); // DEBUG
    }
  });
  
  Logger.log(`Totale attività trovate: ${attivita.length}`);
  
  return {
    attivita: attivita,
    mesiDisponibili: mesiDisponibili
  };
}

// ============================ API PER IL FRONTEND ============================

/**
 * Restituisce i dati del calendario attività per un mese specifico.
 * @param {number} mese - 1-12
 * @param {number} anno - es. 2025
 */
function getCalendarioAttivita(mese, anno) {
  const dati = _caricaAttivita();
  
  // Filtra attività del mese richiesto
  const attivitaMese = dati.attivita.filter(function(a) {
    return a.mese === mese && a.anno === anno;
  });
  
  // Raggruppa per data
  const mappa = {};
  attivitaMese.forEach(function(a) {
    if (!mappa[a.data]) mappa[a.data] = [];
    mappa[a.data].push(a.attivita);
  });
  
  return {
    mese: mese,
    anno: anno,
    attivita: attivitaMese,
    attivitaPerGiorno: mappa,
    mesiDisponibili: dati.mesiDisponibili
  };
}

/**
 * Dato il nome di un'attività, restituisce tutte le date in cui compare.
 */
function getDettaglioAttivita(nomeAttivita) {
  const dati = _caricaAttivita();
  const nome = String(nomeAttivita || "").trim().toUpperCase();
  
  if (!nome) return { trovata: false, nome: nomeAttivita, date: [] };
  
  const date = dati.attivita.filter(function(a) {
    return a.attivita.toUpperCase() === nome;
  });
  
  // Raggruppa per mese
  const perMese = {};
  date.forEach(function(d) {
    const key = d.anno + "-" + ("0"+d.mese).slice(-2);
    if (!perMese[key]) perMese[key] = { mese: d.meseNome, anno: d.anno, giorni: [] };
    perMese[key].giorni.push({
      data: d.data,
      dataIT: d.dataIT,
      giornoSettimana: d.giornoSettimana,
      numeroGiorno: d.numeroGiorno
    });
  });
  
  // Ordina i giorni dentro ogni mese
  Object.keys(perMese).forEach(function(k) {
    perMese[k].giorni.sort(function(a,b) { return a.data.localeCompare(b.data); });
  });
  
  // Ordina i mesi
  const mesiOrdinati = Object.keys(perMese).sort().map(function(k) {
    return perMese[k];
  });
  
  return {
    trovata: date.length > 0,
    nome: nomeAttivita,
    totale: date.length,
    mesi: mesiOrdinati
  };
}

/**
 * Restituisce l'elenco di tutte le attività distinte presenti nel calendario.
 */
function getElencoAttivita() {
  const dati = _caricaAttivita();
  const mappa = {};
  dati.attivita.forEach(function(a) {
    if (a.attivita) mappa[a.attivita] = (mappa[a.attivita] || 0) + 1;
  });
  
  const elenco = Object.keys(mappa).map(function(nome) {
    return { nome: nome, conteggio: mappa[nome] };
  });
  elenco.sort(function(a,b) { return b.conteggio - a.conteggio; });
  
  return elenco;
}

// Aggiorna anche la funzione diagnostica per includere le attività
// (sostituisci la funzione diagnostica esistente con questa)

function diagnostica() {
  _openDB();
  Logger.log("DB OK: " + SPREADSHEET_ID);

  try {
    const r = _caricaRubricaRaw();
    Logger.log("Rubrica: " + Object.keys(r).length + " docenti caricati.");
  } catch (e) { Logger.log("ERRORE Rubrica: " + e); }

  try {
    const c = _caricaCattedreRaw();
    Logger.log("Cattedre: " + c.length + " righe valide.");
  } catch (e) { Logger.log("ERRORE Cattedre: " + e); }

  try {
    const cal = _caricaCalendarioRaw();
    Logger.log("Calendario CdC: " + cal.length + " CdC trovati.");
  } catch (e) { Logger.log("ERRORE Calendario CdC: " + e); }

  try {
    const co = _caricaCoordinatoriRaw();
    Logger.log("Coordinatori: " + co.length + " classi.");
  } catch (e) { Logger.log("ERRORE Coordinatori: " + e); }
  
  try {
    const att = _caricaAttivitaRaw();
    Logger.log("Calendario Attività: " + att.attivita.length + " attività trovate in " + att.mesiDisponibili.length + " mesi.");
    if (att.attivita.length > 0) {
      Logger.log("Prime 5 attività: " + JSON.stringify(att.attivita.slice(0,5)));
    }
  } catch (e) { Logger.log("ERRORE Calendario Attività: " + e); }
}

// Aggiorna anche invalidaCache per includere le attività
function invalidaCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(["RUBRICA","CATTEDRE","CALENDARIO","COORDINATORI","ATTIVITA"]);
}




// ============================ HELPERS DB ============================

function _openDB() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  _ensureSheet(ss, "UTENTI",         ["EMAIL","COGNOME","NOME","PASSWORD_HASH","SALT","ORE_SETTIMANALI","RUOLO","DATA_REGISTRAZIONE"]);
  _ensureSheet(ss, "PARTECIPAZIONI", ["ID_CDC","EMAIL_DOCENTE","STATO","MARCATO_DA","TIMESTAMP","NOTE"]);
  _ensureSheet(ss, "OTP",            ["EMAIL","CODICE","SCADENZA","USATO"]);
  return ss;
}

function _ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _sheet(name) { return _openDB().getSheetByName(name); }

function _readAll(name) {
  const data = _sheet(name).getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

// ============================ NORMALIZZAZIONE COGNOMI ============================

/**
 * Normalizza un cognome rimuovendo apostrofi, spazi, punti e maiuscolando.
 *  "D'ANGELO" / "D ANGELO" / "DANGELO"  -> "DANGELO"
 *  "DE LUCA"  / "DELUCA"                -> "DELUCA"
 *  "PRAVATA'" / "PRAVATA"               -> "PRAVATA"
 *  "MAUGERI S."                          -> "MAUGERIS"
 */
function _normCognome(s) {
  if (s === null || s === undefined) return "";
  return String(s)
    .toUpperCase()
    .replace(/[''`´]/g, "")
    .replace(/[\.\,]/g, "")
    .replace(/\s+/g, "")
    .trim();
}

/** Estrae cognome dalla parte locale dell'email (nome.cognome@... o nome.secondo.cognome@...). */
function _cognomeDaEmail(email) {
  const local = String(email).toLowerCase().split("@")[0];
  const parts = local.split(".");
  return parts[parts.length - 1].toUpperCase();
}

// ============================ HASH PASSWORD ============================

function _hash(password, salt) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + ":" + salt,
    Utilities.Charset.UTF_8
  );
  return raw.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

function _newSalt() { return Utilities.getUuid(); }

// ============================ CACHE LETTURA FOGLI ============================

/**
 * Le funzioni di lettura dei fogli sorgente possono essere lente. Uso CacheService
 * per ridurre il carico (TTL breve: 60s, sufficiente per ridurre chiamate ripetute
 * nella stessa interazione utente).
 */
function _cachedRead(key, ttlSec, producer) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch(e) { /* cache corrotta, ricalcolo */ }
  }
  const val = producer();
  try {
    const s = JSON.stringify(val);
    if (s.length < 90000) cache.put(key, s, ttlSec); // CacheService limit 100KB per key
  } catch(e) { /* ignoro errori cache */ }
  return val;
}

function invalidaCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(["RUBRICA","CATTEDRE","CALENDARIO","COORDINATORI"]);
}

// ============================ INGESTION DA RUBRICA ============================

function _caricaRubrica() {
  return _cachedRead("RUBRICA", 60, function() { return _caricaRubricaRaw(); });
}

function _caricaRubricaRaw() {
  if (!ID_RUBRICA) {
    throw new Error("ID_RUBRICA non configurato. Inserisci l'ID del file Rubrica nel codice.");
  }
  const ss = SpreadsheetApp.openById(ID_RUBRICA);
  const sh = ss.getSheetByName("Rubrica") || ss.getSheets()[0];
  const data = sh.getDataRange().getValues();

  const out = {}; // email -> {cognomeNorm, cognomeFull, nome, displayName}
  for (let i = 1; i < data.length; i++) {
    const nomeFull = String(data[i][0] || "").trim();
    const email    = String(data[i][2] || "").trim().toLowerCase();
    if (!email || email.indexOf("@") < 0) continue;
    if (email.indexOf(DOMINIO_DOCENTI) < 0) continue;

    const cognomeNorm = _normCognome(_cognomeDaEmail(email));

    // Decompongo "COGNOME [COGNOME2] NOME [NOME2...]":
    // tento k da 1 a parts.length-1: la parte iniziale di k parole, normalizzata,
    // deve coincidere col cognome estratto dall'email.
    let displayCognome = nomeFull, nome = "";
    const parts = nomeFull.split(/\s+/);
    if (parts.length >= 2) {
      let found = false;
      for (let k = parts.length - 1; k >= 1 && !found; k--) {
        const tentativo = parts.slice(0, k).join(" ");
        if (_normCognome(tentativo) === cognomeNorm) {
          displayCognome = tentativo;
          nome = parts.slice(k).join(" ");
          found = true;
        }
      }
      if (!found) {
        displayCognome = parts[0];
        nome = parts.slice(1).join(" ");
      }
    }

    out[email] = {
      cognomeNorm: cognomeNorm,
      cognomeFull: displayCognome,
      nome: nome,
      displayName: nome ? (nome + " " + displayCognome) : displayCognome
    };
  }
  return out;
}

/**
 * Stampa nel log la struttura del foglio CATTEDRE: nome dei sheets, intestazioni
 * e prime 20 righe. Da lanciare se diagnostica() segnala "Cattedre: 0 righe valide"
 * o un numero molto basso.
 */
function dumpCattedre() {
  const ss = SpreadsheetApp.openById(ID_CATTEDRE);
  Logger.log("Spreadsheet name: " + ss.getName());
  Logger.log("Spreadsheet URL : " + ss.getUrl());
  const sheets = ss.getSheets();
  Logger.log("Numero di fogli (tab): " + sheets.length);
  sheets.forEach(function(s, idx) {
    Logger.log("  Tab " + idx + ": '" + s.getName() + "' righe=" + s.getLastRow() + " colonne=" + s.getLastColumn());
  });
  const sh = sheets[0];
  Logger.log("Sto leggendo il primo tab: '" + sh.getName() + "'");
  const data = sh.getDataRange().getValues();
  Logger.log("Righe totali (compresa intestazione): " + data.length);
  for (let i = 0; i < Math.min(data.length, 20); i++) {
    Logger.log("  Riga " + i + ": " + JSON.stringify(data[i]));
  }
}

/**
 * Stampa nel log la struttura del foglio COORDINATORI.
 */
function dumpCoordinatori() {
  const ss = SpreadsheetApp.openById(ID_COORDINATORI);
  Logger.log("Spreadsheet name: " + ss.getName());
  const sheets = ss.getSheets();
  sheets.forEach(function(s, idx) {
    Logger.log("  Tab " + idx + ": '" + s.getName() + "' righe=" + s.getLastRow());
  });
  const sh = sheets[0];
  const data = sh.getDataRange().getValues();
  Logger.log("Righe totali: " + data.length);
  for (let i = 0; i < Math.min(data.length, 15); i++) {
    Logger.log("  Riga " + i + ": " + JSON.stringify(data[i]));
  }
}

// ============================ INGESTION DA CATTEDRE ============================

function _caricaCattedre() {
  return _cachedRead("CATTEDRE", 60, function() { return _caricaCattedreRaw(); });
}

function _caricaCattedreRaw() {
  const ss = SpreadsheetApp.openById(ID_CATTEDRE);
  const sh = ss.getSheets()[0];
  const data = sh.getDataRange().getValues();

  const SKIP = ["POTENZ","SUPPL","ELETTSUPPL","DAASSEGNARE"];
  const out = [];
  for (let i = 1; i < data.length; i++) {
    let docente = String(data[i][0] || "").trim();
    const cdc   = String(data[i][1] || "").trim();
    const cls   = String(data[i][2] || "").trim();
    const ore   = parseFloat(data[i][3]);

    if (!docente || !cls) continue;
    if (docente === "(da assegnare)") continue;

    // "8h-ZAPPA" -> "ZAPPA"
    if (docente.indexOf("-") >= 0 && /\d/.test(docente.split("-")[0])) {
      docente = docente.split("-").slice(1).join("-");
    }

    const cognomeNorm = _normCognome(docente);
    if (SKIP.indexOf(cognomeNorm) >= 0) continue;

    out.push({
      cognomeNorm: cognomeNorm,
      classe: cls.toUpperCase(),
      classeConcorso: cdc,
      ore: isNaN(ore) ? 0 : ore
    });
  }
  return out;
}

// ============================ INGESTION DA COORDINATORI ============================

function _caricaCoordinatori() {
  return _cachedRead("COORDINATORI", 60, function() { return _caricaCoordinatoriRaw(); });
}

function _caricaCoordinatoriRaw() {
  const ss = SpreadsheetApp.openById(ID_COORDINATORI);
  const sh = ss.getSheets()[0];
  const data = sh.getDataRange().getValues();

  const out = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const classeRaw = String(r[2] || "").trim();
    const sezione   = String(r[3] || "").trim();
    const cognome   = String(r[4] || "").trim();
    if (!classeRaw || !sezione || !cognome) continue;
    if (!/\d/.test(classeRaw)) continue;
    if (classeRaw.toUpperCase().indexOf("CLASS") >= 0) continue;

    const numClasse = classeRaw.replace(/[^\d]/g, "");
    if (!numClasse) continue;
    const classeNorm = (numClasse + sezione).toUpperCase();

    out.push({ classe: classeNorm, cognomeNorm: _normCognome(cognome) });
  }
  return out;
}

// ============================ INGESTION DA CALENDARIO ============================

/**
 * Converte un valore "ora.minuti" in formato decimale italiano in minuti dalla mezzanotte.
 *   14.3  -> 14:30 -> 870
 *   16.45 -> 16:45 -> 1005
 *   17.0  -> 17:00 -> 1020
 *   18    -> 18:00 -> 1080
 */
function _decimalToMinutes(v) {
  if (v === null || v === undefined || v === "") return null;
  const num = parseFloat(v);
  if (isNaN(num)) return null;
  const hh = Math.floor(num);
  const s = num.toString();
  let mm = 0;
  if (s.indexOf(".") >= 0) {
    let frac = s.split(".")[1];
    if (frac.length === 1) frac = frac + "0"; // ".3" -> "30"
    mm = parseInt(frac, 10);
    if (isNaN(mm) || mm > 59) mm = Math.round((num - hh) * 60);
  }
  return hh * 60 + mm;
}

function _caricaCalendario() {
  return _cachedRead("CALENDARIO", 60, function() { return _caricaCalendarioRaw(); });
}

function _caricaCalendarioRaw() {
  const ss = SpreadsheetApp.openById(ID_CALENDARIO);
  const sheets = ss.getSheets();
  const out = [];

  sheets.forEach(function(sh) {
    const meseFoglio = sh.getName();
    const data = sh.getDataRange().getValues();
    let dataCorrente = null;

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const cellGiorno = r[0];
      if (cellGiorno) {
        if (cellGiorno instanceof Date) {
          dataCorrente = new Date(cellGiorno.getTime());
        } else {
          const s = String(cellGiorno);
          const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
          if (m) {
            let yy = parseInt(m[3], 10);
            if (yy < 100) yy += 2000;
            dataCorrente = new Date(yy, parseInt(m[2],10) - 1, parseInt(m[1],10));
          }
        }
      }
      if (!dataCorrente) continue;

      const oraIni = _decimalToMinutes(r[1]);
      const oraFin = _decimalToMinutes(r[2]);
      const aula   = r[3];
      const classe = String(r[4] || "").trim().toUpperCase();
      if (oraIni === null || oraFin === null) continue;
      if (!classe) continue;
      if (oraFin <= oraIni) continue;

      const durataOre = (oraFin - oraIni) / 60.0;
      out.push({
        data: _formatDateISO(dataCorrente),
        oraInizioMin: oraIni,
        oraFineMin: oraFin,
        durataOre: Math.round(durataOre * 100) / 100,
        classe: classe,
        aula: aula ? String(aula) : "",
        meseFoglio: meseFoglio
      });
    }
  });

  out.sort(function(a,b) {
    if (a.data !== b.data) return a.data < b.data ? -1 : 1;
    return a.oraInizioMin - b.oraInizioMin;
  });

  out.forEach(function(c) {
    c.id = c.data.replace(/-/g,"") + "-" + c.classe + "-" + c.oraInizioMin;
  });

  return out;
}

// ============================ AUTH ============================

function login(email, password) {
  email = String(email || "").trim().toLowerCase();
  password = String(password || "");
  if (!email || !password) return { success: false, message: "Email e password obbligatorie." };

  const isSegreteria = (email === ADMIN_EMAIL.toLowerCase());
  if (!isSegreteria && email.indexOf(DOMINIO_DOCENTI) < 0) {
    return { success: false, message: "Email non istituzionale (atteso " + DOMINIO_DOCENTI + ")." };
  }

  const sh = _sheet("UTENTI");
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === email) {
      const hashStored = rows[i][3];
      const salt = rows[i][4];
      if (_hash(password, salt) === hashStored) {
        // RESYNC: ad ogni login ricalcolo cognome/nome/ore dalla rubrica e cattedre.
        // Questo autocorregge righe salvate sbagliate in passato.
        let cognomeOk = rows[i][1];
        let nomeOk    = rows[i][2];
        let oreOk     = rows[i][5];
        if (!isSegreteria) {
          try {
            const rub = _caricaRubrica();
            const info = rub[email];
            if (info) {
              cognomeOk = info.cognomeFull;
              nomeOk    = info.nome;
              let oreSum = 0;
              const cattedre = _caricaCattedre();
              cattedre.forEach(function(c) {
                if (c.cognomeNorm === info.cognomeNorm) oreSum += c.ore;
              });
              if (oreSum > 0) oreOk = oreSum;
              // Aggiorno la riga se i dati sono cambiati
              if (rows[i][1] !== cognomeOk || rows[i][2] !== nomeOk || rows[i][5] !== oreOk) {
                sh.getRange(i+1, 2).setValue(cognomeOk);
                sh.getRange(i+1, 3).setValue(nomeOk);
                sh.getRange(i+1, 6).setValue(oreOk);
                SpreadsheetApp.flush();
              }
            }
          } catch(e) { /* se la rubrica non è disponibile, uso i dati salvati */ }
        }
        return _buildLoginOK(email, cognomeOk, nomeOk, oreOk, rows[i][6], false);
      }
      return { success: false, message: "Password errata. Per recupero usa 'Password dimenticata'." };
    }
  }

  // Primo accesso
  if (password !== DEFAULT_PWD) {
    return { success: false, message: "Primo accesso? Usa la password predefinita." };
  }

  let cognome = "", nome = "", ore = 0;
  if (isSegreteria) {
    cognome = "SEGRETERIA"; nome = "ITIS RIVA"; ore = 0;
  } else {
    let rub;
    try { rub = _caricaRubrica(); }
    catch(e) { return { success: false, message: "Errore caricamento rubrica: " + e.message }; }
    const info = rub[email];
    if (!info) {
      return { success: false, message: "Email non presente in rubrica scolastica. Contattare segreteria." };
    }
    cognome = info.cognomeFull;
    nome = info.nome;
    const cattedre = _caricaCattedre();
    cattedre.forEach(function(c) {
      if (c.cognomeNorm === info.cognomeNorm) ore += c.ore;
    });
    if (ore === 0) ore = ORE_FULL_TIME;
  }

  const salt = _newSalt();
  const hash = _hash(password, salt);
  sh.appendRow([
    email, cognome, nome, hash, salt, ore,
    isSegreteria ? "SEGRETERIA" : "DOCENTE",
    new Date()
  ]);
  SpreadsheetApp.flush(); // FIX: forza scrittura prima che getDashboard rilegga
  return _buildLoginOK(email, cognome, nome, ore, isSegreteria ? "SEGRETERIA" : "DOCENTE", true);
}

function _buildLoginOK(email, cognome, nome, ore, ruolo, isFirstAccess) {
  return {
    success: true,
    isFirstAccess: !!isFirstAccess,
    user: {
      email: email,
      cognome: cognome,
      nome: nome,
      displayName: (nome ? nome + " " : "") + cognome,
      ore: ore,
      ruolo: ruolo
    }
  };
}

function changePassword(email, oldPwd, newPwd) {
  email = String(email || "").trim().toLowerCase();
  if (!newPwd || String(newPwd).length < 6) {
    return { success: false, message: "La nuova password deve essere di almeno 6 caratteri." };
  }
  const sh = _sheet("UTENTI");
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === email) {
      if (_hash(oldPwd, rows[i][4]) !== rows[i][3]) {
        return { success: false, message: "Password attuale errata." };
      }
      const newSalt = _newSalt();
      sh.getRange(i+1, 4).setValue(_hash(newPwd, newSalt));
      sh.getRange(i+1, 5).setValue(newSalt);
      return { success: true };
    }
  }
  return { success: false, message: "Utente non trovato." };
}

// ============================ RESET PASSWORD VIA OTP ============================

function richiediResetOTP(email) {
  email = String(email || "").trim().toLowerCase();
  if (!email) return { success: false, message: "Email mancante." };

  const sh = _sheet("UTENTI");
  const rows = sh.getDataRange().getValues();
  let exists = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === email) { exists = true; break; }
  }

  if (exists) {
    const codice = ("" + Math.floor(100000 + Math.random() * 900000));
    const scadenza = new Date(Date.now() + OTP_VALID_MIN * 60 * 1000);
    _sheet("OTP").appendRow([email, codice, scadenza, false]);
    try {
      MailApp.sendEmail({
        to: email,
        subject: "Codice reset password - Gestione CdC ITIS Riva",
        body: "Ciao,\n\n" +
              "il tuo codice di verifica per reimpostare la password e':\n\n" +
              "    " + codice + "\n\n" +
              "Valido per " + OTP_VALID_MIN + " minuti.\n" +
              "Se non hai richiesto tu il reset, ignora questa email.\n\n" +
              "Sistema Gestione CdC - ITIS G. Riva"
      });
    } catch (err) {
      Logger.log("Errore invio OTP a " + email + ": " + err);
    }
  }

  // Risposta uniforme indipendentemente dall'esistenza dell'utente (anti enumeration)
  return { success: true, message: "Se l'email è registrata, riceverai un codice entro pochi minuti." };
}

function confermaResetOTP(email, codice, newPwd) {
  email = String(email || "").trim().toLowerCase();
  codice = String(codice || "").trim();
  if (!newPwd || String(newPwd).length < 6) {
    return { success: false, message: "La nuova password deve essere di almeno 6 caratteri." };
  }
  const shOtp = _sheet("OTP");
  const data = shOtp.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).toLowerCase() === email && String(data[i][1]) === codice && !data[i][3]) {
      const scad = new Date(data[i][2]);
      if (scad < new Date()) return { success: false, message: "Codice scaduto. Richiedine uno nuovo." };

      const shU = _sheet("UTENTI");
      const u = shU.getDataRange().getValues();
      for (let j = 1; j < u.length; j++) {
        if (String(u[j][0]).toLowerCase() === email) {
          const newSalt = _newSalt();
          shU.getRange(j+1, 4).setValue(_hash(newPwd, newSalt));
          shU.getRange(j+1, 5).setValue(newSalt);
          shOtp.getRange(i+1, 4).setValue(true);
          return { success: true, message: "Password aggiornata. Effettua il login." };
        }
      }
      return { success: false, message: "Utente non trovato." };
    }
  }
  return { success: false, message: "Codice non valido." };
}

// ============================ DASHBOARD DOCENTE ============================

function getDashboard(email) {
  email = String(email || "").trim().toLowerCase();
  const utente = _getUtente(email);
  if (!utente) return { error: "Utente non trovato." };

  const isSegreteria = utente.RUOLO === "SEGRETERIA";
  if (isSegreteria) {
    return { isSegreteria: true, segreteria: getDashboardSegreteria() };
  }

  const oreSett = parseFloat(utente.ORE_SETTIMANALI) || ORE_FULL_TIME;
  const soglia = (ORE_CDC_DOVUTE / ORE_FULL_TIME) * oreSett;

  const cognomeNorm = _normCognome(utente.COGNOME);
  const cattedre = _caricaCattedre();
  const mieClassi = {};
  cattedre.forEach(function(c) { if (c.cognomeNorm === cognomeNorm) mieClassi[c.classe] = true; });

  // FIX: aggiungo anche le classi coordinate, perché il coordinatore deve sempre
  // vedere il CdC della propria classe anche se non vi insegna direttamente.
  const coords = _caricaCoordinatori();
  const classiCoordinate = coords
    .filter(function(c) { return c.cognomeNorm === cognomeNorm; })
    .map(function(c) { return c.classe; });
  const isCoordinatore = classiCoordinate.length > 0;
  classiCoordinate.forEach(function(cl) { mieClassi[cl] = true; });

  const tuttiCdC = _caricaCalendario();
  const mieiCdc = tuttiCdC.filter(function(c) { return mieClassi[c.classe]; });

  const partecipazioni = _readAll("PARTECIPAZIONI");
  const mappa = {};
  partecipazioni.forEach(function(p) {
    if (String(p.EMAIL_DOCENTE).toLowerCase() === email) mappa[p.ID_CDC] = p.STATO;
  });

  let oreTotali = 0;
  let orePartecipa = 0;
  const now = new Date();
  const cdcOut = mieiCdc.map(function(c) {
    const stato = mappa[c.id] || "PARTECIPA";
    oreTotali += c.durataOre;
    if (stato === "PARTECIPA") orePartecipa += c.durataOre;
    const fineCdc = _parseDateTime(c.data, c.oraFineMin);
    return {
      id: c.id,
      data: c.data,
      dataIT: _formatDateIT(c.data),
      oraInizio: _minToHHMM(c.oraInizioMin),
      oraFine: _minToHHMM(c.oraFineMin),
      durataOre: c.durataOre,
      classe: c.classe,
      aula: c.aula,
      stato: stato,
      passato: fineCdc < now,
      coordinata: classiCoordinate.indexOf(c.classe) >= 0
    };
  });

  return {
    user: {
      email: email,
      displayName: (utente.NOME ? utente.NOME + " " : "") + utente.COGNOME,
      ore: oreSett,
      ruolo: utente.RUOLO
    },
    cdc: cdcOut,
    soglia: Math.round(soglia * 100) / 100,
    oreTotali: Math.round(oreTotali * 100) / 100,
    orePartecipazione: Math.round(orePartecipa * 100) / 100,
    isCoordinatore: isCoordinatore,
    classiCoordinate: classiCoordinate,
    isSegreteria: false
  };
}

function toggleEsonero(email, idCdc, nuovoStato, forzaSottosoglia) {
  email = String(email || "").trim().toLowerCase();
  const STATI_OK = { "PARTECIPA": true, "ESONERATO": true };
  if (!STATI_OK[nuovoStato]) return { success: false, message: "Stato non valido." };

  // Validazione fuori dal lock (solo letture)
  const utente = _getUtente(email);
  if (!utente) return { success: false, message: "Utente non trovato." };

  const cdc = _caricaCalendario().find(function(c) { return c.id === idCdc; });
  if (!cdc) return { success: false, message: "CdC non trovato." };
  const fineCdc = _parseDateTime(cdc.data, cdc.oraFineMin);
  if (fineCdc < new Date()) {
    return { success: false, message: "Non è più possibile modificare uno stato di CdC già concluso." };
  }

  // Controllo soglia (solo letture, niente lock)
  if (nuovoStato === "ESONERATO" && !forzaSottosoglia) {
    const sim = _simulaContatoreLight(email, utente, idCdc, "ESONERATO");
    if (sim.orePartecipazione < sim.soglia) {
      return {
        success: false,
        requireConfirm: true,
        message: "Esonerandoti scenderesti sotto la soglia minima.",
        oreFuture: sim.orePartecipazione,
        soglia: sim.soglia
      };
    }
  }

  // Lock SOLO intorno alla scrittura
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) {
    Logger.log("toggleEsonero LOCK FAIL: " + e + " | stack: " + (e.stack || "n/a"));
    return { success: false, message: "Lock non disponibile: " + e.toString() };
  }
  try {
    _scriviPartecipazione(idCdc, email, nuovoStato, email, "");
    return { success: true };
  } catch(e2) {
    Logger.log("toggleEsonero WRITE FAIL: " + e2 + " | stack: " + (e2.stack || "n/a"));
    return { success: false, message: "Errore scrittura: " + e2.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * Simulazione leggera del contatore ore: ricalcola le ore di partecipazione del docente
 * applicando una modifica ipotetica a un singolo CdC, senza richiamare getDashboard.
 * Dato che _caricaCattedre/_caricaCalendario/_caricaCoordinatori sono cacheati 60s,
 * questa funzione è molto più veloce (riusa la cache invece di rifare tutto).
 */
function _simulaContatoreLight(email, utente, idCdcModificato, statoModificato) {
  const cognomeNorm = _normCognome(utente.COGNOME);
  const oreSett = parseFloat(utente.ORE_SETTIMANALI) || ORE_FULL_TIME;
  const soglia = (ORE_CDC_DOVUTE / ORE_FULL_TIME) * oreSett;

  const cattedre = _caricaCattedre();
  const mieClassi = {};
  cattedre.forEach(function(c) { if (c.cognomeNorm === cognomeNorm) mieClassi[c.classe] = true; });
  const coords = _caricaCoordinatori();
  coords.forEach(function(c) { if (c.cognomeNorm === cognomeNorm) mieClassi[c.classe] = true; });

  const tuttiCdC = _caricaCalendario();
  const mieiCdc = tuttiCdC.filter(function(c) { return mieClassi[c.classe]; });

  const partecipazioni = _readAll("PARTECIPAZIONI");
  const mappa = {};
  partecipazioni.forEach(function(p) {
    if (String(p.EMAIL_DOCENTE).toLowerCase() === email) mappa[p.ID_CDC] = p.STATO;
  });

  let ore = 0;
  mieiCdc.forEach(function(c) {
    let stato = mappa[c.id] || "PARTECIPA";
    if (c.id === idCdcModificato) stato = statoModificato;
    if (stato === "PARTECIPA") ore += c.durataOre;
  });

  return {
    orePartecipazione: Math.round(ore * 100) / 100,
    soglia: Math.round(soglia * 100) / 100
  };
}

function _scriviPartecipazione(idCdc, emailDocente, stato, marcatoDa, note) {
  const sh = _sheet("PARTECIPAZIONI");
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idCdc) &&
        String(data[i][1]).toLowerCase() === String(emailDocente).toLowerCase()) {
      sh.getRange(i+1, 3).setValue(stato);
      sh.getRange(i+1, 4).setValue(marcatoDa);
      sh.getRange(i+1, 5).setValue(new Date());
      sh.getRange(i+1, 6).setValue(note || "");
      return;
    }
  }
  sh.appendRow([idCdc, emailDocente, stato, marcatoDa, new Date(), note || ""]);
}

function _getUtente(email) {
  if (!email) return null;
  const target = String(email).trim().toLowerCase();
  // Lettura diretta dal range, evita problemi di cache/header su _readAll
  const sh = _sheet("UTENTI");
  const last = sh.getLastRow();
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, 8).getValues();
  for (let i = 0; i < rows.length; i++) {
    const em = String(rows[i][0] || "").trim().toLowerCase();
    if (em && em === target) {
      return {
        EMAIL: rows[i][0],
        COGNOME: rows[i][1],
        NOME: rows[i][2],
        PASSWORD_HASH: rows[i][3],
        SALT: rows[i][4],
        ORE_SETTIMANALI: rows[i][5],
        RUOLO: rows[i][6],
        DATA_REGISTRAZIONE: rows[i][7]
      };
    }
  }
  return null;
}

// ============================ VISTA COORDINATORE ============================

function getCdcDettaglio(emailRichiedente, idCdc) {
  emailRichiedente = String(emailRichiedente || "").trim().toLowerCase();
  const cdc = _caricaCalendario().find(function(c) { return c.id === idCdc; });
  if (!cdc) return { error: "CdC non trovato." };

  const utente = _getUtente(emailRichiedente);
  if (!utente) return { error: "Utente non trovato." };

  let permesso = (utente.RUOLO === "SEGRETERIA");
  if (!permesso) {
    const cognomeNorm = _normCognome(utente.COGNOME);
    const coords = _caricaCoordinatori();
    permesso = coords.some(function(c) { return c.classe === cdc.classe && c.cognomeNorm === cognomeNorm; });
  }
  if (!permesso) return { error: "Non hai i permessi per visualizzare questo CdC." };

  const cattedre = _caricaCattedre();
  const cognomiClasse = {};
  cattedre.forEach(function(c) { if (c.classe === cdc.classe) cognomiClasse[c.cognomeNorm] = true; });

  const rub = _caricaRubrica();
  const cognomeToInfo = {};
  Object.keys(rub).forEach(function(em) {
    const inf = rub[em];
    if (cognomiClasse[inf.cognomeNorm]) {
      cognomeToInfo[inf.cognomeNorm] = { email: em, displayName: inf.displayName };
    }
  });

  const part = _readAll("PARTECIPAZIONI").filter(function(p) { return String(p.ID_CDC) === String(idCdc); });
  const mappaStati = {};
  part.forEach(function(p) { mappaStati[String(p.EMAIL_DOCENTE).toLowerCase()] = p.STATO; });

  const docenti = Object.keys(cognomiClasse).map(function(cn) {
    const info = cognomeToInfo[cn];
    if (!info) {
      return { cognomeNorm: cn, displayName: cn + " (non in rubrica)", email: null, stato: "—" };
    }
    return {
      cognomeNorm: cn,
      displayName: info.displayName,
      email: info.email,
      stato: mappaStati[info.email] || "PARTECIPA"
    };
  });
  docenti.sort(function(a,b) { return a.displayName.localeCompare(b.displayName); });

  const fineCdc = _parseDateTime(cdc.data, cdc.oraFineMin);
  return {
    cdc: {
      id: cdc.id,
      classe: cdc.classe,
      data: cdc.data,
      dataIT: _formatDateIT(cdc.data),
      oraInizio: _minToHHMM(cdc.oraInizioMin),
      oraFine: _minToHHMM(cdc.oraFineMin),
      durataOre: cdc.durataOre,
      passato: fineCdc < new Date()
    },
    docenti: docenti
  };
}

function coordSetStato(emailCoordinatore, idCdc, emailDocente, nuovoStato) {
  const STATI_OK = { "PARTECIPA": true, "ASSENTE": true, "ESONERATO": true };
  if (!STATI_OK[nuovoStato]) return { success: false, message: "Stato non valido." };

  // Validazione permessi e CdC fuori dal lock (solo letture)
  const dett = getCdcDettaglio(emailCoordinatore, idCdc);
  if (dett.error) return { success: false, message: dett.error };
  if (dett.cdc && dett.cdc.passato) {
    return { success: false, message: "Non è più possibile modificare uno stato di CdC già concluso." };
  }

  // Lock SOLO intorno alla scrittura
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) {
    Logger.log("coordSetStato LOCK FAIL: " + e + " | stack: " + (e.stack || "n/a"));
    return { success: false, message: "Lock non disponibile: " + e.toString() };
  }
  try {
    _scriviPartecipazione(idCdc, emailDocente, nuovoStato, emailCoordinatore, "");
    return { success: true };
  } catch(e2) {
    Logger.log("coordSetStato WRITE FAIL: " + e2 + " | stack: " + (e2.stack || "n/a"));
    return { success: false, message: "Errore scrittura: " + e2.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ============================ VISTA SEGRETERIA ============================

function getDashboardSegreteria() {
  // Parto dalla RUBRICA: questa è la lista completa di TUTTI i docenti dell'istituto,
  // indipendentemente dal fatto che abbiano effettuato il login.
  let rubrica;
  try {
    rubrica = _caricaRubrica();
  } catch (e) {
    return { docenti: [], error: "Errore lettura rubrica: " + e.toString() };
  }

  const cattedre = _caricaCattedre();
  const calendario = _caricaCalendario();
  const coordinatori = _caricaCoordinatori();
  const partec = _readAll("PARTECIPAZIONI");
  const utentiRegistrati = _readAll("UTENTI");

  // Mappa partecipazioni per email
  const partByEmail = {};
  partec.forEach(function(p) {
    const em = String(p.EMAIL_DOCENTE).toLowerCase();
    if (!partByEmail[em]) partByEmail[em] = {};
    partByEmail[em][String(p.ID_CDC)] = p.STATO;
  });

  // Mappa utenti registrati per recuperare ore_settimanali memorizzate (override).
  // Nota: il campo "ore" calcolato dalle cattedre è la fonte di verita primaria;
  // utentiRegistrati lo usiamo solo per sapere se l'utente ha gia fatto login.
  const utentiByEmail = {};
  utentiRegistrati.forEach(function(u) {
    utentiByEmail[String(u.EMAIL).toLowerCase()] = u;
  });

  // Per ogni docente in rubrica, calcolo ore settimanali, classi, CdC e partecipazione
  const out = [];
  Object.keys(rubrica).forEach(function(email) {
    const info = rubrica[email];
    const cognomeNorm = info.cognomeNorm;

    // Ore settimanali = somma ore cattedre con stesso cognomeNorm
    let oreSett = 0;
    const mieClassi = {};
    cattedre.forEach(function(c) {
      if (c.cognomeNorm === cognomeNorm) {
        oreSett += c.ore;
        mieClassi[c.classe] = true;
      }
    });
    // Aggiungo le classi coordinate (anche se non vi insegna)
    coordinatori.forEach(function(c) {
      if (c.cognomeNorm === cognomeNorm) mieClassi[c.classe] = true;
    });

    // Se ore settimanali = 0 (docente in rubrica ma senza cattedre note),
    // lo escludo dal report di segreteria (probabile docente di altro plesso o personale ATA)
    if (oreSett === 0 && Object.keys(mieClassi).length === 0) return;

    // Se per qualche motivo non ho ore ma ha classi (raro), default a 18
    if (oreSett === 0) oreSett = ORE_FULL_TIME;

    const soglia = (ORE_CDC_DOVUTE / ORE_FULL_TIME) * oreSett;

    const mieiCdc = calendario.filter(function(c) { return mieClassi[c.classe]; });

    let oreTot = 0, orePart = 0;
    const stati = partByEmail[email] || {};
    mieiCdc.forEach(function(c) {
      oreTot += c.durataOre;
      const st = stati[c.id] || "PARTECIPA";
      if (st === "PARTECIPA") orePart += c.durataOre;
    });

    const haAccount = !!utentiByEmail[email];

    out.push({
      email: email,
      displayName: info.displayName,
      oreSett: oreSett,
      soglia: Math.round(soglia * 100) / 100,
      oreTot: Math.round(oreTot * 100) / 100,
      orePart: Math.round(orePart * 100) / 100,
      sotto: orePart < soglia,
      haAccount: haAccount
    });
  });

  out.sort(function(a,b) { return a.displayName.localeCompare(b.displayName); });
  return { docenti: out };
}

// ============================ UTILITY DATE/ORE ============================

function _minToHHMM(min) {
  const hh = Math.floor(min / 60);
  const mm = min % 60;
  return ("0"+hh).slice(-2) + ":" + ("0"+mm).slice(-2);
}

function _parseDateTime(isoDate, minOfDay) {
  const p = isoDate.split("-");
  return new Date(parseInt(p[0],10), parseInt(p[1],10) - 1, parseInt(p[2],10),
                  Math.floor(minOfDay/60), minOfDay % 60, 0);
}

function _formatDateIT(isoDate) {
  const p = isoDate.split("-");
  return p[2] + "/" + p[1] + "/" + p[0];
}

function _formatDateISO(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone() || "Europe/Rome", "yyyy-MM-dd");
}

// ============================ DIAGNOSTICA ============================

/**
 * Lanciare manualmente una tantum dall'editor Apps Script per verificare
 * che tutti i fogli sorgente siano leggibili. Stampa nei log un riepilogo.
 */
function diagnostica() {
  _openDB();
  Logger.log("DB OK: " + SPREADSHEET_ID);

  try {
    const r = _caricaRubricaRaw();
    Logger.log("Rubrica: " + Object.keys(r).length + " docenti caricati.");
  } catch (e) { Logger.log("ERRORE Rubrica: " + e); }

  try {
    const c = _caricaCattedreRaw();
    Logger.log("Cattedre: " + c.length + " righe valide.");
  } catch (e) { Logger.log("ERRORE Cattedre: " + e); }

  try {
    const cal = _caricaCalendarioRaw();
    Logger.log("Calendario: " + cal.length + " CdC trovati.");
  } catch (e) { Logger.log("ERRORE Calendario: " + e); }

  try {
    const co = _caricaCoordinatoriRaw();
    Logger.log("Coordinatori: " + co.length + " classi.");
  } catch (e) { Logger.log("ERRORE Coordinatori: " + e); }
}

/**
 * Diagnostica specifica: dato un email, dice esattamente cosa trova nel foglio UTENTI.
 * Da lanciare dall'editor (modificando la variabile EMAIL_TEST) se "Utente non trovato"
 * persiste dopo le fix.
 */
function diagnosticaUtente() {
  const EMAIL_TEST = "vatf02006@istruzione.edu.it"; // <- modifica con l'email da testare
  Logger.log("Cerco: '" + EMAIL_TEST + "'");

  const sh = _sheet("UTENTI");
  const last = sh.getLastRow();
  Logger.log("Foglio UTENTI: ultima riga = " + last);
  if (last < 2) { Logger.log("Foglio UTENTI vuoto."); return; }

  const rows = sh.getRange(2, 1, last - 1, 8).getValues();
  Logger.log("Righe trovate: " + rows.length);
  rows.forEach(function(r, i) {
    Logger.log("Riga " + (i+2) + " EMAIL='" + r[0] + "' (lower='" + String(r[0] || '').trim().toLowerCase() + "') COGNOME='" + r[1] + "' RUOLO='" + r[6] + "'");
  });

  const u = _getUtente(EMAIL_TEST);
  Logger.log("Risultato _getUtente: " + (u ? JSON.stringify(u) : "NULL"));
}

function diagnosticaAttivita() {
  const ss = SpreadsheetApp.openById(ID_ATTIVITA);
  Logger.log("=== DIAGNOSTICA ATTIVITÀ ===");
  Logger.log("Nome spreadsheet: " + ss.getName());
  
  const sheets = ss.getSheets();
  Logger.log("Numero fogli: " + sheets.length);
  
  sheets.forEach(function(sh, idx) {
    Logger.log("Foglio " + idx + ": nome='" + sh.getName() + "', righe=" + sh.getLastRow() + ", colonne=" + sh.getLastColumn());
  });
  
  // Leggi il foglio con più colonne
  let shDati = null;
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getLastColumn() > 10) {
      shDati = sheets[i];
      break;
    }
  }
  if (!shDati) shDati = sheets[sheets.length - 1];
  
  Logger.log("Foglio selezionato: '" + shDati.getName() + "'");
  
  const data = shDati.getDataRange().getValues();
  Logger.log("Righe totali lette: " + data.length);
  Logger.log("Colonne totali lette: " + (data.length > 0 ? data[0].length : 0));
  
  // Stampa la riga 0 (intestazioni)
  if (data.length > 0) {
    Logger.log("=== RIGA 0 (intestazioni) ===");
    for (let col = 0; col < data[0].length; col++) {
      Logger.log("  Col " + col + ": '" + data[0][col] + "'");
    }
  }
  
  // Stampa le righe 1-5 per ogni mese (esempio)
  Logger.log("=== PRIME RIGHE (1-5) ===");
  for (let riga = 1; riga <= Math.min(5, data.length - 1); riga++) {
    Logger.log("--- Riga " + riga + " ---");
    for (let col = 0; col < Math.min(data[riga].length, 36); col++) {
      const val = data[riga][col];
      if (val !== "" && val !== null && val !== undefined) {
        Logger.log("  Col " + col + ": '" + val + "' (tipo: " + typeof val + ")");
      }
    }
  }
  
  // Prova a individuare i blocchi mese
  Logger.log("=== RICERCA MESI NELLA RIGA 0 ===");
  const headerRow = data[0];
  const mesi = ["SETTEMBRE","OTTOBRE","NOVEMBRE","DICEMBRE","GENNAIO","FEBBRAIO",
                "MARZO","APRILE","MAGGIO","GIUGNO","LUGLIO","AGOSTO"];
  
  for (let col = 0; col < headerRow.length; col++) {
    const cell = String(headerRow[col] || "").trim().toUpperCase();
    if (mesi.indexOf(cell) >= 0) {
      Logger.log("  TROVATO MESE '" + cell + "' alla colonna " + col);
      // Stampa le prime 3 righe di questo blocco
      for (let r = 1; r <= Math.min(3, data.length - 1); r++) {
        const g = data[r][col];      // giorno lettera
        const n = data[r][col + 1];  // numero
        const a = data[r][col + 2];  // attività
        Logger.log("    Riga " + r + ": g='" + g + "' n='" + n + "' a='" + a + "'");
      }
    }
  }
}

function testCaricaAttivita() {
  try {
    const result = _caricaAttivitaRaw();
    Logger.log("=== RISULTATO _caricaAttivitaRaw ===");
    Logger.log("Numero attività: " + result.attivita.length);
    Logger.log("Numero mesi: " + result.mesiDisponibili.length);
    
    if (result.mesiDisponibili.length > 0) {
      Logger.log("Mesi disponibili: " + JSON.stringify(result.mesiDisponibili));
    }
    
    if (result.attivita.length > 0) {
      Logger.log("Prime 10 attività: " + JSON.stringify(result.attivita.slice(0, 10)));
    } else {
      Logger.log("NESSuna attività trovata!");
    }
  } catch(e) {
    Logger.log("ERRORE: " + e.toString());
    Logger.log("Stack: " + (e.stack || "n/a"));
  }
}

/**
 * Reset completo: cancella UTENTI, PARTECIPAZIONI, OTP. NON tocca i fogli sorgente.
 * Usare con cautela. Da eseguire solo dall'editor Apps Script.
 */
function resetDatabase() {
  const ss = _openDB();
  ["UTENTI","PARTECIPAZIONI","OTP"].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).clearContent();
    }
  });
  Logger.log("Database resettato.");
}
