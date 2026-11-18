/**
 * GESTIONE CALCETTO PRO - Backend Google (Foglio Google condiviso)
 * Da incollare in script.google.com come file "Codice.gs".
 *
 * Mantiene TUTTA la logica originale (saldi, quote, partite, referti).
 * Aggiunge un'API di lettura/scrittura usata dall'app installabile sul telefono:
 * - doGet  con parametri fn/args/callback  -> risposta JSONP
 * - doPost con corpo {fn, args}            -> risposta JSON
 * Nessun modello AI, nessuna dipendenza esterna.
 */

var FOGLIO_ID = "1Bt6RZkR0Qmn6_8TBl4xBqLtc7qxk_NMR2uehVQe1V2o"; // SOSTITUISCI CON L'ID REALE DEL TUO FOGLIO

/* ================= API PER L'APP MOBILE (PWA) ================= */

var API_FUNZIONI = {
  "ottieniDati": ottieniDati,
  "aggiungiGiocatore": aggiungiGiocatore,
  "eliminaGiocatore": eliminaGiocatore,
  "versaQuota": versaQuota,
  "versaQuoteMassive": versaQuoteMassive,
  "versaQuotaIscrizione": versaQuotaIscrizione,
  "cambiaStatoGiocatore": cambiaStatoGiocatore,
  "creaPartita": creaPartita,
  "creaAmichevole": creaAmichevole,
  "chiudiPartita": chiudiPartita,
  "annullaPartita": annullaPartita,
  "annullaPartitaArchiviata": annullaPartitaArchiviata,
  "modificaPartitaArchiviata": modificaPartitaArchiviata,
  "registraSpesaCondivisa": registraSpesaCondivisa,
  "modificaGiocatore": modificaGiocatore
};

function eseguiAPI(nomeFunzione, argomenti) {
  var fn = API_FUNZIONI[nomeFunzione];
  if (!fn) {
    return JSON.stringify({ erroreCritico: "Funzione non trovata: " + nomeFunzione });
  }
  var argomentiSicuri = [];
  for (var i = 0; i < argomenti.length; i++) {
    argomentiSicuri.push(argomenti[i] === undefined ? null : argomenti[i]);
  }
  // Le funzioni originali restituiscono una stringa JSON di ottieniDati();
  // la si ripara a oggetto per evitare doppia serializzazione.
  var risultato = fn.apply(null, argomentiSicuri);
  if (typeof risultato === "string") {
    try { return JSON.parse(risultato); } catch (err) { return risultato; }
  }
  return risultato;
}

function doGet(e) {
  // 1) Pagina web originale, se richiesta senza parametri
  if (!e || !e.parameter || !e.parameter.fn) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Gestione Calcetto Pro')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no');
  }

  // 2) API JSONP usata dall'app installabile (lettura/scrittura sul Foglio condiviso)
  var argomenti = [];
  try {
    if (e.parameter.args) argomenti = JSON.parse(e.parameter.args);
  } catch (errArg) {
    argomenti = [];
  }
  var dati;
  try {
    dati = eseguiAPI(e.parameter.fn, argomenti);
  } catch (errore) {
    dati = { erroreCritico: "Errore dal Foglio Google: " + errore.message };
  }
  var payload = JSON.stringify(dati);
  if (e.parameter.callback) {
    return ContentService.createTextOutput(e.parameter.callback + "(" + payload + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var corpo = {};
  try {
    if (e && e.postData && e.postData.contents) corpo = JSON.parse(e.postData.contents);
  } catch (errPost) {
    corpo = {};
  }
  var dati;
  try {
    dati = eseguiAPI(String(corpo.fn || ""), corpo.args || []);
  } catch (errore) {
    dati = { erroreCritico: "Errore dal Foglio Google: " + errore.message };
  }
  return ContentService.createTextOutput(JSON.stringify(dati))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupIniziale() {
  var ss = getTargetSpreadsheet();

  var sheetGiocatori = ss.getSheetByName("Giocatori");
  if (!sheetGiocatori) {
    sheetGiocatori = ss.insertSheet("Giocatori");
    sheetGiocatori.appendRow(["Nome Giocatore", "Saldo (€)", "Gol", "Ammonizioni", "Ultimo Ruolo", "Assist", "Espulsioni", "Gol Subiti (Azione)", "Gol Subiti (Rigore)", "Gol Subiti (Piazzato)", "Quota Iscrizione", "Attivo"]);
  } else {
    if (sheetGiocatori.getLastColumn() < 5) {
      sheetGiocatori.getRange("C1").setValue("Gol");
      sheetGiocatori.getRange("D1").setValue("Ammonizioni");
      sheetGiocatori.getRange("E1").setValue("Ultimo Ruolo");
    }
    if (sheetGiocatori.getLastColumn() < 6) {
      sheetGiocatori.getRange("F1").setValue("Assist");
    }
    if (sheetGiocatori.getLastColumn() < 7) {
      sheetGiocatori.getRange("G1").setValue("Espulsioni");
    }
    if (sheetGiocatori.getLastColumn() < 10) {
      sheetGiocatori.getRange("H1").setValue("Gol Subiti (Azione)");
      sheetGiocatori.getRange("I1").setValue("Gol Subiti (Rigore)");
      sheetGiocatori.getRange("J1").setValue("Gol Subiti (Piazzato)");
    }
    if (sheetGiocatori.getLastColumn() < 12) {
      sheetGiocatori.getRange("K1").setValue("Quota Iscrizione");
      sheetGiocatori.getRange("L1").setValue("Attivo");
      
      var numRows = sheetGiocatori.getLastRow();
      if (numRows > 1) {
        sheetGiocatori.getRange(2, 11, numRows - 1, 1).setValue(0);
        sheetGiocatori.getRange(2, 12, numRows - 1, 1).setValue(true);
      }
    }
  }
  sheetGiocatori.getRange("A1:L1").setFontWeight("bold");

  if (!ss.getSheetByName("Log")) {
    var sheetLog = ss.insertSheet("Log");
    sheetLog.appendRow(["Data", "Operazione", "Importo/Dato", "Dettagli"]);
    sheetLog.getRange("A1:D1").setFontWeight("bold");
  }

  if (!ss.getSheetByName("Campi")) {
    var sheetCampi = ss.insertSheet("Campi");
    sheetCampi.appendRow(["Nome Campo"]);
    sheetCampi.getRange("A1").setFontWeight("bold");
  }

  var sheetPartite = ss.getSheetByName("Partite");
  if (!sheetPartite) {
    sheetPartite = ss.insertSheet("Partite");
    sheetPartite.appendRow(["ID", "Data Inserimento", "Dettagli Partita", "Costo", "Convocati", "Stato", "Risultato", "Referto"]);
    sheetPartite.getRange("A1:H1").setFontWeight("bold");
  }
}


function ottieniDati() {
  try {
    var ss = getTargetSpreadsheet();

    var foglioGiocatori = ss.getSheetByName("Giocatori");
    var giocatori = [];
    if (foglioGiocatori) {
      var dati = foglioGiocatori.getDataRange().getValues();
      for (var i = 1; i < dati.length; i++) {
        if (dati[i][0] !== "") {
          var attivoVal = dati[i][11];
          var isAttivo = (attivoVal === true || String(attivoVal).toLowerCase() === "true" || attivoVal === "Vero");

          giocatori.push({
            nome: String(dati[i][0]),
            saldo: parseFloat(dati[i][1]) || 0,
            gol: parseInt(dati[i][2]) || 0,
            ammonizioni: parseInt(dati[i][3]) || 0,
            ultimoRuolo: dati[i][4] ? String(dati[i][4]) : "",
            assist: parseInt(dati[i][5]) || 0,
            espulsioni: parseInt(dati[i][6]) || 0,
            golSubitiAzione: parseInt(dati[i][7]) || 0,
            golSubitiRigore: parseInt(dati[i][8]) || 0,
            golSubitiPiazzato: parseInt(dati[i][9]) || 0,
            quotaIscrizione: parseFloat(dati[i][10]) || 0,
            attivo: isAttivo
          });
        }
      }
    }

    var fondoCassa = 0;
    for (var f = 0; f < giocatori.length; f++) {
      fondoCassa += giocatori[f].saldo;
    }

    var foglioCampi = ss.getSheetByName("Campi");
    var campi = [];
    if (foglioCampi) {
      var datiCampi = foglioCampi.getDataRange().getValues();
      for (var k = 1; k < datiCampi.length; k++) {
        if (datiCampi[k][0] !== "") campi.push(String(datiCampi[k][0]));
      }
    }

    var foglioPartite = ss.getSheetByName("Partite");
    var partiteAperte = [];
    var partiteChiuse = [];

    if (foglioPartite) {
      var datiPartite = foglioPartite.getDataRange().getValues();
      for (var j = 1; j < datiPartite.length; j++) {
        if (datiPartite[j][0] === "") continue;

        var dataSicura = "";
        if (datiPartite[j][1] instanceof Date) {
          dataSicura = datiPartite[j][1].toISOString();
        } else {
          dataSicura = String(datiPartite[j][1]);
        }

        var p = {
          id: String(datiPartite[j][0]),
          data: dataSicura,
          dettagli: String(datiPartite[j][2]),
          costo: parseFloat(datiPartite[j][3]) || 0,
          convocati: datiPartite[j][4] ? JSON.parse(datiPartite[j][4]) : [],
          risultato: String(datiPartite[j][6] || ""),
          referto: datiPartite[j][7] ? JSON.parse(datiPartite[j][7]) : []
        };

        if (datiPartite[j][5] === "Aperta") partiteAperte.push(p);
        else if (datiPartite[j][5] === "Chiusa") partiteChiuse.push(p);
      }
    }

    var payload = { giocatori: giocatori, fondoCassa: fondoCassa, partiteAperte: partiteAperte, partiteChiuse: partiteChiuse, campi: campi };
    return JSON.stringify(payload);
    
  } catch (errore) {
    return JSON.stringify({ erroreCritico: "Errore dal Foglio Google: " + errore.message });
  }
}

function aggiungiGiocatore(nome) {
  var ss = getTargetSpreadsheet();
  ss.getSheetByName("Giocatori").appendRow([nome, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, false]);
  ss.getSheetByName("Log").appendRow([new Date(), "Rosa", "-", "Aggiunto giocatore: " + nome + " (Inattivo)"]);
  return ottieniDati();
}

function eliminaGiocatore(nome) {
  var ss = getTargetSpreadsheet();
  var foglioGiocatori = ss.getSheetByName("Giocatori");
  var dati = foglioGiocatori.getDataRange().getValues();
  
  for (var i = 1; i < dati.length; i++) {
    if (dati[i][0] === nome) {
      foglioGiocatori.deleteRow(i + 1);
      break;
    }
  }
  ss.getSheetByName("Log").appendRow([new Date(), "Rosa", "-", "Eliminato giocatore: " + nome]);
  return ottieniDati();
}

function versaQuota(nome, importo) {
  var ss = getTargetSpreadsheet();
  var foglioGiocatori = ss.getSheetByName("Giocatori");
  var dati = foglioGiocatori.getDataRange().getValues();
  
  for (var i = 1; i < dati.length; i++) {
    if (dati[i][0] === nome) {
      var saldoAttuale = parseFloat(dati[i][1]) || 0;
      foglioGiocatori.getRange(i + 1, 2).setValue(saldoAttuale + parseFloat(importo));
      break;
    }
  }
  ss.getSheetByName("Log").appendRow([new Date(), "Ricarica", importo, "Versamento da: " + nome]);
  return ottieniDati();
}

function versaQuoteMassive(versamenti) {
  try {
    var ss = getTargetSpreadsheet();
    var foglioGiocatori = ss.getSheetByName("Giocatori");
    var dati = foglioGiocatori.getDataRange().getValues();
    var foglioLog = ss.getSheetByName("Log");

    for (var v = 0; v < versamenti.length; v++) {
      var nome = versamenti[v].nome;
      var importo = parseFloat(versamenti[v].importo);
      
      for (var i = 1; i < dati.length; i++) {
        if (dati[i][0] === nome) {
          var saldoAttuale = parseFloat(dati[i][1]) || 0;
          foglioGiocatori.getRange(i + 1, 2).setValue(saldoAttuale + importo);
          foglioLog.appendRow([new Date(), "Ricarica Massiva", importo, "Versamento da: " + nome]);
          dati[i][1] = saldoAttuale + importo;
          break;
        }
      }
    }
    return ottieniDati();
  } catch (e) {
    return JSON.stringify({ erroreCritico: "Errore durante il versamento massivo: " + e.message });
  }
}

function versaQuotaIscrizione(nome, importo) {
  var ss = getTargetSpreadsheet();
  var foglioGiocatori = ss.getSheetByName("Giocatori");
  var dati = foglioGiocatori.getDataRange().getValues();
  
  for (var i = 1; i < dati.length; i++) {
    if (dati[i][0] === nome) {
      var quotaAttuale = parseFloat(dati[i][10]) || 0;
      foglioGiocatori.getRange(i + 1, 11).setValue(quotaAttuale + parseFloat(importo));
      break;
    }
  }
  ss.getSheetByName("Log").appendRow([new Date(), "Quota Iscrizione", importo, "Versamento iscrizione da: " + nome]);
  return ottieniDati();
}

function cambiaStatoGiocatore(nome, nuovoStato) {
  var ss = getTargetSpreadsheet();
  var foglioGiocatori = ss.getSheetByName("Giocatori");
  var dati = foglioGiocatori.getDataRange().getValues();
  
  for (var i = 1; i < dati.length; i++) {
    if (dati[i][0] === nome) {
      foglioGiocatori.getRange(i + 1, 12).setValue(nuovoStato);
      break;
    }
  }
  var statoTxt = nuovoStato ? "Attivato" : "Disattivato";
  ss.getSheetByName("Log").appendRow([new Date(), "Stato Giocatore", "-", "Giocatore " + nome + " impostato come: " + statoTxt]);
  return ottieniDati();
}

function creaPartita(costoTotale, giocatoriSelezionati, dettagliPartita, nomeCampo, mappaRuoli) {
  var ss = getTargetSpreadsheet();
  
  if (nomeCampo) {
    var foglioCampi = ss.getSheetByName("Campi");
    if (!foglioCampi) {
      foglioCampi = ss.insertSheet("Campi");
      foglioCampi.appendRow(["Nome Campo"]);
    }
    var datiCampi = foglioCampi.getDataRange().getValues().flat();
    if (datiCampi.indexOf(nomeCampo) === -1) {
      foglioCampi.appendRow([nomeCampo]);
    }
  }

  var idPartita = Utilities.getUuid();
  ss.getSheetByName("Partite").appendRow([
    idPartita, new Date(), dettagliPartita, costoTotale, JSON.stringify(giocatoriSelezionati), "Aperta", "", ""
  ]);

  if (mappaRuoli) {
    var foglioGiocatori = ss.getSheetByName("Giocatori");
    var datiGiocatori = foglioGiocatori.getDataRange().getValues();
    for (var j = 1; j < datiGiocatori.length; j++) {
      var nomeG = String(datiGiocatori[j][0]);
      if (mappaRuoli[nomeG]) {
        foglioGiocatori.getRange(j + 1, 5).setValue(mappaRuoli[nomeG]);
      }
    }
  }

  return ottieniDati();
}

function creaAmichevole(dati) {
  try {
    var ss = getTargetSpreadsheet();
    var sheetPartite = ss.getSheetByName("Partite");
    var idPartita = Utilities.getUuid();
    
    if (dati.campo) {
      var sheetCampi = ss.getSheetByName("Campi");
      if (!sheetCampi) {
        sheetCampi = ss.insertSheet("Campi");
        sheetCampi.appendRow(["Nome Campo"]);
      }
      var datiCampi = sheetCampi.getDataRange().getValues().flat();
      if (datiCampi.indexOf(dati.campo) === -1) sheetCampi.appendRow([dati.campo]);
    }

    var dataObj = new Date(dati.data + 'T12:00:00');
    var opzioniData = { weekday: 'long', day: 'numeric', month: 'long' };
    var dataLeggibile = dataObj.toLocaleDateString('it-IT', opzioniData);
    dataLeggibile = dataLeggibile.charAt(0).toUpperCase() + dataLeggibile.slice(1);

    var dettaglioFormazione = "";
    if(dati.formazioni) {
      dettaglioFormazione = "\n\n👕 " + dati.squadraA + ":\n" + dati.formazioni.A.join(", ") + "\n\n👕 " + dati.squadraB + ":\n" + dati.formazioni.B.join(", ");
    }

    var dettagli = dataLeggibile + " ore " + dati.ora + " - " + dati.campo + " (" + dati.squadraA + " vs " + dati.squadraB + ") [Amichevole]" + dettaglioFormazione;
    
    sheetPartite.appendRow([
      idPartita, 
      new Date(), 
      dettagli, 
      parseFloat(dati.costo) || 0, 
      JSON.stringify(dati.giocatori), 
      "Aperta", 
      "", 
      ""
    ]);
    
    ss.getSheetByName("Log").appendRow([new Date(), "Creazione Amichevole", dati.costo || 0, "Creata con formazioni e referto in attesa"]);
    return ottieniDati();
  } catch (e) {
    return JSON.stringify({ erroreCritico: "Errore durante la creazione dell'amichevole: " + e.message });
  }
}

function chiudiPartita(idPartita, costoFinale, presenti, risultato, referto) {
  var ss = getTargetSpreadsheet();
  var foglioPartite = ss.getSheetByName("Partite");
  var foglioGiocatori = ss.getSheetByName("Giocatori");
  
  var datiPartite = foglioPartite.getDataRange().getValues();
  var dettagliPartita = "";
  var rigaPartita = -1;
  
  for (var i = 1; i < datiPartite.length; i++) {
    if (datiPartite[i][0] === idPartita) {
      rigaPartita = i + 1;
      dettagliPartita = datiPartite[i][2].split('\n\n')[0]; 
      foglioPartite.getRange(rigaPartita, 6).setValue("Chiusa");
      foglioPartite.getRange(rigaPartita, 7).setValue(risultato);
      foglioPartite.getRange(rigaPartita, 8).setValue(JSON.stringify(referto));
      foglioPartite.getRange(rigaPartita, 4).setValue(parseFloat(costoFinale) || 0);
      break;
    }
  }

  if (rigaPartita === -1) return ottieniDati();

  var paganti = 0;
  for (var k = 0; k < referto.length; k++) {
    if (presenti.indexOf(referto[k].nome) !== -1 && referto[k].pagaQuota === true) {
      paganti++;
    }
  }
  var quotaIndividuale = paganti > 0 ? parseFloat(costoFinale) / paganti : 0;

  var datiGiocatori = foglioGiocatori.getDataRange().getValues();
  
  for (var j = 1; j < datiGiocatori.length; j++) {
    var nome = datiGiocatori[j][0];
    if (presenti.indexOf(nome) !== -1 && !nome.includes("(Esterno)")) {
      var saldoAttuale = parseFloat(datiGiocatori[j][1]) || 0;
      var golAttuali = parseInt(datiGiocatori[j][2]) || 0;
      var ammAttuali = parseInt(datiGiocatori[j][3]) || 0;
      var assistAttuali = parseInt(datiGiocatori[j][5]) || 0;
      var rossiAttuali = parseInt(datiGiocatori[j][6]) || 0;
      var subAzioneAttuali = parseInt(datiGiocatori[j][7]) || 0;
      var subRigoreAttuali = parseInt(datiGiocatori[j][8]) || 0;
      var subPiazzatoAttuali = parseInt(datiGiocatori[j][9]) || 0;
      
      var datiReferto = referto.find(r => r.nome === nome);
      var golPartita = datiReferto ? parseInt(datiReferto.gol) : 0;
      var ammPartita = datiReferto ? parseInt(datiReferto.amm) : 0;
      var assistPartita = datiReferto ? parseInt(datiReferto.assist) : 0;
      var rossiPartita = datiReferto ? parseInt(datiReferto.rossi) : 0;
      var subAzionePartita = datiReferto ? parseInt(datiReferto.subitiAzione) : 0;
      var subRigorePartita = datiReferto ? parseInt(datiReferto.subitiRigore) : 0;
      var subPiazzatoPartita = datiReferto ? parseInt(datiReferto.subitiPiazzato) : 0;
      var pagaQuota = datiReferto ? datiReferto.pagaQuota : true;
      
      if (pagaQuota) {
        foglioGiocatori.getRange(j + 1, 2).setValue(saldoAttuale - quotaIndividuale);
      }
      
      foglioGiocatori.getRange(j + 1, 3).setValue(golAttuali + golPartita);
      foglioGiocatori.getRange(j + 1, 4).setValue(ammAttuali + ammPartita);
      foglioGiocatori.getRange(j + 1, 6).setValue(assistAttuali + assistPartita);
      foglioGiocatori.getRange(j + 1, 7).setValue(rossiAttuali + rossiPartita);
      foglioGiocatori.getRange(j + 1, 8).setValue(subAzioneAttuali + subAzionePartita);
      foglioGiocatori.getRange(j + 1, 9).setValue(subRigoreAttuali + subRigorePartita);
      foglioGiocatori.getRange(j + 1, 10).setValue(subPiazzatoAttuali + subPiazzatoPartita);
    }
  }
  
  var infoDettagli = dettagliPartita ? "Dettagli: " + dettagliPartita + " | " : "";
  ss.getSheetByName("Log").appendRow([new Date(), "Chiusura Partita", costoFinale, infoDettagli + "Presenti: " + presenti.length + " (" + paganti + " paganti) (Quota: " + quotaIndividuale.toFixed(2) + "€)"]);
  return ottieniDati();
}

function annullaPartita(idPartita) {
  var ss = getTargetSpreadsheet();
  var foglioPartite = ss.getSheetByName("Partite");
  var dati = foglioPartite.getDataRange().getValues();
  
  for (var i = 1; i < dati.length; i++) {
    if (dati[i][0] === idPartita) {
      var dettagli = dati[i][2].split('\n\n')[0];
      foglioPartite.deleteRow(i + 1);
      ss.getSheetByName("Log").appendRow([new Date(), "Annullamento", "-", "Annullata partita in attesa: " + dettagli]);
      break;
    }
  }
  return ottieniDati();
}

function function_stornaEffettiReferto(foglioGiocatori, oldReferto, oldCosto) {
  var pagantiOld = 0;
  for(var k=0; k<oldReferto.length; k++) {
    if(oldReferto[k].pagaQuota) pagantiOld++;
  }
  var oldQuota = pagantiOld > 0 ? parseFloat(oldCosto) / pagantiOld : 0;
  var datiGiocatori = foglioGiocatori.getDataRange().getValues();

  for (var j = 1; j < datiGiocatori.length; j++) {
    var nome = datiGiocatori[j][0];
    var oldRef = oldReferto.find(r => r.nome === nome);
    
    if (oldRef && !nome.includes("(Esterno)")) {
      var saldoAttuale = parseFloat(datiGiocatori[j][1]) || 0;
      var golAttuali = parseInt(datiGiocatori[j][2]) || 0;
      var ammAttuali = parseInt(datiGiocatori[j][3]) || 0;
      var assistAttuali = parseInt(datiGiocatori[j][5]) || 0;
      var rossiAttuali = parseInt(datiGiocatori[j][6]) || 0;
      var subAzioneAttuali = parseInt(datiGiocatori[j][7]) || 0;
      var subRigoreAttuali = parseInt(datiGiocatori[j][8]) || 0;
      var subPiazzatoAttuali = parseInt(datiGiocatori[j][9]) || 0;

      if (oldRef.pagaQuota) {
        foglioGiocatori.getRange(j + 1, 2).setValue(saldoAttuale + oldQuota); 
      }
      
      foglioGiocatori.getRange(j + 1, 3).setValue(golAttuali - (parseInt(oldRef.gol) || 0));
      foglioGiocatori.getRange(j + 1, 4).setValue(ammAttuali - (parseInt(oldRef.amm) || 0));
      foglioGiocatori.getRange(j + 1, 6).setValue(assistAttuali - (parseInt(oldRef.assist) || 0));
      foglioGiocatori.getRange(j + 1, 7).setValue(rossiAttuali - (parseInt(oldRef.rossi) || 0));
      foglioGiocatori.getRange(j + 1, 8).setValue(subAzioneAttuali - (parseInt(oldRef.subitiAzione) || 0));
      foglioGiocatori.getRange(j + 1, 9).setValue(subRigoreAttuali - (parseInt(oldRef.subitiRigore) || 0));
      foglioGiocatori.getRange(j + 1, 10).setValue(subPiazzatoAttuali - (parseInt(oldRef.subitiPiazzato) || 0));
    }
  }
}

function annullaPartitaArchiviata(idPartita) {
  try {
    var ss = getTargetSpreadsheet();
    var foglioPartite = ss.getSheetByName("Partite");
    var foglioGiocatori = ss.getSheetByName("Giocatori");
    var datiPartite = foglioPartite.getDataRange().getValues();
    
    for (var i = 1; i < datiPartite.length; i++) {
      if (datiPartite[i][0] === idPartita) {
        var dettagli = datiPartite[i][2].split('\n\n')[0];
        var oldCosto = parseFloat(datiPartite[i][3]) || 0;
        var oldReferto = datiPartite[i][7] ? JSON.parse(datiPartite[i][7]) : [];
        
        function_stornaEffettiReferto(foglioGiocatori, oldReferto, oldCosto);
        
        foglioPartite.deleteRow(i + 1);
        ss.getSheetByName("Log").appendRow([new Date(), "Storno e Annullamento", oldCosto, "Annullata partita archiviata e ripristinati saldi/statistiche: " + dettagli]);
        break;
      }
    }
    return ottieniDati();
  } catch (e) {
    return JSON.stringify({ erroreCritico: "Errore durante l'annullamento archiviato: " + e.message });
  }
}

function modificaPartitaArchiviata(idPartita, nuovoCosto, nuovoRisultato, nuovoReferto) {
  try {
    var ss = getTargetSpreadsheet();
    var foglioPartite = ss.getSheetByName("Partite");
    var foglioGiocatori = ss.getSheetByName("Giocatori");
    var datiPartite = foglioPartite.getDataRange().getValues();
    
    for (var i = 1; i < datiPartite.length; i++) {
      if (datiPartite[i][0] === idPartita) {
        var dettagli = datiPartite[i][2].split('\n\n')[0];
        var oldCosto = parseFloat(datiPartite[i][3]) || 0;
        var oldReferto = datiPartite[i][7] ? JSON.parse(datiPartite[i][7]) : [];
        
        function_stornaEffettiReferto(foglioGiocatori, oldReferto, oldCosto);

        var pagantiNuovi = 0;
        for(var k=0; k<nuovoReferto.length; k++) {
          if(nuovoReferto[k].pagaQuota) pagantiNuovi++;
        }
        var nuovaQuota = pagantiNuovi > 0 ? parseFloat(nuovoCosto) / pagantiNuovi : 0;
        var datiGiocatori = foglioGiocatori.getDataRange().getValues();

        for (var j = 1; j < datiGiocatori.length; j++) {
          var nome = datiGiocatori[j][0];
          var newRef = nuovoReferto.find(r => r.nome === nome);
          
          if (newRef && !nome.includes("(Esterno)")) {
            var saldoAttuale = parseFloat(datiGiocatori[j][1]) || 0;
            if (newRef.pagaQuota) {
              foglioGiocatori.getRange(j + 1, 2).setValue(saldoAttuale - nuovaQuota); 
            }
            
            foglioGiocatori.getRange(j + 1, 3).setValue((parseInt(datiGiocatori[j][2]) || 0) + (parseInt(newRef.gol) || 0));
            foglioGiocatori.getRange(j + 1, 4).setValue((parseInt(datiGiocatori[j][3]) || 0) + (parseInt(newRef.amm) || 0));
            foglioGiocatori.getRange(j + 1, 6).setValue((parseInt(datiGiocatori[j][5]) || 0) + (parseInt(newRef.assist) || 0));
            foglioGiocatori.getRange(j + 1, 7).setValue((parseInt(datiGiocatori[j][6]) || 0) + (parseInt(newRef.rossi) || 0));
            foglioGiocatori.getRange(j + 1, 8).setValue((parseInt(datiGiocatori[j][7]) || 0) + (parseInt(newRef.subitiAzione) || 0));
            foglioGiocatori.getRange(j + 1, 9).setValue((parseInt(datiGiocatori[j][8]) || 0) + (parseInt(newRef.subitiRigore) || 0));
            foglioGiocatori.getRange(j + 1, 10).setValue((parseInt(datiGiocatori[j][9]) || 0) + (parseInt(newRef.subitiPiazzato) || 0));
          }
        }

        foglioPartite.getRange(i + 1, 4).setValue(parseFloat(nuovoCosto) || 0);
        foglioPartite.getRange(i + 1, 7).setValue(nuovoRisultato);
        foglioPartite.getRange(i + 1, 8).setValue(JSON.stringify(nuovoReferto));

        ss.getSheetByName("Log").appendRow([new Date(), "Modifica Partita", nuovoCosto, "Modificata partita in archivio (Saldi e Stat ricalcolati): " + dettagli]);
        break;
      }
    }
    return ottieniDati();
  } catch (e) {
    return JSON.stringify({ erroreCritico: "Errore durante la modifica archiviata: " + e.message });
  }
}

function registraSpesaCondivisa(importoTotale, causale, giocatoriSelezionati) {
  var ss = getTargetSpreadsheet();
  var foglioGiocatori = ss.getSheetByName("Giocatori");
  var dati = foglioGiocatori.getDataRange().getValues();
  
  var quotaIndividuale = parseFloat(importoTotale) / giocatoriSelezionati.length;
  
  for (var i = 1; i < dati.length; i++) {
    var nome = dati[i][0];
    if (giocatoriSelezionati.indexOf(nome) !== -1) {
      var saldoAttuale = parseFloat(dati[i][1]) || 0;
      foglioGiocatori.getRange(i + 1, 2).setValue(saldoAttuale - quotaIndividuale);
    }
  }
  
  ss.getSheetByName("Log").appendRow([new Date(), "Spesa Condivisa", importoTotale, causale + " (Diviso tra " + giocatoriSelezionati.length + " giocatori: " + quotaIndividuale.toFixed(2) + "€ a testa)"]);
  return ottieniDati();
}

function modificaGiocatore(nomeOriginale, datiAggiornati) {
  var ss = getTargetSpreadsheet();
  var foglio = ss.getSheetByName("Giocatori");
  var values = foglio.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === nomeOriginale) {
      foglio.getRange(i + 1, 1).setValue(datiAggiornati.nome);
      foglio.getRange(i + 1, 2).setValue(parseFloat(datiAggiornati.saldo) || 0);
      foglio.getRange(i + 1, 3).setValue(parseInt(datiAggiornati.gol) || 0);
      foglio.getRange(i + 1, 4).setValue(parseInt(datiAggiornati.ammonizioni) || 0);
      foglio.getRange(i + 1, 6).setValue(parseInt(datiAggiornati.assist) || 0);
      foglio.getRange(i + 1, 7).setValue(parseInt(datiAggiornati.espulsioni) || 0);
      foglio.getRange(i + 1, 8).setValue(parseInt(datiAggiornati.golSubitiAzione) || 0);
      foglio.getRange(i + 1, 9).setValue(parseInt(datiAggiornati.golSubitiRigore) || 0);
      foglio.getRange(i + 1, 10).setValue(parseInt(datiAggiornati.golSubitiPiazzato) || 0);
      break;
    }
  }
  
  var dettaglioModifica = "Modificati dati per: " + nomeOriginale;
  if (nomeOriginale !== datiAggiornati.nome) {
    dettaglioModifica += " (Nome cambiato in " + datiAggiornati.nome + ")";
  }
  ss.getSheetByName("Log").appendRow([new Date(), "Modifica Rosa", datiAggiornati.saldo, dettaglioModifica]);
  
  return ottieniDati();
}