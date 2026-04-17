// ============================================================
//  INVENTAIRE STOCK FLACONNAGE — Normec Abiolab
//  Google Apps Script — Code.gs
// ============================================================

var SHEET_ID        = '1npxsBujw471ryLc3w4e2Qwt6zhVVJ5XE-JHn0Q5UYR0';
var SHEET_CATALOGUE = 'Catalogue';
var SHEET_GESTION   = 'Commande gestion stock';
var SHEET_NORD      = 'CommandeNord';

// ─── doGet : sert la page HTML ────────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Inventaire Stock — Abiolab')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── getCatalogueData : appelée par google.script.run ────────────────────────
function getCatalogueData() {
  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var ws       = ss.getSheetByName(SHEET_CATALOGUE);
  var data     = ws.getDataRange().getValues();
  var produits = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var nom = row[0] ? String(row[0]).trim() : '';
    if (!nom) continue;

    var mini    = (row[3] !== '' && row[3] !== null) ? Number(row[3]) : null;
    var max     = (row[4] !== '' && row[4] !== null) ? Number(row[4]) : null;
    var numNord = (row[10] !== '' && row[10] !== null && !isNaN(Number(row[10]))) ? Number(row[10]) : null;

    produits.push({
      nom:         nom,
      unite:       row[2] ? String(row[2]).trim() : '',
      mini:        mini,
      max:         max,
      circuit:     row[6] ? String(row[6]).trim() : '',
      fournisseur: row[7] ? String(row[7]).trim() : '',
      reference:   row[8] ? String(row[8]).trim() : '',
      categorie:   row[9] ? String(row[9]).trim() : 'Autres',
      numNord:     numNord
    });
  }

  return { ok: true, produits: produits };
}

// ─── sauvegarderEtEnvoyer : tout en une seule fonction ───────────────────────
function sauvegarderEtEnvoyer(inventaire) {
  try {
    if (!inventaire || !Array.isArray(inventaire)) {
      return { ok: false, error: 'Données manquantes' };
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);

    // 1. Écrire les stocks dans le Catalogue
    majCatalogue(ss, inventaire);

    // 2. Générer les bons de commande + calculer alertes/zéros
    var result = genererBonCommandes(ss);

    // Force la validation des écritures dans les feuilles avant l'export PDF
    SpreadsheetApp.flush();

    // 3. Envoyer l'email
    envoyerMails(result.alertes, result.zeros);

    return { ok: true, nbAlertes: result.alertes.length, nbZeros: result.zeros.length };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ─── majCatalogue ─────────────────────────────────────────────────────────────
function majCatalogue(ss, inventaire) {
  var ws    = ss.getSheetByName(SHEET_CATALOGUE);
  var data  = ws.getDataRange().getValues();
  var index = {};

  for (var i = 1; i < data.length; i++) {
    var nom = data[i][0] ? String(data[i][0]).trim() : '';
    if (nom) index[nom] = i + 1;
  }

  for (var j = 0; j < inventaire.length; j++) {
    var item = inventaire[j];
    if (item.stock === null || item.stock === undefined) continue;

    var row = index[item.nom];
    if (row) ws.getRange(row, 2).setValue(item.stock);
  }
}

// ─── genererBonCommandes ──────────────────────────────────────────────────────
function genererBonCommandes(ss) {
  SpreadsheetApp.flush();

  // Rouvrir pour forcer la relecture des valeurs à jour
  var ws      = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_CATALOGUE);
  var data    = ws.getDataRange().getValues();
  var alertes = [];
  var zeros   = [];

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var nom     = row[0] ? String(row[0]).trim() : '';
    var stock   = (row[1] !== '' && row[1] !== null) ? Number(row[1]) : null;
    var unite   = row[2] ? String(row[2]).trim() : '';
    var mini    = (row[3] !== '' && row[3] !== null) ? Number(row[3]) : null;
    var max     = (row[4] !== '' && row[4] !== null) ? Number(row[4]) : null;
    var circuit = row[6] ? String(row[6]).trim() : '';
    var fourn   = row[7] ? String(row[7]).trim() : '';
    var ref     = row[8] ? String(row[8]).trim() : '';
    var cat     = row[9] ? String(row[9]).trim() : '';
    var numNord = (row[10] !== '' && row[10] !== null && !isNaN(Number(row[10]))) ? Number(row[10]) : null;

    if (!nom) continue;

    // Stock à zéro
    if (stock === 0) {
      zeros.push({
        nom: nom,
        unite: unite,
        circuit: circuit,
        categorie: cat,
        fournisseur: fourn,
        reference: ref
      });
    }

    // Stock faible : calculé directement en JS (stock <= mini)
    if (stock !== null && mini !== null && max !== null && stock <= mini) {
      var aCmd = max + ' ' + unite + (unite.slice(-1) !== 's' ? 's' : '');
      alertes.push({
        nom: nom,
        aCmd: aCmd,
        maxQte: max,
        circuit: circuit,
        numNord: numNord,
        fournisseur: fourn,
        reference: ref,
        unite: unite,
        categorie: cat
      });
    }
  }

  majBonGestion(ss, alertes);
  majCommandeNord(ss, alertes);

  return { alertes: alertes, zeros: zeros, nbAlertes: alertes.length };
}

// ─── majBonGestion ────────────────────────────────────────────────────────────
function majBonGestion(ss, alertes) {
  var ws        = ss.getSheetByName(SHEET_GESTION);
  var allData   = ws.getDataRange().getValues();
  var headerRow = 3;

  for (var i = 0; i < allData.length; i++) {
    if (String(allData[i][0]).indexOf('NOM') !== -1) {
      headerRow = i + 1;
      break;
    }
  }

  var lastRow = ws.getLastRow();
  if (lastRow > headerRow) {
    ws.getRange(headerRow + 1, 1, lastRow - headerRow, 4).clearContent();
  }

  var idx = 0;
  for (var j = 0; j < alertes.length; j++) {
    var item = alertes[j];
    if (item.circuit.toLowerCase().indexOf('gestion') === -1) continue;

    var row = headerRow + 1 + idx;
    ws.getRange(row, 1).setValue(item.nom);
    ws.getRange(row, 2).setValue(item.reference);
    ws.getRange(row, 3).setValue(item.fournisseur);
    ws.getRange(row, 4).setValue(item.aCmd);
    idx++;
  }
}

// ─── majCommandeNord ──────────────────────────────────────────────────────────
function majCommandeNord(ss, alertes) {
  var ws         = ss.getSheetByName(SHEET_NORD);
  var allData    = ws.getDataRange().getValues();
  var ligneIndex = {};

  // Indexation des lignes par numéro (colonne A)
  for (var i = 0; i < allData.length; i++) {
    var num = allData[i][0];
    if (num !== null && num !== '' && !isNaN(Number(num))) {
      ligneIndex[Number(num)] = i + 1;
    }
  }

  // Nettoyage colonne Quantité (colonne C = 3)
  var rows = Object.keys(ligneIndex);
  for (var r = 0; r < rows.length; r++) {
    ws.getRange(ligneIndex[Number(rows[r])], 3).clearContent();
  }

  // Injection des quantités
  for (var j = 0; j < alertes.length; j++) {
    var item = alertes[j];
    if (!item.numNord) continue;

    var row = ligneIndex[item.numNord];
    if (!row) continue;

    if (item.circuit.toLowerCase().indexOf('nord') !== -1) {
      // Circuit Nord → quantité complète
      ws.getRange(row, 3).setValue(item.aCmd);
    } else if (item.circuit.toLowerCase().indexOf('gestion') !== -1) {
      // Circuit Gestion avec N° Nord → 10% arrondi à l'entier supérieur
      var qte10 = Math.ceil(item.maxQte * 0.1);
      if (qte10 < 1) qte10 = 1;
      ws.getRange(row, 3).setValue(qte10 + ' ' + item.unite);
    }
  }
}

// ─── envoyerMails ─────────────────────────────────────────────────────────────
function envoyerMails(alertes, zeros) {
  var DESTINATAIRES = 'jerome.bourgois@normecgroup.com,julie.thollet@normecgroup.com';
  var now           = new Date();
  var dateStr       = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var semaine       = getSemaine(now);

  // Section 1 — stocks faibles
  var htmlAlertes = '';
  if (alertes.length > 0) {
    var gestion = [];
    var nord    = [];

    for (var i = 0; i < alertes.length; i++) {
      if (alertes[i].circuit.toLowerCase().indexOf('gestion') !== -1) gestion.push(alertes[i]);
      else nord.push(alertes[i]);
    }

    if (gestion.length > 0) {
      htmlAlertes += '<h3 style="color:#6B35A8;margin:16px 0 8px;font-size:13px">Commande Gestion stock</h3>';
      htmlAlertes += tableauAlertes(gestion);
    }

    if (nord.length > 0) {
      htmlAlertes += '<h3 style="color:#2456B0;margin:16px 0 8px;font-size:13px">Commande Nord</h3>';
      htmlAlertes += tableauAlertes(nord);
    }
  } else {
    htmlAlertes = '<p style="color:#1E6B3C;font-size:13px">Tous les stocks sont au-dessus des seuils minimum.</p>';
  }

  // Section 2 — stocks à zéro
  var htmlZeros = '';
  if (zeros.length > 0) {
    htmlZeros = '<table style="border-collapse:collapse;width:100%;font-size:13px">'
      + '<tr style="background:#FDF0E8">'
      + '<th style="text-align:left;padding:8px 12px;border:1px solid #F0C4A8">Produit</th>'
      + '<th style="text-align:left;padding:8px 12px;border:1px solid #F0C4A8">Catégorie</th>'
      + '<th style="text-align:left;padding:8px 12px;border:1px solid #F0C4A8">Circuit</th>'
      + '</tr>';

    for (var z = 0; z < zeros.length; z++) {
      htmlZeros += '<tr>'
        + '<td style="padding:7px 12px;border:1px solid #E2E0DA">' + zeros[z].nom + '</td>'
        + '<td style="padding:7px 12px;border:1px solid #E2E0DA">' + zeros[z].categorie + '</td>'
        + '<td style="padding:7px 12px;border:1px solid #E2E0DA">' + zeros[z].circuit + '</td>'
        + '</tr>';
    }

    htmlZeros += '</table>';
  } else {
    htmlZeros = '<p style="color:#1E6B3C;font-size:13px">Aucun stock à zéro cette semaine.</p>';
  }

  var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;color:#1A1916">'
    + '<div style="background:#1E6B3C;padding:24px 28px;border-radius:8px 8px 0 0">'
    + '<div style="color:rgba(255,255,255,0.7);font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:6px">Normec Abiolab</div>'
    + '<div style="color:white;font-size:20px;font-weight:700">Inventaire Stock — Semaine ' + semaine + '</div>'
    + '<div style="color:rgba(255,255,255,0.65);font-size:13px;margin-top:4px">Réalisé le ' + dateStr + '</div>'
    + '</div>'
    + '<div style="background:#F7F6F3;padding:20px 28px;border:1px solid #E2E0DA;border-top:none;border-radius:0 0 8px 8px">'
    + '<h2 style="font-size:14px;font-weight:700;color:#1A1916;margin:0 0 12px;padding-bottom:8px;border-bottom:2px solid #1E6B3C">'
    + 'Stocks faibles — Commandes à déclencher (' + alertes.length + ' produit' + (alertes.length > 1 ? 's' : '') + ')'
    + '</h2>'
    + htmlAlertes
    + '<h2 style="font-size:14px;font-weight:700;color:#C4521A;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #C4521A">'
    + 'Stocks à zéro — Ruptures (' + zeros.length + ' produit' + (zeros.length > 1 ? 's' : '') + ')'
    + '</h2>'
    + htmlZeros
    + '<p style="font-size:11px;color:#7A7770;margin-top:28px;padding-top:16px;border-top:1px solid #E2E0DA;line-height:1.6">'
    + 'Email généré automatiquement depuis l\'interface de saisie d\'inventaire Abiolab.<br>'
    + 'Les bons de commande ont été mis à jour dans le Google Sheet.'
    + '</p>'
    + '</div></div>';

  var sujet = '[Abiolab] Inventaire stock — S' + semaine + ' ' + dateStr
    + ' — ' + alertes.length + ' commande(s) - ' + zeros.length + ' rupture(s)';

  // ─── Génération des PDF des bons de commande ───
  var ss             = SpreadsheetApp.openById(SHEET_ID);
  var ssId           = ss.getId();
  var piecesJointes  = [];

  // Trouver les IDs des onglets
  var wsGestion = ss.getSheetByName(SHEET_GESTION);
  var wsNord    = ss.getSheetByName(SHEET_NORD);

  // PDF "Commande gestion stock"
  if (wsGestion && alertes.filter(function(a) { return a.circuit.toLowerCase().indexOf('gestion') !== -1; }).length > 0) {
    var urlGestion = 'https://docs.google.com/spreadsheets/d/' + ssId
      + '/export?format=pdf'
      + '&gid=' + wsGestion.getSheetId()
      + '&portrait=true'
      + '&fitw=true'
      + '&gridlines=false'
      + '&printtitle=false'
      + '&sheetnames=false'
      + '&fzr=false';

    var tokenGestion = ScriptApp.getOAuthToken();
    var respGestion  = UrlFetchApp.fetch(urlGestion, { headers: { Authorization: 'Bearer ' + tokenGestion } });

    piecesJointes.push({
      fileName: 'Commande_Gestion_Stock_S' + semaine + '_' + dateStr.replace(/\//g, '-') + '.pdf',
      content:  respGestion.getBlob().getBytes(),
      mimeType: 'application/pdf'
    });
  }

  // PDF "CommandeNord"
  if (wsNord && alertes.filter(function(a) { return a.circuit.toLowerCase().indexOf('nord') !== -1; }).length > 0) {
    var urlNord = 'https://docs.google.com/spreadsheets/d/' + ssId
      + '/export?format=pdf'
      + '&gid=' + wsNord.getSheetId()
      + '&portrait=true'
      + '&fitw=true'
      + '&gridlines=false'
      + '&printtitle=false'
      + '&sheetnames=false'
      + '&fzr=false';

    var tokenNord = ScriptApp.getOAuthToken();
    var respNord  = UrlFetchApp.fetch(urlNord, { headers: { Authorization: 'Bearer ' + tokenNord } });

    piecesJointes.push({
      fileName: 'Commande_Nord_S' + semaine + '_' + dateStr.replace(/\//g, '-') + '.pdf',
      content:  respNord.getBlob().getBytes(),
      mimeType: 'application/pdf'
    });
  }

  MailApp.sendEmail({
    to:          DESTINATAIRES,
    subject:     sujet,
    htmlBody:    htmlBody,
    attachments: piecesJointes
  });
}

function tableauAlertes(liste) {
  var html = '<table style="border-collapse:collapse;width:100%;font-size:13px;margin-bottom:8px">'
    + '<tr style="background:#F2F1EE">'
    + '<th style="text-align:left;padding:8px 12px;border:1px solid #E2E0DA">Produit</th>'
    + '<th style="text-align:left;padding:8px 12px;border:1px solid #E2E0DA">À commander</th>'
    + '<th style="text-align:left;padding:8px 12px;border:1px solid #E2E0DA">Fournisseur</th>'
    + '<th style="text-align:left;padding:8px 12px;border:1px solid #E2E0DA">Référence</th>'
    + '</tr>';

  for (var i = 0; i < liste.length; i++) {
    var item = liste[i];
    html += '<tr>'
      + '<td style="padding:7px 12px;border:1px solid #E2E0DA">' + item.nom + '</td>'
      + '<td style="padding:7px 12px;border:1px solid #E2E0DA;font-weight:700;color:#1E6B3C">' + item.aCmd + '</td>'
      + '<td style="padding:7px 12px;border:1px solid #E2E0DA">' + (item.fournisseur || '—') + '</td>'
      + '<td style="padding:7px 12px;border:1px solid #E2E0DA">' + (item.reference || '—') + '</td>'
      + '</tr>';
  }

  return html + '</table>';
}

function getSemaine(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}
