/**
 * Google Apps Script — Récepteur des soumissions du Baromètre FRBTP
 * Déploiement : Extensions > Apps Script sur un Google Sheet vide,
 * coller ce code, puis Déployer > Nouveau déploiement > Application Web
 * (Exécuter en tant que "Moi", accès "Tout le monde").
 */

const SHEET_NAME = "Reponses";

const COLUMNS = [
  "horodatage","secteur","taille_entreprise","type_client","regime_juridique","nom_donneur",
  "montant_ht","date_reception","delai_type","delai_jours_personnalise","date_echeance",
  "date_paiement","toujours_impaye","frequence_retards","penalites_reclamees","consequences",
  "email","telephone","consentement",
  "jours_retard","taux_moyen_pct","base_legale","interets_moratoires","indemnite_forfaitaire",
  "total_reclamable","statut"
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const r = payload.resultat_calcul || {};
    const sheet = getSheet_();

    const row = [
      payload.horodatage || new Date().toISOString(),
      payload.secteur, payload.taille_entreprise, payload.type_client, payload.regime_juridique,
      payload.nom_donneur || "",
      payload.montant_ht, payload.date_reception, payload.delai_type, payload.delai_jours_personnalise || "",
      payload.date_echeance, payload.date_paiement || "", payload.toujours_impaye,
      payload.frequence_retards || "", payload.penalites_reclamees || "",
      (payload.consequences || []).join(";"),
      payload.email || "", payload.telephone || "", payload.consentement,
      r.jours_retard, r.taux_moyen_pct, r.base_legale, r.interets_moratoires,
      r.indemnite_forfaitaire, r.total_reclamable, r.statut
    ];
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = (e.parameter.action || "stats");
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  if (action === "csv") {
    const csv = [headers.join(",")].concat(
      data.map(row => row.map(v => `"${String(v).replace(/"/g,'""')}"`).join(","))
    ).join("\n");
    return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
  }

  const idx = name => headers.indexOf(name);
  const stats = {};

  data.forEach(row => {
    const client = row[idx("type_client")];
    if (!client) return;
    if (!stats[client]) stats[client] = { count: 0, sum_jours: 0, sum_total: 0, depassements: 0 };
    stats[client].count++;
    stats[client].sum_jours += Number(row[idx("jours_retard")]) || 0;
    stats[client].sum_total += Number(row[idx("total_reclamable")]) || 0;
    if ((Number(row[idx("jours_retard")]) || 0) > 0) stats[client].depassements++;
  });

  const result = Object.entries(stats).map(([client, s]) => ({
    type_client: client,
    nb_reponses: s.count,
    delai_retard_moyen_jours: Math.round(s.sum_jours / s.count),
    total_interets_du: Math.round(s.sum_total * 100) / 100,
    moyenne_interets_du: Math.round((s.sum_total / s.count) * 100) / 100,
    pct_depassement_delai_legal: Math.round((s.depassements / s.count) * 1000) / 10
  }));

  return ContentService.createTextOutput(JSON.stringify({ ok: true, stats: result, total_reponses: data.length }))
    .setMimeType(ContentService.MimeType.JSON);
}
