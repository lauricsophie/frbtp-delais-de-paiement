/**
 * FRBTP — Délais de paiement — Backend anonymisé
 * Ne stocke aucune donnée nominative d'entreprise (ni nom, ni SIRET).
 * Regroupement des factures d'une même saisie via un idSession (UUID) technique,
 * sans lien avec une identité réelle.
 */

const SHEET_ID = 'REMPLACER_PAR_ID_DU_GOOGLE_SHEET';
const SHEET_NAME = 'Signalements';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

    if (!Array.isArray(data.factures) || data.factures.length === 0) {
      return jsonResponse({ status: 'error', message: 'Aucune facture recue' });
    }

    data.factures.forEach(f => {
      const dateEcheance = new Date(f.dateEcheance);
      const aujourdHui = new Date();
      const joursRetard = f.statut === 'en_attente'
        ? Math.max(0, Math.floor((aujourdHui - dateEcheance) / 86400000))
        : '';

      sheet.appendRow([
        data.idSession,
        data.horodatage,
        data.secteur,
        data.effectif,
        data.zone,
        data.donneurOrdre,
        data.typeDonneurOrdre,
        f.reference || '',
        f.dateEmission,
        f.dateEcheance,
        f.montant,
        f.statut,
        joursRetard
      ]);
    });

    return jsonResponse({ status: 'ok', factures_enregistrees: data.factures.length });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (e.parameter.action === 'liste_donneurs') {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    const colDonneurOrdre = 5;
    const set = new Set();
    for (let i = 1; i < values.length; i++) {
      if (values[i][colDonneurOrdre]) set.add(values[i][colDonneurOrdre]);
    }
    return jsonResponse([...set].sort());
  }
  return jsonResponse({ status: 'ok', message: 'API FRBTP delais de paiement' });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/*
 NOTE POUR L'INTEGRATION :
 - Le calcul des penalites de retard base sur taux.json (taux d'interet legal / indemnite forfaitaire)
   present dans la version precedente du script doit etre reintegre ici (fonction de calcul du montant
   de penalite par facture) : le contenu exact de l'ancien apps-script.gs n'a pas pu etre recupere
   automatiquement lors de la preparation de cette pull request. Merci de relire et fusionner avant merge.
 - Penser a creer les en-tetes de colonnes dans l'onglet "Signalements" :
   idSession | horodatage | secteur | effectif | zone | donneurOrdre | typeDonneurOrdre |
   reference | dateEmission | dateEcheance | montant | statut | joursRetard
*/
