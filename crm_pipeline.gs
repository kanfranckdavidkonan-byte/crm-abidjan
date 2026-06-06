// ============================================================
// CRM PIPELINE — LIVRAISON ABIDJAN
// Google Apps Script — Version terrain opérationnelle
// ============================================================

// ============================================================
// CONSTANTES CONFIGURABLES
// ============================================================

const COMMERCIAL_EMAIL   = "ton-email@gmail.com";
const FORM_TITLE         = "Nouveau Prospect — Livraison Abidjan";
const SHEET_TITLE        = "CRM Pipeline — Livraison Abidjan";
const CODE_CLIENT_PREFIX = "ABJ";

// Étapes du pipeline
const ETAPES = [
  "Étape 1 — Prospect Identifié",
  "Étape 2 — Prise de Contact",
  "Étape 3 — Devis Envoyé",
  "Étape 4 — Négociation",
  "Étape 5 — Contrat Signé",
  "Étape 6 — Onboarding",
  "Étape 7 — Client Actif",
  "Étape 8 — Fidélisation"
];

// Couleurs pipeline
const COULEUR_VERT   = "#d9ead3"; // priorité haute ≥ 70
const COULEUR_ORANGE = "#fce5cd"; // priorité moyenne 40-69
const COULEUR_GRIS   = "#f3f3f3"; // priorité basse < 40
const COULEUR_HEADER = "#1a73e8";

// ============================================================
// MENU PERSONNALISÉ
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 CRM Pipeline")
    .addItem("➡️ Avancer à l'étape suivante",  "avancerEtapeSuivante")
    .addItem("📄 Générer un devis",             "genererDevis")
    .addItem("✅ Marquer Contrat Signé",         "marquerContratSigne")
    .addItem("❌ Marquer comme Perdu",           "marquerPerdu")
    .addSeparator()
    .addItem("📊 Rafraîchir les KPIs",          "refreshKPIs")
    .addItem("📍 Voir position client sur Maps", "voirPositionMaps")
    .addSeparator()
    .addItem("⚙️ Réinstaller les déclencheurs", "installTriggers")
    .addToUi();
}

// ============================================================
// FONCTION PRINCIPALE — INSTALLATION DU CRM
// ============================================================

function setupCRM() {
  try {
    Logger.log("Démarrage installation CRM...");

    // 1. Créer le Google Form
    var form = creerForm_();

    // 2. Créer le Google Spreadsheet
    var ss = creerSpreadsheet_(form);

    // 3. Lier le Form à la feuille Prospects
    var sheetProspects = ss.getSheetByName("Prospects");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    // 4. Supprimer la feuille "Form Responses 1" créée automatiquement
    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    var formSheet = ss.getSheetByName("Form Responses 1");
    if (formSheet) ss.deleteSheet(formSheet);

    // 5. Remplir la feuille Config
    remplirConfig_(ss);

    // 6. Remplir la feuille Pipeline (vue kanban)
    remplirPipeline_(ss);

    // 7. Initialiser les KPIs
    refreshKPIs();

    // 8. Stocker l'URL du form dans Config
    var configSheet = ss.getSheetByName("Config");
    configSheet.getRange("A20").setValue("URL du Form");
    configSheet.getRange("B20").setValue(form.getPublishedUrl());
    configSheet.getRange("A21").setValue("URL du Spreadsheet");
    configSheet.getRange("B21").setValue(ss.getUrl());

    Logger.log("Installation terminée. Form URL: " + form.getPublishedUrl());
    SpreadsheetApp.getUi().alert(
      "✅ CRM installé avec succès !\n\n" +
      "Form URL :\n" + form.getPublishedUrl() + "\n\n" +
      "Exécutez maintenant installTriggers() pour les déclencheurs automatiques."
    );
  } catch(e) {
    logErreur_("setupCRM", e.toString());
    throw e;
  }
}

// ============================================================
// CRÉATION DU GOOGLE FORM
// ============================================================

function creerForm_() {
  var form = FormApp.create(FORM_TITLE);
  form.setDescription("Enregistrement d'un nouveau prospect — Livraison Abidjan");
  form.setConfirmationMessage("✅ Prospect enregistré avec succès ! Un commercial vous contactera rapidement.");

  // Nom du prospect
  form.addTextItem()
    .setTitle("Nom du prospect")
    .setHelpText("Nom complet ou raison sociale")
    .setRequired(true);

  // Téléphone WhatsApp
  form.addTextItem()
    .setTitle("Téléphone WhatsApp")
    .setHelpText("Format : +225XXXXXXXXXX")
    .setRequired(true);

  // Type de client
  form.addMultipleChoiceItem()
    .setTitle("Type de client")
    .setChoiceValues(["Entreprise", "Boutique en ligne", "Particulier"])
    .setRequired(true);

  // Localisation / Quartier
  form.addListItem()
    .setTitle("Localisation / Quartier")
    .setChoiceValues(["Yopougon","Abobo","Cocody","Plateau","Marcory","Adjamé","Port-Bouët","Autre"])
    .setRequired(true);

  // Adresse exacte
  form.addTextItem()
    .setTitle("Adresse exacte")
    .setHelpText("Ex : Rue des Jardins face pharmacie, Cocody")
    .setRequired(false);

  // Latitude (GPS)
  form.addTextItem()
    .setTitle("Latitude")
    .setHelpText("Rempli depuis la page de géolocalisation terrain")
    .setRequired(false);

  // Longitude (GPS)
  form.addTextItem()
    .setTitle("Longitude")
    .setHelpText("Rempli depuis la page de géolocalisation terrain")
    .setRequired(false);

  // Précision GPS
  form.addTextItem()
    .setTitle("Précision GPS (mètres)")
    .setHelpText("Précision de la détection GPS en mètres")
    .setRequired(false);

  // Volume estimé
  form.addMultipleChoiceItem()
    .setTitle("Volume estimé colis/mois")
    .setChoiceValues(["1-10","11-50","51-100","100+"])
    .setRequired(true);

  // Source du prospect
  form.addMultipleChoiceItem()
    .setTitle("Source du prospect")
    .setChoiceValues(["Bouche-à-oreille","Réseaux sociaux","Marché Adjamé","Référence client","Autre"])
    .setRequired(true);

  // Notes additionnelles
  form.addParagraphTextItem()
    .setTitle("Notes additionnelles")
    .setHelpText("Informations complémentaires sur le prospect")
    .setRequired(false);

  return form;
}

// ============================================================
// CRÉATION DU SPREADSHEET ET DE SES FEUILLES
// ============================================================

function creerSpreadsheet_(form) {
  var ss = SpreadsheetApp.create(SHEET_TITLE);

  // Renommer la feuille par défaut
  var sheet1 = ss.getSheets()[0];
  sheet1.setName("Prospects");

  // Créer les autres feuilles
  ss.insertSheet("Pipeline");
  ss.insertSheet("Devis");
  ss.insertSheet("KPIs");
  ss.insertSheet("Config");
  ss.insertSheet("Logs");

  // Configurer la feuille Prospects
  configurerFeuilleProspects_(ss.getSheetByName("Prospects"));

  // Configurer la feuille Devis
  configurerFeuilleDevis_(ss.getSheetByName("Devis"));

  // Configurer la feuille Logs
  configurerFeuilleLogs_(ss.getSheetByName("Logs"));

  return ss;
}

function configurerFeuilleProspects_(sheet) {
  var headers = [
    "ID Prospect","Horodatage","Nom","Téléphone","Type Client","Quartier",
    "Adresse Exacte","📍 Maps","Zone Auto","Latitude","Longitude",
    "Précision (m)","Adresse Détectée","Volume/Mois","Source",
    "Score Priorité","Étape Actuelle","Date Étape","Délai (j)",
    "Statut Devis","Code Client","Tarif Mensuel FCFA","Notes","Commercial"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  sheet.setFrozenRows(1);

  // Largeurs colonnes (pixels approximatifs)
  var largeurs = [120,140,160,130,120,100,200,80,100,90,90,90,180,90,130,90,180,110,80,110,110,130,200,120];
  for (var i = 0; i < largeurs.length; i++) {
    sheet.setColumnWidth(i + 1, largeurs[i]);
  }
}

function configurerFeuilleDevis_(sheet) {
  var headers = [
    "ID Devis","ID Prospect","Nom Client","Zone","Volume",
    "Tarif Estimé FCFA","Date Envoi","Date Rappel J+3",
    "Date Expiration J+15","Statut"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  sheet.setFrozenRows(1);
}

function configurerFeuilleLogs_(sheet) {
  var headers = ["Horodatage","Fonction","Type","Message"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  sheet.setFrozenRows(1);
}

// ============================================================
// DÉCLENCHEUR SOUMISSION FORMULAIRE
// ============================================================

function onFormSubmit(e) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet  = ss.getSheetByName("Prospects");
    var now    = new Date();

    // Récupérer les réponses du formulaire
    var reponses = e.values; // [horodatage, nom, tel, type, quartier, adresse, lat, lng, precision, volume, source, notes]
    // Index : 0=timestamp, 1=nom, 2=tel, 3=type, 4=quartier, 5=adresse, 6=lat, 7=lng, 8=precision, 9=volume, 10=source, 11=notes

    var nom       = reponses[1]  || "";
    var tel       = reponses[2]  || "";
    var typeClient= reponses[3]  || "";
    var quartier  = reponses[4]  || "";
    var adresse   = reponses[5]  || "";
    var latStr    = reponses[6]  || "";
    var lngStr    = reponses[7]  || "";
    var precision = reponses[8]  || "";
    var volume    = reponses[9]  || "";
    var source    = reponses[10] || "";
    var notes     = reponses[11] || "";

    // Générer ID unique
    var lastRow  = sheet.getLastRow();
    var numSeq   = String(lastRow).padStart(3, "0");
    var dateStr  = Utilities.formatDate(now, "Africa/Abidjan", "yyyyMMdd");
    var idProspect = "PRO-" + dateStr + "-" + numSeq;

    // Calculer score priorité
    var score = calculerScore_(volume, typeClient, source);

    // Géolocalisation
    var lat       = parseFloat(latStr)  || 0;
    var lng       = parseFloat(lngStr)  || 0;
    var lienMaps  = "";
    var zoneAuto  = quartier; // fallback sur quartier sélectionné
    var adresseDetectee = "";

    if (lat !== 0 && lng !== 0) {
      lienMaps = '=HYPERLINK("https://maps.google.com/?q=' + lat + ',' + lng + '","📍 Voir")';
      zoneAuto = detecterZone_(lat, lng);
    }

    // Écrire la ligne dans Prospects
    var rowData = [
      idProspect,                                  // A - ID Prospect
      now,                                         // B - Horodatage
      nom,                                         // C - Nom
      tel,                                         // D - Téléphone
      typeClient,                                  // E - Type Client
      quartier,                                    // F - Quartier
      adresse,                                     // G - Adresse Exacte
      lienMaps,                                    // H - Maps
      zoneAuto,                                    // I - Zone Auto
      lat || "",                                   // J - Latitude
      lng || "",                                   // K - Longitude
      precision,                                   // L - Précision (m)
      adresseDetectee,                             // M - Adresse Détectée
      volume,                                      // N - Volume/Mois
      source,                                      // O - Source
      score,                                       // P - Score Priorité
      ETAPES[0],                                   // Q - Étape Actuelle
      now,                                         // R - Date Étape
      0,                                           // S - Délai (j)
      "",                                          // T - Statut Devis
      "",                                          // U - Code Client
      "",                                          // V - Tarif Mensuel FCFA
      notes,                                       // W - Notes
      COMMERCIAL_EMAIL                             // X - Commercial
    ];

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    // Formule Maps (si coordonnées)
    if (lienMaps) {
      sheet.getRange(newRow, 8).setFormula(
        '=HYPERLINK("https://maps.google.com/?q=' + lat + ',' + lng + '","📍 Voir")'
      );
    }

    // Couleur de fond selon score
    var couleur = score >= 70 ? COULEUR_VERT : (score >= 40 ? COULEUR_ORANGE : COULEUR_GRIS);
    sheet.getRange(newRow, 1, 1, rowData.length).setBackground(couleur);

    // Rafraîchir les KPIs et le Pipeline
    refreshKPIs();
    remplirPipeline_(SpreadsheetApp.getActiveSpreadsheet());

    // Envoyer email de notification
    envoyerEmailNotification_(nom, tel, typeClient, zoneAuto, volume, score, lat, lng);

    logAction_("onFormSubmit", "INFO", "Nouveau prospect créé : " + idProspect);

  } catch(err) {
    logErreur_("onFormSubmit", err.toString());
  }
}

// ============================================================
// CALCUL SCORE DE PRIORITÉ (0-100)
// ============================================================

function calculerScore_(volume, typeClient, source) {
  var score = 0;

  // Score volume
  if (volume === "100+")    score += 40;
  else if (volume === "51-100") score += 30;
  else if (volume === "11-50")  score += 15;
  else if (volume === "1-10")   score += 5;

  // Score type client
  if (typeClient === "Entreprise")        score += 30;
  else if (typeClient === "Boutique en ligne") score += 20;
  else if (typeClient === "Particulier")  score += 10;

  // Score source
  if (source === "Référence client")    score += 30;
  else if (source === "Bouche-à-oreille") score += 20;
  else if (source === "Réseaux sociaux")  score += 15;
  else if (source === "Marché Adjamé")    score += 10;
  else score += 5;

  return Math.min(score, 100);
}

// ============================================================
// DÉTECTION DE ZONE GPS (polygones Abidjan)
// ============================================================

function detecterZone_(lat, lng) {
  if (lat >= 5.316 && lat <= 5.322 && lng >= -4.023 && lng <= -4.012) return "Plateau";
  if (lat >= 5.330 && lat <= 5.380 && lng >= -3.990 && lng <= -3.940) return "Cocody";
  if (lat >= 5.290 && lat <= 5.315 && lng >= -3.990 && lng <= -3.960) return "Marcory";
  if (lat >= 5.320 && lat <= 5.380 && lng >= -4.090 && lng <= -4.040) return "Yopougon";
  if (lat >= 5.390 && lat <= 5.440 && lng >= -4.030 && lng <= -3.990) return "Abobo";
  if (lat >= 5.340 && lat <= 5.360 && lng >= -4.030 && lng <= -4.010) return "Adjamé";
  if (lat >= 5.240 && lat <= 5.290 && lng >= -3.990 && lng <= -3.940) return "Port-Bouët";
  return "Autre / Grand Abidjan";
}

// ============================================================
// MISE À JOUR D'ÉTAPE
// ============================================================

function updateStage(prospectId, nouvelleEtape) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Prospects");
    var data  = sheet.getDataRange().getValues();
    var now   = new Date();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === prospectId) {
        var dateCreation = new Date(data[i][1]);
        var delaiJours   = Math.round((now - dateCreation) / (1000 * 60 * 60 * 24));

        sheet.getRange(i + 1, 17).setValue(nouvelleEtape);  // Étape Actuelle
        sheet.getRange(i + 1, 18).setValue(now);            // Date Étape
        sheet.getRange(i + 1, 19).setValue(delaiJours);     // Délai (j)

        // Actions spéciales selon étape
        if (nouvelleEtape === ETAPES[2]) { // Devis Envoyé
          creerLigneDevis_(ss, data[i], i + 1, now);
        }

        if (nouvelleEtape === ETAPES[4]) { // Contrat Signé
          var codeClient = genererCodeClient_(ss);
          var tarif      = calculerTarif_(data[i][8], data[i][13]); // zone, volume
          sheet.getRange(i + 1, 21).setValue(codeClient);
          sheet.getRange(i + 1, 22).setValue(tarif);
        }

        logAction_("updateStage", "INFO", prospectId + " → " + nouvelleEtape);
        refreshKPIs();
        remplirPipeline_(ss);
        return;
      }
    }
    Logger.log("Prospect non trouvé : " + prospectId);
  } catch(e) {
    logErreur_("updateStage", e.toString());
  }
}

// ============================================================
// CRÉATION LIGNE DEVIS
// ============================================================

function creerLigneDevis_(ss, rowData, rowIndex, now) {
  var sheetDevis = ss.getSheetByName("Devis");
  var lastRow    = sheetDevis.getLastRow() + 1;
  var numDevis   = String(lastRow - 1).padStart(3, "0");
  var idDevis    = "DEV-" + Utilities.formatDate(now, "Africa/Abidjan", "yyyyMMdd") + "-" + numDevis;

  var rappelDate    = new Date(now); rappelDate.setDate(now.getDate() + 3);
  var expirationDate= new Date(now); expirationDate.setDate(now.getDate() + 15);

  var zone   = rowData[8];  // Zone Auto
  var volume = rowData[13]; // Volume/Mois
  var tarif  = calculerTarif_(zone, volume);

  sheetDevis.appendRow([
    idDevis,          // ID Devis
    rowData[0],       // ID Prospect
    rowData[2],       // Nom Client
    zone,             // Zone
    volume,           // Volume
    tarif,            // Tarif Estimé FCFA
    now,              // Date Envoi
    rappelDate,       // Rappel J+3
    expirationDate,   // Expiration J+15
    "En attente"      // Statut
  ]);

  // Mettre à jour statut devis dans Prospects
  var sheetProspects = ss.getSheetByName("Prospects");
  sheetProspects.getRange(rowIndex, 20).setValue("Devis envoyé — En attente");
}

// ============================================================
// CALCUL TARIF DEPUIS GRILLE CONFIG
// ============================================================

function calculerTarif_(zone, volume) {
  var tarifParColis = 2500; // défaut

  var estIntramuros = ["Plateau","Cocody","Marcory","Yopougon","Abobo","Adjamé","Port-Bouët"].indexOf(zone) >= 0;

  if (estIntramuros) {
    if (volume === "1-10")   tarifParColis = 2500;
    else if (volume === "11-50")  tarifParColis = 2000;
    else if (volume === "51-100") tarifParColis = 1500;
    else if (volume === "100+")   tarifParColis = 1200;
  } else {
    if (volume === "1-10")   tarifParColis = 3500;
    else if (volume === "11-50")  tarifParColis = 3000;
    else if (volume === "51-100") tarifParColis = 2500;
    else if (volume === "100+")   tarifParColis = 2000;
  }

  // Estimer volume moyen
  var volMin = {"1-10":5,"11-50":30,"51-100":75,"100+":120};
  var volMoyen = volMin[volume] || 30;

  // Remise volume > 50/mois
  var remise = volMoyen > 50 ? 0.90 : 1.0;

  return Math.round(tarifParColis * volMoyen * remise);
}

// ============================================================
// GÉNÉRATION CODE CLIENT
// ============================================================

function genererCodeClient_(ss) {
  var sheet    = ss.getSheetByName("Prospects");
  var data     = sheet.getDataRange().getValues();
  var annee    = new Date().getFullYear();
  var count    = 1;

  for (var i = 1; i < data.length; i++) {
    var code = data[i][20]; // Code Client
    if (code && code.toString().indexOf(CODE_CLIENT_PREFIX + "-" + annee) === 0) count++;
  }

  return CODE_CLIENT_PREFIX + "-" + annee + "-" + String(count).padStart(3, "0");
}

// ============================================================
// EMAIL DE NOTIFICATION
// ============================================================

function envoyerEmailNotification_(nom, tel, type, zone, volume, score, lat, lng) {
  try {
    var lienMaps = (lat && lng) ? "\n📍 Position : https://maps.google.com/?q=" + lat + "," + lng : "";
    var sujet  = "🆕 Nouveau prospect [Score " + score + "] — " + nom;
    var corps  =
      "Nouveau prospect enregistré dans le CRM :\n\n" +
      "👤 Nom       : " + nom     + "\n" +
      "📱 WhatsApp  : " + tel     + "\n" +
      "🏢 Type      : " + type    + "\n" +
      "📍 Zone      : " + zone    + "\n" +
      "📦 Volume    : " + volume  + " colis/mois\n" +
      "⭐ Score     : " + score   + "/100\n" +
      lienMaps + "\n\n" +
      "Accédez au CRM pour traiter ce prospect rapidement.";

    MailApp.sendEmail(COMMERCIAL_EMAIL, sujet, corps);
  } catch(e) {
    logErreur_("envoyerEmailNotification", e.toString());
  }
}

// ============================================================
// ACTIONS MENU — AVANCER ÉTAPE
// ============================================================

function avancerEtapeSuivante() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveRange().getRow();

  if (row < 2) { ui.alert("Sélectionnez une ligne prospect."); return; }

  var data   = sheet.getRange(row, 1, 1, 24).getValues()[0];
  var id     = data[0];
  var etape  = data[16];

  if (!id || !etape) { ui.alert("Ligne invalide. Sélectionnez une ligne prospect."); return; }

  var indexActuel = ETAPES.indexOf(etape);
  if (indexActuel === ETAPES.length - 1) {
    ui.alert("Ce prospect est déjà à la dernière étape : " + etape);
    return;
  }

  var prochaine = ETAPES[indexActuel + 1];
  var confirm   = ui.alert("Avancer « " + data[2] + " » vers :\n" + prochaine + " ?", ui.ButtonSet.YES_NO);
  if (confirm === ui.Button.YES) {
    updateStage(id, prochaine);
    ui.alert("✅ Étape mise à jour !");
  }
}

function marquerContratSigne() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveRange().getRow();
  if (row < 2) { ui.alert("Sélectionnez une ligne prospect."); return; }
  var id   = sheet.getRange(row, 1).getValue();
  var nom  = sheet.getRange(row, 3).getValue();
  var conf = ui.alert("Marquer « " + nom + " » comme Contrat Signé ?", ui.ButtonSet.YES_NO);
  if (conf === ui.Button.YES) {
    updateStage(id, ETAPES[4]);
    ui.alert("✅ Contrat signé enregistré !");
  }
}

function marquerPerdu() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveRange().getRow();
  if (row < 2) { ui.alert("Sélectionnez une ligne prospect."); return; }
  var id   = sheet.getRange(row, 1).getValue();
  var nom  = sheet.getRange(row, 3).getValue();
  var conf = ui.alert("Marquer « " + nom + " » comme Perdu ?", ui.ButtonSet.YES_NO);
  if (conf === ui.Button.YES) {
    updateStage(id, "❌ Perdu");
    ui.alert("Prospect marqué comme perdu.");
  }
}

function genererDevis() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveRange().getRow();
  if (row < 2) { ui.alert("Sélectionnez une ligne prospect."); return; }
  var id  = sheet.getRange(row, 1).getValue();
  var nom = sheet.getRange(row, 3).getValue();
  var conf= ui.alert("Générer un devis pour « " + nom + " » et passer en Étape 3 ?", ui.ButtonSet.YES_NO);
  if (conf === ui.Button.YES) {
    updateStage(id, ETAPES[2]);
    ui.alert("📄 Devis créé et enregistré dans la feuille Devis !");
  }
}

function voirPositionMaps() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveRange().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert("Sélectionnez une ligne prospect."); return; }
  var lat = sheet.getRange(row, 10).getValue();
  var lng = sheet.getRange(row, 11).getValue();
  if (!lat || !lng) {
    SpreadsheetApp.getUi().alert("Aucune coordonnée GPS pour ce prospect. Utilisez le Quartier sélectionné.");
    return;
  }
  var url = "https://maps.google.com/?q=" + lat + "," + lng;
  SpreadsheetApp.getUi().alert("Position GPS :\n" + url + "\n\nCopiez ce lien dans votre navigateur.");
}

// ============================================================
// RAFRAÎCHISSEMENT KPIs
// ============================================================

function refreshKPIs() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheetKPI      = ss.getSheetByName("KPIs");
    var sheetProspects= ss.getSheetByName("Prospects");
    if (!sheetKPI || !sheetProspects) return;

    sheetKPI.clear();
    var data = sheetProspects.getDataRange().getValues();
    if (data.length <= 1) return;

    // Compter par étape
    var compteurs = {};
    ETAPES.forEach(function(e) { compteurs[e] = 0; });
    var perdus = 0, totalScore = 0, nbActifs = 0;
    var valeurPipeline = 0;
    var parZone = {};
    var now = new Date();
    var prospectsSansActivite = [];

    for (var i = 1; i < data.length; i++) {
      var etape = data[i][16];
      var score = data[i][15] || 0;
      var zone  = data[i][8]  || "Inconnu";
      var tarif = data[i][21] || 0;
      var dateEtape = data[i][17] ? new Date(data[i][17]) : null;

      if (etape === "❌ Perdu") { perdus++; continue; }
      if (compteurs[etape] !== undefined) compteurs[etape]++;

      totalScore += score;
      nbActifs++;
      valeurPipeline += tarif;

      parZone[zone] = (parZone[zone] || 0) + 1;

      // Prospects sans activité depuis 7j+
      if (dateEtape) {
        var joursInactif = Math.round((now - dateEtape) / (1000 * 60 * 60 * 24));
        if (joursInactif >= 7) prospectsSansActivite.push([data[i][0], data[i][2], joursInactif]);
      }
    }

    var scoresMoyen = nbActifs > 0 ? Math.round(totalScore / nbActifs) : 0;
    var contratsSigns= compteurs[ETAPES[4]] || 0;
    var tauxConversion= nbActifs > 0 ? Math.round((contratsSigns / nbActifs) * 100) : 0;

    // En-tête KPIs
    sheetKPI.getRange("A1").setValue("📊 TABLEAU DE BORD CRM — " + Utilities.formatDate(now, "Africa/Abidjan", "dd/MM/yyyy HH:mm"));
    sheetKPI.getRange("A1").setFontWeight("bold").setBackground(COULEUR_HEADER).setFontColor("#ffffff");
    sheetKPI.getRange("A1:D1").merge();

    // Métriques principales
    var ligneMetriques = [
      ["Métrique","Valeur","Cible","Statut"],
      ["Total Prospects Actifs", nbActifs, "—",""],
      ["Taux Conversion → Contrat", tauxConversion + "%", "> 25%", tauxConversion >= 25 ? "✅" : "⚠️"],
      ["Score Moyen Prospects", scoresMoyen + "/100", "> 50", scoresMoyen >= 50 ? "✅" : "⚠️"],
      ["Valeur Pipeline Estimée (FCFA)", valeurPipeline, "—", ""],
      ["Prospects Perdus", perdus, "—", ""]
    ];
    sheetKPI.getRange(3, 1, ligneMetriques.length, 4).setValues(ligneMetriques);
    styleHeader_(sheetKPI, 4, 3);

    // Compteurs par étape
    sheetKPI.getRange("A10").setValue("📌 PROSPECTS PAR ÉTAPE");
    sheetKPI.getRange("A10:B10").merge().setFontWeight("bold").setBackground("#f4f4f4");
    var ligneEtapes = ETAPES.map(function(e) { return [e, compteurs[e]]; });
    sheetKPI.getRange(11, 1, ligneEtapes.length, 2).setValues(ligneEtapes);

    // Répartition géographique
    var ligneGeo = Object.keys(parZone).map(function(z) { return [z, parZone[z]]; });
    sheetKPI.getRange("A21").setValue("🗺️ RÉPARTITION GÉOGRAPHIQUE");
    sheetKPI.getRange("A21:B21").merge().setFontWeight("bold").setBackground("#f4f4f4");
    if (ligneGeo.length > 0) {
      sheetKPI.getRange(22, 1, ligneGeo.length, 2).setValues(ligneGeo);
    }

    // Prospects sans activité
    if (prospectsSansActivite.length > 0) {
      sheetKPI.getRange("A30").setValue("⚠️ PROSPECTS SANS ACTIVITÉ (7j+)");
      sheetKPI.getRange("A30:C30").merge().setFontWeight("bold").setBackground("#f4cccc");
      sheetKPI.getRange(31, 1, prospectsSansActivite.length, 3).setValues(prospectsSansActivite);
      sheetKPI.getRange(31, 1, prospectsSansActivite.length, 3).setBackground("#f4cccc");
    }

    sheetKPI.setFrozenRows(1);
    logAction_("refreshKPIs", "INFO", "KPIs mis à jour");

  } catch(e) {
    logErreur_("refreshKPIs", e.toString());
  }
}

// ============================================================
// VUE PIPELINE (KANBAN TEXTUEL)
// ============================================================

function remplirPipeline_(ss) {
  try {
    var sheetPipeline  = ss.getSheetByName("Pipeline");
    var sheetProspects = ss.getSheetByName("Prospects");
    if (!sheetPipeline || !sheetProspects) return;

    sheetPipeline.clear();

    var data = sheetProspects.getDataRange().getValues();
    var parEtape = {};
    ETAPES.forEach(function(e) { parEtape[e] = []; });

    for (var i = 1; i < data.length; i++) {
      var etape = data[i][16];
      if (parEtape[etape]) {
        parEtape[etape].push([data[i][2], data[i][15], data[i][8]]); // nom, score, zone
      }
    }

    // Écriture kanban horizontale
    sheetPipeline.getRange(1, 1, 1, ETAPES.length).setValues([ETAPES]);
    sheetPipeline.getRange(1, 1, 1, ETAPES.length)
      .setBackground(COULEUR_HEADER).setFontColor("#ffffff").setFontWeight("bold");

    // Compteurs
    var compteurs = ETAPES.map(function(e) { return [parEtape[e].length + " prospect(s)"]; });
    for (var c = 0; c < ETAPES.length; c++) {
      sheetPipeline.getRange(2, c + 1).setValue(compteurs[c][0]).setFontStyle("italic");
    }

    // Noms des prospects par colonne
    var maxProspects = Math.max.apply(null, ETAPES.map(function(e) { return parEtape[e].length; }));
    for (var col = 0; col < ETAPES.length; col++) {
      var prospects = parEtape[ETAPES[col]];
      for (var p = 0; p < prospects.length; p++) {
        var cell = sheetPipeline.getRange(3 + p, col + 1);
        cell.setValue(prospects[p][0] + "\n⭐" + prospects[p][1] + " | " + prospects[p][2]);
        cell.setBackground(prospects[p][1] >= 70 ? COULEUR_VERT : (prospects[p][1] >= 40 ? COULEUR_ORANGE : COULEUR_GRIS));
        cell.setWrap(true);
      }
      sheetPipeline.setColumnWidth(col + 1, 160);
    }

    sheetPipeline.setFrozenRows(2);
  } catch(e) {
    logErreur_("remplirPipeline", e.toString());
  }
}

// ============================================================
// FEUILLE CONFIG — GRILLE TARIFAIRE
// ============================================================

function remplirConfig_(ss) {
  var sheet = ss.getSheetByName("Config");
  sheet.clear();

  var titreHeader = [["⚙️ CONFIGURATION CRM — LIVRAISON ABIDJAN","","","",""]];
  sheet.getRange(1, 1, 1, 5).setValues(titreHeader).merge()
    .setBackground(COULEUR_HEADER).setFontColor("#ffffff").setFontWeight("bold");

  // Grille tarifaire
  var grille = [
    ["GRILLE TARIFAIRE (FCFA/colis)","1-10 colis","11-50 colis","51-100 colis","100+ colis"],
    ["Abidjan intramuros","2500","2000","1500","1200"],
    ["Grand Abidjan","3500","3000","2500","2000"],
    ["Intérieur du pays","Sur devis (min. 5000)","","",""]
  ];
  sheet.getRange(3, 1, grille.length, 5).setValues(grille);
  styleHeader_(sheet, 5, 3);

  // Remises
  var remises = [
    ["REMISES & CONDITIONS",""],
    ["Volume > 50 colis/mois","-10% automatique"],
    ["Renouvellement anticipé","-5% sur tarif mensuel"],
    ["Paiement mensuel","Virement ou Mobile Money (MTN/Moov/Orange)"]
  ];
  sheet.getRange(9, 1, remises.length, 2).setValues(remises);
  sheet.getRange(9, 1, 1, 2).merge().setFontWeight("bold").setBackground("#f4f4f4");

  // Infos CRM
  sheet.getRange(15, 1).setValue("INFORMATIONS CRM");
  sheet.getRange(15, 1, 1, 2).merge().setFontWeight("bold").setBackground("#f4f4f4");
  sheet.getRange(16, 1).setValue("Email Commercial");
  sheet.getRange(16, 2).setValue(COMMERCIAL_EMAIL);
  sheet.getRange(17, 1).setValue("Prefix Code Client");
  sheet.getRange(17, 2).setValue(CODE_CLIENT_PREFIX);
  sheet.getRange(18, 1).setValue("Date Installation");
  sheet.getRange(18, 2).setValue(new Date());

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 200);
  sheet.setFrozenRows(1);
}

// ============================================================
// RAPPELS AUTOMATIQUES
// ============================================================

function checkReminders() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Prospects");
    var sheetDevis = ss.getSheetByName("Devis");
    var now   = new Date();
    var rappels = [], relances = [], expires = [];

    // Vérifier prospects sans mise à jour 48h
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var etape     = data[i][16];
      var dateEtape = data[i][17] ? new Date(data[i][17]) : null;
      if (!dateEtape || etape === "❌ Perdu" || etape === ETAPES[7]) continue;
      var heuresInactif = (now - dateEtape) / (1000 * 60 * 60);
      if (heuresInactif >= 48) {
        rappels.push(data[i][2] + " (" + etape + ") — " + Math.round(heuresInactif / 24) + "j inactif");
      }
    }

    // Vérifier devis sans réponse J+3 et expirés J+15
    var dataDevis = sheetDevis.getDataRange().getValues();
    for (var j = 1; j < dataDevis.length; j++) {
      var statut    = dataDevis[j][9];
      var dateEnvoi = dataDevis[j][6] ? new Date(dataDevis[j][6]) : null;
      if (!dateEnvoi || statut !== "En attente") continue;
      var joursDevis = (now - dateEnvoi) / (1000 * 60 * 60 * 24);
      if (joursDevis >= 15) {
        expires.push(dataDevis[j][2] + " — Devis expiré (" + dataDevis[j][0] + ")");
        sheetDevis.getRange(j + 1, 10).setValue("Expiré");
      } else if (joursDevis >= 3) {
        relances.push(dataDevis[j][2] + " — Relance J+" + Math.floor(joursDevis) + " (" + dataDevis[j][0] + ")");
      }
    }

    // Envoyer rapport
    if (rappels.length + relances.length + expires.length > 0) {
      var corps = "📋 RAPPORT QUOTIDIEN CRM — " + Utilities.formatDate(now, "Africa/Abidjan", "dd/MM/yyyy") + "\n\n";
      if (rappels.length)  corps += "⏰ Prospects à relancer (" + rappels.length + ") :\n" + rappels.join("\n") + "\n\n";
      if (relances.length) corps += "📄 Devis sans réponse (" + relances.length + ") :\n" + relances.join("\n") + "\n\n";
      if (expires.length)  corps += "❌ Devis expirés (" + expires.length + ") :\n" + expires.join("\n");
      MailApp.sendEmail(COMMERCIAL_EMAIL, "📋 Rapport CRM — " + Utilities.formatDate(now, "Africa/Abidjan", "dd/MM/yyyy"), corps);
    }

    logAction_("checkReminders", "INFO", "Rappels envoyés : " + (rappels.length + relances.length + expires.length));
  } catch(e) {
    logErreur_("checkReminders", e.toString());
  }
}

// ============================================================
// INSTALLATION DES DÉCLENCHEURS
// ============================================================

function installTriggers() {
  try {
    // Supprimer les anciens triggers
    ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Trigger onFormSubmit
    ScriptApp.newTrigger("onFormSubmit")
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();

    // Trigger checkReminders à 8h00
    ScriptApp.newTrigger("checkReminders")
      .timeBased()
      .atHour(8)
      .everyDays(1)
      .create();

    // Trigger refreshKPIs à 7h55
    ScriptApp.newTrigger("refreshKPIs")
      .timeBased()
      .atHour(7)
      .nearMinute(55)
      .everyDays(1)
      .create();

    var triggers = ScriptApp.getProjectTriggers();
    var ids = triggers.map(function(t) { return t.getHandlerFunction() + " : " + t.getUniqueId(); });
    SpreadsheetApp.getUi().alert(
      "✅ Déclencheurs installés (" + triggers.length + ") :\n\n" + ids.join("\n")
    );
    logAction_("installTriggers", "INFO", triggers.length + " triggers installés");
  } catch(e) {
    logErreur_("installTriggers", e.toString());
  }
}

// ============================================================
// WEB APP — RÉCEPTION DES DONNÉES DEPUIS terrain.html
// ============================================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "ping";
  try {
    if (action === "ping")  return jsonResponse_({ status: "ok", message: "CRM Web App actif" });
    if (action === "list")  return jsonResponse_(listProspects_());
    if (action === "kpis")  return jsonResponse_(getKPIs_());
    return jsonResponse_({ status: "error", message: "Action inconnue : " + action });
  } catch(err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action || "create";
    if (action === "create")  return createProspect_(payload);
    if (action === "advance") return advanceStage_(payload);
    if (action === "note")    return addNote_(payload);
    return jsonResponse_({ status: "error", message: "Action inconnue : " + action });
  } catch(err) {
    Logger.log("[ERREUR doPost] " + err.toString());
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

// ---- Création prospect (ancienne logique doPost) ----
function createProspect_(payload) {
  var ss    = SpreadsheetApp.openById(getCRMSpreadsheetId_());
  var sheet = ss.getSheetByName("Prospects");
  var now   = new Date();

  var nom        = payload.nom       || "";
  var tel        = payload.tel       || "";
  var typeClient = payload.type      || "";
  var quartier   = payload.quartier  || "";
  var adresse    = payload.adresse   || "";
  var volume     = payload.volume    || "";
  var source     = payload.source    || "";
  var notes      = payload.notes     || "";
  var latStr     = payload.lat       || "";
  var lngStr     = payload.lng       || "";
  var precision  = payload.precision || "";

  var lastRow    = sheet.getLastRow();
  var numSeq     = String(lastRow).padStart(3, "0");
  var dateStr    = Utilities.formatDate(now, "Africa/Abidjan", "yyyyMMdd");
  var idProspect = "PRO-" + dateStr + "-" + numSeq;

  var score    = calculerScore_(volume, typeClient, source);
  var lat      = parseFloat(latStr) || 0;
  var lng      = parseFloat(lngStr) || 0;
  var lienMaps = "";
  var zoneAuto = quartier;

  if (lat !== 0 && lng !== 0) {
    lienMaps = "https://maps.google.com/?q=" + lat + "," + lng;
    zoneAuto = detecterZone_(lat, lng);
  }

  var rowData = [
    idProspect, now, nom, tel, typeClient, quartier,
    adresse, lienMaps ? '=HYPERLINK("' + lienMaps + '","📍 Voir")' : "",
    zoneAuto, lat || "", lng || "", precision, "",
    volume, source, score, ETAPES[0], now, 0,
    "", "", "", notes, COMMERCIAL_EMAIL
  ];

  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  if (lienMaps) sheet.getRange(newRow, 8).setFormula('=HYPERLINK("' + lienMaps + '","📍 Voir")');

  var couleur = score >= 70 ? COULEUR_VERT : (score >= 40 ? COULEUR_ORANGE : COULEUR_GRIS);
  sheet.getRange(newRow, 1, 1, rowData.length).setBackground(couleur);

  refreshKPIs();
  remplirPipeline_(ss);
  envoyerEmailNotification_(nom, tel, typeClient, zoneAuto, volume, score, lat, lng);
  logAction_("createProspect", "INFO", "Prospect créé via terrain.html : " + idProspect);

  return jsonResponse_({ status: "ok", id: idProspect, score: score, zone: zoneAuto });
}

// ---- Liste tous les prospects ----
function listProspects_() {
  var ss      = SpreadsheetApp.openById(getCRMSpreadsheetId_());
  var sheet   = ss.getSheetByName("Prospects");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: "ok", prospects: [] };

  var data = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
  var prospects = [];

  data.forEach(function(row) {
    if (!row[0]) return;
    var lat = row[9], lng = row[10];
    prospects.push({
      id:        row[0],
      date:      row[1] ? Utilities.formatDate(new Date(row[1]), "Africa/Abidjan", "dd/MM/yyyy") : "",
      nom:       row[2],
      tel:       String(row[3]),
      type:      row[4],
      quartier:  row[5],
      adresse:   row[6],
      zone:      row[8],
      lat:       lat,
      lng:       lng,
      volume:    row[13],
      source:    row[14],
      score:     row[15] || 0,
      etape:     row[16],
      dateEtape: row[17] ? Utilities.formatDate(new Date(row[17]), "Africa/Abidjan", "dd/MM/yyyy") : "",
      relances:  row[18] || 0,
      notes:     row[22] || "",
      commercial: row[23] || ""
    });
  });

  // Plus récents en premier
  prospects.reverse();
  return { status: "ok", prospects: prospects };
}

// ---- Avancer à l'étape suivante ----
function advanceStage_(payload) {
  var id = payload.id;
  if (!id) return jsonResponse_({ status: "error", message: "ID manquant" });

  var ss    = SpreadsheetApp.openById(getCRMSpreadsheetId_());
  var sheet = ss.getSheetByName("Prospects");
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 24).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== id) continue;

    var rowNum      = i + 2;
    var currentEtape = data[i][16];
    var currentIdx  = ETAPES.indexOf(currentEtape);

    if (currentIdx < 0 || currentIdx >= ETAPES.length - 1) {
      return jsonResponse_({ status: "error", message: "Étape déjà finale ou inconnue" });
    }

    var nextEtape = ETAPES[currentIdx + 1];
    sheet.getRange(rowNum, 17).setValue(nextEtape);
    sheet.getRange(rowNum, 18).setValue(new Date());
    sheet.getRange(rowNum, 19).setValue((data[i][18] || 0) + 1);

    var score   = data[i][15] || 0;
    var couleur = score >= 70 ? COULEUR_VERT : (score >= 40 ? COULEUR_ORANGE : COULEUR_GRIS);
    sheet.getRange(rowNum, 1, 1, 24).setBackground(couleur);

    remplirPipeline_(ss);
    logAction_("advanceStage", "INFO", id + " : " + currentEtape + " → " + nextEtape);

    return jsonResponse_({ status: "ok", etape: nextEtape, etapeIdx: currentIdx + 1 });
  }
  return jsonResponse_({ status: "error", message: "Prospect introuvable : " + id });
}

// ---- Ajouter une note ----
function addNote_(payload) {
  var id   = payload.id;
  var note = payload.note;
  if (!id || !note) return jsonResponse_({ status: "error", message: "ID ou note manquant" });

  var ss    = SpreadsheetApp.openById(getCRMSpreadsheetId_());
  var sheet = ss.getSheetByName("Prospects");
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 24).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== id) continue;
    var rowNum   = i + 2;
    var dateStr  = Utilities.formatDate(new Date(), "Africa/Abidjan", "dd/MM HH:mm");
    var existing = data[i][22] || "";
    var newNotes = existing ? existing + "\n[" + dateStr + "] " + note : "[" + dateStr + "] " + note;
    sheet.getRange(rowNum, 23).setValue(newNotes);
    logAction_("addNote", "INFO", id + " : note ajoutée");
    return jsonResponse_({ status: "ok", notes: newNotes });
  }
  return jsonResponse_({ status: "error", message: "Prospect introuvable : " + id });
}

// ---- KPIs dashboard ----
function getKPIs_() {
  var ss      = SpreadsheetApp.openById(getCRMSpreadsheetId_());
  var sheet   = ss.getSheetByName("Prospects");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: "ok", kpis: { total: 0, parEtape: {}, scoreMoyen: 0, topZones: [], rappels: [] } };

  var data       = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
  var parEtape   = {};
  ETAPES.forEach(function(e) { parEtape[e] = 0; });

  var totalScore = 0, countScore = 0;
  var zones      = {};
  var today      = new Date();
  var rappels    = [];

  data.forEach(function(row) {
    if (!row[0]) return;
    var etape = row[16];
    if (parEtape.hasOwnProperty(etape)) parEtape[etape]++;

    if (row[15]) { totalScore += row[15]; countScore++; }

    var zone = row[8];
    if (zone) zones[zone] = (zones[zone] || 0) + 1;

    var etapesFermes = ["Gagné", "Perdu", "Fidélisation"];
    var dateEtape = row[17];
    if (dateEtape && etapesFermes.indexOf(etape) < 0) {
      var daysSince = Math.floor((today - new Date(dateEtape)) / 86400000);
      if (daysSince >= 3) {
        rappels.push({ id: row[0], nom: row[2], tel: String(row[3]), etape: etape, jours: daysSince });
      }
    }
  });

  var topZones = [];
  for (var z in zones) topZones.push({ zone: z, count: zones[z] });
  topZones.sort(function(a, b) { return b.count - a.count; });
  rappels.sort(function(a, b) { return b.jours - a.jours; });

  return {
    status: "ok",
    kpis: {
      total:      countScore,
      parEtape:   parEtape,
      scoreMoyen: countScore > 0 ? Math.round(totalScore / countScore) : 0,
      topZones:   topZones.slice(0, 5),
      rappels:    rappels.slice(0, 10)
    }
  };
}

// ---- Utilitaire réponse JSON ----
function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Récupère l'ID du Spreadsheet CRM depuis la feuille Config
function getCRMSpreadsheetId_() {
  // Cherche parmi tous les fichiers Drive le spreadsheet CRM
  var files = DriveApp.getFilesByName(SHEET_TITLE);
  if (files.hasNext()) return files.next().getId();
  throw new Error("Spreadsheet CRM introuvable. Exécutez d'abord setupCRM().");
}

// ============================================================
// UTILITAIRES — STYLE ET LOGS
// ============================================================

function styleHeader_(sheet, nbCols, ligne) {
  ligne = ligne || 1;
  sheet.getRange(ligne, 1, 1, nbCols)
    .setBackground(COULEUR_HEADER)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function logAction_(fn, type, message) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss ? ss.getSheetByName("Logs") : null;
    if (!sheet) return;
    sheet.appendRow([new Date(), fn, type, message]);
  } catch(e) { Logger.log("Erreur log : " + e.toString()); }
}

function logErreur_(fn, message) {
  Logger.log("[ERREUR] " + fn + " : " + message);
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss ? ss.getSheetByName("Logs") : null;
    if (!sheet) return;
    sheet.appendRow([new Date(), fn, "ERREUR", message]);
    sheet.getRange(sheet.getLastRow(), 1, 1, 4).setBackground("#f4cccc");
  } catch(e) {}
}
