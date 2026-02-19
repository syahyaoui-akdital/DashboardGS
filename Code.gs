/**
 * =======================================================================
 * Fichier : Code.gs
 * Description : Backend complet sécurisé (Flash, Dépenses, Recouvrement, etc.)
 * =======================================================================
 */


// =======================================================================
// CONFIGURATION SÉCURITÉ & BASE DE DONNÉES
// =======================================================================

// ⚠️ REMPLACEZ PAR L'ID DE VOTRE NOUVEAU SHEET SÉCURITÉ
const SECURITY_DB_ID = "1otCm507BxaxcK2rxingHL3TOcjciIopgl6xruG0zg4o"; 

function getSecurityDb() {
  return SpreadsheetApp.openById(SECURITY_DB_ID);
}

// 1. RÉCUPÉRER LE CONTEXTE (Remplaçant de votre ancienne fonction)
// Cette fonction est utilisée par getDashboardData pour filtrer les données
// Remplacez la fonction getUserContext() par celle-ci :
// 1. RÉCUPÉRER LE CONTEXTE
function getUserContext() {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSecurityDb();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  let userConfig = null;
  
  // Recherche de l'utilisateur par Email dans le Sheet
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === email) {
      userConfig = { 
        role: data[i][3], 
        entity: data[i][4],
        poste: data[i][5] || 'Utilisateur',
        photo: data[i][6] || '' 
      };
      break; // Ce break est maintenant bien à l'intérieur de la boucle for
    }
  }

  if (!userConfig) {
    return { email: email, hasAccess: false, role: 'NONE', entity: 'NONE', poste: '', photo: '' };
  }

  return {
    email: email,
    hasAccess: true,
    role: userConfig.role,
    entity: userConfig.entity,
    poste: userConfig.poste,
    photo: userConfig.photo
  };
}

// 2. VÉRIFICATION LOGIN AVEC LOGS (Pour l'écran de connexion)
function verifyCustomLogin(username, password, clientIp) {
  const ss = getSecurityDb();
  const userSheet = ss.getSheetByName('Users');
  const logSheet = ss.getSheetByName('Logs');
  
  const activeEmail = Session.getActiveUser().getEmail().toLowerCase();
  username = String(username).trim().toLowerCase();
  password = String(password).trim();
  const data = userSheet.getDataRange().getValues();
  
  let userData = null;

  // Recherche utilisateur (Col A = Username)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username) {
      userData = { 
        pass: data[i][1], 
        email: String(data[i][2]).toLowerCase(), 
        role: data[i][3], 
        entity: data[i][4],
        poste: data[i][5] || 'Utilisateur',
        photo: data[i][6] || ''
      };
      break; // Ce break est également bien protégé dans sa boucle
    }
  }

  // Fonction Helper pour écrire dans l'onglet Logs
  const log = (status, details) => {
    logSheet.appendRow([new Date(), username, clientIp, "LOGIN", status, details]);
  };

  // Règle 1 : Cohérence Email Google vs Username
  let emailPrefix = activeEmail.split('@')[0];
  if (activeEmail && username !== emailPrefix) {
      log("FAIL", "Incohérence Compte Google (" + activeEmail + ")");
      return { success: false, message: "Sécurité : L'utilisateur '" + username + "' ne correspond pas au compte Google connecté." };
  }

  // Règle 2 : Utilisateur existe ?
  if (!userData) {
    log("FAIL", "Utilisateur inconnu");
    return { success: false, message: "Identifiant ou mot de passe incorrect." };
  }

  // Règle 3 : Mot de passe
  if (String(userData.pass) !== password) {
    log("FAIL", "Mot de passe incorrect");
    return { success: false, message: "Identifiant ou mot de passe incorrect." };
  }

  log("SUCCESS", "OK");
  return { 
    success: true, 
    role: userData.role, 
    entity: userData.entity, 
    username: username, 
    email: activeEmail,
    poste: userData.poste,
    photo: userData.photo
  };
}

// 3. NOUVEAU : CHANGER MOT DE PASSE
function changeUserPassword(username, oldPass, newPass, clientIp) {
  const ss = getSecurityDb();
  const userSheet = ss.getSheetByName('Users');
  const logSheet = ss.getSheetByName('Logs');
  
  const data = userSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username.toLowerCase()) {
      // Vérification ancien mot de passe
      if (String(data[i][1]) !== String(oldPass)) {
        logSheet.appendRow([new Date(), username, clientIp, "CHANGE_PASS", "FAIL", "Ancien MDP faux"]);
        return { success: false, message: "L'ancien mot de passe est incorrect." };
      }
      
      // Mise à jour (Colonne B = Password, index 2 pour getRange)
      userSheet.getRange(i + 1, 2).setValue(newPass);
      logSheet.appendRow([new Date(), username, clientIp, "CHANGE_PASS", "SUCCESS", "MDP Modifié"]);
      return { success: true, message: "Mot de passe modifié avec succès." };
    }
  }
  return { success: false, message: "Utilisateur introuvable." };
}

// 4. NOUVEAU : MOT DE PASSE OUBLIÉ (Envoi Email)
function resetUserPassword(email) {
  const ss = getSecurityDb();
  const userSheet = ss.getSheetByName('Users');
  const logSheet = ss.getSheetByName('Logs');
  
  const data = userSheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === email.trim().toLowerCase()) {
      // Génération mot de passe temporaire
      const newPass = Math.random().toString(36).slice(-8).toUpperCase();
      
      // Sauvegarde
      userSheet.getRange(i + 1, 2).setValue(newPass);
      
      try {
        MailApp.sendEmail(email, "AKDITAL Dashboard - Réinitialisation", 
          "Bonjour,\n\nVotre nouveau mot de passe temporaire est : " + newPass + 
          "\n\nVeuillez le changer dès votre connexion.\n\nCordialement.");
          
        logSheet.appendRow([new Date(), data[i][0], "SYSTEM", "RESET_PASS", "SUCCESS", "Email envoyé à " + email]);
        found = true;
      } catch(e) {
        return { success: false, message: "Erreur technique lors de l'envoi de l'email." };
      }
      break;
    }
  }
  
  // Message générique par sécurité
  return { success: true, message: "Si cet email existe dans la base, un nouveau mot de passe a été envoyé." };
}

// =======================================================================
// 1. CONFIGURATION ET MENU
// =======================================================================

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // 1. Appel de votre fonction existante pour récupérer le contexte utilisateur
  var userContext = getUserContext();
  
  // 2. Transmission des variables au template HTML (Index.html)
  template.userEmail = userContext.email;
  template.userName = userContext.email.split('@')[0].replace('.', ' '); // Déduit le prénom/nom
  
  // 3. Récupération du Poste (avec une valeur par défaut de sécurité)
  template.userPoste = userContext.poste ? userContext.poste : "Collaborateur Akdital";
  
  // 4. Récupération de la Photo (avec avatar dynamique par défaut si vide)
  template.userPhoto = userContext.photo ? userContext.photo : "https://ui-avatars.com/api/?name=" + template.userName + "&background=e0e7ff&color=005c97&bold=true";
  
  return template.evaluate()
      .setTitle('DASHBOARD - REGION GRAND SUD')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Administration')
    .addItem('🔄 Mettre à jour Flash (BDD)', 'clientNormalizeData')
    .addItem('📥 Import Détail Organismes', 'importDetailsOrganisme')
    .addToUi();
}

// =======================================================================
// 2. MODULE DASHBOARD (FLASH & BUDGET) - LOGIQUE SÉPARÉE
// =======================================================================

function getDashboardData() {
  const user = getUserContext();
  
  // 1. SÉCURITÉ
  if (!user.hasAccess) throw new Error("Accès refusé pour : " + user.email);

  // 2. RÉCUPÉRATION DES DONNÉES BRUTES
  const rawData = { 
    normalizedData: getNormalizedData(), // Données Flash (Tableau)
    contactData: getExtraData().contact,
    caData: getExtraData().ca,           // Données Budget (Objet JSON)
    capacity: getExtraData().capacity    // Données Capacité (Objet JSON)
  };

  // 3. LOGIQUE DE FILTRAGE
  if (user.entity !== 'ALL') {
    
    // -----------------------------------------------------------
    // A. FILTRAGE PARTIE FLASH (Strict)
    // -----------------------------------------------------------
    // HIA CIOA ne voit que "HIA CIOA" (pas de fusion ici)
    const header = rawData.normalizedData[0];
    const filteredRows = rawData.normalizedData.slice(1).filter(row => {
      const rowEnt = String(row[0]).trim();
      return rowEnt === user.entity; 
    });
    rawData.normalizedData = [header, ...filteredRows];

    // -----------------------------------------------------------
    // B. FILTRAGE PARTIE CA VS BUDGET (Exception HIA CIOA)
    // -----------------------------------------------------------
    const filteredCA = {};
    
    // On filtre les clés (noms des entités) avec votre logique spécifique
    const keptKeys = Object.keys(rawData.caData).filter(rowEnt => {
        rowEnt = String(rowEnt).trim(); // Nettoyage du nom de l'entité (Clé)

        // VOTRE CODE DEMANDÉ :
        if (user.entity === 'HIA CIOA') return rowEnt === 'HIA' || rowEnt === 'CIOA';
        return rowEnt === user.entity;
    });

    // Reconstruction de l'objet avec les clés conservées
    keptKeys.forEach(key => {
        filteredCA[key] = rawData.caData[key];
    });
    rawData.caData = filteredCA;

    // -----------------------------------------------------------
    // C. FILTRAGE PARTIE CAPACITÉ (Aucun Filtre)
    // -----------------------------------------------------------
    // On laisse rawData.capacity tel quel pour que les calculs fonctionnent.
  }

  return { ...rawData, userContext: user };
}

// ... (Fonctions utilitaires d'import standard inchangées ci-dessous) ...

function normalizeData() {
  var allData = [];
  var headers = ["Entité","Mois","Famille","Sous-famille","Élément","Valeur", "Année"];
  allData.push(headers);
  try {
    var ss2025 = SpreadsheetApp.getActiveSpreadsheet();
    var data2025 = processSheetData(ss2025, 'BDD', 2025);
    if (data2025.length > 0) allData = allData.concat(data2025);
  } catch(e) {}
  try {
    var ss2026 = SpreadsheetApp.openById("1rQGq8iLm5MmPEBS8YArTGJyON588Yy9dEyZgGKhsYNY");
    var data2026 = processSheetData(ss2026, 'BDD', 2026);
    if (data2026.length > 0) allData = allData.concat(data2026);
  } catch(e) {}
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var normSheet = ss.getSheetByName("NORMALIZED");
  if (!normSheet) normSheet = ss.insertSheet("NORMALIZED");
  normSheet.clear();
  if (allData.length > 1) normSheet.getRange(1, 1, allData.length, allData[0].length).setValues(allData);
  return allData;
}

function processSheetData(ss, sheetName, year) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if(data.length < 4) return [];
  var normalized = [];
  var moisNoms = ["JAN", "FEV", "MAR", "AVR", "MAI", "JUIN", "JUIL", "AOU", "SEP", "OCT", "NOV", "DEC"];
  var lastEntite = "";
  for (var i = 3; i < data.length; i++) {
    var entite = data[i][0];
    if (entite && String(entite).trim() !== '') { lastEntite = entite; } else { entite = lastEntite; }
    var element = data[i][2]; var famille = data[i][17]; var sousFamille = data[i][18];
    if (!entite || !element || String(element).trim() === '') continue;
    for (var j = 3; j <= 14; j++) {
      var headerVal = data[0][j];
      var moisFormatte = "DEC";
      if (headerVal instanceof Date) { moisFormatte = moisNoms[headerVal.getMonth()]; } 
      else {
        var s = String(headerVal).toUpperCase().trim();
        if (s.match(/\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4}/)) { var parts = s.split(/[\/\-\.]/); var mIdx = parseInt(parts[1], 10) - 1; if (mIdx >= 0 && mIdx < 12) moisFormatte = moisNoms[mIdx]; } 
        else { if (s.startsWith("JUIN")) moisFormatte = "JUIN"; else if (s.startsWith("JUIL")) moisFormatte = "JUIL"; else if (moisNoms.includes(s.substring(0, 3))) moisFormatte = s.substring(0, 3); else moisFormatte = s.substring(0, 3); }
      }
      var val = data[i][j]; if (val === "" || val == null) val = 0;
      normalized.push([entite, moisFormatte, famille, sousFamille, element, val, year]);
    }
  }
  return normalized;
}

function getExtraData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var extraData = { contact: {}, ca: {}, capacity: {} };
  var moisNoms = ["JAN", "FEV", "MAR", "AVR", "MAI", "JUIN", "JUIL", "AOU", "SEP", "OCT", "NOV", "DEC"];
  try {
    var lienFlash = ss.getSheetByName("LIEN FLASH");
    if (lienFlash && lienFlash.getLastRow() > 1) {
      lienFlash.getDataRange().getValues().slice(1).forEach(function(row) { 
        if (row[0]) extraData.contact[row[0]] = { respExploit: row[18] || '-', dirMedical: row[22] || '-' };
      });
    }
  } catch(e) {}
  var processExtraSheets = function(sourceSS, sheetNameCA, sheetNameCap, year) {
    var sheetCA = sourceSS.getSheetByName(sheetNameCA);
    if (sheetCA && sheetCA.getLastRow() > 1) {
      var dataCA = sheetCA.getDataRange().getValues();
      dataCA.slice(1).forEach(function(row) { 
        var ent = row[16]; 
        if (ent) {
          var mCA = {}; for (var k = 0; k < 12; k++) mCA[moisNoms[k]] = parseMoney(row[17 + k]); 
          if (!extraData.ca[ent]) extraData.ca[ent] = {};
          extraData.ca[ent][year] = { caMensuel: mCA, caAnnuel: row[29] || 0, budgetAnnuel: parseMoney(row[30]) };
        }
      });
    }
    var sheetCap = sourceSS.getSheetByName(sheetNameCap);
    if (sheetCap && sheetCap.getLastRow() > 2) {
      var dataCap = sheetCap.getDataRange().getValues();
      dataCap.slice(2).forEach(function(row) {
           if (row[0] && row[3]) {
             var key = String(row[0]).trim() + "_" + String(row[3]).trim();
             var c = {}; for (var m = 0; m < 12; m++) c[moisNoms[m]] = row[4 + m] || 0;
             if (!extraData.capacity[key]) extraData.capacity[key] = {};
             extraData.capacity[key][year] = c;
           }
      });
    }
  };
  processExtraSheets(ss, "RECAP", "CAPACITE", 2025);
  try {
    var ss2026 = SpreadsheetApp.openById("1rQGq8iLm5MmPEBS8YArTGJyON588Yy9dEyZgGKhsYNY");
    processExtraSheets(ss2026, "RECAP", "CAPACITE", 2026);
  } catch(e) {}
  return extraData;
}

function parseMoney(v) { return typeof v === 'number' ? v : parseFloat(String(v).replace(/[€,\s\u00A0]/g, '').replace(',','.')) || 0; }
function getNormalizedData() { const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName("NORMALIZED"); if (!sheet) return normalizeData(); return sheet.getDataRange().getValues(); }
function clientNormalizeData() { const user = getUserContext(); if (user.role !== 'ADMIN') return "⛔ ERREUR"; try { normalizeData(); return "Mise à jour Flash réussie !"; } catch (e) { return "Erreur"; } }

// =======================================================================
// 3. MODULE DÉPENSES CAISSES (SÉCURISÉ + EXCEPTION HIA CIOA)
// =======================================================================

function saveExpensesData(cleanData) {
  const user = getUserContext();
  if (user.role !== 'ADMIN') return "Droit insuffisant";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("DEPENSES");
  const headers = ["Entité", "Jour", "Mois", "Année", "Bénéficiaire", "Montant (MAD)", "Nature", "Catégorie", "Ordre"];
  
  if (!sheet) {
    sheet = ss.insertSheet("DEPENSES");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (currentHeaders.length < 9 || currentHeaders[1] !== "Jour") {
       sheet.clear();
       sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
       sheet.setFrozenRows(1);
    }
  }
  
  if (!cleanData || cleanData.length === 0) return "Aucune donnée.";
  
  const entityToReplace = String(cleanData[0][0]).trim();
  const lastRow = sheet.getLastRow();
  let existingData = [];
  if (lastRow > 1) {
    existingData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  }
  
  const keptData = existingData.filter(row => String(row[0]).trim() !== entityToReplace);
  const finalData = [...keptData, ...cleanData];
  
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, finalData[0].length).setValues(finalData);
  }
  
  const today = new Date();
  const dateStr = today.getDate().toString().padStart(2, '0') + '/' + (today.getMonth() + 1).toString().padStart(2, '0') + '/' + today.getFullYear();
  PropertiesService.getScriptProperties().setProperty('LAST_DEPENSE_UPDATE', dateStr);
  
  return "Succès : Dépenses mises à jour pour " + entityToReplace;
}

function getExpensesDashboardData() {
  const user = getUserContext();
  if (!user.hasAccess) throw new Error("Accès refusé.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DEPENSES");
  const lastUpdate = PropertiesService.getScriptProperties().getProperty('LAST_DEPENSE_UPDATE') || "--/--/----";
  
  if (!sheet) return { data: [], lastUpdate: lastUpdate }; 
  
  const rawData = sheet.getDataRange().getValues();

  // FILTRAGE (AVEC EXCEPTION HIA CIOA)
  if (user.entity !== 'ALL') {
    const header = rawData[0];
    const filteredRows = rawData.slice(1).filter(row => {
      const rowEnt = String(row[0]).trim();
      // EXCEPTION HIA CIOA
      if (user.entity === 'HIA CIOA') return rowEnt === 'HIA' || rowEnt === 'CIOA';
      return rowEnt === user.entity;
    });
    return { data: [header, ...filteredRows], lastUpdate: lastUpdate };
  }

  return { data: rawData, lastUpdate: lastUpdate };
}

// =======================================================================
// 4. MODULE RECOUVREMENT (SÉCURISÉ + EXCEPTION HIA CIOA)
// =======================================================================

function getDelaisData() {
  const SPREADSHEET_ID = '19dHq5rKIgldsyvxc6nI7hP-lfdCZP7zPRKvFE1ksijU';
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Délais Obj Rec");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    return data.slice(1);
  } catch (e) {
    Logger.log("Erreur lecture délais : " + e.message);
    return [];
  }
}

function saveRecouvFocusData(cleanData) {
  const user = getUserContext();
  if (user.role !== 'ADMIN') return "Droit insuffisant";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("RECOUV_FOCUS");
  const headers = [
    "Entité", "Famille", "Num. Dossier", "Patient", "N° Facture", 
    "J_Recep", "M_Recep", "A_Recep", "Expédié Par", 
    "J_Exp", "M_Exp", "A_Exp", "J_Ret", "M_Ret", "A_Ret", 
    "Motif Retour", "Organisme", "Matricule", "Montant"
  ];
  if (!sheet) {
    sheet = ss.insertSheet("RECOUV_FOCUS");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  
  if (!cleanData || cleanData.length === 0) return "Aucune donnée reçue.";
  const entitiesToReplace = [...new Set(cleanData.map(r => String(r[0]).trim()))];
  const lastRow = sheet.getLastRow();
  let existingData = [];
  if (lastRow > 1) {
    const lastCol = sheet.getLastColumn();
    const rawData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    existingData = rawData.map(row => {
      if (row.length > 19) return row.slice(0, 19);
      if (row.length < 19) {
        let newRow = [...row];
        while (newRow.length < 19) newRow.push("");
        return newRow;
      }
      return row;
    });
  }
  
  const keptData = existingData.filter(row => {
    const rowEnt = String(row[0]).trim();
    return !entitiesToReplace.includes(rowEnt);
  });
  const finalData = [...keptData, ...cleanData];
  
  sheet.clearContents(); 
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, 19).setValues(finalData);
  }
  
  const today = new Date();
  const dateStr = today.getDate().toString().padStart(2, '0') + '/' + (today.getMonth() + 1).toString().padStart(2, '0') + '/' + today.getFullYear() + ' à ' + today.getHours().toString().padStart(2, '0') + ':' + today.getMinutes().toString().padStart(2, '0');
  PropertiesService.getScriptProperties().setProperty('LAST_RECOUV_UPDATE', dateStr);

  return `Succès Recouvrement. Mises à jour : ${entitiesToReplace.join(', ')}`;
}

function getRecouvFocusData() {
  const user = getUserContext();
  if (!user.hasAccess) throw new Error("Accès refusé.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lastUpdate = "--/--/----";
  try {
      const prop = PropertiesService.getScriptProperties().getProperty('LAST_RECOUV_UPDATE');
      if (prop) lastUpdate = prop;
  } catch(e) {}

  try {
      const sheet = ss.getSheetByName("RECOUV_FOCUS");
      if (!sheet || sheet.getLastRow() < 2) return { data: [], lastUpdate: lastUpdate };
      
      const rawData = sheet.getDataRange().getValues();

      // FILTRAGE (AVEC EXCEPTION HIA CIOA)
      if (user.entity !== 'ALL') {
        const header = rawData[0];
        const filteredRows = rawData.slice(1).filter(row => {
          const rowEnt = String(row[0]).trim();
          // EXCEPTION HIA CIOA
          if (user.entity === 'HIA CIOA') return rowEnt === 'HIA' || rowEnt === 'CIOA';
          return rowEnt === user.entity;
        });
        return { data: [header, ...filteredRows], lastUpdate: lastUpdate };
      }

      return { data: rawData, lastUpdate: lastUpdate };
  } catch (e) {
      return { data: [], lastUpdate: lastUpdate };
  }
}

// =======================================================================
// 5. MODULE EXPEDITION (SÉCURISÉ + EXCEPTION HIA CIOA)
// =======================================================================

function saveExpeditionData(cleanData) {
  const user = getUserContext();
  if (user.role !== 'ADMIN') return "Droit insuffisant";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("EXPEDITION_FOCUS");
  const headers = [
    "Entité", "Num. Dossier", "Patient", "Motif", 
    "J_Sortie", "M_Sortie", "A_Sortie", 
    "Organisme", "N° Facture", "Nature",
    "J_PEC", "M_PEC", "A_PEC", 
    "Montant PEC",
    "J_Entree", "M_Entree", "A_Entree",
    "Responsable PEC"
  ];
  if (!sheet) {
    sheet = ss.insertSheet("EXPEDITION_FOCUS");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  
  if (!cleanData || cleanData.length === 0) return "Aucune donnée.";
  
  const entitiesToReplace = [...new Set(cleanData.map(r => String(r[0]).trim()))];
  const lastRow = sheet.getLastRow();
  
  let existingData = [];
  if (lastRow > 1) {
     const lastCol = sheet.getLastColumn();
     const readWidth = lastCol < headers.length ? headers.length : lastCol;
     if (readWidth > 0) {
       const rawData = sheet.getRange(2, 1, lastRow - 1, readWidth).getValues();
       existingData = rawData.map(row => {
          let newRow = row.slice(0, headers.length);
          while(newRow.length < headers.length) newRow.push("");
          return newRow;
       });
     }
  }
  
  const keptData = existingData.filter(row => !entitiesToReplace.includes(String(row[0]).trim()));
  const finalData = [...keptData, ...cleanData];
  
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, headers.length).setValues(finalData);
  }
  
  const today = new Date();
  PropertiesService.getScriptProperties().setProperty('LAST_EXPEDITION_UPDATE', today.toLocaleString());
  
  return `Succès Expédition. ${finalData.length} dossiers traités.`;
}

function getExpeditionData() {
  const user = getUserContext();
  if (!user.hasAccess) throw new Error("Accès refusé.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lastUpdate = "--/--/----";
  try { lastUpdate = PropertiesService.getScriptProperties().getProperty('LAST_EXPEDITION_UPDATE') || lastUpdate;
  } catch(e) {}

  const sheet = ss.getSheetByName("EXPEDITION_FOCUS");
  if (!sheet || sheet.getLastRow() < 2) return { data: [], lastUpdate: lastUpdate };
  
  const rawData = sheet.getDataRange().getValues();

  // FILTRAGE (AVEC EXCEPTION HIA CIOA)
  if (user.entity !== 'ALL') {
    const header = rawData[0];
    const filteredRows = rawData.slice(1).filter(row => {
      const rowEnt = String(row[0]).trim();
      // EXCEPTION HIA CIOA
      if (user.entity === 'HIA CIOA') return rowEnt === 'HIA' || rowEnt === 'CIOA';
      return rowEnt === user.entity;
    });
    return { data: [header, ...filteredRows], lastUpdate: lastUpdate };
  }

  return { data: rawData, lastUpdate: lastUpdate };
}

// =======================================================================
// 6. MODULE SUIVI DES RETOURS (SÉCURISÉ + EXCEPTION HIA CIOA)
// =======================================================================

function getSuiviRetoursData() {
  const user = getUserContext();
  if (!user.hasAccess) throw new Error("Accès refusé.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const COL_CONFIG = {
    "HIA":     { patient: "B", reponse: "C", facture: "D", organisme: "G", montant: "I", observation: "J", motif: "K" },
    "CIOA":    { patient: "B", reponse: "C", facture: "D", organisme: "H", montant: "J", observation: "K", motif: "L" },
    "CIT":     { patient: "B", reponse: "C", facture: "D", organisme: "H", montant: "K", observation: "L", motif: "M" },
    "CID":     { patient: "B", reponse: "C", facture: "D", organisme: "H", montant: "K", observation: "L", motif: "M" },
    "PIL":     { patient: "B", reponse: "C", facture: "D", organisme: "G", montant: "J", observation: "K", motif: "L" },
    "HPG":     { patient: "B", reponse: "C", facture: "D", organisme: "G", montant: "J", observation: "K", motif: "L" },
    "ALHIKMA": { patient: "B", reponse: "C", facture: "D", organisme: "G", montant: "J", observation: "K", motif: "L" }
  };
  const sources = [
    { entite: "HIA", id: "1YR53J83Z9RMNrDp8ZPbLyRSpo3yKq58qecRSQAvja3s", sheetName: "HIA" },
    { entite: "CIOA", id: "1YR53J83Z9RMNrDp8ZPbLyRSpo3yKq58qecRSQAvja3s", sheetName: "CIOA" },
    { entite: "CIT", id: "1ToPy2i7ndUm7qnJLBtD4dac91EsBG4FxPAkfAy_3Hmg", sheetName: "CIT" },
    { entite: "CID", id: "1ujygPJyOflsYNMRa1QgtjftWnOCHGOrdjKD_CdkOX78", sheetName: "CID" },
    { entite: "PIL", id: "11WT_Bn96P4HGI3z4rN6AIN2GqIePfMXkj5IxvUlMbhk", sheetName: "PIL" },
    { entite: "HPG", id: "1Y0YuO0aLbh9GmpFCbU2AEy9M3DYBiOtnU-mj5U_CQgU", sheetName: "HPG" },
    { entite: "ALHIKMA", id: "1KYD12mBFiHanr_rPJG5CvI9L0f3O2tqiPE3EL1Hpwr4", sheetName: "HIKMA" }
  ];
  let compiledData = [];

  sources.forEach(source => {
    // FILTRAGE SOURCE AVANT IMPORT (AVEC EXCEPTION HIA CIOA)
    if (user.entity !== 'ALL') {
        let isAllowed = (source.entite === user.entity);
        if (user.entity === 'HIA CIOA' && (source.entite === 'HIA' || source.entite === 'CIOA')) isAllowed = true;
        if (!isAllowed) return; 
    }

    try {
      const extSs = SpreadsheetApp.openById(source.id);
      const sheet = extSs.getSheetByName(source.sheetName);
      
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 5) { 
          const conf = COL_CONFIG[source.entite];
          if (conf) {
            const dataValues = sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).getValues();
            
            const idxPatient = letterToColumn(conf.patient) - 1;
            const idxReponse = letterToColumn(conf.reponse) - 1;
            const idxFacture = letterToColumn(conf.facture) - 1;
            const idxOrganisme = letterToColumn(conf.organisme) - 1;
            const idxMontant = letterToColumn(conf.montant) - 1;
            const idxObs = letterToColumn(conf.observation) - 1;
            const idxMotif = letterToColumn(conf.motif) - 1;

            dataValues.forEach(row => {
              const valPatient = row[idxPatient];
              const valFacture = row[idxFacture];

              if ((valPatient && String(valPatient).trim() !== "") || (valFacture && String(valFacture).trim() !== "")) {
                const motifOriginal = String(row[idxMotif] || "").trim();
                compiledData.push({
                  entite: source.entite,
                  patient: String(valPatient || "").trim(),
                  reponse: String(row[idxReponse] || "").trim(),
                  facture: String(valFacture || "").trim(),
                  organisme: String(row[idxOrganisme] || "").trim(),
                  montant: cleanImportAmount(row[idxMontant]),
                  observation: String(row[idxObs] || "").trim(),
                  motif: motifOriginal,
                  motifStd: standardizeMotif(motifOriginal)
                });
              }
            });
          }
        }
      }
    } catch (e) {
      Logger.log("Erreur import " + source.entite + " : " + e.message);
    }
  });

  return compiledData;
}

function letterToColumn(letter) {
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function cleanImportAmount(val) {
  if (val == null || val === "") return 0;
  if (typeof val === 'number') return val;
  let s = String(val).replace(/[^0-9.,-]/g, '');
  s = s.replace(',', '.');
  const parsed = parseFloat(s);
  return isNaN(parsed) ? 0 : parsed;
}

function standardizeMotif(rawVal) {
  if (!rawVal) return "";
  const s = String(rawVal).trim();
  const upper = s.toUpperCase();

  const RULES = [
    { label: "Surfacturation", keywords: ["SURFACTURATION", "SUFACTURATION"] },
    { label: "Non Conforme PEC", keywords: ["NON CONFORME A LA PEC", "NON CONFORME A L'ACCORD", "NON CONFORME A LA PRISE EN CHARGE", "ACTE NON COUVERT", "CODE DE L'ACTE", "PAS FAIT OBJET D UN ACCORD", "PAS CONFORME A LA PEC"] },
    { label: "Hors Délai / Validité", keywords: ["HORS DELAIS", "HORS DÉLAIS", "VALIDITE", "DATE DU DECES", "DATE DE COUVERTURE", "POSTERIEURS A LA DATE", "DELA DE LA DATE", "PRESCRIPTION EST ERRONEE"] },
    { label: "Problème Facture / ICE", keywords: ["FACTURE ERRONEE", "FACTURE ERRONÉE", "MANQUE FACTURE", "ORIGINAL DE LA FACTURE", "NOM JURIDIQUE", "ICE", "IDENTIFICATION FISCAL", "TICKET MODERATEUR", "PART CNOPS"] },
    { label: "Prescription Manquante", keywords: ["MANQUE PRESCRIPTION", "MQ PRESCRIPTION", "MANQUE ORDONNANCE", "MQ ORDONNANCE", "ABSENCE DE LA PRESCRIPTION"] },
    { label: "Détail Médicaments / Vignette", keywords: ["DETAIL DES MEDICAMENTS", "DETAIL MEDICAMENTS", "DECOMPTE PHARMACIE", "DECOMPTE DES MEDICAMENTS", "VIGNETTE", "CODE A BARRE", "PHARMACIE"] },
    { label: "Compte Rendu Manquant", keywords: ["COMPTE RENDU", "CR OPERATOIRE", "CR HOSPITALISATION", "CR ANAPATH", "CR RADIO", "RAPPORT"] },
    { label: "CD / Imagerie", keywords: ["CD DE LA", "CLICHE", "ECHO"] },
    { label: "Signature / Cachet / INPE", keywords: ["SIGNATURE", "CACHET", "INPE", "IDENTIFICATION DU MEDECIN", "IDENTIFICATION DU MEDCIN"] },
    { label: "CIN / Identité", keywords: ["CIN", "CARTE NATIONALE", "CARTE D'IDENTITE", "NOM DU BENEFICIAIRE", "NOM ADHERENT", "LIEN AVEC L'ASSURE", "ASSUREE ERRONE", "ADHERENT SUR FACTURE"] },
    { label: "Accident / PV", keywords: ["ACCIDENT", "PROCES VERBAL", "PV"] },
    { label: "Droits Fermés / Annulé", keywords: ["FERMETURE DROIT", "DROIT FERME", "ANNULEE", "ANNULÉE"] },
    { label: "DMI / Prothèse", keywords: ["DISPOSITIF MEDICAL", "PROTHESE", "STENT", "DMI"] },
    { label: "Problème Date (Divers)", keywords: ["DATE", "DISCORDANCE", "RECTIFICATION DATE"] }
  ];
  for (let i = 0; i < RULES.length; i++) {
    if (RULES[i].keywords.some(k => upper.includes(k))) {
      return RULES[i].label;
    }
  }
  return "Autre";
}

// =======================================================================
// 7. MODULE OBJECTIF VS RÉALISATION (SÉCURISÉ + EXCEPTION HIA CIOA)
// =======================================================================

function importObjRealData() {
  const user = getUserContext();
  if (user.role !== 'ADMIN') return "Droit insuffisant";

  const SOURCE_ID = '19dHq5rKIgldsyvxc6nI7hP-lfdCZP7zPRKvFE1ksijU';
  const SHEET_NAME = 'RECOUVREMENT & EXPEDITION';
  const TARGET_SHEET_NAME = 'OBJ_REAL_DATA';
  try {
    const sourceSS = SpreadsheetApp.openById(SOURCE_ID);
    const sourceSheet = sourceSS.getSheetByName(SHEET_NAME);
    if (!sourceSheet) return "Erreur : La feuille 'RECOUVREMENT & EXPEDITION' est introuvable.";
    const data = sourceSheet.getDataRange().getValues();
    if (data.length < 2) return "Aucune donnée à importer.";
    
    const processedData = [];
    processedData.push(["Famille", "Jour", "Mois", "Année", "Entité", "Objectif", "Réalisation"]); 

    const moisNoms = ["JAN", "FEV", "MAR", "AVR", "MAI", "JUIN", "JUIL", "AOU", "SEP", "OCT", "NOV", "DEC"];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const famille = row[1]; 
      const dateRaw = row[2];
      const entite = row[4];
      const obj = row[5];
      const real = row[6];

      if (!famille && !entite) continue;
      let j = "", m = "", a = "";
      if (dateRaw instanceof Date) {
        j = dateRaw.getDate();
        m = moisNoms[dateRaw.getMonth()];
        a = dateRaw.getFullYear();
      } else if (String(dateRaw).trim() !== "") {
         try {
           const parts = String(dateRaw).split('/');
           if(parts.length === 3) {
             j = parseInt(parts[0], 10);
             let mIdx = parseInt(parts[1], 10) - 1;
             m = moisNoms[mIdx] || parts[1];
             a = parseInt(parts[2], 10);
           }
         } catch(e) {}
      }
      processedData.push([famille, j, m, a, entite, obj, real]);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
    } else {
      targetSheet.clear();
    }
    
    if (processedData.length > 0) {
      targetSheet.getRange(1, 1, processedData.length, processedData[0].length).setValues(processedData);
    }

    return "Succès : " + (processedData.length - 1) + " lignes importées.";
  } catch (e) {
    return "Erreur lors de l'importation : " + e.message;
  }
}

function getObjRealDashboardData() {
  const user = getUserContext();
  if (!user.hasAccess) throw new Error("Accès refusé.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Assurez-vous que le nom est EXACTEMENT le même que sur l'onglet (espaces compris)
  const sheet = ss.getSheetByName("OBJ_REAL_DATA");
  
  if (!sheet) return [];
  
  // Récupère tout : Headers + Données
  const rawData = sheet.getDataRange().getValues();
  
  // Petit filtre de sécurité pour l'utilisateur
  if (user.entity !== 'ALL') {
    const header = rawData[0];
    const filteredRows = rawData.slice(1).filter(row => {
      // Dans votre image, l'entité est à la colonne E (index 4)
      const rowEnt = String(row[4]).trim(); 
      if (user.entity === 'HIA CIOA') return rowEnt === 'HIA' || rowEnt === 'CIOA';
      return rowEnt === user.entity;
    });
    return [header, ...filteredRows];
  }

  return rawData;
}

// =======================================================================
// 8. MODULE HONORAIRES MEDECINS (SÉCURISÉ + EXCEPTION HIA CIOA)
// =======================================================================

function clearHonorairesFile(entity, year) {
  const user = getUserContext();
  if (user.role !== 'ADMIN') return "Droit insuffisant";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let indexSheet = ss.getSheetByName("DB_INDEX");
  if (!indexSheet) return "Index introuvable";

  const indexData = indexSheet.getDataRange().getValues();
  let targetId = null;
  for (let i = 1; i < indexData.length; i++) {
    if (String(indexData[i][0]) === String(entity) && String(indexData[i][1]) === String(year)) {
      targetId = indexData[i][2];
      break;
    }
  }

  if (targetId) {
    try {
      const targetSS = SpreadsheetApp.openById(targetId);
      const targetSheet = targetSS.getSheets()[0];
      if (targetSheet.getLastRow() > 1) {
        targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
      }
      return "CLEARED";
    } catch (e) {
      return "Erreur clear: " + e.message;
    }
  }
  return "NOT_FOUND";
}

function saveHonorairesBatch(data, entity, year) {
  const user = getUserContext();
  if (user.role !== 'ADMIN') return "Droit insuffisant";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let indexSheet = ss.getSheetByName("DB_INDEX");
  if (!indexSheet) {
    indexSheet = ss.insertSheet("DB_INDEX");
    indexSheet.appendRow(["ENTITE", "ANNEE", "SPREADSHEET_ID", "FILENAME", "LAST_UPDATE"]);
    indexSheet.getRange(1,1,1,5).setFontWeight("bold");
  }

  const indexData = indexSheet.getDataRange().getValues();
  let targetId = null;
  let rowIndex = -1;

  for (let i = 1; i < indexData.length; i++) {
    if (String(indexData[i][0]) === String(entity) && String(indexData[i][1]) === String(year)) {
      targetId = indexData[i][2];
      rowIndex = i + 1;
      break;
    }
  }

  let targetSS;
  if (targetId) {
    try {
      targetSS = SpreadsheetApp.openById(targetId);
    } catch(e) {
      return "Erreur : Impossible d'ouvrir le fichier d'archive pour " + entity + " " + year;
    }
  } else {
    const fileName = "DB_HONO_" + entity + "_" + year;
    targetSS = SpreadsheetApp.create(fileName);
    const headers = [
      "Entité", "J_Sortie", "M_Sortie", "A_Sortie", "Famille", "Bénéficiaire", "Spécialité", "Solde",
      "Type Paie", "Organisme", "Num PEC", "Dossier", "Réglé", "J_Paie", "M_Paie", "A_Paie",
      "Type Paiement", "Montant Brut", "Montant Net", "Montant Retenu", "Patente", "J_Envoi", "M_Envoi", "A_Envoi",
      "Type Organisme", "Eligible", "Mode Paie"
    ];
    targetSS.getSheets()[0].setName("DATA");
    targetSS.getSheets()[0].appendRow(headers);
    indexSheet.appendRow([entity, year, targetSS.getId(), fileName, new Date()]);
  }

  const targetSheet = targetSS.getSheets()[0];
  if (data.length > 0) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
  }

  if(rowIndex > -1) {
    indexSheet.getRange(rowIndex, 5).setValue(new Date());
  }

  return "OK";
}

function getHonorairesData() {
  const user = getUserContext();
  if (!user.hasAccess) throw new Error("Accès refusé.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName("DB_INDEX");
  let lastUpdate = "--/--/----";
  
  if (!indexSheet || indexSheet.getLastRow() < 2) {
    return { data: [], lastUpdate: "Aucune archive" };
  }

  const sources = indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 5).getValues();
  let aggregData = [];
  const dummyHeader = [
      "Entité", "J_Sortie", "M_Sortie", "A_Sortie", "Famille", "Bénéficiaire", "Spécialité", "Solde",
      "Type Paie", "Organisme", "Num PEC", "Dossier", "Réglé", "J_Paie", "M_Paie", "A_Paie",
      "Type Paiement", "Montant Brut", "Montant Net", "Montant Retenu", "Patente", "J_Envoi", "M_Envoi", "A_Envoi",
      "Type Organisme", "Eligible", "Mode Paie"
  ];
  aggregData.push(dummyHeader);

  for (let i = 0; i < sources.length; i++) {
    const entityFile = String(sources[i][0]).trim();
    
    // FILTRAGE FICHIERS SOURCES AVEC EXCEPTION HIA CIOA
    if (user.entity !== 'ALL') {
        let isAllowed = (user.entity === entityFile);
        // EXCEPTION HIA CIOA
        if (user.entity === 'HIA CIOA' && (entityFile === 'HIA' || entityFile === 'CIOA')) isAllowed = true;
        if (!isAllowed) continue;
    }

    const id = sources[i][2];
    try {
      const extSS = SpreadsheetApp.openById(id);
      const extSheet = extSS.getSheets()[0];
      const lr = extSheet.getLastRow();
      if (lr > 1) {
        const vals = extSheet.getRange(2, 1, lr - 1, 27).getValues();
        aggregData = aggregData.concat(vals);
      }
    } catch (e) {
      console.error("Erreur lecture ID " + id);
    }
  }
  
  try {
    const dates = sources.map(r => r[4]).filter(d => d instanceof Date);
    if(dates.length > 0) {
      const maxDate = new Date(Math.max.apply(null, dates));
      lastUpdate = maxDate.toLocaleString();
    }
  } catch(e){}

  return { data: aggregData, lastUpdate: lastUpdate };
}

// =======================================================================
// 9. MODULE DOSSIERS NON SOLDÉS (SÉCURISÉ + EXCEPTION HIA CIOA)
// =======================================================================

function saveNonSoldesData(cleanData) {
  const user = getUserContext();
  if (user.role !== 'ADMIN') return "Droit insuffisant";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("NON_SOLDE_FOCUS");
  
  // NOUVELLES COLONNES DEMANDÉES
  const headers = [
    "Entité", "Num. Dossier", "ID Séjour", "Patient", "Service", 
    "J_Sortie", "M_Sortie", "A_Sortie", "Organisme", "N° facture", 
    "Total patient", "Total Paiement", "Reste", 
    "Chèques Impayés", "Chèques Caution", "État Dossier", "Note Dossier"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("NON_SOLDE_FOCUS");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  
  if (!cleanData || cleanData.length === 0) return "Aucune donnée.";
  const entitiesToReplace = [...new Set(cleanData.map(r => String(r[0]).trim()))];
  const lastRow = sheet.getLastRow();
  
  let existingData = [];
  if (lastRow > 1) {
     const lastCol = sheet.getLastColumn();
     const readWidth = lastCol < headers.length ? headers.length : lastCol;
     const rawData = sheet.getRange(2, 1, lastRow - 1, readWidth).getValues();
     existingData = rawData.map(row => {
        let newRow = row.slice(0, headers.length);
        while(newRow.length < headers.length) newRow.push("");
        return newRow;
     });
  }
  
  const keptData = existingData.filter(row => !entitiesToReplace.includes(String(row[0]).trim()));
  const finalData = [...keptData, ...cleanData];
  
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, headers.length).setValues(finalData);
  }
  
  const today = new Date();
  PropertiesService.getScriptProperties().setProperty('LAST_NONSOLDE_UPDATE', today.toLocaleString());
  return `Succès : Dossiers non soldés mis à jour pour ${entitiesToReplace.join(', ')}`;
}

function getNonSoldesData() {
  const user = getUserContext();
  if (!user.hasAccess) throw new Error("Accès refusé.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lastUpdate = "--/--/----";
  try { 
    lastUpdate = PropertiesService.getScriptProperties().getProperty('LAST_NONSOLDE_UPDATE') || lastUpdate;
  } catch(e) {}

  const sheet = ss.getSheetByName("NON_SOLDE_FOCUS");
  if (!sheet || sheet.getLastRow() < 2) return { data: [], lastUpdate: lastUpdate };
  
  const rawData = sheet.getDataRange().getValues();

  // Filtrage entité (AVEC EXCEPTION HIA CIOA)
  if (user.entity !== 'ALL') {
    const header = rawData[0];
    const filteredRows = rawData.slice(1).filter(row => {
      // Index 0 = Entité dans NON_SOLDE_FOCUS
      const rowEnt = String(row[0]).trim();
      // EXCEPTION HIA CIOA
      if (user.entity === 'HIA CIOA') return rowEnt === 'HIA' || rowEnt === 'CIOA';
      return rowEnt === user.entity;
    });
    return { data: [header, ...filteredRows], lastUpdate: lastUpdate };
  }

  return { data: rawData, lastUpdate: lastUpdate };
}

// ... (Votre code existant jusqu'à la fin) ...

// =======================================================================
// 10. MODULE RH
// =======================================================================

function testConnexionRH() {
  const id = PropertiesService.getScriptProperties().getProperty("RH_SHEET_ID");
  if (!id) throw new Error("RH_SHEET_ID manquant");

  const ss = SpreadsheetApp.openById(id);
  Logger.log("Fichier RH : " + ss.getName());

  const sheets = ss.getSheets().map(s => s.getName());
  Logger.log("Onglets RH :\n- " + sheets.join("\n- "));
}
/* =========================================================
   RH BACKEND – STABLE (DATES JJ/MM/AAAA -> MOIS = YYYY-MM)
   EFFECTIF TOTAL = KPI_CND -> "Nombre de personnel"
   + Sous-menu "Masse salariale & rubriques" (apiRHMasse)
   ========================================================= */

const RH_CONF = {
  PROP_ID: "RH_SHEET_ID",
  TABS: {
    KPI: "KPI_CND",
    GENRE: "RH - EFFECTIF - GENRE",
    AGE: "TRANCHE_D_AGE",
    MOUVEMENT: "MOUVEMENT",
    CATEGORIE: "RH - EFFECTIF - NB"
  }
};

// Onglets du sous-menu MASSE
const RH_MASSE_CONF = {
  TABS: {
    MASSE: "RH - MASSE VS BUDGET",
    VARS: "RH - VARIABLES TRAITEES" // fallback géré dans apiRHMasse
  }
};

function normalizeTout_(v) {
  const s = String(v || "").toUpperCase().trim();
  return (s === "TOUT") ? "ALL" : s;
}

function normalizeMoisTout_(v) {
  const s = String(v || "").toUpperCase().trim();
  return (s === "TOUT") ? "" : String(v || "").trim();
}

/* =========================
   API RH – VUE GLOBALE
   ========================= */

function apiRH(params) {
  const sheetId = PropertiesService.getScriptProperties().getProperty(RH_CONF.PROP_ID);
  if (!sheetId) throw new Error("RH_SHEET_ID manquant (Propriétés du script)");

  const ss = SpreadsheetApp.openById(sheetId);

  const entParamRaw = (params && params.entite) ? params.entite : "ALL";
  const moisParamRaw = (params && params.mois) ? params.mois : "";
  const entParam = normalizeTout_(entParamRaw);
  const moisParam = normalizeMoisTout_(moisParamRaw);

  const KPI = readObjects_(ss, RH_CONF.TABS.KPI);
  const GENRE = readObjects_(ss, RH_CONF.TABS.GENRE);
  const AGE = readObjects_(ss, RH_CONF.TABS.AGE);
  const MOUV = readObjects_(ss, RH_CONF.TABS.MOUVEMENT);
  const CAT = readObjects_(ss, RH_CONF.TABS.CATEGORIE);

  const entitesList = uniq_(splitEntites_([
    ...KPI.map(r => r.entite_norm),
    ...GENRE.map(r => r.entite_norm),
    ...AGE.map(r => r.entite_norm),
    ...MOUV.map(r => r.entite_norm),
    ...CAT.map(r => r.entite_norm)
  ])).sort();

  const entites = ["ALL", ...entitesList];

  const moisList = uniq_([
    ...KPI.map(r => r.mois_norm),
    ...GENRE.map(r => r.mois_norm),
    ...AGE.map(r => r.mois_norm),
    ...MOUV.map(r => r.mois_norm),
    ...CAT.map(r => r.mois_norm)
  ]).sort();

  const defaultMois = moisList.length ? moisList[moisList.length - 1] : "";
  const mois = moisParam || defaultMois;

  const isAll = (x) => String(x || "").toUpperCase().trim() === "ALL";

  const filt = (r) =>
    (isAll(entParam) || matchEntite_(r.entite_norm, entParam)) &&
    (mois ? (r.mois_norm === mois) : true);

  const kpis = buildKpisFromCnd_(KPI.filter(filt), entParam, mois);

  const genreF = GENRE.filter(filt);
  const ageF = AGE.filter(filt);
  const mouvF = MOUV.filter(filt);
  const catF = CAT.filter(filt);

  const charts = {
    categorie: buildCategorie_(catF),
    flux: buildFlux_(mouvF),
    genre: buildGenre_(genreF),
    age: buildAge_(ageF)
  };

  return {
    filters: { entites, mois: moisList, defaultMois },
    kpis,
    charts
  };
}

// ---------- HISTORIQUE (modal courbe) ----------
function apiRHHistory(params) {
  const sheetId = PropertiesService.getScriptProperties().getProperty(RH_CONF.PROP_ID);
  if (!sheetId) throw new Error("RH_SHEET_ID manquant (Propriétés du script)");

  const ss = SpreadsheetApp.openById(sheetId);

  const ent = normalizeTout_((params && params.entite) ? params.entite : "ALL");
  const kind = String((params && params.kind) || "").trim();
  const key = String((params && params.key) || "").trim();

  const GENRE = readObjects_(ss, RH_CONF.TABS.GENRE);
  const AGE = readObjects_(ss, RH_CONF.TABS.AGE);
  const MOUV = readObjects_(ss, RH_CONF.TABS.MOUVEMENT);
  const CAT = readObjects_(ss, RH_CONF.TABS.CATEGORIE);

  const match = (r) => (ent === "ALL") ? true : matchEntite_(r.entite_norm, ent);

  const months = uniq_([
    ...GENRE.filter(match).map(r => r.mois_norm),
    ...AGE.filter(match).map(r => r.mois_norm),
    ...MOUV.filter(match).map(r => r.mois_norm),
    ...CAT.filter(match).map(r => r.mois_norm)
  ]).sort();

  const byMonth = (rows) => {
    const map = {};
    rows.forEach(r => {
      const m = r.mois_norm || "";
      if (!m) return;
      if (!map[m]) map[m] = [];
      map[m].push(r);
    });
    return map;
  };

  if (!kind) return { months, points: [], label: "" };

  if (kind === "age") {
    const map = byMonth(AGE.filter(match));
    const bucket = mapAgeKey_(key);
    const points = months.map(m => {
      const obj = buildAge_(map[m] || []);
      return { mois: m, value: Number(obj[bucket] || 0) };
    });
    return { months, points, label: key };
  }

  if (kind === "categorie") {
    const map = byMonth(CAT.filter(match));
    const field = mapCategorieField_(key);
    const points = months.map(m => {
      const obj = buildCategorie_(map[m] || []);
      return { mois: m, value: Number(obj[field] || 0) };
    });
    return { months, points, label: key };
  }

  if (kind === "genre") {
    const map = byMonth(GENRE.filter(match));
    const field = mapGenreKey_(key);
    const points = months.map(m => {
      const obj = buildGenre_(map[m] || []);
      return { mois: m, value: Number(obj[field] || 0) };
    });
    return { months, points, label: key };
  }

  if (kind === "flux") {
    const map = byMonth(MOUV.filter(match));
    const field = mapFluxKey_(key);
    const points = months.map(m => {
      const obj = buildFlux_(map[m] || []);
      return { mois: m, value: Number(obj[field] || 0) };
    });
    return { months, points, label: key };
  }

  return { months, points: [], label: "" };
}

/* =========================
   API RH – MASSE SALARIALE & RUBRIQUES
   ========================= */

function apiRHMasse(params) {
  const sheetId = PropertiesService.getScriptProperties().getProperty(RH_CONF.PROP_ID);
  if (!sheetId) throw new Error("RH_SHEET_ID manquant (Propriétés du script)");

  const ss = SpreadsheetApp.openById(sheetId);

  const entParamRaw = (params && params.entite) ? params.entite : "ALL";
  const moisParamRaw = (params && params.mois) ? params.mois : "";
  const entParam = normalizeTout_(entParamRaw);
  const moisParam = normalizeMoisTout_(moisParamRaw);

  const MASSE = readObjects_(ss, RH_MASSE_CONF.TABS.MASSE);

  let VARS = readObjects_(ss, RH_MASSE_CONF.TABS.VARS);
  if (!VARS.length) {
    // fallback si ton onglet porte ce nom
    VARS = readObjects_(ss, "RH - VARIABLES TRAITEES PAR RUBRIQUE");
  }

  const entitesList = uniq_(splitEntites_([
    ...MASSE.map(r => r.entite_norm),
    ...VARS.map(r => r.entite_norm)
  ])).sort();

  const moisList = uniq_([
    ...MASSE.map(r => r.mois_norm),
    ...VARS.map(r => r.mois_norm)
  ]).sort();

  const defaultMois = moisList.length ? moisList[moisList.length - 1] : "";
  const mois = moisParam || defaultMois;

  const isAll = (x) => String(x || "").toUpperCase().trim() === "ALL";
  const entOk = (rowEnt) => isAll(entParam) ? true : matchEntite_(rowEnt, entParam);

  // Séries MASSE (Budget/Réalisé/%)
  const months = moisList.slice();
  const byMonth = {};
  months.forEach(m => byMonth[m] = { real: 0, budget: 0 });

  MASSE.forEach(r => {
    if (!r.mois_norm) return;
    if (!entOk(r.entite_norm)) return;

    const m = r.mois_norm;
    if (!byMonth[m]) byMonth[m] = { real: 0, budget: 0 };

    // Grâce à normKey_:
    // "MS TRAITEE (SANS VARIABLES)" -> mstraiteesansvariables
    // "BUDGET (SANS VARIABLE)" -> budgetsansvariable
    byMonth[m].real += num_(r.mstraiteesansvariables);
    byMonth[m].budget += num_(r.budgetsansvariable);
  });

  const serieReal = months.map(m => +(byMonth[m]?.real || 0));
  const serieBudget = months.map(m => +(byMonth[m]?.budget || 0));
  const seriePct = months.map((m, i) => {
    const b = serieBudget[i];
    const r = serieReal[i];
    if (!b) return null;
    return +((r / b) * 100);
  });

  // Séries VARIABLES (Actes/Garde/DF/Autres + total)
  const varsByMonth = {};
  months.forEach(m => varsByMonth[m] = { actes: 0, garde: 0, df: 0, autres: 0, total: 0 });

  VARS.forEach(r => {
    if (!r.mois_norm) return;
    if (!entOk(r.entite_norm)) return;

    const m = r.mois_norm;
    if (!varsByMonth[m]) varsByMonth[m] = { actes: 0, garde: 0, df: 0, autres: 0, total: 0 };

    varsByMonth[m].actes += num_(r.actes);
    varsByMonth[m].garde += num_(r.garde);
    varsByMonth[m].df += num_(r.df);

    // "Autres (HS, astreintes, primes var …)" -> autreshsastreintesprimesvar
    varsByMonth[m].autres += num_(r.autreshsastreintesprimesvar ?? r.autres);

    // "TOTAL SANS ND" -> totalsansnd
    varsByMonth[m].total += num_(r.totalsansnd);
  });

  const varsSeries = months.map(m => ({
    mois: m,
    actes: +(varsByMonth[m]?.actes || 0),
    garde: +(varsByMonth[m]?.garde || 0),
    df: +(varsByMonth[m]?.df || 0),
    autres: +(varsByMonth[m]?.autres || 0),
    total: +(varsByMonth[m]?.total || 0)
  }));

  const selectedMasse = (() => {
    const r = byMonth[mois] || { real: 0, budget: 0 };
    const pct = (r.budget ? (r.real / r.budget) * 100 : null);
    return { mois, real: +r.real, budget: +r.budget, pct: (pct === null ? null : +pct) };
  })();

  const selectedVars = (() => {
    const v = varsByMonth[mois] || { actes: 0, garde: 0, df: 0, autres: 0, total: 0 };
    const total = +v.total || 0;
    const pct = (x) => total ? (x / total) * 100 : 0;
    return {
      mois,
      actes: +v.actes,
      garde: +v.garde,
      df: +v.df,
      autres: +v.autres,
      total: +v.total,
      pct: {
        actes: +pct(+v.actes),
        garde: +pct(+v.garde),
        df: +pct(+v.df),
        autres: +pct(+v.autres)
      }
    };
  })();

  return {
    ok: true,
    filters: { entites: ["ALL", ...entitesList], mois: moisList, defaultMois },
    state: { entite: entParam, mois },
    series: { months, budget: serieBudget, real: serieReal, pct: seriePct },
    varsSeries,
    selected: { masse: selectedMasse, vars: selectedVars }
  };
}

/* =========================
   BUILDERS
   ========================= */

function buildKpisFromCnd_(rows, ent, mois) {
  const sum = (k) => rows.reduce((a, r) => a + num_(r[k]), 0);

  // effectif total vient de KPI_CND -> "Nombre de personnel"
  const effectifTotal = sum("nombredepersonnel");

  const nbRecrutements = sum("nbrecrutements");
  const nbDeparts = sum("nbdeparts");

  const effectifDebut = sum("effectifdebutmois");
  const effectifFin = sum("effectiffinmois");

  const joursAbs = sum("joursdabsence");
  const joursTheo = sum("jourstheoriques");

  const patients = sum("nombredepatients");
  const personnel = effectifTotal;

  return {
    scope: { entite: ent, mois },
    rows: rows.length,
    effectifTotal,
    effectifDebut,
    effectifFin,
    nbRecrutements,
    nbDeparts,
    turnover: effectifTotal > 0 ? (nbDeparts / effectifTotal) : 0,
    absentisme: joursTheo > 0 ? (joursAbs / joursTheo) : 0,
    joursAbs,
    joursTheo,
    patients,
    personnel,
    ratioPatientsPersonnel: personnel > 0 ? (patients / personnel) : null
  };
}

function buildGenre_(rows) {
  let femmes = 0, hommes = 0;
  rows.forEach(r => {
    femmes += num_(r["femmes"]);
    hommes += num_(r["hommes"]);
  });
  return { femmes, hommes };
}

// AGE (robuste)
function buildAge_(rows) {
  const out = { "20-30": 0, "30-45": 0, "45-55": 0, ">55": 0 };

  rows.forEach(r => {
    const raw = String(
      r["tranchedage"] ??
      r["tranche_d_age"] ??
      r["trancheage"] ??
      r["tranche"] ??
      r["tranche d age"] ??
      ""
    );

    const label = normTxt_(raw);
    const v = num_(r["effectif"] ?? r["effectifs"] ?? r["nb"] ?? r["nombre"] ?? r["valeur"] ?? 0);

    if (/20\D*30/.test(label)) out["20-30"] += v;
    else if (/30\D*45/.test(label)) out["30-45"] += v;
    else if (/45\D*55/.test(label) && !/>\D*55/.test(label)) out["45-55"] += v;
    else if (/>+\D*55/.test(label) || /plus\D*de\D*55/.test(label) || (/55\D*ans/.test(label) && !/45\D*55/.test(label))) out[">55"] += v;
  });

  return out;
}

function buildFlux_(rows) {
  let in_ = 0, out_ = 0;

  rows.forEach(r => {
    if (r["entrants"] !== undefined || r["sortants"] !== undefined) {
      in_ += num_(r["entrants"]);
      out_ += num_(r["sortants"]);
      return;
    }
    const type = String(r["type"] || "").toUpperCase();
    const eff = num_(r["effectif"]);
    if (type.includes("ENTR")) in_ += eff;
    if (type.includes("SORT")) out_ += eff;
  });

  return { in: in_, out: out_ };
}

// CATEGORIE
function buildCategorie_(rows) {
  let paramedical = 0, direction = 0, hebergement = 0, stagiaires = 0;

  rows.forEach(r => {
    paramedical += num_(r["paramedical"]);
    direction += num_(r["direction"]);
    hebergement += num_(r["hebergement"]);
    stagiaires += num_(r["stagiaires"] ?? r["stagiaire"] ?? r["stagiere"] ?? r["stg"] ?? 0);
  });

  return {
    "Paramédical": paramedical,
    "Direction": direction,
    "Hébergement": hebergement,
    "Stagiaires": stagiaires,
    paramedical,
    direction,
    hebergement,
    stagiaires
  };
}

/* =========================
   READERS & HELPERS
   ========================= */

function readObjects_(ss, tabName) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headersRaw = values[0].map(h => String(h || "").trim());
  const headers = headersRaw.map(h => normKey_(h));

  const iEnt = headers.findIndex(h => h.includes("entite"));
  const iMois = headers.findIndex(h => h === "mois" || h.includes("date"));

  const out = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row || row.every(v => v === "" || v === null || v === undefined)) continue;

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const k = headers[c];
      if (!k) continue;
      obj[k] = row[c];
    }

    obj.entite_norm = (iEnt >= 0 ? String(row[iEnt] || "") : "").toUpperCase().trim();
    obj.mois_norm = (iMois >= 0 ? monthKey_(row[iMois]) : "");

    out.push(obj);
  }

  return out;
}

function normKey_(s) {
  return String(s || "")
    .toLowerCase()
    .trim()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function normTxt_(s) {
  return String(s || "")
    .toLowerCase()
    .trim()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .replace(/[^a-z0-9> ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function monthKey_(v) {
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    return `${y}-${m}`;
  }

  const s = String(v || "").trim();
  if (!s) return "";

  if (/^\d{4}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s.substring(0, 7);

  const m1 = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
  if (m1) return `${m1[3]}-${String(m1[2]).padStart(2, "0")}`;

  return s.length >= 7 ? s.substring(0, 7) : s;
}

function num_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/\s/g, "").replace(",", ".");
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function uniq_(arr) {
  return [...new Set((arr || []).filter(Boolean))];
}

function splitEntites_(arr) {
  const out = [];
  (arr || []).forEach(e => {
    const s = String(e || "").toUpperCase().trim();
    if (!s) return;
    if (s.includes("/")) s.split("/").forEach(x => { const v = x.trim(); if (v) out.push(v); });
    else out.push(s);
  });
  return out;
}

function matchEntite_(rowEnt, wanted) {
  const re = String(rowEnt || "").toUpperCase().trim();
  let w = String(wanted || "").toUpperCase().trim();
  if (!re || !w) return false;

  // Normalisations de libellés côté UI
  if (w === "HIA CIOA") w = "HIA/CIOA";

  // On supporte les deux côtés avec "/" :
  // - re = "HIA/CIOA" et w = "HIA"
  // - re = "HIA" et w = "HIA/CIOA"
  const rParts = re.includes("/") ? re.split("/").map(x => x.trim()).filter(Boolean) : [re];
  const wParts = w.includes("/") ? w.split("/").map(x => x.trim()).filter(Boolean) : [w];

  // match si intersection non vide
  return rParts.some(p => wParts.includes(p));
}

function mapCategorieField_(label) {
  const k = normTxt_(label);
  if (k.includes("param")) return "paramedical";
  if (k.includes("dir")) return "direction";
  if (k.includes("heb")) return "hebergement";
  if (k.includes("stag")) return "stagiaires";
  return "paramedical";
}

function mapGenreKey_(label) {
  const k = normTxt_(label);
  if (k.includes("hom")) return "hommes";
  return "femmes";
}

function mapFluxKey_(label) {
  const k = normTxt_(label);
  if (k.includes("sort") || k.includes("out")) return "out";
  return "in";
}

function mapAgeKey_(label) {
  const k = normTxt_(label);
  if (k.includes("20") && k.includes("30")) return "20-30";
  if (k.includes("30") && k.includes("45")) return "30-45";
  if (k.includes("45") && k.includes("55") && !k.includes(">")) return "45-55";
  return ">55";
}

// =======================================================================
// 11. IMPORT DÉTAIL (REC & EXP) - COLONNES FIXES POUR EXP
// =======================================================================

function importDetailsOrganisme() {
  const user = getUserContext();
  if (user.role !== 'ADMIN') return "⛔ Droit insuffisant";

  const SOURCE_ID = '19dHq5rKIgldsyvxc6nI7hP-lfdCZP7zPRKvFE1ksijU'; 
  
  try {
    const sourceSS = SpreadsheetApp.openById(SOURCE_ID);
    const localSS = SpreadsheetApp.getActiveSpreadsheet();

    // Helper Date
    function getDateParts(d) {
      if (!d || !(d instanceof Date)) return [0, 0, 0];
      return [d.getDate(), d.getMonth() + 1, d.getFullYear()];
    }

    // --- LOGIQUE SPÉCIFIQUE EXPÉDITION (COLONNES D, F, M, S) ---
    function processExpeditionFixed() {
      // 1. Chercher l'onglet (plusieurs noms possibles)
      let srcSheet = null;
      const names = ["BDD EXP", "EXPEDITION", "EXP", "DATA EXP"];
      for (let n of names) { srcSheet = sourceSS.getSheetByName(n); if(srcSheet) break; }
      
      if (!srcSheet) return "⚠️ Onglet EXP introuvable";

      const data = srcSheet.getDataRange().getValues();
      if (data.length < 2) return "⚠️ Onglet EXP vide";

      let output = [["Entité", "Jour", "Mois", "Année", "Organisme", "Montant"]];
      let count = 0;

      // 2. Lecture des colonnes fixes (Index commencant à 0 : A=0, D=3, F=5, M=12, S=18)
      const colEntite = 3;   // D
      const colOrg = 5;      // F
      const colMontant = 12; // M
      const colDate = 18;    // S

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Vérification de sécurité pour éviter les erreurs "Out of bounds"
        if (!row[colEntite] && row[colEntite] !== 0) continue; // Skip ligne vide
        
        const entite = row[colEntite];
        const orgVal = row[colOrg] || "Inconnu";
        
        // Nettoyage Montant (M)
        let mntVal = row[colMontant];
        if (typeof mntVal !== 'number') {
           mntVal = parseFloat(String(mntVal).replace(/[\s,]/g, (m) => m === ',' ? '.' : '')) || 0;
        }

        // Nettoyage Date (S)
        let rawDate = row[colDate];
        let finalDate = null;
        if (rawDate instanceof Date) {
            finalDate = rawDate;
        } else if (rawDate) {
            // Tentative de parsing si c'est du texte
            let s = String(rawDate).trim();
            if (s.length >= 8) { // ex: 01/01/2024
               let parts = s.split(/[\/\-\.]/);
               if (parts.length === 3) finalDate = new Date(parts[2], parts[1] - 1, parts[0]);
            }
        }

        if (finalDate && !isNaN(finalDate.getTime())) {
             const [jj, mm, aa] = getDateParts(finalDate);
             output.push([entite, jj, mm, aa, orgVal, mntVal]);
             count++;
        }
      }

      // 3. Écriture
      let targetSheet = localSS.getSheetByName("DETAIL EXP");
      if (!targetSheet) targetSheet = localSS.insertSheet("DETAIL EXP");
      targetSheet.clear();
      
      if (output.length > 1) {
        targetSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
        return `${count} lignes (Colonnes D,F,M,S)`;
      }
      return "0 ligne importée (Vérifiez le format date Col S)";
    }

    // --- LOGIQUE RECOUVREMENT (Standard / Intelligente) ---
    function processRecouvrementSmart() {
       let srcSheet = null;
       const names = ["BDD REC", "REC", "RECOUVREMENT"];
       for (let n of names) { srcSheet = sourceSS.getSheetByName(n); if(srcSheet) break; }
       if (!srcSheet) return "⚠️ Onglet REC introuvable";

       const data = srcSheet.getDataRange().getValues();
       if (data.length < 2) return "⚠️ Onglet REC vide";
       
       // Détection auto ou défaut (Entité=C/2, Date=D/3, Org=E/4, Mnt=F/5)
       // ... (code simplifié pour REC, on suppose que REC fonctionne ou utilise les défauts)
       let output = [["Entité", "Jour", "Mois", "Année", "Organisme", "Montant"]];
       let count = 0;
       
       for (let i = 1; i < data.length; i++) {
         const row = data[i];
         // Défauts BDD REC : C=Entité(2), D=Date(3), E=Org(4), F=Montant(5)
         // Ajustez si votre REC a changé aussi
         let ent = row[2]; 
         let dt = row[3];
         let org = row[4];
         let mnt = row[5];
         
         if(ent) {
            if(typeof mnt !== 'number') mnt = parseFloat(String(mnt).replace(/[\s,]/g, (m) => m === ',' ? '.' : '')) || 0;
            let fDate = (dt instanceof Date) ? dt : null;
            if(!fDate && dt) { let p = String(dt).split('/'); if(p.length===3) fDate = new Date(p[2], p[1]-1, p[0]); }
            
            if(fDate) {
              const [jj,mm,aa] = getDateParts(fDate);
              output.push([ent, jj, mm, aa, org, mnt]);
              count++;
            }
         }
       }
       
       let targetSheet = localSS.getSheetByName("DETAIL REC");
       if (!targetSheet) targetSheet = localSS.insertSheet("DETAIL REC");
       targetSheet.clear();
       if (output.length > 1) targetSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
       return `${count} lignes`;
    }

    const resExp = processExpeditionFixed();
    const resRec = processRecouvrementSmart();

    return `✅ Import Terminé.\nEXP: ${resExp}\nREC: ${resRec}`;

  } catch (e) {
    return "❌ Erreur : " + e.message;
  }
}

// Assurez-vous d'avoir toujours cette fonction aussi :
function getDetailsOrganismeData(type) {
  const sheetName = (type === 'EXP') ? 'DETAIL EXP' : 'DETAIL REC';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data;
}