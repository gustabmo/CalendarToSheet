// 2026-05-02 by Gustavo Exel and claude.ai

// =============================================================================
// CALENDRIER SCOLAIRE — Google Apps Script  v3
// À attacher à un fichier Google Sheets.
// Menu : 📅 Calendrier
//
// Onglets générés (script-managed, ne pas modifier manuellement) :
//   "[année] Général"           — calendrier public
//   "[année] Parents"           — événements non tagués (tous niveaux)
//   "[année] Parents JE"        — parents + événements #jardindenfants
//   "[année] Parents Prim"      — parents + événements #primaire
//   "[année] Parents Sec1"      — parents + événements #secondaire1
//   "[année] Parents Sec2"      — parents + événements #secondaire2
//   "[année] Profs JE"          — parents JE + événements #jardindenfants du cal. profs
//   "[année] Profs Prim"        — parents Prim + événements #primaire du cal. profs
//   "[année] Profs Sec1"        — parents Sec1 + événements #secondaire1 du cal. profs
//   "[année] Profs Sec2"        — parents Sec2 + événements #secondaire2 du cal. profs
// =============================================================================

// ================================================
// CALENDRIER SCOLAIRE — notes pour session future
//
// Contexte : Google Apps Script attaché à un Google Sheets.
// École à Genève (francophone). 3 calendriers : public, parents, professeurs.
// Niveaux : jardindenfants, primaire, secondaire1, secondaire2.
//
// Tags dans le champ "description" des événements Google Calendar :
//   #vacances           → colore la cellule en orange clair
//   #compact:texte      → titre court pour la grille
//   #jardindenfants / #primaire / #secondaire1 / #secondaire2  → filtre par niveau
//   Pas de tag niveau   → apparaît sur tous les onglets
//
// Vacances GE : récupérées automatiquement depuis ge.ch
// Couleurs : bleu = week-end, orange foncé = vacances GE, orange clair = vacances école
// ================================================

// ============================================================================
// CONFIG SHEET LAYOUT  (onglet nommé "Config")
//
// ── CALENDRIERS (colonnes A–B) ────────────────────────────────────────────
//   Ligne 1 : en-têtes (ignorés)
//   Lignes 2+ :
//     A : Rôle du calendrier    "public" | "parents" | "professeurs"
//     B : Google Calendar ID    ex. "abc@group.calendar.google.com"
//
// ── PARAMÈTRES (colonnes D–E) ─────────────────────────────────────────────
//   Clé (col D)             Valeur (col E)
//   FETCH_GE_VACANCES       TRUE  ou  FALSE
//
// ── ANNÉES SCOLAIRES (colonnes G–I) ──────────────────────────────────────
//   Ligne 1 : en-têtes (ignorés)
//   Lignes 2+ :
//     G : Libellé    ex. "2025-26"  → préfixe des noms d'onglets
//     H : Début      ex. "2025-08-01"
//     I : Fin        ex. "2026-08-31"
// ============================================================================

// ---- Palette de couleurs ---------------------------------------------------
var COLOR_WEEKEND        = "#C9DAF8";  // bleu — week-ends
var COLOR_GE_VACATION    = "#E06B0A";  // orange foncé — vacances cantonales GE
var COLOR_OWN_VACATION   = "#FCE5CD";  // orange clair — vacances propres à l'école
var COLOR_HEADER_BG      = "#434343";  // fond en-têtes de mois
var COLOR_HEADER_FG      = "#FFFFFF";  // texte en-têtes de mois
var COLOR_NODAY_BG       = "#EFEFEF";  // jours inexistants (ex. 31 février)

// ---- Seuils de taille de police (en nombre de caractères du texte affiché) -
var FONT_SIZE_NORMAL     = 7;
var FONT_SIZE_MEDIUM     = 6;
var FONT_SIZE_SMALL      = 5;
var FONT_THRESHOLD_MEDIUM = 18;  // > N caractères → taille medium
var FONT_THRESHOLD_SMALL  = 30;  // > N caractères → taille small

// ---- Niveaux scolaires -----------------------------------------------------
// Clé = tag dans la description, Valeur = suffixe utilisé dans le nom d'onglet
var LEVELS = {
  "jardindenfants": "JE",
  "primaire":       "Prim",
  "secondaire1":    "Sec1",
  "secondaire2":    "Sec2"
};
var LEVEL_KEYS = ["jardindenfants", "primaire", "secondaire1", "secondaire2"];

// ---- URL de base vacances GE -----------------------------------------------
var GE_VACATION_BASE = "https://www.ge.ch/vacances-scolaires-jours-feries/vacances-scolaires-";

// ============================================================================
// MENU
// ============================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📅 Calendrier")
    .addItem("Générer les onglets pour une année…", "promptGenerateYear")
    .addSeparator()
    .addItem("Initialiser l'onglet Config",         "setupConfigSheet")
    .addToUi();
}

// ============================================================================
// DIALOGUE DE SÉLECTION D'ANNÉE
// ============================================================================
function promptGenerateYear() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = readConfig_(ss);
  if (!config) return;

  if (config.years.length === 0) {
    SpreadsheetApp.getUi().alert("Aucune année scolaire trouvée dans l'onglet Config (colonnes G–I).");
    return;
  }

  var labels = config.years.map(function(y){ return y.label; });
  var ui     = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "Générer les onglets",
    "Entrez le libellé de l'année à générer :\n" + labels.join("  |  "),
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var chosen = result.getResponseText().trim();
  var year   = config.years.filter(function(y){ return y.label === chosen; })[0];
  if (!year) {
    ui.alert("Année « " + chosen + " » introuvable dans la Config.");
    return;
  }

  generateAllViewsForYear_(ss, year, config);
}

// ============================================================================
// GÉNÈRE TOUS LES ONGLETS D'UNE ANNÉE
// ============================================================================
function generateAllViewsForYear_(ss, year, config) {
  var geVacations = [];
  if (config.fetchGe) {
    geVacations = fetchGeVacations_(year.start.getFullYear(), year.end.getFullYear());
    Logger.log("Vacances GE récupérées : " + geVacations.length + " périodes");
  }

  // Collecte tous les événements des 3 calendriers
  var calPublic = collectCalendarEvents_(config.calPublicId,  year.start, year.end);
  var calParents= collectCalendarEvents_(config.calParentsId, year.start, year.end);
  var calProfs  = collectCalendarEvents_(config.calProfsId,   year.start, year.end);

  var label = year.label;

  // ── 1. Général (public uniquement, tous événements) ──────────────────────
  generateSheet_(ss, label + " Général", year, geVacations,
    mergeEventMaps_([calPublic]),
    { mode: "général" });

  // ── 2. Parents général (public + parents, événements NON tagués niveau) ──
  generateSheet_(ss, label + " Parents", year, geVacations,
    mergeEventMaps_([calPublic, calParents]),
    { mode: "parents-général" });

  // ── 3–6. Parents par niveau ───────────────────────────────────────────────
  LEVEL_KEYS.forEach(function(levelKey) {
    var suffix = LEVELS[levelKey];
    generateSheet_(ss, label + " Parents " + suffix, year, geVacations,
      mergeEventMaps_([calPublic, calParents]),
      { mode: "parents-niveau", level: levelKey });
  });

  // ── 7–10. Profs par niveau ────────────────────────────────────────────────
  LEVEL_KEYS.forEach(function(levelKey) {
    var suffix = LEVELS[levelKey];
    generateSheet_(ss, label + " Profs " + suffix, year, geVacations,
      mergeEventMaps_([calPublic, calParents, calProfs]),
      { mode: "profs-niveau", level: levelKey });
  });

  SpreadsheetApp.getUi().alert("Onglets générés pour l'année " + label + " ✓");
}

// ============================================================================
// RÈGLES DE VISIBILITÉ D'UN ÉVÉNEMENT SELON LE MODE DE L'ONGLET
//
// Modes :
//   "général"        — calendrier public, tout passe
//   "parents-général"— public + parents, uniquement événements SANS tag de niveau
//   "parents-niveau" — public + parents, événements sans tag OU avec ce niveau
//   "profs-niveau"   — public + parents + profs, idem parents-niveau
//
// Retourne : false = invisible | true = visible
// ============================================================================
function isEventVisible_(evt, viewCfg) {
  var hasLevelTag = evt.levels.length > 0;

  switch (viewCfg.mode) {
    case "général":
      return true;

    case "parents-général":
      // Uniquement les événements sans aucun tag de niveau
      return !hasLevelTag;

    case "parents-niveau":
    case "profs-niveau":
      // Visible si pas de tag de niveau, ou si ce niveau est parmi les tags
      return !hasLevelTag || evt.levels.indexOf(viewCfg.level) !== -1;
  }
  return false;
}

// ============================================================================
// GÉNÈRE UN ONGLET
// ============================================================================
function generateSheet_(ss, sheetName, year, geVacations, events, viewCfg) {
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet(sheetName);

  var months    = buildMonthList_(year.start, year.end);
  var NUM_MONTHS = months.length;
  var DAY_ROWS   = 31;
  var TITLE_ROW  = 1;
  var HEADER_ROW = 2;
  var DATA_START = 3;

  // Hauteurs de lignes
  sheet.setRowHeight(TITLE_ROW,  28);
  sheet.setRowHeight(HEADER_ROW, 20);
  for (var r = DATA_START; r < DATA_START + DAY_ROWS; r++) {
    sheet.setRowHeight(r, 18);
  }

  // Titre
  sheet.getRange(TITLE_ROW, 1, 1, NUM_MONTHS * 2).merge()
    .setValue(sheetName)
    .setFontSize(13)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#FFFFFF");

  // Colonnes par mois
  for (var m = 0; m < NUM_MONTHS; m++) {
    var mo      = months[m];
    var colDay  = m * 2 + 1;
    var colEvt  = colDay + 1;

    // En-tête de mois
    sheet.getRange(HEADER_ROW, colDay, 1, 2).merge()
      .setValue(mo.label)
      .setFontWeight("bold")
      .setFontSize(10)
      .setBackground(COLOR_HEADER_BG)
      .setFontColor(COLOR_HEADER_FG)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    var daysInMonth = new Date(mo.year, mo.month + 1, 0).getDate();

    for (var d = 1; d <= DAY_ROWS; d++) {
      var row = DATA_START + d - 1;

      if (d > daysInMonth) {
        sheet.getRange(row, colDay, 1, 2)
          .setValue("").setBackground(COLOR_NODAY_BG);
        continue;
      }

      var dateObj   = new Date(mo.year, mo.month, d);
      var dateKey   = formatDateKey_(dateObj);
      var dow       = dateObj.getDay();
      var isWeekend = dow === 0 || dow === 6;
      var isGeVac   = isInRanges_(dateObj, geVacations);
      var isFirstDayOfMonth = (d === 1);

      // Événements du jour filtrés selon le mode de l'onglet
      var dayEvts = (events[dateKey] || []).filter(function(e){
        return isEventVisible_(e, viewCfg);
      });

      // Sépare vacances et événements ordinaires
      var vacEvts    = dayEvts.filter(function(e){ return e.isVacance; });
      var normalEvts = dayEvts.filter(function(e){ return !e.isVacance; });

      // ---- Couleur de fond (priorité : GE > propre école > week-end) -------
      var isSchoolVac = vacEvts.length > 0;
      var bg = null;
      if      (isGeVac)      bg = COLOR_GE_VACATION;
      else if (isSchoolVac)  bg = COLOR_OWN_VACATION;
      else if (isWeekend)    bg = COLOR_WEEKEND;

      if (bg) sheet.getRange(row, colDay, 1, 2).setBackground(bg);

      // ---- Numéro du jour --------------------------------------------------
      sheet.getRange(row, colDay)
        .setValue(d)
        .setFontSize(8)
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle");

      // ---- Texte de l'événement --------------------------------------------
      // Pour les vacances : texte uniquement le 1er jour du mois (ou 1er jour absolu)
      var textParts = [];

      // Vacances : affiche le titre seulement au 1er jour visible du mois
      if (vacEvts.length > 0) {
        var isFirstDayOfVac = isFirstDayOfRange_(dateObj, vacEvts[0], mo.month);
        if (isFirstDayOfVac) {
          textParts.push(vacEvts[0].displayTitle);
        }
      }

      // Événements normaux : toujours affichés
      normalEvts.forEach(function(e){ textParts.push(e.displayTitle); });

      if (textParts.length > 0) {
        var text     = textParts.join(" / ");
        var fontSize = fontSizeForLength_(text.length);
        sheet.getRange(row, colEvt)
          .setValue(text)
          .setFontSize(fontSize)
          .setWrap(true)
          .setHorizontalAlignment("left")
          .setVerticalAlignment("middle");
      } else {
        sheet.getRange(row, colEvt).setValue("");
      }
    }

    sheet.setColumnWidth(colDay,  26);
    sheet.setColumnWidth(colEvt,  88);
  }

  // Bordures
  sheet.getRange(DATA_START, 1, DAY_ROWS, NUM_MONTHS * 2)
    .setBorder(true, true, true, true, true, true,
               "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);

  // Légende
  var legendRow = DATA_START + DAY_ROWS + 1;
  sheet.getRange(legendRow, 1).setValue("Légende").setFontWeight("bold").setFontSize(9);
  [
    [COLOR_GE_VACATION,  "Vacances cantonales GE"],
    [COLOR_OWN_VACATION, "Vacances propres à l'école"],
    [COLOR_WEEKEND,      "Week-end"],
  ].forEach(function(item, i) {
    var lr = legendRow + 1 + i;
    sheet.getRange(lr, 1, 1, 2).setBackground(item[0]);
    sheet.getRange(lr, 2).setValue(item[1]).setFontSize(8);
    sheet.setRowHeight(lr, 16);
  });

  sheet.setFrozenRows(2);
  protectSheet_(sheet);
}

// ============================================================================
// COLLECTE LES ÉVÉNEMENTS D'UN CALENDRIER
// Retourne : { "YYYY-MM-DD": [ eventObj, ... ], ... }
//
// eventObj = {
//   displayTitle : string  (titre compact ou titre brut, tags supprimés)
//   isVacance    : bool
//   levels       : [string]  ex. ["primaire", "secondaire1"]
//   startDate    : Date  (minuit, heure locale)
//   endDate      : Date  (minuit, exclusif)
// }
// ============================================================================
function collectCalendarEvents_(calId, startDate, endDate) {
  var result = {};
  if (!calId) return result;

  var cal;
  try {
    cal = (calId === "primary")
      ? CalendarApp.getDefaultCalendar()
      : CalendarApp.getCalendarById(calId);
  } catch(e) {
    Logger.log("Impossible d'accéder au calendrier : " + calId + " — " + e.message);
    return result;
  }
  if (!cal) { Logger.log("Calendrier introuvable : " + calId); return result; }

  var calEvents = cal.getEvents(startDate, endDate);
  calEvents.forEach(function(evt) {
    var isAllDay = evt.isAllDayEvent();
    var rawTitle = evt.getTitle();
    var desc     = evt.getDescription() || "";

    // Parse la description pour extraire les tags
    var parsed   = parseDescription_(desc, rawTitle);

    var cursor, endCursor;
    if (isAllDay) {
      cursor    = new Date(evt.getAllDayStartDate()); cursor.setHours(0,0,0,0);
      endCursor = new Date(evt.getAllDayEndDate());   endCursor.setHours(0,0,0,0);
    } else {
      // Événements avec heure : traités comme mono-journée (heures ignorées)
      cursor    = new Date(evt.getStartTime()); cursor.setHours(0,0,0,0);
      endCursor = new Date(cursor);
      endCursor.setDate(endCursor.getDate() + 1);
    }

    var evtStartDate = new Date(cursor);
    var evtEndDate   = new Date(endCursor);

    while (cursor < endCursor) {
      var key = formatDateKey_(cursor);
      if (!result[key]) result[key] = [];
      result[key].push({
        displayTitle : parsed.displayTitle,
        isVacance    : parsed.isVacance,
        levels       : parsed.levels,
        startDate    : evtStartDate,
        endDate      : evtEndDate
      });
      cursor.setDate(cursor.getDate() + 1);
    }
  });

  return result;
}

// ============================================================================
// PARSE LA DESCRIPTION D'UN ÉVÉNEMENT
// Extrait : #vacances, #compact:..., #jardindenfants, #primaire, #secondaire1, #secondaire2
// ============================================================================
function parseDescription_(desc, rawTitle) {
  var isVacance    = false;
  var levels       = [];
  var compactTitle = null;

  // #vacances
  if (/#vacances\b/i.test(desc)) isVacance = true;

  // #compact:texte court
  var compactMatch = desc.match(/#compact:([^#\n\r]+)/i);
  if (compactMatch) compactTitle = compactMatch[1].trim();

  // Tags de niveaux
  LEVEL_KEYS.forEach(function(k) {
    if (new RegExp("#" + k + "\\b", "i").test(desc)) levels.push(k);
  });

  // Titre affiché = compact si dispo, sinon titre brut
  var displayTitle = compactTitle || rawTitle;

  return { isVacance: isVacance, levels: levels, displayTitle: displayTitle };
}

// ============================================================================
// FUSIONNE PLUSIEURS MAPS D'ÉVÉNEMENTS (un par calendrier) EN UNE SEULE
// ============================================================================
function mergeEventMaps_(maps) {
  var merged = {};
  maps.forEach(function(m) {
    Object.keys(m).forEach(function(key) {
      if (!merged[key]) merged[key] = [];
      merged[key] = merged[key].concat(m[key]);
    });
  });
  return merged;
}

// ============================================================================
// PREMIER JOUR VISIBLE D'UNE VACATION DANS UN MOIS DONNÉ
// Retourne true si dateObj est le 1er jour de la vacation dans ce mois
// (soit le 1er du mois si la vacation a démarré avant, soit le 1er jour de la vacation)
// ============================================================================
function isFirstDayOfRange_(dateObj, evt, currentMonth) {
  var d = new Date(dateObj);
  // Premier jour du mois courant
  var firstOfMonth = new Date(d.getFullYear(), currentMonth, 1);
  // Début effectif de la vacation dans ce mois
  var effectiveStart = evt.startDate > firstOfMonth ? evt.startDate : firstOfMonth;
  effectiveStart.setHours(0,0,0,0);
  d.setHours(0,0,0,0);
  return d.getTime() === effectiveStart.getTime();
}

// ============================================================================
// TAILLE DE POLICE SELON LA LONGUEUR DU TEXTE
// ============================================================================
function fontSizeForLength_(len) {
  if (len > FONT_THRESHOLD_SMALL)  return FONT_SIZE_SMALL;
  if (len > FONT_THRESHOLD_MEDIUM) return FONT_SIZE_MEDIUM;
  return FONT_SIZE_NORMAL;
}

// ============================================================================
// RÉCUPÈRE LES VACANCES SCOLAIRES GE
// ============================================================================
function fetchGeVacations_(startYear, endYear) {
  var vacations = [];
  for (var y = startYear; y < endYear; y++) {
    var slug = y + "-" + (y + 1);
    var url  = GE_VACATION_BASE + slug;
    try {
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) {
        Logger.log("ge.ch HTTP " + resp.getResponseCode() + " pour " + url);
        continue;
      }
      var parsed = parseGeVacationsHtml_(resp.getContentText());
      Logger.log("  " + slug + " : " + parsed.length + " périodes");
      vacations = vacations.concat(parsed);
    } catch(e) {
      Logger.log("Erreur ge.ch pour " + slug + " : " + e.message);
    }
  }
  return vacations;
}

function parseGeVacationsHtml_(html) {
  var MONTHS_FR = {
    "janvier":1,"février":2,"mars":3,"avril":4,"mai":5,"juin":6,
    "juillet":7,"août":8,"septembre":9,"octobre":10,"novembre":11,"décembre":12
  };
  var text      = html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ");
  var vacations = [];
  var re = /du\s+(?:\w+\s+)?(\d{1,2})(?:er|ème)?\s+(\w+)\s+(\d{4})\s+au\s+(?:\w+\s+)?(\d{1,2})(?:er|ème)?\s+(\w+)\s+(\d{4})/gi;
  var m;
  while ((m = re.exec(text)) !== null) {
    var m1 = MONTHS_FR[m[2].toLowerCase()];
    var m2 = MONTHS_FR[m[5].toLowerCase()];
    if (!m1 || !m2) continue;
    var start = new Date(parseInt(m[3]), m1-1, parseInt(m[1]));
    var end   = new Date(parseInt(m[6]), m2-1, parseInt(m[4]) + 1);
    if (!isNaN(start) && !isNaN(end) && end > start) vacations.push({start:start, end:end});
  }
  return vacations;
}

// ============================================================================
// LIRE LA CONFIG
// ============================================================================
function readConfig_(ss) {
  var cfgSheet = ss.getSheetByName("Config");
  if (!cfgSheet) {
    SpreadsheetApp.getUi().alert('Onglet "Config" introuvable.\nUtilisez 📅 Calendrier → Initialiser l\'onglet Config.');
    return null;
  }
  var data = cfgSheet.getDataRange().getValues();

  // Calendriers (col A=0, B=1)
  var calPublicId = "", calParentsId = "", calProfsId = "";
  for (var r = 1; r < data.length; r++) {
    var role = String(data[r][0] || "").trim().toLowerCase();
    var id   = String(data[r][1] || "").trim();
    if (!role || !id) continue;
    if (role === "public")      calPublicId  = id;
    if (role === "parents")     calParentsId = id;
    if (role === "professeurs") calProfsId   = id;
  }

  // Paramètres (col D=3, E=4)
  var settings = {};
  for (var r = 1; r < data.length; r++) {
    var k = String(data[r][3] || "").trim();
    var v = data[r][4];
    if (k) settings[k] = v;
  }
  var fetchGe = String(settings["FETCH_GE_VACANCES"] || "FALSE").toUpperCase() === "TRUE";

  // Années scolaires (col G=6, H=7, I=8)
  var years = [];
  for (var r = 1; r < data.length; r++) {
    var label = String(data[r][6] || "").trim();
    var start = data[r][7];
    var end   = data[r][8];
    if (!label || !start || !end) continue;
    var sd = (start instanceof Date) ? start : new Date(start);
    var ed = (end   instanceof Date) ? end   : new Date(end);
    if (isNaN(sd) || isNaN(ed)) continue;
    years.push({ label: label, start: sd, end: ed });
  }

  return {
    calPublicId:  calPublicId,
    calParentsId: calParentsId,
    calProfsId:   calProfsId,
    fetchGe:      fetchGe,
    years:        years
  };
}

// ============================================================================
// INITIALISATION DE L'ONGLET CONFIG
// ============================================================================
function setupConfigSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Config") || ss.insertSheet("Config");
  sheet.clearContents();
  sheet.clearFormats();

  var hCells = [
    [1,1,"Rôle calendrier"],[1,2,"Google Calendar ID"],
    [1,4,"Paramètre"],[1,5,"Valeur"],
    [1,7,"Libellé année"],[1,8,"Début"],[1,9,"Fin"]
  ];
  hCells.forEach(function(h){
    sheet.getRange(h[0],h[1]).setValue(h[2])
      .setFontWeight("bold").setBackground("#444444").setFontColor("#FFFFFF");
  });

  // Exemples calendriers
  [
    ["public",      "votre-cal-public@group.calendar.google.com"],
    ["parents",     "votre-cal-parents@group.calendar.google.com"],
    ["professeurs", "votre-cal-profs@group.calendar.google.com"],
  ].forEach(function(r,i){ sheet.getRange(i+2,1,1,2).setValues([r]); });

  // Paramètres
  sheet.getRange(2,4,1,2).setValues([["FETCH_GE_VACANCES","TRUE"]]);

  // Années scolaires
  [
    ["2025-26","2025-08-01","2026-08-31"],
    ["2026-27","2026-08-01","2027-08-31"],
  ].forEach(function(r,i){ sheet.getRange(i+2,7,1,3).setValues([r]); });

  // Largeurs de colonnes
  [160,280,20,190,80,20,100,100,100].forEach(function(w,i){
    sheet.setColumnWidth(i+1,w);
  });

  // Note
  sheet.getRange(7,1,1,5).merge()
    .setValue("⚠️  Les onglets générés sont gérés par le script. Ne pas les modifier manuellement.")
    .setFontStyle("italic").setFontSize(9).setFontColor("#888888");

  SpreadsheetApp.getUi().alert(
    "Onglet Config prêt.\n\n" +
    "1. Remplacez les IDs de calendrier par les vrais\n" +
    "   (Google Agenda → Paramètres → [calendrier] → ID du calendrier)\n" +
    "2. Ajustez les années scolaires (colonnes G–I)\n" +
    "3. Utilisez 📅 Calendrier → Générer les onglets pour une année…"
  );
}

// ============================================================================
// PROTECTION DE L'ONGLET (warning uniquement)
// ============================================================================
function protectSheet_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(function(p){ p.remove(); });
  sheet.protect()
    .setDescription("Géré par le script — utiliser 📅 Calendrier pour mettre à jour.")
    .setWarningOnly(true);
}

// ============================================================================
// UTILITAIRES
// ============================================================================
function buildMonthList_(startDate, endDate) {
  var NOMS = ["Janvier","Février","Mars","Avril","Mai","Juin",
              "Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
  var months = [];
  var cursor = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  var limit  = new Date(endDate.getFullYear(),   endDate.getMonth(),   1);
  while (cursor <= limit) {
    months.push({ label: NOMS[cursor.getMonth()], month: cursor.getMonth(), year: cursor.getFullYear() });
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return months;
}

function formatDateKey_(date) {
  return date.getFullYear() + "-"
    + String(date.getMonth()+1).padStart(2,"0") + "-"
    + String(date.getDate()).padStart(2,"0");
}

function isInRanges_(dateObj, ranges) {
  for (var i = 0; i < ranges.length; i++) {
    if (dateObj >= ranges[i].start && dateObj < ranges[i].end) return true;
  }
  return false;
}
