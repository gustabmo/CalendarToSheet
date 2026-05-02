// 2026-05-02 by Gustavo Exel and claude.ai

// =============================================================================
// CALENDRIER SCOLAIRE — Google Apps Script  v4
// À attacher à un fichier Google Sheets.
// Menu : Calendrier
//
// Onglets générés (script-managed, ne pas modifier manuellement) :
//   "[année] Public"            — calendrier public
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
// ── RÈGLE DES COULEURS DE FOND (priorité décroissante) ───────────────────
//   1. BLEU        (COLOR_WEEKEND)      — samedi et dimanche
//   2. ORANGE FONCÉ (COLOR_GE_VACATION) — jours sans école dans le calendrier
//                                         officiel DIP / Canton de Genève :
//                                           • Périodes de vacances scolaires
//                                             (automne, Noël, février, Pâques, été)
//                                           • Jours fériés officiels GE chômés
//                                             à l'école : Jeûne genevois, Fête du
//                                             travail, Ascension, Lundi de Pentecôte,
//                                             Restauration de la République (31 déc),
//                                             Nouvel An, Vendredi-Saint, Lundi de
//                                             Pâques, Fête nationale (1er août), Noël
//                                           • Ponts accordés par le DIP
//                                             (ex. Pont de l'Ascension = jeudi+vendredi)
//                                         Source : https://www.ge.ch/vacances-scolaires-jours-feries
//                                         Scraped chaque génération depuis les pages
//                                         "vacances-scolaires-YYYY-YYYY" et
//                                         "jours-feries-officiels-2023-2027"
//   3. ORANGE CLAIR (COLOR_OWN_VACATION) — jours de congé supplémentaires propres
//                                          à notre école (marqués #vacances dans
//                                          le calendrier Google "parents" ou "public")
//                                          Notre école est un SUPERSET du calendrier
//                                          GE : tous les jours GE sont off + quelques
//                                          jours additionnels propres à notre école.
//
// Les jours GE sont récupérés automatiquement si FETCH_GE_VACANCES = TRUE.
// ── FIN RÈGLE COULEURS ───────────────────────────────────────────────────
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
// ── ANNÉES SCOLAIRES (colonnes G–J) ──────────────────────────────────────
//   Ligne 1 : en-têtes (ignorés)
//   Lignes 2+ :
//     G : Libellé    ex. "2025-26"  → préfixe des noms d'onglets
//     H : Début      ex. "2025-08-01"
//     I : Fin        ex. "2026-08-31"
//     J : Générer    laisser non-vide (ex. "X" ou "OUI") pour l'année à générer
//                    → exactement UNE ligne doit avoir cette colonne non-vide
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
    .createMenu("Calendrier")
    .addItem("Générer les onglets (année marquée dans Config)", "generateMarkedYear")
    .addSeparator()
    .addItem("Initialiser l'onglet Config", "setupConfigSheet")
    .addToUi();
}

// ============================================================================
// GÉNÉRATION DEPUIS LA CONFIG (colonne J)
// Lit la colonne J (index 9) des années scolaires.
// Exactement une ligne doit avoir une valeur non-vide dans J.
// Erreur si zéro ou plus d'une ligne est marquée.
// ============================================================================
function generateMarkedYear() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = readConfig_(ss);
  if (!config) return;

  var ui = SpreadsheetApp.getUi();

  // Filtre les années ayant la colonne "Générer" non-vide
  var marked = config.years.filter(function(y) { return y.generate; });

  if (marked.length === 0) {
    ui.alert(
      "Erreur : aucune année à générer.\n\n" +
      "Mettez une valeur (ex. « X ») dans la colonne J (Générer) " +
      "de l'année souhaitée dans l'onglet Config."
    );
    return;
  }
  if (marked.length > 1) {
    ui.alert(
      "Erreur : plusieurs années marquées pour génération (" +
      marked.map(function(y){ return y.label; }).join(", ") + ").\n\n" +
      "Ne laissez qu'une seule valeur non-vide dans la colonne J (Générer)."
    );
    return;
  }

  generateAllViewsForYear_(ss, marked[0], config);
}

// ============================================================================
// GÉNÈRE TOUS LES ONGLETS D'UNE ANNÉE
// ============================================================================
function generateAllViewsForYear_(ss, year, config) {
  var geVacations = [];
  if (config.fetchGe) {
    geVacations = fetchGeVacations_(year.start.getFullYear(), year.end.getFullYear(), year.start);
    Logger.log("Vacances GE récupérées : " + geVacations.length + " périodes");
  }

  // Collecte tous les événements des 3 calendriers
  var calPublic = collectCalendarEvents_(config.calPublicId,  year.start, year.end);
  var calParents= collectCalendarEvents_(config.calParentsId, year.start, year.end);
  var calProfs  = collectCalendarEvents_(config.calProfsId,   year.start, year.end);

  var label = year.label;

  // ── 1. Public (public uniquement, tous événements) ──────────────────────
  generateSheet_(ss, label + " Public", year, geVacations,
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
  if (evt.horsAnnuel) return false;

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
// Efface le contenu et reformate l'onglet existant (sans le supprimer, pour
// préserver l'URL de publication web). Crée l'onglet s'il n'existe pas encore.
//
// Stratégie de performance : toutes les propriétés de cellule sont d'abord
// calculées en mémoire dans des tableaux 2D, puis écrites en un minimum
// d'appels batch à l'API Sheets (setValues, setBackgrounds, setFontSizes,
// setFontWeights, setHorizontalAlignments, setVerticalAlignments,
// setFontColors, setWraps). Cela réduit les appels API de plusieurs milliers
// à une dizaine, ce qui est le principal levier de performance sur Apps Script.
// ============================================================================
function generateSheet_(ss, sheetName, year, geVacations, events, viewCfg) {
  // ── Récupère ou crée l'onglet sans le supprimer ──────────────────────────
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    // Efface contenu, formats et fusions pour repartir à zéro
    // (breakApart sur toute la plage pour dissoudre les cellules fusionnées
    //  avant clearFormats, qui échouerait sur des plages fusionnées)
    sheet.getDataRange().breakApart();
    sheet.clearContents();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var months     = buildMonthList_(year.start, year.end);
  var NUM_MONTHS = months.length;
  var NUM_COLS   = NUM_MONTHS * 2;
  var DAY_ROWS   = 31;
  var TITLE_ROW  = 1;   // 1-indexed for Sheets API
  var HEADER_ROW = 2;
  var DATA_START = 3;
  // Total rows in our managed area: title(1) + header(1) + data(31)
  var TOTAL_ROWS = 1 + 1 + DAY_ROWS;

  // ── Hauteurs de lignes (peu d'appels, pas de gain à batcher) ────────────
  sheet.setRowHeight(TITLE_ROW,  28);
  sheet.setRowHeight(HEADER_ROW, 20);
  for (var r = DATA_START; r < DATA_START + DAY_ROWS; r++) {
    sheet.setRowHeight(r, 18);
  }

  // ── Largeurs de colonnes ─────────────────────────────────────────────────
  for (var m = 0; m < NUM_MONTHS; m++) {
    sheet.setColumnWidth(m * 2 + 1, 26);
    sheet.setColumnWidth(m * 2 + 2, 88);
  }

  // ── Initialise les tableaux 2D (indexés [rowIndex][colIndex], 0-based) ──
  // Couvrent les lignes TITLE_ROW … DATA_START+DAY_ROWS-1
  var values     = [];   // cell text / number
  var backgrounds= [];   // background colour string
  var fontSizes  = [];   // font size (pt)
  var fontWeights= [];   // "bold" | "normal"
  var hAligns    = [];   // "left" | "center" | "right"
  var vAligns    = [];   // "top" | "middle" | "bottom"
  var fontColors = [];   // foreground colour string
  var wraps      = [];   // true | false

  for (var ri = 0; ri < TOTAL_ROWS; ri++) {
    values[ri]      = [];
    backgrounds[ri] = [];
    fontSizes[ri]   = [];
    fontWeights[ri] = [];
    hAligns[ri]     = [];
    vAligns[ri]     = [];
    fontColors[ri]  = [];
    wraps[ri]       = [];
    for (var ci = 0; ci < NUM_COLS; ci++) {
      values[ri][ci]      = "";
      backgrounds[ri][ci] = null;      // null = keep default
      fontSizes[ri][ci]   = 10;
      fontWeights[ri][ci] = "normal";
      hAligns[ri][ci]     = "left";
      vAligns[ri][ci]     = "middle";
      fontColors[ri][ci]  = "#000000";
      wraps[ri][ci]       = false;
    }
  }

  // Row 0 = TITLE_ROW, Row 1 = HEADER_ROW, Rows 2…32 = data days
  var TITLE_RI  = 0;
  var HEADER_RI = 1;

  // ── Ligne de titre (sera fusionnée après le batch) ───────────────────────
  values[TITLE_RI][0]      = sheetName;
  fontSizes[TITLE_RI][0]   = 13;
  fontWeights[TITLE_RI][0] = "bold";
  hAligns[TITLE_RI][0]     = "center";
  vAligns[TITLE_RI][0]     = "middle";
  backgrounds[TITLE_RI][0] = "#FFFFFF";

  // ── Calcul des cellules par mois ─────────────────────────────────────────
  for (var m = 0; m < NUM_MONTHS; m++) {
    var mo      = months[m];
    var colDay  = m * 2;        // 0-based column index for the day number
    var colEvt  = colDay + 1;   // 0-based column index for the event text

    // En-tête de mois (HEADER_ROW, sera fusionné après le batch)
    values[HEADER_RI][colDay]      = mo.label;
    fontWeights[HEADER_RI][colDay] = "bold";
    fontSizes[HEADER_RI][colDay]   = 10;
    backgrounds[HEADER_RI][colDay] = COLOR_HEADER_BG;
    fontColors[HEADER_RI][colDay]  = COLOR_HEADER_FG;
    hAligns[HEADER_RI][colDay]     = "center";
    vAligns[HEADER_RI][colDay]     = "middle";

    var daysInMonth = new Date(mo.year, mo.month + 1, 0).getDate();

    for (var d = 1; d <= DAY_ROWS; d++) {
      var ri  = HEADER_RI + d;   // data row index (0-based)

      if (d > daysInMonth) {
        // Jour inexistant (ex. 31 dans un mois de 30 jours)
        backgrounds[ri][colDay] = COLOR_NODAY_BG;
        backgrounds[ri][colEvt] = COLOR_NODAY_BG;
        continue;
      }

      var dateObj   = new Date(mo.year, mo.month, d);
      var dateKey   = formatDateKey_(dateObj);
      var dow       = dateObj.getDay();
      var isWeekend = dow === 0 || dow === 6;
      var isGeVac   = isInRanges_(dateObj, geVacations);

      // Événements du jour filtrés selon le mode de l'onglet
      var dayEvts = (events[dateKey] || []).filter(function(e){
        return isEventVisible_(e, viewCfg);
      });

      var vacEvts    = dayEvts.filter(function(e){ return  e.isVacance; });
      var normalEvts = dayEvts.filter(function(e){ return !e.isVacance; });

      // ---- Couleur de fond (priorité : week-end > GE > école) -------------
      var isSchoolVac = vacEvts.length > 0;
      var bg = null;
      if      (isWeekend)   bg = COLOR_WEEKEND;
      else if (isGeVac)     bg = COLOR_GE_VACATION;
      else if (isSchoolVac) bg = COLOR_OWN_VACATION;

      if (bg) {
        backgrounds[ri][colDay] = bg;
        backgrounds[ri][colEvt] = bg;
      }

      // ---- Numéro du jour --------------------------------------------------
      values[ri][colDay]     = d;
      fontSizes[ri][colDay]  = 8;
      hAligns[ri][colDay]    = "left";
      vAligns[ri][colDay]    = "middle";

      // ---- Texte de l'événement --------------------------------------------
      var textParts = [];

      if (vacEvts.length > 0) {
        if (isFirstWeekdayOfVacInMonth_(dateObj, vacEvts[0], mo.month)) {
          textParts.push(vacEvts[0].displayTitle);
        }
      }
      normalEvts.forEach(function(e){ textParts.push(e.displayTitle); });

      if (textParts.length > 0) {
        var text = textParts.join(" / ");
        values[ri][colEvt]     = text;
        fontSizes[ri][colEvt]  = fontSizeForLength_(text.length);
        wraps[ri][colEvt]      = true;
        hAligns[ri][colEvt]    = "left";
        vAligns[ri][colEvt]    = "middle";
      }
    }
  }

  // ── Écriture batch dans la feuille ───────────────────────────────────────
  // Un seul bloc couvre toutes les lignes gérées (titre + en-tête + données).
  var dataRange = sheet.getRange(TITLE_ROW, 1, TOTAL_ROWS, NUM_COLS);
  dataRange.setValues(values);
  dataRange.setFontSizes(fontSizes);
  dataRange.setFontWeights(fontWeights);
  dataRange.setHorizontalAlignments(hAligns);
  dataRange.setVerticalAlignments(vAligns);
  dataRange.setFontColors(fontColors);
  dataRange.setWraps(wraps);

  // setBackgrounds ne peut pas recevoir null (contrairement aux autres setters
  // qui acceptent les valeurs par défaut). On remplace les null par la couleur
  // blanche ("") pour les cellules sans couleur spécifique, puis on applique.
  var bgClean = backgrounds.map(function(row) {
    return row.map(function(c) { return c || "white"; });
  });
  dataRange.setBackgrounds(bgClean);

  // ── Fusions (doit se faire après setValues) ──────────────────────────────
  sheet.getRange(TITLE_ROW, 1, 1, NUM_COLS).merge();
  for (var m = 0; m < NUM_MONTHS; m++) {
    sheet.getRange(HEADER_ROW, m * 2 + 1, 1, 2).merge();
  }

  // ── Bordures (une seule plage) ────────────────────────────────────────────
  sheet.getRange(DATA_START, 1, DAY_ROWS, NUM_COLS)
    .setBorder(true, true, true, true, true, true,
               "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);

  // ── Légende (une seule ligne) ─────────────────────────────────────────────
  // Légende — week-end volontairement omis : le bleu est suffisamment intuitif
  // et sa présence alourdirait la légende sans apporter d'information utile.
  var legendRow = DATA_START + DAY_ROWS + 1;
  sheet.setRowHeight(legendRow, 16);

  // "Légende" label
  sheet.getRange(legendRow, 1)
    .setValue("Légende").setFontWeight("bold").setFontSize(9);

  // Colored badge + label for each vacation type, side by side starting col 2
  [
    [COLOR_GE_VACATION,  "Vacances cantonales GE"],
    [COLOR_OWN_VACATION, "Vacances propres à l'école"]
  ].forEach(function(item, i) {
    var col = 2 + i;   // cols 2, 3
    sheet.getRange(legendRow, col)
      .setValue(item[1]).setFontSize(7)
      .setBackground(item[0])
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  });

  // "Dernière mise à jour" — timestamp of this tab's generation
  var now       = new Date();
  var pad       = function(n){ return String(n).padStart(2, "0"); };
  var timestamp = now.getFullYear() + "-" + pad(now.getMonth()+1) + "-" + pad(now.getDate()) +
                  " " + pad(now.getHours()) + ":" + pad(now.getMinutes());
  // Place it after the two badges (col 6) — leaves space regardless of NUM_MONTHS
  var tsCol = Math.max(5, NUM_COLS - 3);
  sheet.getRange(legendRow, tsCol, 1, NUM_COLS - tsCol + 1).merge()
    .setValue("Dernière mise à jour : " + timestamp)
    .setFontSize(8).setFontStyle("italic").setFontColor("#888888")
    .setHorizontalAlignment("right").setVerticalAlignment("middle");

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
        horsAnnuel   : parsed.horsAnnuel,
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
  var horsAnnuel   = false;
  var isVacance    = false;
  var levels       = [];
  var compactTitle = null;

  // #horsAnnuel
  if (/#horsAnnuel\b/i.test(desc)) horsAnnuel = true;

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

  return { horsAnnuel: horsAnnuel, isVacance: isVacance, levels: levels, displayTitle: displayTitle };
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
// PREMIER JOUR NON-WEEK-END D'UNE VACATION DANS UN MOIS DONNÉ
//
// Retourne true si dateObj est le premier jour ouvré (lundi–vendredi) de
// la vacation dans le mois courant. La vacation peut avoir démarré avant
// le 1er du mois ; dans ce cas l'effectiveStart est le 1er du mois.
// Si tous les jours restants de la vacation dans le mois sont des week-ends,
// aucun titre ne sera affiché (edge-case extrêmement rare).
// ============================================================================
function isFirstWeekdayOfVacInMonth_(dateObj, evt, currentMonth) {
  var d = new Date(dateObj); d.setHours(0,0,0,0);

  // Début effectif de la vacation dans ce mois
  var firstOfMonth = new Date(d.getFullYear(), currentMonth, 1);
  var effectiveStart = (evt.startDate > firstOfMonth) ? new Date(evt.startDate) : new Date(firstOfMonth);
  effectiveStart.setHours(0,0,0,0);

  // Avance effectiveStart jusqu'au premier jour non-week-end
  while (effectiveStart < evt.endDate) {
    var dow = effectiveStart.getDay();
    if (dow !== 0 && dow !== 6) break;          // lundi–vendredi trouvé
    effectiveStart.setDate(effectiveStart.getDate() + 1);
  }

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
// RÉCUPÈRE LES JOURS SANS ÉCOLE DU CANTON DE GENÈVE
//
// Combine deux sources depuis ge.ch :
//   1. Pages "vacances-scolaires-YYYY-YYYY" — contiennent à la fois les
//      périodes de vacances (format "du X au Y") ET des entrées mono-jour
//      comme "Jeûne genevois le JJ mois YYYY", "Fête du travail le ...",
//      "Pentecôte le ...", "Pont de l'Ascension les JJ et JJ mois YYYY".
//   2. Page "jours-feries-officiels-2023-2027" — liste des jours fériés
//      officiels (Nouvel An, Vendredi-Saint, Lundi de Pâques, Ascension,
//      Pentecôte, Fête nationale, Jeûne genevois, Noël, Restauration de la
//      République). Sert de filet de sécurité si la page vacances est
//      incomplète.
//
// Retourne un tableau de {start: Date, end: Date} où end est exclusif
// (convention : end = lendemain à minuit, même pour les mono-jour).
// ============================================================================
// yearStart : Date — début de l'année scolaire (ex. 2025-08-01).
//             Utilisé pour construire la période des grandes vacances d'été
//             d'ouverture. Hypothèse : yearStart = 01.08.YYYY (voir ci-dessous).
function fetchGeVacations_(startYear, endYear, yearStart) {
  var vacations = [];

  // ── Construit la liste de toutes les URLs à récupérer en parallèle ──────
  // Ordre : pages annuelles d'abord (indices 0..N-1), puis jours fériés (index N).
  var slugs    = [];   // slug correspondant à chaque page annuelle
  var requests = [];   // tableau passé à fetchAll

  for (var y = startYear; y < endYear; y++) {
    var slug = y + "-" + (y + 1);
    slugs.push(slug);
    requests.push({ url: GE_VACATION_BASE + slug, muteHttpExceptions: true });
  }

  var urlFeries   = "https://www.ge.ch/vacances-scolaires-jours-feries/jours-feries-officiels-2023-2027";
  var feriesIndex = requests.length;   // index de la page jours fériés dans responses
  requests.push({ url: urlFeries, muteHttpExceptions: true });

  // ── Requêtes parallèles ──────────────────────────────────────────────────
  var responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch(e) {
    Logger.log("fetchAll ge.ch échoué : " + e.message);
    return vacations;
  }

  // ── 1. Pages annuelles vacances scolaires ──────────────────────────────
  for (var i = 0; i < slugs.length; i++) {
    var slug = slugs[i];
    try {
      var resp = responses[i];
      if (resp.getResponseCode() !== 200) {
        Logger.log("ge.ch HTTP " + resp.getResponseCode() + " pour " + slug);
        continue;
      }
      var html   = resp.getContentText();
      var parsed = parseGeVacationsHtml_(html);
      Logger.log("  " + slug + " : " + parsed.length + " périodes/jours");
      vacations = vacations.concat(parsed);

      // ── Grandes vacances d'été d'OUVERTURE ──────────────────────────────
      // La page ge.ch indique la rentrée ("Rentrée scolaire le JJ mois YYYY")
      // mais pas quand commencent les vacances d'été qui la précèdent.
      // Hypothèse valide pour toutes nos années scolaires : le libellé de
      // l'année commence le 01.08 → on pose que les grandes vacances d'été
      // commencent à yearStart (01.08.YYYY) et se terminent la veille de la
      // rentrée. Si cette hypothèse change un jour, réviser ici.
      var rentree = parseRentreeScolaire_(html);
      if (rentree) {
        var summerStart = new Date(yearStart); summerStart.setHours(0,0,0,0);
        var summerEnd   = new Date(rentree);   // minuit du jour de rentrée (exclu)
        if (summerEnd > summerStart) {
          Logger.log("  Grandes vacances d'été : " +
            formatDateKey_(summerStart) + " → " + formatDateKey_(summerEnd) + " (exclu)");
          vacations.push({start: summerStart, end: summerEnd});
        }
      } else {
        Logger.log("  Rentrée scolaire introuvable pour " + slug +
          " — grandes vacances d'été non générées");
      }
    } catch(e) {
      Logger.log("Erreur ge.ch pour " + slug + " : " + e.message);
    }
  }

  // ── 2. Page jours fériés officiels (filet de sécurité) ─────────────────
  try {
    var resp2 = responses[feriesIndex];
    if (resp2.getResponseCode() === 200) {
      var parsedFeries = parseGeFeriesHtml_(resp2.getContentText(), startYear, endYear);
      Logger.log("  Jours fériés officiels : " + parsedFeries.length + " jours");
      vacations = vacations.concat(parsedFeries);
    }
  } catch(e) {
    Logger.log("Erreur ge.ch jours fériés : " + e.message);
  }

  // Déduplique par date de début (évite doublons entre les deux sources)
  vacations = deduplicateDates_(vacations);
  Logger.log("Total GE jours sans école (après dédup) : " + vacations.length);
  return vacations;
}

// ============================================================================
// PARSE LA PAGE VACANCES SCOLAIRES ANNUELLE DE GE.CH
//
// Formats capturés (le texte HTML est d'abord aplati en texte brut) :
//
//   A) Période standard :
//        "du [lundi] JJ [mois] [YYYY] au [vendredi] JJ mois YYYY"
//        ex. "du lundi 20 octobre 2025 au vendredi 24 octobre 2025"
//            "du lundi 22 décembre au vendredi 2 janvier 2026"   ← année début absente
//
//   B) Mono-jour :
//        "le [jour] JJ mois YYYY"
//        ex. "le jeudi 11 septembre 2025"  (Jeûne genevois, Fête du travail…)
//        ex. "le lundi 25 mai 2026"        (Pentecôte)
//
//   C) Deux jours consécutifs (même mois) :
//        "les [jour] JJ et [jour] JJ mois YYYY"
//        ex. "les jeudi 14 et vendredi 15 mai 2026"  (Pont de l'Ascension)
//
// Note : les grandes vacances d'été ne sont PAS parsées ici (la page indique
// seulement "dès le JJ mois" sans date de fin). Elles sont construites dans
// fetchGeVacations_() via parseRentreeScolaire_() qui lit la date de rentrée.
//
// Note sur end exclusif : new Date(y, mo-1, d+1) est sûr même en fin de mois
// ============================================================================
function parseGeVacationsHtml_(html) {
  var MONTHS_FR = {
    "janvier":1,"février":2,"mars":3,"avril":4,"mai":5,"juin":6,
    "juillet":7,"août":8,"septembre":9,"octobre":10,"novembre":11,"décembre":12
  };
  var text      = html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ");
  var vacations = [];
  var m;

  // ── Format A : périodes "du ... au ..." ──────────────────────────────
  // Groupes : 1=jour début, 2=mois début, 3=année début (optionnel),
  //           4=jour fin, 5=mois fin, 6=année fin
  var reRange = /du\s+(?:\w+\s+)?(\d{1,2})(?:er|ème|e)?\s+(\w+)\s+(\d{4}\s+)?au\s+(?:\w+\s+)?(\d{1,2})(?:er|ème|e)?\s+(\w+)\s+(\d{4})/gi;
  while ((m = reRange.exec(text)) !== null) {
    var m1 = MONTHS_FR[m[2].toLowerCase()];
    var m2 = MONTHS_FR[m[5].toLowerCase()];
    if (!m1 || !m2) continue;
    var y2 = parseInt(m[6]);
    // Si l'année de début est absente, la déduire : si mois début > mois fin → année précédente
    var y1 = m[3] ? parseInt(m[3]) : (m1 > m2 ? y2 - 1 : y2);
    var start = new Date(y1, m1-1, parseInt(m[1]));
    // new Date(y, mo, d+1) gère correctement le débordement de fin de mois
    var end   = new Date(y2, m2-1, parseInt(m[4]) + 1);
    if (!isNaN(start.getTime()) && !isNaN(end.getTime()) && end > start)
      vacations.push({start:start, end:end});
  }

  // ── Format B : mono-jour "le [jour] JJ mois YYYY" ────────────────────
  var reSingle = /\ble\s+(?:\w+\s+)?(\d{1,2})(?:er|ème|e)?\s+(\w+)\s+(\d{4})/gi;
  while ((m = reSingle.exec(text)) !== null) {
    var mo = MONTHS_FR[m[2].toLowerCase()];
    if (!mo) continue;
    var y  = parseInt(m[3]);
    var d  = parseInt(m[1]);
    var start = new Date(y, mo-1, d);
    var end   = new Date(y, mo-1, d + 1);
    if (!isNaN(start.getTime())) vacations.push({start:start, end:end});
  }

  // ── Format C : deux jours "les [jour] JJ et [jour] JJ mois YYYY" ─────
  var reTwo = /\bles\s+(?:\w+\s+)?(\d{1,2})(?:er|ème|e)?\s+et\s+(?:\w+\s+)?(\d{1,2})(?:er|ème|e)?\s+(\w+)\s+(\d{4})/gi;
  while ((m = reTwo.exec(text)) !== null) {
    var mo = MONTHS_FR[m[3].toLowerCase()];
    if (!mo) continue;
    var y  = parseInt(m[4]);
    var d1 = parseInt(m[1]), d2 = parseInt(m[2]);
    var start = new Date(y, mo-1, d1);
    var end   = new Date(y, mo-1, d2 + 1);
    if (!isNaN(start.getTime()) && end > start) vacations.push({start:start, end:end});
  }

  return vacations;
}

// ============================================================================
// EXTRAIT LA DATE DE RENTRÉE SCOLAIRE DEPUIS UNE PAGE VACANCES GE.CH
// Cherche "Rentrée scolaire le [jour] JJ mois YYYY".
// Retourne un objet Date à minuit (heure locale), ou null si non trouvé.
// ============================================================================
function parseRentreeScolaire_(html) {
  var MONTHS_FR = {
    "janvier":1,"février":2,"mars":3,"avril":4,"mai":5,"juin":6,
    "juillet":7,"août":8,"septembre":9,"octobre":10,"novembre":11,"décembre":12
  };
  var text = html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ");
  // "Rentrée scolaire le lundi 18 août 2025" ou "Rentrée scolaire le jeudi 20 août 2026"
  var re = /rentr[ée]{1,2}e\s+scolaire\s+le\s+(?:\w+\s+)?(\d{1,2})(?:er|ème|e)?\s+(\w+)\s+(\d{4})/i;
  var m  = re.exec(text);
  if (!m) return null;
  var mo = MONTHS_FR[m[2].toLowerCase()];
  if (!mo) return null;
  var d = new Date(parseInt(m[3]), mo-1, parseInt(m[1]));
  d.setHours(0,0,0,0);
  return isNaN(d.getTime()) ? null : d;
}

// ============================================================================
// PARSE LA PAGE JOURS FÉRIÉS OFFICIELS DE GE.CH (2023-2027)
// Filtre uniquement les années entre startYear et endYear (exclusif).
// Format attendu dans le texte : "NomFérié[jour] JJ mois YYYY"
// ============================================================================
function parseGeFeriesHtml_(html, startYear, endYear) {
  var MONTHS_FR = {
    "janvier":1,"février":2,"mars":3,"avril":4,"mai":5,"juin":6,
    "juillet":7,"août":8,"septembre":9,"octobre":10,"novembre":11,"décembre":12
  };
  var text      = html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ");
  var vacations = [];

  // Cherche toute date de la forme "JJ mois YYYY" dans la page
  var re = /(\d{1,2})(?:er|ème|e)?\s+(\w+)\s+(20\d{2})/gi;
  var m;
  while ((m = re.exec(text)) !== null) {
    var mo = MONTHS_FR[m[2].toLowerCase()];
    if (!mo) continue;
    var y = parseInt(m[3]);
    if (y < startYear || y >= endYear) continue;  // hors de la plage demandée
    var start = new Date(y, mo-1, parseInt(m[1]));
    var end   = new Date(y, mo-1, parseInt(m[1]) + 1);
    if (!isNaN(start)) vacations.push({start:start, end:end});
  }
  return vacations;
}

// ============================================================================
// DÉDUPLIQUE UN TABLEAU DE {start, end} PAR DATE DE DÉBUT
// Conserve la plage la plus longue pour chaque date de début identique.
// ============================================================================
function deduplicateDates_(ranges) {
  var seen = {};
  ranges.forEach(function(r) {
    var key = r.start.getTime();
    if (!seen[key] || r.end > seen[key].end) seen[key] = r;
  });
  return Object.keys(seen).map(function(k){ return seen[k]; });
}

// ============================================================================
// LIRE LA CONFIG
// ============================================================================
function readConfig_(ss) {
  var cfgSheet = ss.getSheetByName("Config");
  if (!cfgSheet) {
    SpreadsheetApp.getUi().alert('Onglet "Config" introuvable.\nUtilisez Calendrier → Initialiser l\'onglet Config.');
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

  // Années scolaires (col G=6, H=7, I=8, J=9)
  // Col J : valeur non-vide = marquer cette année pour génération
  var years = [];
  for (var r = 1; r < data.length; r++) {
    var label    = String(data[r][6] || "").trim();
    var start    = data[r][7];
    var end      = data[r][8];
    var generate = String(data[r][9] || "").trim() !== "";  // col J non-vide
    if (!label || !start || !end) continue;
    // Convertit en Date locale (minuit heure locale) quelle que soit la source :
    //   - Sheets renvoie un objet Date UTC-minuit quand la cellule est formatée
    //     en date → getFullYear/Month/Date() en heure locale = correct.
    //   - Sheets renvoie une chaîne "YYYY-MM-DD" quand la cellule est du texte
    //     → new Date("YYYY-MM-DD") = UTC minuit → en Zurich (UTC+2) getDate()
    //     donne le jour PRÉCÉDENT. On parse donc manuellement via split("-").
    var sd = parseDateLocal_(start);
    var ed = parseDateLocal_(end);
    if (!sd || !ed) continue;
    years.push({ label: label, start: sd, end: ed, generate: generate });
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
    [1,7,"Libellé année"],[1,8,"Début"],[1,9,"Fin"],[1,10,"Générer"]
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

  // Années scolaires — col J : mettre "X" sur la ligne à générer
  [
    ["2025-26","2025-08-01","2026-08-31","X"],
    ["2026-27","2026-08-01","2027-08-31",""],
  ].forEach(function(r,i){ sheet.getRange(i+2,7,1,4).setValues([r]); });

  // Largeurs de colonnes (ajout col J)
  [160,280,20,190,80,20,100,100,100,80].forEach(function(w,i){
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
    "3. Mettez « X » (ou toute valeur) dans la colonne J (Générer)\n" +
    "   de l'année souhaitée — exactement une ligne doit être marquée\n" +
    "4. Utilisez Calendrier → Générer les onglets (année marquée dans Config)"
  );
}

// ============================================================================
// PROTECTION DE L'ONGLET (warning uniquement)
// ============================================================================
function protectSheet_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(function(p){ p.remove(); });
  sheet.protect()
    .setDescription("Géré par le script — utiliser Calendrier pour mettre à jour.")
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

// ============================================================================
// PARSE UNE DATE EN HEURE LOCALE (minuit), quelle que soit la source Sheets.
//
// Problème : new Date("2025-08-01") est interprété comme UTC minuit.
// En Zurich (UTC+2), getDate() retourne 31 (juillet !). Pour éviter ce
// décalage, on parse les chaînes "YYYY-MM-DD" manuellement avec split("-"),
// ce qui construit new Date(y, m-1, d) = minuit heure locale garanti.
// Les objets Date natifs (retournés par Sheets pour les cellules date) sont
// déjà en heure locale : on extrait y/m/d avec les getters locaux.
// Retourne null si la valeur est invalide.
// ============================================================================
function parseDateLocal_(val) {
  if (!val) return null;
  if (val instanceof Date) {
    // Date native Sheets : getters locaux = correct
    var d = new Date(val.getFullYear(), val.getMonth(), val.getDate());
    return isNaN(d.getTime()) ? null : d;
  }
  var s = String(val).trim();
  // Format "YYYY-MM-DD" ou "YYYY/MM/DD"
  var parts = s.split(/[-\/]/);
  if (parts.length === 3) {
    var y = parseInt(parts[0]), mo = parseInt(parts[1]), day = parseInt(parts[2]);
    if (!isNaN(y) && !isNaN(mo) && !isNaN(day)) {
      var d = new Date(y, mo - 1, day);
      return isNaN(d.getTime()) ? null : d;
    }
  }
  return null;
}

function isInRanges_(dateObj, ranges) {
  for (var i = 0; i < ranges.length; i++) {
    if (dateObj >= ranges[i].start && dateObj < ranges[i].end) return true;
  }
  return false;
}