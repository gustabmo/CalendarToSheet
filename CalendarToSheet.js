// 2026-05-02 CalendarToSheet.js(.gs) by Gustavo Exel and claude.ai - first version
//
// Goal: to creat several different versions of yearly calendars based on events that 
// are on google calendars Public / Parents / Profs
//
// =============================================================================
// CALENDRIER SCOLAIRE — Google Apps Script
// À attacher à un fichier Google Sheets.
// Menu : Calendrier
//
// Onglets générés (script-managed, ne pas modifier manuellement) :
//   "[année] Public"            — calendrier public
//   "[année] Parents"           — public + parents, non-tagués ou #general
//   "[année] Parents tout"      — public + parents, tout
//   "[année] Parents JE"        — public + parents, #jardindenfants
//   "[année] Parents Prim"      — public + parents, #primaire
//   "[année] Parents Sec1"      — public + parents, #secondaire1
//   "[année] Parents Sec2"      — public + parents, #secondaire2
//   "[année] Profs"             — public + parents + profs, non-tagués ou #general
//   "[année] Profs tout"        — public + parents + profs, tout
//   "[année] Profs JE"          — public + parents + profs, #jardindenfants
//   "[année] Profs Prim"        — public + parents + profs, #primaire
//   "[année] Profs Sec1"        — public + parents + profs, #secondaire1
//   "[année] Profs Sec2"        — public + parents + profs, #secondaire2
// =============================================================================
//
// ================================================
// CALENDRIER SCOLAIRE — notes pour session future
//
// Contexte : Google Apps Script attaché à un Google Sheets.
// École à Genève (francophone). 3 calendriers : public, parents, professeurs.
// Niveaux : jardindenfants, primaire, secondaire1, secondaire2.
//
// Tags dans le champ "description" des événements Google Calendar, function parseDescription_() :
//   #vacances           → colore la cellule en COLOR_OWN_VACATION
//   #vacancesDIP[:dd.mm.yyyy-dd.mm.yyyy] → colore la cellule en COLOR_GE_VACATION
//   #compact:texte      → titre court pour la grille
//   #jardindenfants / #primaire / #secondaire1 / #secondaire2  → filtre par niveau
//   Pas de tag niveau   → voir "non-tagués" sur la description des tabs
//   #general            → a un tag de niveau mais apparait quand même avec les non-tagués
//   #horsAnnuel         → n'apparait pas sur ces calendriers
//   #multiday
//   
//
// ── RÈGLE DES COULEURS DE FOND (priorité décroissante) ───────────────────
//   1. BLEU        (COLOR_WEEKEND)      — samedi et dimanche
//   2. ORANGE FONCÉ (COLOR_GE_VACATION) — jours sans école dans le calendrier
//                                         officiel DIP / Canton de Genève
//                                         voir #vacancesDIP
//   3. ORANGE CLAIR (COLOR_OWN_VACATION) — jours de congé supplémentaires propres
//                                          à notre école (marqués #vacances dans
//                                          le calendrier Google "parents" ou "public")
//                                          Notre école est un SUPERSET du calendrier
//                                          GE : tous les jours GE sont off + quelques
//                                          jours additionnels propres à notre école.
// ── FIN RÈGLE COULEURS ───────────────────────────────────────────────────
// ================================================
//
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
//   FETCH_GE_VACANCES       TRUE  ou  FALSE ... caduque!
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

// ---- Couleurs des pistes multiday (6 teintes de vert) ----------------------
var TRACK_COLORS = [
  "#2E7D32",  // 1 vert foncé
  "#66BB6A",  // 2 vert moyen
  "#26A69A",  // 3 vert-sarcelle
  "#AED581",  // 4 vert-jaune
  "#80DEEA",  // 5 cyan clair
  "#C8E6C9"   // 6 vert très clair
];

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

// ============================================================================
// MENU
// ============================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Calendrier")
    .addItem("Générer les onglets (année marquée dans Config)", "generateMarkedYear")
    // .addSeparator()
    // .addItem("Initialiser l'onglet Config", "setupConfigSheet")
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
  // Collecte tous les événements des 3 calendriers
  var calPublic = collectCalendarEvents_(config.calPublicId,  year.start, year.end);
  var calParents= collectCalendarEvents_(config.calParentsId, year.start, year.end);
  var calProfs  = collectCalendarEvents_(config.calProfsId,   year.start, year.end);

  // Plages GE extraites des tags #vacancesDIP:dd.mm.yyyy-dd.mm.yyyy dans les événements
  var geVacations = extractGeVacationsFromEvents_([calPublic, calParents, calProfs]);
  Logger.log("Plages GE extraites des calendriers : " + geVacations.length + " périodes");

  var label = year.label;

  // ── "[année] Public"            — calendrier public
  generateSheet_(ss, label + " Public", year, geVacations,
    mergeEventMaps_([calPublic]),
    { mode: "public" });

  // "[année] Parents"           — public + parents, non-tagués ou #general
  generateSheet_(ss, label + " Parents", year, geVacations,
    mergeEventMaps_([calPublic, calParents]),
    { mode: "parents-général" });

  // "[année] Parents tout"      — public + parents, tout
  generateSheet_(ss, label + " Parents tout", year, geVacations,
    mergeEventMaps_([calPublic, calParents]),
    { mode: "parents-tout" });

  // "[année] Profs"             — public + parents + profs, non-tagués ou #general
  generateSheet_(ss, label + " Profs", year, geVacations,
    mergeEventMaps_([calPublic, calParents, calProfs]),
    { mode: "profs-général" });

  // "[année] Profs tout"        — public + parents + profs, tout
  generateSheet_(ss, label + " Profs tout", year, geVacations,
    mergeEventMaps_([calPublic, calParents, calProfs]),
    { mode: "profs-tout" });

  // "[année] Parents [niveau]"        — public + parents, #[niveau]
  LEVEL_KEYS.forEach(function(levelKey) {
    var suffix = LEVELS[levelKey];
    generateSheet_(ss, label + " Parents " + suffix, year, geVacations,
      mergeEventMaps_([calPublic, calParents]),
      { mode: "parents-niveau", level: levelKey });
  });

  // "[année] Profs [niveau]"          — public + parents + profs, #[niveau]
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
// Retourne : false = invisible | true = visible
// ============================================================================
function isEventVisible_(evt, viewCfg) {
  if (evt.horsAnnuel) return false;

  var hasLevelTag = evt.levels.length > 0;

  switch (viewCfg.mode) {
    case "public":
    case "parents-tout":
    case "profs-tout":
      return true;

    case "parents-général":
    case "profs-général":
      // Uniquement les événements sans aucun tag de niveau
      return !hasLevelTag || evt.isGeneral;

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
// UTILITAIRES ÉVÉNEMENTS MULTIDAY
// ============================================================================
function getEventKey_(evt) {
  return evt.startDate.getTime() + '|' + evt.endDate.getTime() + '|' + evt.displayTitle;
}

function collectMultidayEvents_(eventsMap, viewCfg) {
  var seen = {};
  var result = [];
  Object.keys(eventsMap).forEach(function(dateKey) {
    eventsMap[dateKey].forEach(function(evt) {
      if (!evt.isMultiday || evt.isVacance) return;
      if (!isEventVisible_(evt, viewCfg)) return;
      var key = getEventKey_(evt);
      if (!seen[key]) { seen[key] = true; result.push(evt); }
    });
  });
  result.sort(function(a, b) { return a.startDate - b.startDate; });
  return result;
}

function buildMultidayColorMap_(multidayEvts) {
  var map = {};
  multidayEvts.forEach(function(evt, i) {
    map[getEventKey_(evt)] = TRACK_COLORS[i % TRACK_COLORS.length];
  });
  return map;
}

// Affecte une piste (1-indexed) à chaque événement multiday actif dans le mois.
// Tri par date de début globale → les événements démarrés plus tôt occupent
// les pistes les plus à gauche ; les pistes se réaffectent à chaque nouveau mois.
function assignTracksForMonth_(multidayEvts, monthStart, nextMonthStart) {
  var active = multidayEvts.filter(function(evt) {
    return evt.startDate < nextMonthStart && evt.endDate > monthStart;
  });
  var trackEnds   = [];   // trackEnds[t] = fin de l'événement occupant la piste t
  var assignments = {};
  active.forEach(function(evt) {
    var effStart = evt.startDate > monthStart ? evt.startDate : monthStart;
    var effEnd   = evt.endDate   < nextMonthStart ? evt.endDate : nextMonthStart;
    var assigned = -1;
    for (var t = 0; t < trackEnds.length; t++) {
      if (trackEnds[t] <= effStart) { trackEnds[t] = effEnd; assigned = t + 1; break; }
    }
    if (assigned === -1) { trackEnds.push(effEnd); assigned = trackEnds.length; }
    assignments[getEventKey_(evt)] = assigned;
  });
  return assignments;
}

// ============================================================================
function generateSheet_(ss, sheetName, year, geVacations, events, viewCfg) {
  // ── Récupère ou crée l'onglet sans le supprimer ──────────────────────────
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    // Efface contenu, formats et fusions pour repartir à zéro
    // (breakApart sur toute la plage pour dissoudre les cellules fusionnées
    //  avant clearFormats, qui échouerait sur des plages fusionnées)
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clearContents();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var months     = buildMonthList_(year.start, year.end);
  var NUM_MONTHS = months.length;
  var EVT_COLS       = 6;
  var COLS_PER_MONTH = 1 + EVT_COLS;
  var NUM_COLS       = NUM_MONTHS * COLS_PER_MONTH;
  var DAY_ROWS   = 31;
  var TITLE_ROW  = 1;   // 1-indexed for Sheets API
  var HEADER_ROW = 2;
  var DATA_START = 3;
  // Total rows in our managed area: title(1) + header(1) + data(31)
  var TOTAL_ROWS = 1 + 1 + DAY_ROWS;

  // ── Redimensionne la feuille si nécessaire ───────────────────────────────
  var currentCols = sheet.getMaxColumns();
  if (currentCols < NUM_COLS) {
    sheet.insertColumnsAfter(currentCols, NUM_COLS - currentCols);
  } else if (currentCols > NUM_COLS) {
    sheet.deleteColumns(NUM_COLS + 1, currentCols - NUM_COLS);
  }

  // ── Hauteurs de lignes (peu d'appels, pas de gain à batcher) ────────────
  sheet.setRowHeight(TITLE_ROW,  45);
  sheet.setRowHeight(HEADER_ROW, 32);
  for (var r = DATA_START; r < DATA_START + DAY_ROWS; r++) {
    sheet.setRowHeight(r, 29);
  }

  // ── Largeurs de colonnes ─────────────────────────────────────────────────
  var evtColWidths = [15, 15, 15, 15, 14, 14]; // 6 cols summing to 88px
  for (var m = 0; m < NUM_MONTHS; m++) {
    sheet.setColumnWidth(m * COLS_PER_MONTH + 1, 26);
    for (var ec = 0; ec < EVT_COLS; ec++) {
      sheet.setColumnWidth(m * COLS_PER_MONTH + 2 + ec, evtColWidths[ec]);
    }
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

  // ── Prétraitement des événements multiday ────────────────────────────────
  var multidayEvts     = collectMultidayEvents_(events, viewCfg);
  var multidayColorMap = buildMultidayColorMap_(multidayEvts);
  var monthTrackCounts        = [];  // [m][d0] = nb de pistes actives
  var monthDayLastTrackExpand = [];  // [m][d0] = nb colonnes empruntées par la dernière piste si elle démarre

  // ── Calcul des cellules par mois ─────────────────────────────────────────
  for (var m = 0; m < NUM_MONTHS; m++) {
    var mo      = months[m];
    var colDay  = m * COLS_PER_MONTH;   // 0-based column index for the day number
    var colEvt  = colDay + 1;           // 0-based column index for the first event column

    var monthStart     = new Date(mo.year, mo.month, 1);
    var nextMonthStart = new Date(mo.year, mo.month + 1, 1);
    var monthTracks    = assignTracksForMonth_(multidayEvts, monthStart, nextMonthStart);
    monthTrackCounts[m]        = new Array(DAY_ROWS).fill(0);
    monthDayLastTrackExpand[m] = new Array(DAY_ROWS).fill(0);

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
        for (var ec = 0; ec < EVT_COLS; ec++) {
          backgrounds[ri][colEvt + ec] = COLOR_NODAY_BG;
        }
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

      var vacEvts         = dayEvts.filter(function(e){ return  e.isVacance; });
      var normalEvts      = dayEvts.filter(function(e){ return !e.isVacance && !e.isMultiday; });
      var multidayDayEvts = dayEvts.filter(function(e){ return  e.isMultiday && !e.isVacance; });

      // ---- Couleur de fond (priorité : week-end > GE > école) -------------
      var isDipVacEvt = vacEvts.some(function(e){ return e.isVacanceDIP; });
      var isSchoolVac = vacEvts.some(function(e){ return !e.isVacanceDIP; });
      var bg = null;
      if      (isWeekend)              bg = COLOR_WEEKEND;
      else if (isGeVac || isDipVacEvt) bg = COLOR_GE_VACATION;
      else if (isSchoolVac)            bg = COLOR_OWN_VACATION;

      if (bg) {
        backgrounds[ri][colDay] = bg;
        for (var ec = 0; ec < EVT_COLS; ec++) {
          backgrounds[ri][colEvt + ec] = bg;
        }
      }

      // ---- Numéro du jour --------------------------------------------------
      values[ri][colDay]     = d;
      fontSizes[ri][colDay]  = 8;
      hAligns[ri][colDay]    = "left";
      vAligns[ri][colDay]    = "middle";

      // ---- Pistes multiday actives ce jour ------------------------------------
      var activeTracks   = {};
      var maxActiveTrack = 0;
      multidayDayEvts.forEach(function(e) {
        var key   = getEventKey_(e);
        var track = monthTracks[key];
        if (track && track <= EVT_COLS) {
          activeTracks[track] = e;
          if (track > maxActiveTrack) maxActiveTrack = track;
        }
      });
      monthTrackCounts[m][d - 1] = maxActiveTrack;

      // ---- Pistes multiday : couleur + texte (le texte reste dans la cellule colorée)
      for (var t = 1; t <= maxActiveTrack; t++) {
        var tColIdx = colEvt + t - 1;
        var tEvt    = activeTracks[t];
        if (tEvt) {
          backgrounds[ri][tColIdx] = multidayColorMap[getEventKey_(tEvt)];
          if (isFirstWeekdayOfVacInMonth_(dateObj, tEvt, mo.month)) {
            values[ri][tColIdx]    = tEvt.displayTitle;
            fontSizes[ri][tColIdx] = fontSizeForLength_(tEvt.displayTitle.length);
            wraps[ri][tColIdx]     = true;
            hAligns[ri][tColIdx]   = "left";
            vAligns[ri][tColIdx]   = "middle";
          }
        }
        // gap track : le fond par défaut déjà posé reste intact
      }

      // ---- Zone restante : événements mono-jour, + extension de la dernière piste si elle démarre
      var remaining  = EVT_COLS - maxActiveTrack;
      var remColBase = colEvt + maxActiveTrack;   // 0-based col index of first remaining column

      var lastEvt        = maxActiveTrack > 0 ? activeTracks[maxActiveTrack] : null;
      var lastIsStarting = lastEvt != null && isFirstWeekdayOfVacInMonth_(dateObj, lastEvt, mo.month);

      var textParts = [];
      if (vacEvts.length > 0 && isFirstWeekdayOfVacInMonth_(dateObj, vacEvts[0], mo.month)) {
        textParts.push(vacEvts[0].displayTitle);
      }
      normalEvts.forEach(function(e){ textParts.push(e.displayTitle); });
      var hasText = textParts.length > 0;

      // La dernière piste emprunte une fraction de la zone restante pour élargir sa cellule
      var lastTrackExpand = 0;
      if (lastIsStarting && remaining > 0) {
        var numCompeting = 1 + (hasText ? 1 : 0);
        lastTrackExpand  = Math.floor(remaining / numCompeting);
        var expandColor  = multidayColorMap[getEventKey_(lastEvt)];
        for (var bc = 0; bc < lastTrackExpand; bc++) {
          backgrounds[ri][remColBase + bc] = expandColor;
        }
      }
      monthDayLastTrackExpand[m][d - 1] = lastTrackExpand;

      if (hasText) {
        var normalColStart = remColBase + lastTrackExpand;
        var text = textParts.join(" / ");
        values[ri][normalColStart]      = text;
        fontSizes[ri][normalColStart]   = fontSizeForLength_(text.length);
        wraps[ri][normalColStart]       = true;
        hAligns[ri][normalColStart]     = "left";
        vAligns[ri][normalColStart]     = "middle";
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
    sheet.getRange(HEADER_ROW, m * COLS_PER_MONTH + 1, 1, COLS_PER_MONTH).merge();
    var d = 0;
    while (d < DAY_ROWS) {
      var maxTrack = monthTrackCounts[m][d] || 0;
      var remaining = EVT_COLS - maxTrack;
      var expand   = monthDayLastTrackExpand[m][d] || 0;

      if (expand > 0) {
        // La dernière piste fusionne avec les colonnes empruntées (1 seule ligne)
        // 1-based col of last track cell: m*COLS_PER_MONTH + maxTrack + 1
        var trackCol1  = m * COLS_PER_MONTH + maxTrack + 1;
        var mergedCols = 1 + expand;
        sheet.getRange(DATA_START + d, trackCol1, 1, mergedCols).merge();
        // Zone restante après l'emprunt → événements normaux
        var normalCols = remaining - expand;
        if (normalCols >= 2) {
          sheet.getRange(DATA_START + d, trackCol1 + mergedCols, 1, normalCols).merge();
        }
        d++;
      } else {
        // Regroupe les lignes consécutives avec le même maxTrack et pas d'extension
        var runEnd = d + 1;
        while (runEnd < DAY_ROWS &&
               (monthTrackCounts[m][runEnd] || 0) === maxTrack &&
               (monthDayLastTrackExpand[m][runEnd] || 0) === 0) {
          runEnd++;
        }
        if (remaining >= 2) {
          sheet.getRange(DATA_START + d, m * COLS_PER_MONTH + maxTrack + 2, runEnd - d, remaining)
               .mergeAcross();
        }
        d = runEnd;
      }
    }
  }

  // ── Bordures (une seule plage) ────────────────────────────────────────────
  sheet.getRange(DATA_START, 1, DAY_ROWS, NUM_COLS)
    .setBorder(true, true, true, true, true, true,
               "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_THIN);

  // ── Légende (une seule ligne) ─────────────────────────────────────────────
  // Légende — week-end volontairement omis : le bleu est suffisamment intuitif
  // et sa présence alourdirait la légende sans apporter d'information utile.
  var legendRow = DATA_START + DAY_ROWS + 1;

  // saut de ligne avant la légende pour aérer un peu
  sheet.setRowHeight(legendRow - 1, 10);
  sheet.getRange(legendRow - 1, 1, 1, NUM_COLS)
    .merge();

  sheet.setRowHeight(legendRow, 25);

  let nextCol=1

  // "Légende" label
  sheet.getRange(legendRow, nextCol, 1, COLS_PER_MONTH)
    .merge()
    .setValue("Légende :").setFontWeight("bold").setFontSize(9);
  nextCol += COLS_PER_MONTH;

  // Colored badge + label for each vacation type, side by side starting col 2
  [
    [COLOR_GE_VACATION,  "Vacances cantonales GE"],
    [COLOR_OWN_VACATION, "Vacances propres à l'école"]
  ].forEach(function(item, i) {
    sheet.getRange(legendRow, nextCol, 1, COLS_PER_MONTH )
      .merge()
      .setBackground(item[0])
      .setValue(item[1]).setFontSize(7)
      .setVerticalAlignment("middle");
    nextCol += COLS_PER_MONTH;
  });

  sheet.getRange(legendRow, nextCol, 1, COLS_PER_MONTH )
    .merge();
  nextCol += COLS_PER_MONTH;

  // "Dernière mise à jour" — timestamp of this tab's generation
  var now       = new Date();
  var pad       = function(n){ return String(n).padStart(2, "0"); };
  var timestamp =  pad(now.getDate()) + "-" + pad(now.getMonth()+1) + "-" + now.getFullYear() +
                  " " + pad(now.getHours()) + ":" + pad(now.getMinutes());
  sheet.getRange(legendRow, nextCol,1,COLS_PER_MONTH*2)
    .merge()
    .setValue("Dernière mise à jour : " + timestamp)
    .setFontSize(8).setFontStyle("italic").setFontColor("#888888")
    .setVerticalAlignment("middle");
  nextCol += COLS_PER_MONTH*2;

  sheet.getRange(legendRow, nextCol, 1, NUM_COLS-nextCol+1 )
    .merge();


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
        isVacanceDIP : parsed.isVacanceDIP,
        geDates      : parsed.geDates,
        isMultiday   : parsed.isMultiday,
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
// Tags reconnus : voire Tags dans le champ "description" des événements Google Calendar : 
// ============================================================================
function parseDescription_(desc, rawTitle) {
  var horsAnnuel   = false;
  var isVacance    = false;
  var isVacanceDIP = false;
  var geDates      = null;
  var levels       = [];
  var compactTitle = null;

  // #horsAnnuel
  if (/#horsAnnuel\b/i.test(desc)) horsAnnuel = true;

  // #vacancesDIP:dd.mm.yyyy-dd.mm.yyyy  →  OWN_VACATION + plage GE aux dates du tag
  var dipRangeMatch = desc.match(/#vacancesDIP:(\d{2}\.\d{2}\.\d{4})-(\d{2}\.\d{2}\.\d{4})/i);
  if (dipRangeMatch) {
    isVacance = true;
    var parseDmy = function(s) {
      var p = s.split('.'); return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    };
    var geStart = parseDmy(dipRangeMatch[1]);
    var geEnd   = parseDmy(dipRangeMatch[2]);
    geEnd.setDate(geEnd.getDate() + 1);  // fin exclusive
    geDates = { start: geStart, end: geEnd };
  } else if (/#vacancesDIP\b/i.test(desc)) {
    // #vacancesDIP seul → GE_VACATION
    isVacance    = true;
    isVacanceDIP = true;
  } else if (/#vacances\b/i.test(desc)) {
    // #vacances → OWN_VACATION
    isVacance = true;
  }

  // #multiday
  var isMultiday = /#multiday\b/i.test(desc);

  // #multiday
  var isGeneral = /#g[eé]n[eé]ral\b/i.test(desc);

  // #compact:texte court
  var compactMatch = desc.match(/#compact:([^#\n\r]+)/i);
  if (compactMatch) compactTitle = compactMatch[1].trim();

  // Tags de niveaux
  LEVEL_KEYS.forEach(function(k) {
    if (new RegExp("#" + k + "\\b", "i").test(desc)) levels.push(k);
  });

  var displayTitle = compactTitle || rawTitle;

  return { 
    horsAnnuel: horsAnnuel,  isMultiday: isMultiday,
    isVacance: isVacance,  isVacanceDIP: isVacanceDIP,  geDates: geDates, 
    levels: levels,  isGeneral: isGeneral, 
    displayTitle: displayTitle 
  };
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
// EXTRAIT LES PLAGES GE_VACATION DES TAGS #vacancesDIP:dd.mm.yyyy-dd.mm.yyyy
// ============================================================================
function extractGeVacationsFromEvents_(eventMaps) {
  var ranges = [];
  eventMaps.forEach(function(map) {
    Object.keys(map).forEach(function(dateKey) {
      map[dateKey].forEach(function(evt) {
        if (evt.geDates) ranges.push(evt.geDates);
      });
    });
  });
  return deduplicateDates_(ranges);
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