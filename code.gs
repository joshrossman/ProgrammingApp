// ==========================================
// GLOBAL CONFIGURATION
// ==========================================

var SPREADSHEET_ID = "1C0HT_ekmKpyXiYOAwGycyuBtwIXMo9dwFBkq2ahSH3g";
var SHEET_NAME = "Sheet1";

// ==========================================
// WEB APP ENTRY POINT
// ==========================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Tasks Matrix')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// LOCAL DATA HELPERS
// ==========================================

function getLocalData(key) {
  var props = PropertiesService.getScriptProperties();
  var data = props.getProperty(key);
  return data ? JSON.parse(data) : null;
}

function saveLocalData(key, data) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(key, JSON.stringify(data));
}

// ==========================================
// INITIAL DATA LOAD
// ==========================================

function getInitialData() {
  var eventsData = getLocalData("local_events_db") || {};
  return {
    events: eventsData,
    tasks: getLocalData("local_tasks_db") || {},
    expandedEvents: getLocalData("ui_expanded_events") || {},
    eventOrder: getLocalData("ui_event_order") || Object.keys(eventsData),
    templates: getLocalData("matrix_templates") || {},
    globalQuickTasks: getLocalData("global_quick_tasks") || [],
    links: getLocalData("important_links") || []
  };
}

// ==========================================
// UI STATE PERSISTENCE
// ==========================================

function toggleEventSyncAsync(eventId) {
  var exp = getLocalData("ui_expanded_events") || {};
  exp[eventId] = !exp[eventId];
  saveLocalData("ui_expanded_events", exp);
}

function saveEventOrderAsync(order) {
  saveLocalData("ui_event_order", order);
}

// ==========================================
// TASK CRUD OPERATIONS
// ==========================================

function toggleTaskStatusAsync(eventId, taskId) {
  var tasksData = getLocalData("local_tasks_db") || {};
  if (tasksData[eventId] && tasksData[eventId][taskId]) {
    tasksData[eventId][taskId].complete = !tasksData[eventId][taskId].complete;
    saveLocalData("local_tasks_db", tasksData);
  }
  return tasksData;
}

function updateTaskFieldAsync(eventId, taskId, field, newValue) {
  var tasksData = getLocalData("local_tasks_db") || {};
  if (tasksData[eventId] && tasksData[eventId][taskId]) {
    tasksData[eventId][taskId][field] = newValue;
    saveLocalData("local_tasks_db", tasksData);
  }
  return tasksData;
}

function saveSpecificTaskAsync(eventId, taskId, taskObj) {
  var tasksData = getLocalData("local_tasks_db") || {};
  if (!tasksData[eventId]) tasksData[eventId] = {};
  tasksData[eventId][taskId] = taskObj;
  saveLocalData("local_tasks_db", tasksData);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var eventsData = getLocalData("local_events_db") || {};
    var eventName = eventsData[eventId] ? eventsData[eventId].name : "Unknown Event";
    if (taskObj.text) sheet.appendRow([new Date(), eventName, taskObj.text, taskObj.dueDate, "Open", taskObj.details, taskObj.assignee, ""]);
  } catch (err) {}
  return tasksData;
}

function deleteSpecificTaskAsync(eventId, taskId) {
  var tasksData = getLocalData("local_tasks_db") || {};
  if (tasksData[eventId] && tasksData[eventId][taskId]) {
    delete tasksData[eventId][taskId];
    saveLocalData("local_tasks_db", tasksData);
  }
  return tasksData;
}

// ==========================================
// EVENT CRUD OPERATIONS
// ==========================================

function saveEventWithDetailsAsync(eventName, eventDate, emails, locations, quickTasks) {
  var eventsData = getLocalData("local_events_db") || {};
  var tasksData  = getLocalData("local_tasks_db")  || {};
  var uniqueEventId = "EVT-" + new Date().getTime();
  eventsData[uniqueEventId] = { name: eventName, date: eventDate, time: "TBD" };
  saveLocalData("local_events_db", eventsData);
  tasksData[uniqueEventId] = {};
  for (var i = 0; i < quickTasks.length; i++) {
    var tId = "TSK-" + new Date().getTime() + "-" + i;
    var qt = quickTasks[i];
    // qt may be a plain string or a template task object {name, subtasks}
    var taskText = (typeof qt === 'object' && qt !== null) ? (qt.name || "") : String(qt || "");
    var taskSubtasks = (typeof qt === 'object' && qt !== null && Array.isArray(qt.subtasks))
      ? qt.subtasks.map(function(st) {
          var stText = typeof st === 'string' ? st : (st && st.name ? st.name : "");
          return { text: stText, complete: false, dueDate: eventDate };
        })
      : [];
    tasksData[uniqueEventId][tId] = { text: taskText, dueDate: eventDate, complete: false, details: "", assignee: emails, subtasks: taskSubtasks };
  }
  if (locations && locations.length > 0) {
    var locTaskId = "TSK-" + new Date().getTime() + "-LOC";
    tasksData[uniqueEventId][locTaskId] = {
      text: "Assigned Locations", dueDate: eventDate, complete: false,
      details: "(" + locations.join(", ") + ")", assignee: emails,
      subtasks: locations.map(function(loc){ return { text: loc, complete: false }; })
    };
  }
  saveLocalData("local_tasks_db", tasksData);
  var order = getLocalData("ui_event_order") || [];
  order.push(uniqueEventId);
  saveLocalData("ui_event_order", order);
  return { events: eventsData, tasks: tasksData,
           expandedEvents: getLocalData("ui_expanded_events") || {},
           eventOrder: order, templates: getLocalData("matrix_templates") || {},
           globalQuickTasks: getLocalData("global_quick_tasks") || [] };
}

function updateEventNameAsync(eventId, newName) {
  var eventsData = getLocalData("local_events_db") || {};
  if (eventsData[eventId]) { eventsData[eventId].name = newName; saveLocalData("local_events_db", eventsData); }
  return eventsData;
}

function updateEventDateAsync(eventId, newDate) {
  var eventsData = getLocalData("local_events_db") || {};
  if (eventsData[eventId]) { eventsData[eventId].date = newDate; saveLocalData("local_events_db", eventsData); }
  return eventsData;
}

function updateEventTypeAsync(eventId, type) {
  var eventsData = getLocalData("local_events_db") || {};
  if (eventsData[eventId]) { eventsData[eventId].eventType = type; saveLocalData("local_events_db", eventsData); }
  return eventsData;
}

function updateEventArchiveStatusAsync(eventId, isArchived) {
  var eventsData = getLocalData("local_events_db") || {};
  if (eventsData[eventId]) { eventsData[eventId].isArchived = isArchived; saveLocalData("local_events_db", eventsData); }
  return { events: eventsData, tasks: getLocalData("local_tasks_db") || {},
           expandedEvents: getLocalData("ui_expanded_events") || {},
           eventOrder: getLocalData("ui_event_order") || Object.keys(eventsData),
           templates: getLocalData("matrix_templates") || {}, globalQuickTasks: getLocalData("global_quick_tasks") || [] };
}

function deleteEventAsync(eventId) {
  var eventsData = getLocalData("local_events_db") || {};
  if (eventsData[eventId]) { delete eventsData[eventId]; saveLocalData("local_events_db", eventsData); }
  var tasksData = getLocalData("local_tasks_db") || {};
  if (tasksData[eventId]) { delete tasksData[eventId]; saveLocalData("local_tasks_db", tasksData); }
  var order = getLocalData("ui_event_order") || [];
  var idx = order.indexOf(eventId);
  if (idx > -1) { order.splice(idx, 1); saveLocalData("ui_event_order", order); }
  var exp = getLocalData("ui_expanded_events") || {};
  if (exp[eventId] !== undefined) { delete exp[eventId]; saveLocalData("ui_expanded_events", exp); }
  return { events: eventsData, tasks: tasksData, expandedEvents: exp, eventOrder: order,
           templates: getLocalData("matrix_templates") || {}, globalQuickTasks: getLocalData("global_quick_tasks") || [] };
}

function hardDeleteEventPermanentlyAsync(eventId) {
  var cleanId = String(eventId).replace(/['"`]/g, "").trim();
  var eventsData = getLocalData("local_events_db") || {};
  if (eventsData[cleanId]) delete eventsData[cleanId];
  saveLocalData("local_events_db", eventsData);
  var tasksData = getLocalData("local_tasks_db") || {};
  if (tasksData[cleanId]) delete tasksData[cleanId];
  saveLocalData("local_tasks_db", tasksData);
  var order = getLocalData("ui_event_order") || [];
  var idx = order.indexOf(cleanId);
  if (idx > -1) { order.splice(idx, 1); saveLocalData("ui_event_order", order); }
  var exp = getLocalData("ui_expanded_events") || {};
  if (exp[cleanId] !== undefined) delete exp[cleanId];
  saveLocalData("ui_expanded_events", exp);
  return { events: eventsData, tasks: tasksData, expandedEvents: exp, eventOrder: order,
           templates: getLocalData("matrix_templates") || {}, globalQuickTasks: getLocalData("global_quick_tasks") || [] };
}

function deleteEvent(rowNumOrId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("sheet one");
  if (!sourceSheet) throw new Error("Primary event sheet not found.");
  const sheets = ss.getSheets();
  let archiveSheet = sheets.find(s => s.getName().toLowerCase() === "archives");
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet("Archives");
    archiveSheet.getRange(1,1,1,sourceSheet.getLastColumn())
      .setValues(sourceSheet.getRange(1,1,1,sourceSheet.getLastColumn()).getValues());
  }
  const targetRow = parseInt(rowNumOrId, 10);
  if (isNaN(targetRow) || targetRow <= 1 || targetRow > sourceSheet.getLastRow())
    throw new Error("Invalid row.");
  archiveSheet.appendRow(sourceSheet.getRange(targetRow,1,1,sourceSheet.getLastColumn()).getValues()[0]);
  sourceSheet.deleteRow(targetRow);
  return true;
}

// ==========================================
// TEMPLATE & QUICK TASK MANAGEMENT
// ==========================================

function updateTemplateStorageAsync(updatedTemplates) {
  saveLocalData("matrix_templates", updatedTemplates);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Matrix_Config") || ss.insertSheet("Matrix_Config");
    logSheet.getRange("A1").setValue("Master Templates Layout Array");
    logSheet.getRange("A2").setValue(JSON.stringify(updatedTemplates));
  } catch(e) { Logger.log("Template backup error: " + e.toString()); }
  return updatedTemplates;
}

function updateGlobalQuickTasksAsync(tasksArray) {
  saveLocalData("global_quick_tasks", tasksArray);
  return tasksArray;
}

// ==========================================
// SCHEDULED BACKUP
// ==========================================

function runDailySpreadsheetBackup() {
  var now = new Date(), day = now.getDay(), hour = now.getHours();
  if (day === 5 && hour >= 16) { Logger.log("Backup skipped: Friday after 4 PM."); return; }
  if (day === 6)               { Logger.log("Backup skipped: Saturday."); return; }
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("System_Backups");
    if (!sheet) {
      sheet = ss.insertSheet("System_Backups");
      sheet.appendRow(["Timestamp","Events","Tasks","Order","Templates"]);
      sheet.getRange(1,1,1,5).setFontWeight("bold");
    }
    var p = PropertiesService.getScriptProperties();
    sheet.appendRow([now,
      p.getProperty("local_events_db")  || "{}",
      p.getProperty("local_tasks_db")   || "{}",
      p.getProperty("ui_event_order")   || "[]",
      p.getProperty("matrix_templates") || "{}"]);
  } catch (err) { Logger.log("Backup error: " + err.toString()); }
}

// ==========================================
// TRIGGER MANAGEMENT
// ==========================================

// Run this ONCE manually from the Apps Script editor to set up the daily backup trigger.
// After running, you'll see it in the Triggers page (clock icon in left sidebar).
// Do NOT run it multiple times or you'll get duplicate triggers.
function setupDailyBackupTrigger() {
  // Delete any existing triggers for the backup function first to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runDailySpreadsheetBackup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  // Create a new time-based trigger that fires daily between 2–3 AM
  ScriptApp.newTrigger('runDailySpreadsheetBackup')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  Logger.log('Daily backup trigger created — runs every day at 2–3 AM.');
}

// Run this if you want to check what triggers currently exist.
function listTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    Logger.log(
      trigger.getHandlerFunction() + 
      ' — type: ' + trigger.getEventType() + 
      ' — source: ' + trigger.getTriggerSource()
    );
  });
}

// Run this to remove the daily backup trigger entirely.
function removeDailyBackupTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runDailySpreadsheetBackup') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  Logger.log('Removed ' + removed + ' backup trigger(s).');
}

var ASSIGNEE_MAP = {
  'josh.rossman@yavnehacademy.org':    'Josh',
  'shoshana.kattan@yavnehacademy.org': 'Shoshana',
  'russi.shor@yavnehacademy.org':      'Russi'
};
var ALL_ASSIGNEE_EMAILS = Object.keys(ASSIGNEE_MAP);

// Updated URL
var WEB_APP_URL = "https://script.google.com/macros/s/AKfycbw2DJsFVqEG1w4s1Njy9_Dg9y0Xoq5nAF-7Onsj8a79CIqE3a7BdO4RO6HGUq0wMEIq_g/exec";

// ── Helpers ───────────────────────────────────────────────────────────────────

function _parseDate(dateStr) {
  if (!dateStr || dateStr === "No Date" || dateStr === "TBD") return null;
  var d = new Date(String(dateStr).split('T')[0] + 'T00:00:00');
  return isNaN(d.getTime()) ? null : d;
}

function _formatDate(dateStr) {
  var d = _parseDate(dateStr);
  if (!d) return "No Date";
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function _msToIso(ms) {
  var d = new Date(parseInt(ms, 10));
  return d.getFullYear() + "-" +
    String(d.getMonth() + 1).padStart(2, '0') + "-" +
    String(d.getDate()).padStart(2, '0') + "T00:00:00";
}

function _getTaskTitle(textField) {
  if (!textField) return "Unnamed Task";
  if (typeof textField === 'object') return String(textField.name || "Unnamed Task");
  return String(textField);
}

function _getAssigneeLabel(assignee) {
  if (!assignee) return "Unassigned";
  return String(assignee).split(',').map(function(p) {
    var t = p.trim(); return ASSIGNEE_MAP[t] || t;
  }).filter(Boolean).join(', ');
}

function _urgencyIcon(dueDate, isComplete) {
  if (isComplete) return "";
  var d = _parseDate(dueDate);
  if (!d) return "";
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var diff  = Math.round((d - today) / 86400000);
  if (diff <  0)  return "🚨 ";
  if (diff <= 1)  return "⚠️ ";
  if (diff <= 7)  return "🔴 ";
  if (diff <= 14) return "🟠 ";
  return "";
}

// ── Addon persisted state ─────────────────────────────────────────────────────

// expand_mode: "all_open" | "all_closed" | "filter"
// In "filter" mode, sections auto-open when they have tasks matching the active filter.

function _getExpandMode()  { return getLocalData("addon_expand_mode") || "all_closed"; }
function _setExpandMode(m) { saveLocalData("addon_expand_mode", m); }

function _getFilters() {
  return getLocalData("addon_filters") ||
    { dueFilter: "all", typeFilter: "all", assigneeFilter: "all" };
}
function _setFilters(f) { saveLocalData("addon_filters", f); }

function _getEditState() {
  return getLocalData("addon_edit_state") ||
    { mode: false, addTasks: true, editDates: false, editAssignees: false };
}
function _setEditState(s) { saveLocalData("addon_edit_state", s); }

// Controls panel open/closed state — tracks which sub-panels are expanded.
// { controlsOpen: bool, linksOpen: bool, editOpen: bool, filtersOpen: bool }
function _getControlsState() {
  return getLocalData("addon_controls_state") || 
    { controlsOpen: false, linksOpen: false, editOpen: false, filtersOpen: true };
}
function _setControlsState(s) { saveLocalData("addon_controls_state", s); }

// ── Filter logic ──────────────────────────────────────────────────────────────
//
// KEY DESIGN: filter values are passed via action .setParameters(), NOT read
// from e.formInput. This is because:
//   1. formInput for a DROPDOWN onChange only reliably contains the widget that
//      changed, not the full card state.
//   2. Collapsible sections that are currently collapsed may omit their widgets
//      from formInput entirely.
// By encoding the selected value in parameters, we always have the right value.

function _taskPassesFilters(task, filters) {
  // Date filter — only applies to incomplete tasks
  if (filters.dueFilter !== "all" && !(task.complete === true || task.complete === "true")) {
    var d     = _parseDate(task.dueDate);
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var diff  = d ? Math.round((d - today) / 86400000) : 9999;
    if (filters.dueFilter === "overdue" && diff >= 0)  return false;
    if (filters.dueFilter === "today"   && diff !== 0) return false;
    if (filters.dueFilter === "2days"   && diff > 2)   return false;
    if (filters.dueFilter === "week"    && diff > 7)   return false;
    if (filters.dueFilter === "2weeks"  && diff > 14)  return false;
    if (filters.dueFilter === "3weeks"  && diff > 21)  return false;
    if (filters.dueFilter === "month"   && diff > 30)  return false;
  }
  // Assignee filter
  if (filters.assigneeFilter !== "all") {
    var assigneeRaw = String(task.assignee || "").trim();
    if (!assigneeRaw) return false;
    var filterVal    = filters.assigneeFilter.trim().toLowerCase();
    var assigneeParts = assigneeRaw.split(',').map(function(s){ return s.trim().toLowerCase(); });
    var matched = assigneeParts.some(function(a) {
      return a === filterVal || a.indexOf(filterVal) > -1;
    });
    if (!matched) return false;
  }
  return true;
}

// Returns true if this event has at least one incomplete task passing all filters.
function _eventHasMatchingTasks(tasksMap, filters) {
  return Object.keys(tasksMap).some(function(tId) {
    var t = tasksMap[tId];
    return !(t.complete === true || t.complete === "true") && _taskPassesFilters(t, filters);
  });
}

// Returns true if any filter that affects individual task visibility is active.
// typeFilter works at the event level (hiding whole events) so it doesn't
// trigger the auto-open behavior — only due date and assignee do.
function _anyFilterActive(filters) {
  return filters.dueFilter      !== "all" ||
         filters.assigneeFilter !== "all";
}

// ── Card builder ──────────────────────────────────────────────────────────────

function _buildCard() {
  var data        = getInitialData();
  var eventsData  = data.events || {};
  var tasksData   = data.tasks  || {};
  var filters     = _getFilters();
  var expandMode  = _getExpandMode();
  var editState   = _getEditState();

  var order = Array.isArray(data.eventOrder)
    ? data.eventOrder.slice()
    : Object.keys(eventsData);

  // All active (non-archived) events
  var activeIds = order.filter(function(id) {
    var ev = eventsData[id];
    return ev && !(ev.isArchived === true || ev.isArchived === "true");
  });

  // Apply type filter at the event level
  if (filters.typeFilter !== "all") {
    activeIds = activeIds.filter(function(id) {
      var t = eventsData[id].eventType === 'academic' ? 'academic'
            : eventsData[id].eventType === 'personal' ? 'personal'
            : 'programming';
      return t === filters.typeFilter;
    });
  }

  // ── Card header ───────────────────────────────────────────────────────────
  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setImageAltText("Tasks"));

  // ── Top section: web app link + expand/collapse all ───────────────────────
  var topSection = CardService.newCardSection();
  topSection.addWidget(
    CardService.newDecoratedText()
      .setText("🌐 Open Web App")
      .setBottomLabel("Open the full task manager")
      .setOpenLink(CardService.newOpenLink()
        .setUrl(WEB_APP_URL)
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.NOTHING))
  );
  topSection.addWidget(
    CardService.newButtonSet()
      .addButton(CardService.newTextButton().setText("Expand All")
        .setOnClickAction(CardService.newAction().setFunctionName("handleExpandAll")))
      .addButton(CardService.newTextButton().setText("Collapse All")
        .setOnClickAction(CardService.newAction().setFunctionName("handleCollapseAll")))
  );
  card.addSection(topSection);

  // ── Controls panel (outer collapsible) ───────────────────────────────────
  // Contains three sub-panels: Links, Edit Mode, Filters.
  // Each sub-panel has its own open/closed toggle button stored in ScriptProperties
  // so it persists across card rebuilds (e.g. after a filter dropdown change).
  var ctrl = _getControlsState();
  var links = getLinksData();

  // Build the full widget list for the controls panel.
  // Sub-panels use a "▶ Show / ▼ Hide" button pattern since Card Service
  // does not support nested collapsible sections.
  var ctrlWidgets = [];

  // ── Sub-panel: Important Links ───────────────────────────────────────────
  ctrlWidgets.push(
    CardService.newTextButton()
      .setText((ctrl.linksOpen ? "▼" : "▶") + "  🔗 Important Links")
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
      .setOnClickAction(CardService.newAction()
        .setFunctionName("handleToggleSubPanel").setParameters({ panel: "links" }))
  );
  if (ctrl.linksOpen) {
    if (links.length === 0) {
      ctrlWidgets.push(CardService.newTextParagraph().setText("   No links saved yet. Add them in the web app."));
    } else {
      links.forEach(function(link) {
        ctrlWidgets.push(
          CardService.newDecoratedText()
            .setText("   🔗 " + String(link.label || "Untitled"))
            .setBottomLabel("   " + String(link.url || ""))
            .setOpenLink(CardService.newOpenLink()
              .setUrl(String(link.url || "https://google.com"))
              .setOpenAs(CardService.OpenAs.FULL_SIZE)
              .setOnClose(CardService.OnClose.NOTHING))
            .setWrapText(true)
        );
      });
    }
    ctrlWidgets.push(CardService.newDivider());
  }

  // ── Sub-panel: Edit Mode ─────────────────────────────────────────────────
  ctrlWidgets.push(
    CardService.newTextButton()
      .setText((ctrl.editOpen ? "▼" : "▶") + "  ✏️ Edit Mode")
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
      .setOnClickAction(CardService.newAction()
        .setFunctionName("handleToggleSubPanel").setParameters({ panel: "edit" }))
  );
  if (ctrl.editOpen) {
    ctrlWidgets.push(CardService.newDecoratedText().setText("Enable Edit Mode")
      .setSwitchControl(CardService.newSwitch()
        .setFieldName("toggle_edit_mode").setValue("true").setSelected(editState.mode)
        .setOnChangeAction(CardService.newAction()
          .setFunctionName("handleEditStateChange").setParameters({ toggle: "mode" }))));
    if (editState.mode) {
      ctrlWidgets.push(CardService.newDecoratedText().setText("Add Tasks / Subtasks")
        .setSwitchControl(CardService.newSwitch()
          .setFieldName("toggle_edit_add").setValue("true").setSelected(editState.addTasks)
          .setOnChangeAction(CardService.newAction()
            .setFunctionName("handleEditStateChange").setParameters({ toggle: "addTasks" }))));
      ctrlWidgets.push(CardService.newDecoratedText().setText("Edit Due Dates")
        .setSwitchControl(CardService.newSwitch()
          .setFieldName("toggle_edit_dates").setValue("true").setSelected(editState.editDates)
          .setOnChangeAction(CardService.newAction()
            .setFunctionName("handleEditStateChange").setParameters({ toggle: "editDates" }))));
      ctrlWidgets.push(CardService.newDecoratedText().setText("Edit Assignees")
        .setSwitchControl(CardService.newSwitch()
          .setFieldName("toggle_edit_assignees").setValue("true").setSelected(editState.editAssignees)
          .setOnChangeAction(CardService.newAction()
            .setFunctionName("handleEditStateChange").setParameters({ toggle: "editAssignees" }))));
    }
    ctrlWidgets.push(CardService.newDivider());
  }

  // ── Sub-panel: Filters ───────────────────────────────────────────────────
  ctrlWidgets.push(
    CardService.newTextButton()
      .setText((ctrl.filtersOpen ? "▼" : "▶") + "  🔍 Filters")
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
      .setOnClickAction(CardService.newAction()
        .setFunctionName("handleToggleSubPanel").setParameters({ panel: "filters" }))
  );
  if (ctrl.filtersOpen) {
    var dueSel = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName("due_filter").setTitle("Due Date")
      .addItem("All tasks",         "all",    filters.dueFilter === "all")
      .addItem("🚨 Overdue",        "overdue",filters.dueFilter === "overdue")
      .addItem("🔴 Due today",      "today",  filters.dueFilter === "today")
      .addItem("🔴 Due in 2 days",  "2days",  filters.dueFilter === "2days")
      .addItem("⚠️ Due this week",  "week",   filters.dueFilter === "week")
      .addItem("🟠 Due in 2 weeks", "2weeks", filters.dueFilter === "2weeks")
      .addItem("🟠 Due in 3 weeks", "3weeks", filters.dueFilter === "3weeks")
      .addItem("📅 Due this month", "month",  filters.dueFilter === "month")
      .setOnChangeAction(CardService.newAction().setFunctionName("handleFilterChange")
        .setParameters({ filterType: "due" }));
    ctrlWidgets.push(dueSel);

    var typeSel = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName("type_filter").setTitle("Event Type")
      .addItem("All types",       "all",         filters.typeFilter === "all")
      .addItem("💻 Programming",  "programming", filters.typeFilter === "programming")
      .addItem("📚 Academic",     "academic",    filters.typeFilter === "academic")
      .addItem("🔒 Josh Personal","personal",    filters.typeFilter === "personal")
      .setOnChangeAction(CardService.newAction().setFunctionName("handleFilterChange")
        .setParameters({ filterType: "type" }));
    ctrlWidgets.push(typeSel);

    var asnSel = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName("assignee_filter").setTitle("Assignee")
      .addItem("All assignees", "all", filters.assigneeFilter === "all")
      .setOnChangeAction(CardService.newAction().setFunctionName("handleFilterChange")
        .setParameters({ filterType: "assignee" }));
    ALL_ASSIGNEE_EMAILS.forEach(function(email) {
      asnSel.addItem(ASSIGNEE_MAP[email], email, filters.assigneeFilter === email);
    });
    ctrlWidgets.push(asnSel);
  }

  // Outer controls section — collapsible via native Show more/less.
  // setNumUncollapsibleWidgets(0) = all hidden when collapsed.
  // setNumUncollapsibleWidgets(total) = all visible when open.
  var ctrlSection = CardService.newCardSection()
    .setHeader("⚙️ Controls")
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(ctrl.controlsOpen ? ctrlWidgets.length : 0);
  ctrlWidgets.forEach(function(w) { ctrlSection.addWidget(w); });
  card.addSection(ctrlSection);

  // ── Event sections ────────────────────────────────────────────────────────
  if (activeIds.length === 0) {
    card.addSection(CardService.newCardSection()
      .setCollapsible(true).setNumUncollapsibleWidgets(0)
      .addWidget(CardService.newTextParagraph().setText("No events match your current filters.")));
  } else {
    var anyActive = _anyFilterActive(filters);
    activeIds.forEach(function(eventId) {
      var shouldOpen;
      if (expandMode === "all_open") {
        shouldOpen = true;
      } else if (expandMode === "all_closed") {
        shouldOpen = false;
      } else {
        // "filter" mode: open this event if it has tasks matching the active filters
        shouldOpen = anyActive && _eventHasMatchingTasks(tasksData[eventId] || {}, filters);
      }
      card.addSection(_buildEventSection(
        eventId, eventsData[eventId], tasksData[eventId] || {},
        filters, shouldOpen, editState
      ));
    });
  }

  return card.build();
}

// ── Event section ─────────────────────────────────────────────────────────────

function _buildEventSection(eventId, ev, tasksMap, filters, isOpen, editState) {
  var taskIds   = Object.keys(tasksMap);
  var total     = taskIds.length;
  var done      = taskIds.filter(function(t) {
    return tasksMap[t].complete === true || tasksMap[t].complete === "true";
  }).length;
  var remaining = total - done;
  var today     = new Date(); today.setHours(0, 0, 0, 0);
  var overdue   = taskIds.filter(function(t) {
    if (tasksMap[t].complete === true || tasksMap[t].complete === "true") return false;
    var d = _parseDate(tasksMap[t].dueDate); return d && d < today;
  }).length;

  var headerLine = String(ev.name || "Unnamed") +
    "  ·  📅 " + _formatDate(ev.date) +
    "  ·  " + remaining + " remaining" +
    (overdue > 0 ? "  🚨" + overdue : "");

  var taskWidgets = _buildTaskWidgets(eventId, ev, tasksMap, filters, editState);

  var section = CardService.newCardSection()
    .setHeader(headerLine)
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(isOpen ? taskWidgets.length : 0);

  taskWidgets.forEach(function(w) { section.addWidget(w); });
  return section;
}

// ── Task widget list ──────────────────────────────────────────────────────────

function _buildTaskWidgets(eventId, ev, tasksMap, filters, editState) {
  var widgets = [];
  var taskIds = Object.keys(tasksMap);

  taskIds.sort(function(a, b) {
    var dA = _parseDate(tasksMap[a].dueDate), dB = _parseDate(tasksMap[b].dueDate);
    if (!dA && !dB) return 0; if (!dA) return 1; if (!dB) return -1; return dA - dB;
  });

  var incomplete = taskIds.filter(function(t) {
    return !(tasksMap[t].complete === true || tasksMap[t].complete === "true") &&
           _taskPassesFilters(tasksMap[t], filters);
  });
  var complete = taskIds.filter(function(t) {
    return (tasksMap[t].complete === true || tasksMap[t].complete === "true") &&
           _taskPassesFilters(tasksMap[t], filters);
  });

  if (incomplete.length === 0 && complete.length === 0) {
    widgets.push(CardService.newTextParagraph().setText("No tasks match the current filters."));
    return widgets;
  }

  if (editState.mode && editState.editDates) {
    var evDp = CardService.newDatePicker()
      .setFieldName("evdate_" + eventId).setTitle("Event Date");
    var evDate = _parseDate(ev.date);
    if (evDate) evDp.setValueInMsSinceEpoch(evDate.getTime());
    evDp.setOnChangeAction(CardService.newAction()
      .setFunctionName("handleEventDateChange").setParameters({ eventId: String(eventId) }));
    widgets.push(evDp);
  }

  incomplete.forEach(function(taskId) {
    _pushTaskWidgets(widgets, eventId, taskId, tasksMap[taskId], false, editState);
  });

  if (complete.length > 0) {
    widgets.push(CardService.newDecoratedText()
      .setText("── ✅ Completed (" + complete.length + ") ──").setWrapText(false));
    complete.forEach(function(taskId) {
      _pushTaskWidgets(widgets, eventId, taskId, tasksMap[taskId], true, editState);
    });
  }

  if (editState.mode && editState.addTasks) {
    widgets.push(CardService.newDivider());
    widgets.push(CardService.newTextInput()
      .setFieldName("new_task_text_" + eventId).setTitle("➕ New task"));
    widgets.push(CardService.newTextButton().setText("Save Task")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("handleAddTask").setParameters({ eventId: String(eventId) })));
  }

  return widgets;
}

// ── Per-task widgets ──────────────────────────────────────────────────────────

function _pushTaskWidgets(widgets, eventId, taskId, task, isComplete, editState) {
  var title    = _getTaskTitle(task.text);
  var icon     = _urgencyIcon(task.dueDate, isComplete);
  var dateLine = _parseDate(task.dueDate) ? "\nDue: " + _formatDate(task.dueDate) : "";
  var asnLine  = "\n👤 " + _getAssigneeLabel(task.assignee);
  var linkLine = task.link ? "\n🔗 " + task.link : "";
  var label    = icon + title + "\n" + dateLine + asnLine + linkLine;

  widgets.push(
    CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("task__" + eventId + "__" + taskId)
      .addItem(label, taskId, isComplete)
      .setOnChangeAction(CardService.newAction()
        .setFunctionName("handleToggleTask")
        .setParameters({ eventId: String(eventId), taskId: String(taskId) }))
  );

  // If the task has a link, show it as a tappable decorated text row
  if (task.link) {
    widgets.push(
      CardService.newDecoratedText()
        .setText("   🔗 Open Link: " + _getTaskTitle(task.text))
        .setWrapText(true)
        .setOpenLink(CardService.newOpenLink()
          .setUrl(String(task.link))
          .setOpenAs(CardService.OpenAs.FULL_SIZE)
          .setOnClose(CardService.OnClose.NOTHING))
    );
  }

  if (editState.mode && editState.editDates && !isComplete) {
    var dp = CardService.newDatePicker()
      .setFieldName("date_task_" + eventId + "_" + taskId)
      .setTitle("📅 Change date: " + title);
    var curDate = _parseDate(task.dueDate);
    if (curDate) dp.setValueInMsSinceEpoch(curDate.getTime());
    dp.setOnChangeAction(CardService.newAction()
      .setFunctionName("handleTaskDateChange")
      .setParameters({ eventId: String(eventId), taskId: String(taskId) }));
    widgets.push(dp);
  }

  if (editState.mode && editState.editAssignees) {
    var currentAssignees = String(task.assignee || "").split(',')
      .map(function(s){ return s.trim(); }).filter(Boolean);
    var asgnInput = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("asgn__" + eventId + "__" + taskId)
      .setTitle("👤 Assignees: " + title)
      .setOnChangeAction(CardService.newAction()
        .setFunctionName("handleTaskAssigneeChange")
        .setParameters({ eventId: String(eventId), taskId: String(taskId) }));
    ALL_ASSIGNEE_EMAILS.forEach(function(email) {
      asgnInput.addItem(ASSIGNEE_MAP[email], email, currentAssignees.indexOf(email) > -1);
    });
    widgets.push(asgnInput);
  }

  var subtasks = task.subtasks || [];
  subtasks.forEach(function(st, idx) {
    var stComplete = st.complete === true || st.complete === "true";
    var stTitle    = _getTaskTitle(st.text);
    var stIcon     = _urgencyIcon(st.dueDate, stComplete);
    var stDate     = _parseDate(st.dueDate) ? "\n   Due: " + _formatDate(st.dueDate) : "";
    var stLabel    = "   ↳ " + stIcon + stTitle + (stDate ? stDate : "");

    widgets.push(
      CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setFieldName("sub__" + eventId + "__" + taskId + "__" + idx)
        .addItem(stLabel, String(idx), stComplete)
        .setOnChangeAction(CardService.newAction()
          .setFunctionName("handleToggleSubtask")
          .setParameters({ eventId: String(eventId), taskId: String(taskId), subIdx: String(idx) }))
    );

    // Show nested sub-subtasks (one level deep in the add-on for readability)
    var nestedSubs = st.subtasks || [];
    nestedSubs.forEach(function(nst, nIdx) {
      var nComplete = nst.complete === true || nst.complete === "true";
      var nTitle    = _getTaskTitle(nst.text);
      var nIcon     = _urgencyIcon(nst.dueDate, nComplete);
      var nLabel    = "      ↳↳ " + nIcon + nTitle;
      widgets.push(
        CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.CHECK_BOX)
          .setFieldName("nsub__" + eventId + "__" + taskId + "__" + idx + "__" + nIdx)
          .addItem(nLabel, String(nIdx), nComplete)
          .setOnChangeAction(CardService.newAction()
            .setFunctionName("handleToggleNestedSubtask")
            .setParameters({ eventId: String(eventId), taskId: String(taskId), subIdx: String(idx), nSubIdx: String(nIdx) }))
      );
    });

    if (editState.mode && editState.editDates && !stComplete) {
      var stDp = CardService.newDatePicker()
        .setFieldName("date_sub_" + eventId + "_" + taskId + "_" + idx)
        .setTitle("📅 Change date: ↳ " + stTitle);
      var stDateObj = _parseDate(st.dueDate);
      if (stDateObj) stDp.setValueInMsSinceEpoch(stDateObj.getTime());
      stDp.setOnChangeAction(CardService.newAction()
        .setFunctionName("handleSubtaskDateChange")
        .setParameters({ eventId: String(eventId), taskId: String(taskId), subIdx: String(idx) }));
      widgets.push(stDp);
    }
  });

  if (editState.mode && editState.addTasks) {
    widgets.push(CardService.newTextInput()
      .setFieldName("new_sub_text_" + eventId + "_" + taskId)
      .setTitle("   ↳ Add subtask to: " + title));
    widgets.push(CardService.newTextButton().setText("Save Subtask")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("handleAddSubtask")
        .setParameters({ eventId: String(eventId), taskId: String(taskId) })));
  }
}

// ==========================================
// ENTRY POINT
// ==========================================

function buildGmailHomepage(e) {
  try {
    return _buildCard();
  } catch (err) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("⚠️ Error"))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText(err.message + "\n" + (err.stack || ""))))
      .build();
  }
}

// ==========================================
// ACTION HANDLERS
// ==========================================

function handleExpandAll(e) {
  _setExpandMode("all_open");
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleCollapseAll(e) {
  _setExpandMode("all_closed");
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleToggleTask(e) {
  toggleTaskStatusAsync(String(e.parameters.eventId), String(e.parameters.taskId));
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleToggleSubtask(e) {
  var eventId   = String(e.parameters.eventId);
  var taskId    = String(e.parameters.taskId);
  var subIdx    = parseInt(e.parameters.subIdx, 10);
  var tasksData = getLocalData("local_tasks_db") || {};
  if (tasksData[eventId] && tasksData[eventId][taskId] &&
      tasksData[eventId][taskId].subtasks[subIdx]) {
    tasksData[eventId][taskId].subtasks[subIdx].complete =
      !tasksData[eventId][taskId].subtasks[subIdx].complete;
    saveLocalData("local_tasks_db", tasksData);
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleToggleNestedSubtask(e) {
  var eventId  = String(e.parameters.eventId);
  var taskId   = String(e.parameters.taskId);
  var subIdx   = parseInt(e.parameters.subIdx, 10);
  var nSubIdx  = parseInt(e.parameters.nSubIdx, 10);
  var tasksData = getLocalData("local_tasks_db") || {};
  var subtask = tasksData[eventId] &&
                tasksData[eventId][taskId] &&
                tasksData[eventId][taskId].subtasks[subIdx];
  if (subtask && subtask.subtasks && subtask.subtasks[nSubIdx]) {
    subtask.subtasks[nSubIdx].complete = !subtask.subtasks[nSubIdx].complete;
    saveLocalData("local_tasks_db", tasksData);
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleEditStateChange(e) {
  var state  = _getEditState();
  var form   = e.formInput || {};
  var toggle = e.parameters.toggle;
  if (toggle === "mode")          state.mode          = form.toggle_edit_mode      === "true";
  if (toggle === "addTasks")      state.addTasks      = form.toggle_edit_add       === "true";
  if (toggle === "editDates")     state.editDates     = form.toggle_edit_dates     === "true";
  if (toggle === "editAssignees") state.editAssignees = form.toggle_edit_assignees === "true";
  _setEditState(state);
  // Keep the controls panel and edit sub-panel open after toggling
  var ctrl = _getControlsState();
  ctrl.controlsOpen = true;
  ctrl.editOpen     = true;
  _setControlsState(ctrl);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleEventDateChange(e) {
  var eventId = String(e.parameters.eventId);
  var ms      = (e.formInput || {})["evdate_" + eventId];
  if (ms && ms.msSinceEpoch) updateEventDateAsync(eventId, _msToIso(ms.msSinceEpoch));
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleTaskDateChange(e) {
  var eventId = String(e.parameters.eventId), taskId = String(e.parameters.taskId);
  var ms      = (e.formInput || {})["date_task_" + eventId + "_" + taskId];
  if (ms && ms.msSinceEpoch) updateTaskFieldAsync(eventId, taskId, "dueDate", _msToIso(ms.msSinceEpoch));
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleSubtaskDateChange(e) {
  var eventId   = String(e.parameters.eventId), taskId = String(e.parameters.taskId);
  var subIdx    = parseInt(e.parameters.subIdx, 10);
  var ms        = (e.formInput || {})["date_sub_" + eventId + "_" + taskId + "_" + subIdx];
  if (ms && ms.msSinceEpoch) {
    var tasksData = getLocalData("local_tasks_db") || {};
    if (tasksData[eventId] && tasksData[eventId][taskId] &&
        tasksData[eventId][taskId].subtasks[subIdx]) {
      tasksData[eventId][taskId].subtasks[subIdx].dueDate = _msToIso(ms.msSinceEpoch);
      saveLocalData("local_tasks_db", tasksData);
    }
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleTaskAssigneeChange(e) {
  var eventId  = String(e.parameters.eventId), taskId = String(e.parameters.taskId);
  var raw      = (e.formInput || {})["asgn__" + eventId + "__" + taskId];
  var selected = Array.isArray(raw) ? raw : (raw && raw !== "false" ? [raw] : []);
  updateTaskFieldAsync(eventId, taskId, "assignee", selected.filter(Boolean).join(", "));
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleAddTask(e) {
  var eventId  = String(e.parameters.eventId);
  var taskText = ((e.formInput || {})["new_task_text_" + eventId] || "").trim();
  if (taskText) {
    var eventsData = getLocalData("local_events_db") || {};
    var eventDate  = (eventsData[eventId] || {}).date || "No Date";
    saveSpecificTaskAsync(eventId, "TSK-" + new Date().getTime(), {
      text: taskText, dueDate: eventDate,
      complete: false, details: "", assignee: "", subtasks: []
    });
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleAddSubtask(e) {
  var eventId = String(e.parameters.eventId), taskId = String(e.parameters.taskId);
  var subText = ((e.formInput || {})["new_sub_text_" + eventId + "_" + taskId] || "").trim();
  if (subText) {
    var tasksData = getLocalData("local_tasks_db") || {};
    if (!tasksData[eventId]) tasksData[eventId] = {};
    if (!tasksData[eventId][taskId])
      tasksData[eventId][taskId] = { text: "Unnamed", dueDate: "No Date", complete: false, subtasks: [] };
    if (!tasksData[eventId][taskId].subtasks) tasksData[eventId][taskId].subtasks = [];
    var parentDate = tasksData[eventId][taskId].dueDate || "No Date";
    tasksData[eventId][taskId].subtasks.push({
      text: subText, complete: false, dueDate: parentDate
    });
    saveLocalData("local_tasks_db", tasksData);
  }
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

// ── Filter handler ────────────────────────────────────────────────────────────
//
// The selected value comes from formInput[fieldName]. Card Service DROPDOWN
// onChange sends the full card formInput, so the changed dropdown's current
// value IS present. We read it by fieldName which is always predictable.

function handleFilterChange(e) {
  // Guard: if called without a proper event object (e.g. direct editor run), bail out
  if (!e || !e.parameters) {
    Logger.log("handleFilterChange called without event parameters — use debugFilterTest() instead");
    return null;
  }

  var f    = _getFilters();
  var type = e.parameters.filterType;

  var fi  = e.formInput || {};
  var val = "all";
  if (type === "due")      val = fi["due_filter"]      || "all";
  if (type === "type")     val = fi["type_filter"]      || "all";
  if (type === "assignee") val = fi["assignee_filter"]  || "all";
  if (!val || val === "") val = "all";

  // ── DEBUG (lightweight) ───────────────────────────────────────────────────
  Logger.log("handleFilterChange: type=" + type + " val=" + val + " formInput=" + JSON.stringify(fi));
  // ── END DEBUG ─────────────────────────────────────────────────────────────

  if (type === "due")      f.dueFilter      = val;
  if (type === "type")     f.typeFilter     = val;
  if (type === "assignee") f.assigneeFilter = val;
  _setFilters(f);

  Logger.log("  filters AFTER    : " + JSON.stringify(f));

  // Keep the controls panel open and filters sub-panel open after a filter change
  var ctrl = _getControlsState();
  ctrl.controlsOpen = true;
  ctrl.filtersOpen  = true;
  _setControlsState(ctrl);

  // Only switch expand mode for task-level filters (due date, assignee).
  // Type filter hides whole events but doesn't auto-open sections.
  if (type === "due" || type === "assignee") {
    if (_anyFilterActive(f)) {
      _setExpandMode("filter");
    } else {
      _setExpandMode("all_closed");
    }
  }

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

function handleAddOnTaskToggle(e) { return handleToggleTask(e); }

// Toggle a sub-panel (links / edit / filters) open or closed.
// Also ensures the outer controls panel stays open when toggling a sub-panel.
function handleToggleSubPanel(e) {
  var panel = e.parameters.panel;
  var ctrl  = _getControlsState();
  ctrl.controlsOpen = true;  // keep outer panel open
  if (panel === "links")   ctrl.linksOpen   = !ctrl.linksOpen;
  if (panel === "edit")    ctrl.editOpen    = !ctrl.editOpen;
  if (panel === "filters") ctrl.filtersOpen = !ctrl.filtersOpen;
  _setControlsState(ctrl);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildCard()))
    .setStateChanged(true).build();
}

// ==========================================
// DEBUG — run this directly from the editor
// Select debugFilterTest in the dropdown and click Run.
// Output appears in the Execution Log panel at the bottom.
// ==========================================

function debugFilterTest() {
  var data      = getInitialData();
  var eventsData = data.events || {};
  var tasksData  = data.tasks  || {};

  Logger.log("========================================");
  Logger.log("DEBUG: All stored task assignee values");
  Logger.log("========================================");

  Object.keys(tasksData).forEach(function(evId) {
    var evName = (eventsData[evId] || {}).name || evId;
    var taskMap = tasksData[evId] || {};
    Object.keys(taskMap).forEach(function(tId) {
      var t = taskMap[tId];
      Logger.log("Event: " + evName +
                 " | Task: " + _getTaskTitle(t.text) +
                 " | assignee (raw): [" + t.assignee + "]" +
                 " | complete: " + t.complete);
    });
  });

  Logger.log("========================================");
  Logger.log("DEBUG: Filter test for each ASSIGNEE_MAP email");
  Logger.log("========================================");

  var emails = Object.keys(ASSIGNEE_MAP);
  emails.forEach(function(email) {
    var filters = { dueFilter: "all", typeFilter: "all", assigneeFilter: email };
    Logger.log("--- Filtering for: " + email + " (" + ASSIGNEE_MAP[email] + ") ---");
    var matchCount = 0;
    Object.keys(tasksData).forEach(function(evId) {
      var evName = (eventsData[evId] || {}).name || evId;
      Object.keys(tasksData[evId] || {}).forEach(function(tId) {
        var t = tasksData[evId][tId];
        var passes = _taskPassesFilters(t, filters);
        if (passes) {
          matchCount++;
          Logger.log("  PASS → [" + evName + "] " + _getTaskTitle(t.text) +
                     " | assignee: [" + t.assignee + "]");
        }
      });
    });
    if (matchCount === 0) Logger.log("  (no tasks matched for " + email + ")");
  });

  Logger.log("========================================");
  Logger.log("DEBUG: Current saved filters in ScriptProperties");
  Logger.log("========================================");
  Logger.log(JSON.stringify(_getFilters(), null, 2));
}

// ==========================================
// IMPORTANT LINKS — CRUD
// ==========================================

function getLinksData() {
  return getLocalData("important_links") || [];
}

function saveLinksData(links) {
  saveLocalData("important_links", links);
  return links;
}
