/**
 * Ami 4 Mayor Canvassing App -- Google Apps Script backend.
 *
 * Bind this script to a Google Sheet (Extensions > Apps Script from within
 * the Sheet). It creates/uses five tabs:
 *
 *   Users      | Username | PasswordHash | DisplayName | Active | Role |
 *   Precincts  | PrecinctID | Status | CompletedDate | UpdatedBy | UpdatedAt | Notes |
 *   Parcels    | Address | Walked | UpdatedBy | UpdatedAt |
 *   Settings   | Key | Value | UpdatedBy | UpdatedAt |
 *   Presence   | Username | DisplayName | LastSeenAt |
 *
 * Presence (who's currently logged in, shown in the page's topbar) is kept
 * up to date by piggybacking on the "all"/"since" polls the page already
 * makes every 20 seconds (see recordHeartbeatAndGetOnline() below) rather
 * than a separate heartbeat request -- one row per username, upserted on
 * every poll, so the sheet stays bounded by how many people have ever
 * logged in rather than growing with every poll.
 *
 * Notes was added as a trailing 6th column (rather than inserted next to
 * CompletedDate) so an already-deployed Sheet with the original 5 columns
 * keeps working with no manual migration -- existing rows just read back
 * with an empty Notes value until edited. Role works the same way: it's a
 * trailing 5th Users column, so a Sheet from before roles existed just
 * reads back as "editor" (normalizeRole()'s default) for every existing
 * login until an admin sets some to "viewer".
 *
 * Roles: "editor" (default) can set precinct status, mark parcels walked,
 * write notes, and change Map Settings -- i.e. everything a canvasser could
 * always do. "viewer" can log in and look at precinct panels, but the
 * backend itself rejects any save/settings-change attempt from a viewer's
 * username (see handleSaveState()/handleSaveSettings()) -- this is enforced
 * here, not just hidden in the app's UI, since the UI alone can't be
 * trusted to be the only thing calling this script.
 *
 * Deploy: Deploy > New deployment > type "Web app" >
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the resulting /exec URL into CONFIG.APPS_SCRIPT_URL in index.html.
 *
 * Admin logins are managed from the Sheet itself via the "Canvassing Admin"
 * menu this script adds (Add/Update Login, Remove Login) -- no separate
 * password file needed, and no plaintext passwords are ever stored: only a
 * SHA-256 hash, matching the hash the webpage computes client-side.
 *
 * NOTE: an earlier version of this app tracked per-address support status,
 * notes, yard signs, and volunteers in a "Data" tab. That workflow was
 * replaced by precinct-level print/walk status plus a per-parcel "Walked"
 * flag. If your Sheet still has the old "Data" tab, it's left untouched --
 * this script just no longer reads or writes it.
 */

var USERS_SHEET_NAME = "Users";
var PRECINCTS_SHEET_NAME = "Precincts";
var PARCELS_SHEET_NAME = "Parcels";
var SETTINGS_SHEET_NAME = "Settings";
var PRESENCE_SHEET_NAME = "Presence";
var SETTINGS_KEY = "mapSettings"; // single row in the Settings sheet holds the whole settings object as JSON
var VALID_PRECINCT_STATUSES = ["Not Yet Printed", "Printed", "Partially Walked", "Completed"];
// How stale a Presence row can be and still count as "online" -- comfortably
// more than the page's own 20-second sync poll (SYNC_POLL_MS in index.html,
// which is what actually drives how often a heartbeat arrives) so one missed
// or slow poll doesn't make someone flicker offline and back.
var ONLINE_WINDOW_MS = 90 * 1000;

// ================== HTTP ENTRY POINTS ==================

function doGet(e) {
  var action = e.parameter.action;
  try {
    if (action === "all") {
      var settingsRow = getSettingsRow();
      var out = {
        ok: true,
        precincts: getAllPrecincts(),
        parcels: getAllParcels(),
        serverTime: new Date().toISOString()
      };
      if (settingsRow) {
        out.settings = settingsRow.value;
        out.settingsUpdatedAt = settingsRow.updatedAt;
      }
      // Records this caller as just-seen and reports back everyone
      // currently online -- see recordHeartbeatAndGetOnline(). Only happens
      // when the page actually sends a username (it's omitted in demo mode,
      // where "all" is never even called).
      out.online = recordHeartbeatAndGetOnline(e.parameter.username);
      return jsonOut(out);
    }
    if (action === "since") {
      var since = parseSinceParam(e.parameter.since);
      var out2 = {
        ok: true,
        precincts: getPrecincts(since),
        parcels: getParcels(since),
        serverTime: new Date().toISOString()
      };
      // Settings-panel colors/basemap rarely change, so -- like precincts
      // and parcels above -- only include it here when it actually changed
      // since the caller's last check, so a saved change from one
      // canvasser reaches everyone else's next 20-second poll.
      var settingsRow2 = getSettingsRow();
      if (settingsRow2 && (!since || settingsRow2.updatedAtDate.getTime() > since.getTime())) {
        out2.settings = settingsRow2.value;
        out2.settingsUpdatedAt = settingsRow2.updatedAt;
      }
      // Same piggybacked heartbeat as "all" above -- this is what actually
      // keeps presence fresh after boot, since "since" is what the page
      // polls every 20 seconds from then on (see SYNC_POLL_MS/
      // pollForUpdates() in index.html).
      out2.online = recordHeartbeatAndGetOnline(e.parameter.username);
      return jsonOut(out2);
    }
    return jsonOut({ ok: true, message: "Ami 4 Mayor Canvassing App backend is running." });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: "Bad request body" });
  }
  try {
    if (body.action === "login") {
      return jsonOut(handleLogin(body.username, body.passwordHash));
    }
    if (body.action === "saveState") {
      return jsonOut(handleSaveState(body));
    }
    if (body.action === "saveWalked") {
      return jsonOut(handleSaveWalked(body));
    }
    if (body.action === "saveSettings") {
      return jsonOut(handleSaveSettings(body));
    }
    return jsonOut({ ok: false, error: "Unknown action" });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ================== LOGIN ==================

function handleLogin(username, passwordHash) {
  username = String(username || "").trim();
  if (!username || !passwordHash) return { ok: false, error: "Username and password are required." };

  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowUser = String(rows[i][0] || "").trim();
    if (rowUser.toLowerCase() !== username.toLowerCase()) continue;

    var storedHash = String(rows[i][1] || "");
    var displayName = String(rows[i][2] || rowUser);
    var active = rows[i][3];
    var role = normalizeRole(rows[i][4]);
    if (active === false || String(active).toLowerCase() === "false" || String(active).toLowerCase() === "no") {
      return { ok: false, error: "This login has been disabled. Contact your campaign admin." };
    }
    if (storedHash.toLowerCase() === String(passwordHash).toLowerCase()) {
      return { ok: true, displayName: displayName, role: role };
    }
    return { ok: false, error: "Invalid username or password." };
  }
  return { ok: false, error: "Invalid username or password." };
}

// "viewer" if the Users row's Role column says so (any case, extra
// whitespace tolerated); "editor" for anything else, including a blank/
// missing column -- so an existing Sheet from before roles existed, or a
// row an admin just hasn't set a role on yet, defaults to full access
// rather than silently locking someone out.
function normalizeRole(raw) {
  return String(raw || "").trim().toLowerCase() === "viewer" ? "viewer" : "editor";
}

// { displayName, role } for a username, or null if there's no such row.
// lookupDisplayName() below is kept as a thin wrapper for existing callers
// that only need the name.
function lookupUser(username) {
  if (!username) return null;
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim().toLowerCase() === username.toLowerCase()) {
      return { displayName: String(rows[i][2] || rows[i][0]), role: normalizeRole(rows[i][4]) };
    }
  }
  return null;
}

// ================== SAVE (precinct panel: status + walked parcels) ==================

// body: { precinctId, status, completedDate, notes, walked: [{address, walked}], username }
// completedDate is a plain "yyyy-mm-dd" string from the dialog's date field
// (or blank/omitted) -- stored as-is, no timezone math needed. notes is a
// free-text field (also optional/blank) shown to every canvasser who opens
// that precinct's panel.
function handleSaveState(body) {
  var precinctId = String(body.precinctId || "").trim();
  var status = String(body.status || "").trim();
  var completedDate = String(body.completedDate || "").trim();
  var notes = String(body.notes || "").trim();
  var walked = Array.isArray(body.walked) ? body.walked : [];
  var username = String(body.username || "").trim();

  if (!precinctId) return { ok: false, error: "Precinct ID is required." };
  if (VALID_PRECINCT_STATUSES.indexOf(status) === -1) return { ok: false, error: "Invalid precinct status." };

  var user = lookupUser(username);
  // Enforced here, not just hidden in the app's UI -- a viewer's own browser
  // could in principle be made to call this endpoint directly regardless of
  // what buttons the page shows them.
  if (user && user.role === "viewer") {
    return { ok: false, error: "Viewers can't make changes to precincts or parcels." };
  }

  var now = new Date();
  var displayName = (user && user.displayName) || username;

  savePrecinctRow(precinctId, status, status === "Completed" ? completedDate : "", notes, displayName, now);

  if (walked.length) saveParcelRows(walked, displayName, now);

  return { ok: true, updatedAt: now.toISOString() };
}

// body: { walked: [{address, walked}], username }
// Used by the front end's "Canvassed addresses" quick-canvass panel (opened
// from the address search list) rather than the precinct panel -- that panel
// lets a canvasser mark parcels across MANY DIFFERENT precincts in a single
// session, so unlike handleSaveState() above there's no single precinctId
// this save is "for". This only ever writes the Parcels sheet (the exact
// same saveParcelRows() the precinct panel itself uses, so both paths stay
// in sync automatically); it deliberately never touches savePrecinctRow(),
// so it can't accidentally stamp some arbitrary "anchor" precinct's own
// status/notes row with a new UpdatedBy/UpdatedAt just because one of its
// parcels was marked from here.
function handleSaveWalked(body) {
  var walked = Array.isArray(body.walked) ? body.walked : [];
  var username = String(body.username || "").trim();

  var user = lookupUser(username);
  // Enforced here, not just hidden in the app's UI -- same reasoning as
  // handleSaveState() above.
  if (user && user.role === "viewer") {
    return { ok: false, error: "Viewers can't make changes to precincts or parcels." };
  }

  var now = new Date();
  var displayName = (user && user.displayName) || username;

  if (walked.length) saveParcelRows(walked, displayName, now);

  return { ok: true, updatedAt: now.toISOString() };
}

// Writes columns B-F (Status, CompletedDate, UpdatedBy, UpdatedAt, Notes) in
// that order -- Notes stays last, after UpdatedAt, rather than moving next
// to CompletedDate, precisely so this single setValues() call still lines up
// column-for-column on a Sheet created before Notes existed.
function savePrecinctRow(precinctId, status, completedDate, notes, displayName, now) {
  var sheet = getPrecinctsSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() === precinctId) {
      sheet.getRange(i + 1, 2, 1, 5).setValues([[status, completedDate, displayName, now, notes]]);
      return;
    }
  }
  sheet.appendRow([precinctId, status, completedDate, displayName, now, notes]);
}

// The browser already sends every changed parcel from a swipe session in
// one request (see the "walked" array on saveState) -- but this used to
// turn back into N round trips to the *Sheet*: one saveParcelRow() call per
// changed address, each of which re-read the entire Parcels sheet (11k+
// rows) from scratch just to find the one row it needed. Swiping a few dozen
// parcels meant a few dozen full-sheet reads/writes in sequence, which is
// what actually made Save feel slow -- not the browser-to-server trip.
// Reading the sheet once, patching every change into that in-memory copy,
// and writing it back in a single batched call fixes that regardless of how
// many parcels were swiped.
function saveParcelRows(walked, displayName, now) {
  var sheet = getParcelsSheet();
  var range = sheet.getDataRange();
  var rows = range.getValues();

  var indexByAddress = {};
  for (var i = 1; i < rows.length; i++) {
    var addr = String(rows[i][0] || "").trim().toLowerCase();
    if (addr) indexByAddress[addr] = i;
  }

  var toAppend = [];
  var changedExisting = false;
  for (var j = 0; j < walked.length; j++) {
    var address = String(walked[j].address || "").trim();
    if (!address) continue;
    var idx = indexByAddress[address.toLowerCase()];
    var value = !!walked[j].walked;
    if (idx !== undefined) {
      rows[idx][1] = value;
      rows[idx][2] = displayName;
      rows[idx][3] = now;
      changedExisting = true;
    } else {
      // Not already a row in the sheet (shouldn't normally happen -- every
      // address should already exist from the data file -- but handled just
      // in case) -- collect it for a single appended batch below rather
      // than an appendRow() per address.
      toAppend.push([address, value, displayName, now]);
    }
  }

  if (changedExisting) range.setValues(rows);
  if (toAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 4).setValues(toAppend);
  }
}

// ================== SETTINGS (shared map defaults: basemap + colors) ==================

// body: { settings: {basemap, colors: {...}}, username }
// Stores the whole settings object as one JSON blob in a single row, keyed
// by SETTINGS_KEY -- simplest way to keep an arbitrary/evolving set of
// fields (new color keys, etc.) without a schema migration every time the
// settings panel grows a new control.
function handleSaveSettings(body) {
  var settings = body && body.settings;
  if (!settings || typeof settings !== "object" || Array.isArray(settings)) {
    return { ok: false, error: "Settings object is required." };
  }
  var username = String(body.username || "").trim();
  var user = lookupUser(username);
  if (user && user.role === "viewer") {
    return { ok: false, error: "Viewers can't change Map Settings." };
  }
  var displayName = (user && user.displayName) || username || "unknown";
  var now = new Date();

  var sheet = getSettingsSheet();
  var rows = sheet.getDataRange().getValues();
  var json = JSON.stringify(settings);
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() === SETTINGS_KEY) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[json, displayName, now]]);
      return { ok: true, updatedAt: now.toISOString() };
    }
  }
  sheet.appendRow([SETTINGS_KEY, json, displayName, now]);
  return { ok: true, updatedAt: now.toISOString() };
}

// Returns { value: <parsed settings object>, updatedAt: <ISO string>,
// updatedAtDate: <Date> } for the single stored settings row, or null if
// nothing has been saved yet (fresh Sheet, or the row was deleted) -- callers
// treat null as "use the client's built-in defaults".
function getSettingsRow() {
  var sheet = getSettingsSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() !== SETTINGS_KEY) continue;
    var raw = rows[i][1];
    var parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (err) {
      return null; // corrupted cell -- fall back to client defaults rather than throwing
    }
    var updatedAtRaw = rows[i][3];
    var updatedAtDate = (updatedAtRaw instanceof Date) ? updatedAtRaw : new Date(updatedAtRaw);
    if (isNaN(updatedAtDate.getTime())) updatedAtDate = new Date(0);
    return {
      value: parsed,
      updatedAt: updatedAtDate.toISOString(),
      updatedAtDate: updatedAtDate
    };
  }
  return null;
}

// ================== READ (precincts + parcels) ==================

function getAllPrecincts() { return getPrecincts(null); }
function getAllParcels() { return getParcels(null); }

function getPrecincts(sinceDate) {
  var sheet = getPrecinctsSheet();
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (sinceDate) {
      var updatedAt = rows[i][4];
      if (!(updatedAt instanceof Date) || updatedAt.getTime() <= sinceDate.getTime()) continue;
    }
    out.push(rowToPrecinctEntry(rows[i]));
  }
  return out;
}

function getParcels(sinceDate) {
  var sheet = getParcelsSheet();
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (sinceDate) {
      var updatedAt = rows[i][3];
      if (!(updatedAt instanceof Date) || updatedAt.getTime() <= sinceDate.getTime()) continue;
    }
    out.push(rowToParcelEntry(rows[i]));
  }
  return out;
}

function rowToPrecinctEntry(row) {
  var updatedAt = row[4];
  return {
    precinctId: String(row[0]),
    status: row[1],
    completedDate: row[2] instanceof Date ? formatDateOnly(row[2]) : String(row[2] || ""),
    updatedBy: row[3],
    updatedAt: (updatedAt instanceof Date) ? updatedAt.toISOString() : String(updatedAt || ""),
    notes: String(row[5] || "")   // blank for any row saved before Notes existed (see the header comment up top)
  };
}

function rowToParcelEntry(row) {
  var updatedAt = row[3];
  return {
    address: row[0],
    walked: toBool(row[1]),
    updatedBy: row[2],
    updatedAt: (updatedAt instanceof Date) ? updatedAt.toISOString() : String(updatedAt || "")
  };
}

function formatDateOnly(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone() || "America/Los_Angeles", "yyyy-MM-dd");
}

function toBool(v) {
  if (typeof v === "boolean") return v;
  var s = String(v || "").trim().toLowerCase();
  return s === "true" || s === "yes" || s === "1";
}

// Parses the "since" query param (an ISO timestamp) into a Date. Returns
// null on anything missing/unparseable, which getPrecincts()/getParcels()
// treat as "no filter" -- i.e. falls back to returning everything.
function parseSinceParam(raw) {
  var d = new Date(String(raw || ""));
  return isNaN(d.getTime()) ? null : d;
}

function lookupDisplayName(username) {
  var user = lookupUser(username);
  return user ? user.displayName : null;
}

// ================== PRESENCE (who's currently online) ==================

// Upserts a single Presence row for `username` with the current time, then
// returns { username, displayName } for every row seen recently enough
// (within ONLINE_WINDOW_MS) to still count as online -- including the
// caller themselves; the page filters itself out before displaying the
// list (see renderOnlineUsers() in index.html). One row per username, kept
// up to date rather than appended to, so this sheet stays bounded by how
// many distinct people have ever logged in, not by how many polls have
// happened.
//
// Called from doGet's "all"/"since" handlers above, which is what actually
// drives this -- there's no separate heartbeat request. A blank/missing
// username (demo mode never sends one) just skips the upsert and returns
// whoever else is already recorded, so this never throws on that call
// shape.
function recordHeartbeatAndGetOnline(rawUsername) {
  var username = String(rawUsername || "").trim();
  var sheet = getPresenceSheet();
  var rows = sheet.getDataRange().getValues();
  var now = new Date();

  if (username) {
    var user = lookupUser(username);
    var displayName = (user && user.displayName) || username;
    var found = false;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "").trim().toLowerCase() === username.toLowerCase()) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[displayName, now]]);
        rows[i][1] = displayName;
        rows[i][2] = now;
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([username, displayName, now]);
      rows.push([username, displayName, now]);
    }
  }

  var cutoffMs = now.getTime() - ONLINE_WINDOW_MS;
  var online = [];
  for (var j = 1; j < rows.length; j++) {
    var rowUsername = String(rows[j][0] || "").trim();
    if (!rowUsername) continue;
    var seenRaw = rows[j][2];
    var seenDate = (seenRaw instanceof Date) ? seenRaw : new Date(seenRaw);
    if (isNaN(seenDate.getTime()) || seenDate.getTime() < cutoffMs) continue; // stale -- treated as offline
    online.push({ username: rowUsername, displayName: String(rows[j][1] || rowUsername) });
  }
  return online;
}

// ================== SHEET HELPERS ==================

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function getUsersSheet() { return getSheet(USERS_SHEET_NAME, ["Username", "PasswordHash", "DisplayName", "Active", "Role"]); }
function getPrecinctsSheet() { return getSheet(PRECINCTS_SHEET_NAME, ["PrecinctID", "Status", "CompletedDate", "UpdatedBy", "UpdatedAt", "Notes"]); }
function getParcelsSheet() { return getSheet(PARCELS_SHEET_NAME, ["Address", "Walked", "UpdatedBy", "UpdatedAt"]); }
function getSettingsSheet() { return getSheet(SETTINGS_SHEET_NAME, ["Key", "Value", "UpdatedBy", "UpdatedAt"]); }
function getPresenceSheet() { return getSheet(PRESENCE_SHEET_NAME, ["Username", "DisplayName", "LastSeenAt"]); }

function sha256Hex(input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return rawHash.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

// ================== ADMIN MENU (add/remove logins from the Sheet) ==================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Canvassing Admin")
    .addItem("Add or update login...", "adminAddOrUpdateUser")
    .addItem("Set role...", "adminSetRole")
    .addItem("Set chuck/rachel = editor, test = viewer", "adminSeedRequestedRoles")
    .addItem("Remove login...", "adminRemoveUser")
    .addItem("Disable login (keep row)...", "adminDisableUser")
    .addToUi();
}

// One-click convenience for a specific, requested role assignment: chuck and
// rachel as editors, test as a viewer. Only touches usernames that already
// have a login row (it can't create one -- a password is required for
// that, see "Add or update login..."); safe to run more than once.
function adminSeedRequestedRoles() {
  var ui = SpreadsheetApp.getUi();
  var assignments = { chuck: "editor", rachel: "editor", test: "viewer" };
  var applied = [];
  var missing = [];
  Object.keys(assignments).forEach(function (username) {
    if (setUserRole(username, assignments[username])) {
      applied.push(username + " -> " + assignments[username]);
    } else {
      missing.push(username);
    }
  });
  var msg = applied.length ? "Updated:\n" + applied.join("\n") : "No matching logins found.";
  if (missing.length) msg += "\n\nNo login row exists yet for: " + missing.join(", ") + " -- add them first with \"Add or update login...\".";
  ui.alert(msg);
}

function adminAddOrUpdateUser() {
  var ui = SpreadsheetApp.getUi();
  var userResp = ui.prompt("Add or update login", "Username (this is what they'll type to log in):", ui.ButtonSet.OK_CANCEL);
  if (userResp.getSelectedButton() !== ui.Button.OK) return;
  var username = userResp.getResponseText().trim();
  if (!username) { ui.alert("Username cannot be blank."); return; }

  var nameResp = ui.prompt("Display name", "Full name to show in the app for " + username + ":", ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;
  var displayName = nameResp.getResponseText().trim() || username;

  var pwResp = ui.prompt("Password", "New password for " + username + ":", ui.ButtonSet.OK_CANCEL);
  if (pwResp.getSelectedButton() !== ui.Button.OK) return;
  var password = pwResp.getResponseText();
  if (!password) { ui.alert("Password cannot be blank."); return; }

  var roleResp = ui.prompt("Role", "\"editor\" (can set status/mark parcels/notes) or \"viewer\" (look only) for " + username + " -- leave blank for editor:", ui.ButtonSet.OK_CANCEL);
  if (roleResp.getSelectedButton() !== ui.Button.OK) return;
  var role = normalizeRole(roleResp.getResponseText());

  var hash = sha256Hex(password);
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim().toLowerCase() === username.toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[username, hash, displayName, true, role]]);
      ui.alert("Updated login for " + username + " (role: " + role + ").");
      return;
    }
  }
  sheet.appendRow([username, hash, displayName, true, role]);
  ui.alert("Added login for " + username + " (role: " + role + ").");
}

// Changes just the Role column for an existing user, without touching their
// password -- the normal path for promoting/demoting someone between editor
// and viewer.
function adminSetRole() {
  var ui = SpreadsheetApp.getUi();
  var userResp = ui.prompt("Set role", "Username to set the role for:", ui.ButtonSet.OK_CANCEL);
  if (userResp.getSelectedButton() !== ui.Button.OK) return;
  var username = userResp.getResponseText().trim();
  if (!username) { ui.alert("Username cannot be blank."); return; }

  var roleResp = ui.prompt("Role", "New role for " + username + " -- \"editor\" or \"viewer\":", ui.ButtonSet.OK_CANCEL);
  if (roleResp.getSelectedButton() !== ui.Button.OK) return;
  var role = normalizeRole(roleResp.getResponseText());

  if (setUserRole(username, role)) {
    ui.alert("Set " + username + "'s role to " + role + ".");
  } else {
    ui.alert("No login found for " + username + ".");
  }
}

// Sets an existing user's Role column directly (column E); returns false if
// no row matches that username. Used by adminSetRole() and by the one-click
// adminSeedRequestedRoles() convenience below.
function setUserRole(username, role) {
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim().toLowerCase() === username.toLowerCase()) {
      sheet.getRange(i + 1, 5).setValue(role);
      return true;
    }
  }
  return false;
}

function adminRemoveUser() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Remove login", "Username to remove entirely:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var username = resp.getResponseText().trim();
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0] || "").trim().toLowerCase() === username.toLowerCase()) {
      sheet.deleteRow(i + 1);
      ui.alert("Removed " + username + ".");
      return;
    }
  }
  ui.alert("No login found for " + username + ".");
}

function adminDisableUser() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Disable login", "Username to disable (keeps their row/history but blocks login):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var username = resp.getResponseText().trim();
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim().toLowerCase() === username.toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue(false);
      ui.alert("Disabled " + username + ".");
      return;
    }
  }
  ui.alert("No login found for " + username + ".");
}
