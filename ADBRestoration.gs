/**
 * ADB Restoration — photo intake, location index, and map feed.
 *
 * Deploy: Extensions > Apps Script, paste, save, then run setup() once.
 * Then Deploy > New deployment > Web app > Execute as: Me,
 * Who has access: Anyone. Copy the /exec URL into the upload form.
 */

var ROOT_NAME  = 'ADB Restoration';
var SHEET_NAME = 'ADB Restoration Log';
var CATEGORIES = ['Concrete Res.', 'Pothole Res.'];

// The work is all in-state. Geocoder results outside this box are wrong.
var AZ_BOUNDS = { latMin: 31.0, latMax: 37.1, lngMin: -115.0, lngMax: -108.9 };

var HEADERS = [
  'Timestamp', 'PO', 'Category', 'Latitude', 'Longitude',
  'Address', 'Captured At', 'Location Source', 'Note',
  'File Name', 'File ID', 'Photo URL', 'Thumbnail URL'
];

/* ------------------------------------------------------------------ setup */

/**
 * Run once. Creates the root folder and log sheet, makes the folder
 * link-viewable, and stores the IDs. Safe to re-run.
 */
function setup() {
  var props = PropertiesService.getScriptProperties();

  var root = findOrCreateFolder_(DriveApp.getRootFolder(), ROOT_NAME);
  root.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty('ROOT_ID', root.getId());

  var sheetId = props.getProperty('SHEET_ID');
  var ss = null;
  if (sheetId) {
    try { ss = SpreadsheetApp.openById(sheetId); } catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(SHEET_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(root);
    props.setProperty('SHEET_ID', ss.getId());
  }

  var sh = ss.getSheets()[0];
  sh.setName('Log');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#efefef');
    sh.setFrozenRows(1);
  }

  Logger.log('Root folder: ' + root.getUrl());
  Logger.log('Log sheet:   ' + ss.getUrl());
  return { rootId: root.getId(), sheetId: ss.getId() };
}

/* ------------------------------------------------------------------ intake */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    // Corrections come back through the same endpoint, but they carry no
    // image and must not be held to the upload path's rules.
    if (body.action === 'edit') return editRow_(body);

    var po       = normalizePO_(body.po);
    var category = matchCategory_(body.category);
    var lat      = Number(body.lat);
    var lng      = Number(body.lng);

    if (!po)                         throw new Error('Missing PO number');
    if (!category)                   throw new Error('Unknown category: ' + body.category);
    if (!isFinite(lat) || !isFinite(lng)) throw new Error('Missing coordinates');
    if (!body.imageBase64)           throw new Error('Missing image data');

    var root     = DriveApp.getFolderById(getProp_('ROOT_ID'));
    var poFolder = findOrCreateFolder_(root, po);
    var target   = findOrCreateFolder_(poFolder, category);

    var name = body.filename || (po + '_' + category.replace(/\W+/g, '') + '_' +
                                 Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd_HHmmss') + '.jpg');
    var blob = Utilities.newBlob(Utilities.base64Decode(body.imageBase64), 'image/jpeg', name);
    var file = target.createFile(blob);

    var fileId = file.getId();
    appendRow_([
      new Date(),
      po,
      category,
      lat,
      lng,
      body.address || '',
      body.capturedAt || '',
      body.locationSource || 'unknown',
      body.note || '',
      name,
      fileId,
      'https://drive.google.com/file/d/' + fileId + '/view',
      'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w600'
    ]);

    return json_({
      ok: true, fileId: fileId, po: po, category: category,
      poLooksNormal: isWellFormedPO_(po)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* -------------------------------------------------------------- edits */

/**
 * Run once to set the code the crew types before an edit will save:
 *   setEditPin('4821')
 * The web app is deployed to "Anyone", so without a code anybody holding the
 * /exec URL could rewrite the log. Re-run to change it; setEditPin('') turns
 * editing off again. Corrections are recorded in the sheet's own version
 * history (File > Version history), so there is always a way back.
 */
function setEditPin(pin) {
  var props = PropertiesService.getScriptProperties();
  var v = String(pin == null ? '' : pin).trim();
  if (!v) {
    props.deleteProperty('EDIT_PIN');
    Logger.log('Editing is off — no code set.');
    return;
  }
  if (v.length < 4) throw new Error('Use at least 4 characters.');
  props.setProperty('EDIT_PIN', v);
  Logger.log('Edit code set. Type it into the map once per phone.');
}

// Ten wrong guesses in fifteen minutes and this stops answering. The URL is
// public, so a short code would otherwise be a few minutes of guessing.
var PIN_TRIES  = 10;
var PIN_WINDOW = 900;

function requirePin_(given) {
  var want = PropertiesService.getScriptProperties().getProperty('EDIT_PIN');
  if (!want) throw new Error('Editing is off — run setEditPin() in Apps Script first');

  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pin_fails') || 0);
  if (fails >= PIN_TRIES) throw new Error('Too many wrong codes — wait a few minutes');

  if (String(given == null ? '' : given) !== want) {
    cache.put('pin_fails', String(fails + 1), PIN_WINDOW);
    throw new Error('Wrong code');
  }
  cache.remove('pin_fails');
}

/**
 * Corrects the tags on one already-uploaded photo: PO, category, note, and
 * the pin itself. Keyed by Drive file ID, which is the one column nothing
 * else rewrites. Only the fields present in the request are touched.
 *
 * The photo never changes — a wrong tag is a filing mistake, not a bad shot.
 */
function editRow_(body) {
  requirePin_(body.pin);

  var fileId = String(body.fileId || '');
  if (!fileId) throw new Error('Missing photo id');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh   = logSheet_();
    var rows = sh.getDataRange().getValues();

    var at = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][10]) === fileId) { at = i; break; }
    }
    if (at < 0) throw new Error('That photo is not in the log');

    var row    = rows[at];
    var oldPO  = String(row[1]);
    var oldCat = String(row[2]);

    var po = body.po == null ? oldPO : normalizePO_(body.po);
    if (!po) throw new Error('Missing PO number');

    var cat = body.category == null ? oldCat : matchCategory_(body.category);
    if (!cat) throw new Error('Unknown category: ' + body.category);

    var lat     = Number(row[3]);
    var lng     = Number(row[4]);
    var address = row[5];
    var source  = row[7];

    if (body.lat != null && body.lng != null) {
      var newLat = Number(body.lat);
      var newLng = Number(body.lng);
      if (!isFinite(newLat) || !isFinite(newLng)) throw new Error('Bad coordinates');
      if (!inAZ_(newLat, newLng)) throw new Error('That pin is outside Arizona');

      // A pin moved by hand outranks the stamp, and the stamped address it
      // came with is no longer where the pin sits.
      if (newLat !== lat || newLng !== lng) {
        lat = newLat;
        lng = newLng;
        source  = 'manual';
        address = reverseGeocode_(lat, lng);
      }
    }

    var note = body.note == null ? row[8] : String(body.note);
    var name = String(row[9]);

    // The PO and category folders are how the photos are filed, so Drive has
    // to move with the sheet or the correction only half happens.
    if (po !== oldPO || cat !== oldCat) {
      var root   = DriveApp.getFolderById(getProp_('ROOT_ID'));
      var target = findOrCreateFolder_(findOrCreateFolder_(root, po), cat);
      var file   = DriveApp.getFileById(fileId);
      name = renamedFile_(name, oldPO, oldCat, po, cat);
      file.setName(name);
      file.moveTo(target);
    }

    sh.getRange(at + 1, 2, 1, 8)
      .setValues([[po, cat, lat, lng, address, row[6], source, note]]);
    sh.getRange(at + 1, 10).setValue(name);

    return json_({
      ok: true, fileId: fileId, po: po, category: cat,
      lat: lat, lng: lng, address: address, source: source, note: note,
      fileName: name, poLooksNormal: isWellFormedPO_(po)
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Street address for a pin that was dragged by hand, so the sheet and the
 * popup stop showing where the photo was originally stamped. Returns '' when
 * the lookup fails — an empty address beats a wrong one.
 */
function reverseGeocode_(lat, lng) {
  try {
    var res = Maps.newGeocoder().reverseGeocode(lat, lng);
    var r = res && res.status === 'OK' && res.results && res.results[0];
    return r ? (r.formatted_address || '') : '';
  } catch (err) {
    return '';
  }
}

/**
 * Uploads are named PO_Category_timestamp.jpg. A corrected PO or category
 * should show in the file name too, but only when the name actually starts
 * with the old one — a hand-named file keeps the name it was given.
 */
function renamedFile_(name, oldPO, oldCat, po, cat) {
  var from = oldPO + '_' + oldCat.replace(/\W+/g, '') + '_';
  var to   = po + '_' + cat.replace(/\W+/g, '') + '_';
  return name.indexOf(from) === 0 ? to + name.slice(from.length) : name;
}

/* -------------------------------------------------------------- map feed */

/**
 * ?action=points  -> all plotted photos
 * ?action=pos     -> known PO numbers, for the upload form dropdown
 * ?action=kml     -> rebuild the KML file, returns its URL
 * ?action=geocode -> street-level coordinates for a stamped address
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'points';
  try {
    if (action === 'pos')  return json_({ ok: true, pos: listPOs_() });
    if (action === 'kml')  return json_({ ok: true, url: buildKml() });
    if (action === 'geocode') return json_(geocodeAddress_(e.parameter.address));
    return json_({ ok: true, points: listPoints_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function listPoints_() {
  var rows = logSheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[3] || !r[4]) continue;
    out.push({
      po: r[1], category: r[2],
      lat: Number(r[3]), lng: Number(r[4]),
      address: r[5], capturedAt: String(r[6]),
      source: r[7], note: r[8],
      fileId: r[10], photo: r[11], thumb: r[12]
    });
  }
  return out;
}

function listPOs_() {
  var rows = logSheet_().getDataRange().getValues();
  var seen = {}, out = [];
  for (var i = 1; i < rows.length; i++) {
    var po = rows[i][1];
    if (po && !seen[po]) { seen[po] = true; out.push(po); }
  }
  // Folders count too — a PO may exist before its first photo lands.
  var kids = DriveApp.getFolderById(getProp_('ROOT_ID')).getFolders();
  while (kids.hasNext()) {
    var n = kids.next().getName();
    if (!seen[n]) { seen[n] = true; out.push(n); }
  }
  return out.sort();
}

/* ------------------------------------------------------------------ geocode */

/**
 * Coordinates for an address read off a photo stamp that carried no lat/lng.
 * Address-level is coarse, but it lands on the right block — closer than the
 * phone's idea of "here" when the photo was taken months ago and miles away.
 * Uses the built-in Maps service, so there is no API key to manage.
 *
 * Returns { ok: true, found: false } when the address is unusable, resolves
 * outside Arizona, or only resolves to a city or ZIP centroid. A downtown pin
 * is the wrong-side-of-town guess this exists to prevent, so those cases fall
 * through and let the caller reach for phone GPS instead.
 */
function geocodeAddress_(address) {
  var q = String(address || '').trim();
  if (q.length < 6) return { ok: true, found: false };

  // OCR'd stamps repeat across a job, and the geocoder is quota-limited.
  var cache = CacheService.getScriptCache();
  var key = 'geo_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, q));
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  var res = Maps.newGeocoder()
    .setRegion('us')
    .setBounds(AZ_BOUNDS.latMin, AZ_BOUNDS.lngMin, AZ_BOUNDS.latMax, AZ_BOUNDS.lngMax)
    .geocode(q);

  var out = { ok: true, found: false };
  var r = res && res.status === 'OK' && res.results && res.results[0];
  if (r && r.geometry.location_type !== 'APPROXIMATE' &&
      inAZ_(r.geometry.location.lat, r.geometry.location.lng)) {
    out = {
      ok: true,
      found: true,
      lat: r.geometry.location.lat,
      lng: r.geometry.location.lng,
      address: r.formatted_address || q
    };
  }
  cache.put(key, JSON.stringify(out), 21600);  // 6h — addresses do not move
  return out;
}

function inAZ_(lat, lng) {
  return lat >= AZ_BOUNDS.latMin && lat <= AZ_BOUNDS.latMax &&
         lng >= AZ_BOUNDS.lngMin && lng <= AZ_BOUNDS.lngMax;
}

/* ------------------------------------------------------------------- KML */

var KML_COLORS = {
  'Pothole Res.':  'ff3643d9',  // aabbggrr — red
  'Concrete Res.': 'ffd98236'   // blue
};

/** Rebuilds ADB Restoration.kml in the root folder. Returns its URL. */
function buildKml() {
  var pts = listPoints_();
  var x = [];
  x.push('<?xml version="1.0" encoding="UTF-8"?>');
  x.push('<kml xmlns="http://www.opengis.net/kml/2.2"><Document>');
  x.push('<name>ADB Restoration</name>');

  CATEGORIES.forEach(function (cat) {
    x.push('<Style id="' + styleId_(cat) + '"><IconStyle><color>' +
           (KML_COLORS[cat] || 'ff888888') + '</color><Icon><href>' +
           'http://maps.google.com/mapfiles/kml/paddle/wht-circle.png' +
           '</href></Icon></IconStyle></Style>');
  });

  CATEGORIES.forEach(function (cat) {
    var inCat = pts.filter(function (p) { return p.category === cat; });
    if (!inCat.length) return;
    x.push('<Folder><name>' + esc_(cat) + ' (' + inCat.length + ')</name>');
    inCat.forEach(function (p) {
      x.push('<Placemark>');
      x.push('<name>' + esc_(p.po) + '</name>');
      x.push('<description><![CDATA[' +
             '<b>' + p.po + '</b> — ' + p.category + '<br>' +
             (p.address || '') + '<br>' + (p.capturedAt || '') + '<br>' +
             (p.source === 'geocode' ? '<i>Pin from the stamped address — block-level</i><br>' : '') +
             (p.note ? p.note + '<br>' : '') +
             '<img src="' + p.thumb + '" width="320"><br>' +
             '<a href="' + p.photo + '">Open photo</a>' +
             ']]></description>');
      x.push('<styleUrl>#' + styleId_(cat) + '</styleUrl>');
      x.push('<Point><coordinates>' + p.lng + ',' + p.lat + ',0</coordinates></Point>');
      x.push('</Placemark>');
    });
    x.push('</Folder>');
  });

  x.push('</Document></kml>');

  var root = DriveApp.getFolderById(getProp_('ROOT_ID'));
  var name = 'ADB Restoration.kml';
  var existing = root.getFilesByName(name);
  while (existing.hasNext()) existing.next().setTrashed(true);
  var file = root.createFile(Utilities.newBlob(x.join('\n'), 'application/vnd.google-earth.kml+xml', name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

/** Hourly trigger so the KML stays current. Run once to install. */
function installKmlTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buildKml') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildKml').timeBased().everyHours(1).create();
}

/* ----------------------------------------------------------------- utils */

function getProp_(key) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('Not set up yet — run setup() first (missing ' + key + ')');
  return v;
}

function logSheet_() {
  return SpreadsheetApp.openById(getProp_('SHEET_ID')).getSheetByName('Log');
}

function appendRow_(values) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { logSheet_().appendRow(values); } finally { lock.releaseLock(); }
}

function findOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/**
 * POs look like ADB0089505 — a three-letter prefix and seven digits.
 * Typing in a truck is error-prone, so this is forgiving:
 *   "adb0089505", "ADB 0089505", "ADB-0089505" -> ADB0089505
 *   "0089505", "89505"                         -> ADB0089505  (padded)
 * Anything that doesn't fit the pattern is passed through uppercased
 * rather than rejected — a weird PO still beats a lost photo.
 */
var PO_PREFIX  = 'ADB';
var PO_DIGITS  = 7;
var PO_PATTERN = /^[A-Z]{2,4}\d{4,10}$/;

function normalizePO_(raw) {
  if (!raw) return '';
  var s = String(raw).toUpperCase().replace(/[^A-Z0-9]+/g, '');
  if (!s) return '';

  // Bare digits: assume the house prefix and pad out.
  if (/^\d+$/.test(s)) {
    return PO_PREFIX + padStart_(s, PO_DIGITS);
  }

  // Known prefix with a short number: pad the numeric tail.
  var m = s.match(/^([A-Z]{2,4})(\d+)$/);
  if (m && m[2].length < PO_DIGITS) {
    return m[1] + padStart_(m[2], PO_DIGITS);
  }

  return s;
}

/** True when the PO matches the expected shape — the form warns if not. */
function isWellFormedPO_(po) {
  return PO_PATTERN.test(String(po || ''));
}

function padStart_(s, len) {
  s = String(s);
  while (s.length < len) s = '0' + s;
  return s;
}

function matchCategory_(raw) {
  if (!raw) return '';
  var k = String(raw).toLowerCase();
  for (var i = 0; i < CATEGORIES.length; i++) {
    if (CATEGORIES[i].toLowerCase() === k) return CATEGORIES[i];
  }
  if (k.indexOf('pot') === 0)  return 'Pothole Res.';
  if (k.indexOf('con') === 0)  return 'Concrete Res.';
  return '';
}

function styleId_(cat) { return cat.replace(/\W+/g, '_'); }
function tz_() { return Session.getScriptTimeZone() || 'America/Phoenix'; }
function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
