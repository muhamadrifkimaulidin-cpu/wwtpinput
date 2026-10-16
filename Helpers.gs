/**
 * ============================================
 * Helpers.gs — Utilities + Cache Layer
 * 3 cache tiers: CacheService, PropertiesService, invalidation
 * ============================================
 */

// ========== BASIC HELPERS ==========
function getCurrentShift() {
  var hour = new Date().getHours();
  if (hour >= 6 && hour < 14) return 1;
  if (hour >= 14 && hour < 22) return 2;
  return 3;
}
function getTodayDate() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function getServerInfo() {
  return { today: getTodayDate(), shift: getCurrentShift(), timezone: Session.getScriptTimeZone() };
}

// ========== CACHE SERVICE (6hr max, server-side) ==========
var CACHE_TTL_ = 3600;
var CACHE_TTL_SHORT_ = 300;
var CACHE_TTL_LONG_ = 21600;

function cacheGet_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var val = cache.get(key);
    if (val === null) return null;
    return JSON.parse(val);
  } catch (e) { return null; }
}

function cachePut_(key, value, ttl) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(value);
    if (json.length > 95000) return;
    cache.put(key, json, ttl || CACHE_TTL_);
  } catch (e) { Logger.log('cachePut_ error: ' + e.message); }
}

function cacheDelete_(key) {
  try { CacheService.getScriptCache().remove(key); } catch (e) {}
}

function cacheDeletePattern_(prefix) {
  try {
    var cache = CacheService.getScriptCache();
    var keys = propGet_('cache_keys_' + prefix);
    if (keys) {
      var arr = keys.split(',');
      arr.forEach(function(k) { cache.remove(k); });
      PropertiesService.getScriptProperties().deleteProperty('cache_keys_' + prefix);
    }
  } catch (e) {}
}

function cacheTrackKey_(prefix, key) {
  try {
    var propKey = 'cache_keys_' + prefix;
    var existing = propGet_(propKey) || '';
    var arr = existing ? existing.split(',') : [];
    if (arr.indexOf(key) === -1) {
      arr.push(key);
      if (arr.length > 50) arr = arr.slice(-50);
      PropertiesService.getScriptProperties().setProperty(propKey, arr.join(','));
    }
  } catch (e) {}
}

// ========== PROPERTIES SERVICE (persistent, key-value) ==========
function propGet_(key) {
  try { return PropertiesService.getScriptProperties().getProperty(key); } catch (e) { return null; }
}
function propSet_(key, value) {
  try { PropertiesService.getScriptProperties().setProperty(key, value); } catch (e) {}
}
function propGetJson_(key) {
  var val = propGet_(key);
  if (!val) return null;
  try { return JSON.parse(val); } catch (e) { return null; }
}
function propSetJson_(key, value) {
  try { propSet_(key, JSON.stringify(value)); } catch (e) {}
}

// ========== CONFIG CACHE ==========
function getConfig_(key) {
  var cached = propGet_('cfg_' + key);
  if (cached !== null) return cached;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    propSet_('cfg_' + data[i][0], String(data[i][1]));
  }
  return propGet_('cfg_' + key);
}

function invalidateConfig_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    for (var key in all) {
      if (key.indexOf('cfg_') === 0) props.deleteProperty(key);
    }
  } catch (e) {}
}

// ========== OPERATOR LIST CACHE ==========
function getCachedOperatorList_() {
  var cached = propGetJson_('operator_list');
  if (cached) return cached;
  return null;
}
function setCachedOperatorList_(list) { propSetJson_('operator_list', list); }
function invalidateOperatorList_() {
  try { PropertiesService.getScriptProperties().deleteProperty('operator_list'); } catch (e) {}
}

// ========== DATA CACHE INVALIDATION ==========
function invalidateDataCaches_() {
  cacheDeletePattern_('dash');
  cacheDeletePattern_('qual');
  cacheDeletePattern_('records');
  cacheDeletePattern_('lastshift');
  Logger.log('Data caches invalidated');
}
function invalidateRequestsCaches_() {
  cacheDeletePattern_('req');
  Logger.log('Request caches invalidated');
}
function invalidateNotifsCaches_() {
  cacheDeletePattern_('ntf');
  Logger.log('Notif caches invalidated');
}
