/**
 * ============================================
 * Auth.gs — Authentication (fully cached)
 * Login reads from PropertiesService cache
 * Sheet read only on first login or after cache miss
 * ============================================
 */

function loginUser(username, password) {
  try {
    // Try cached users first (instant, no sheet read)
    var users = propGetJson_('users_cache');
    if (!users) {
      // Cache miss — read from sheet, then cache
      users = loadUsersFromSheet_();
      if (!users) return { success: false, message: 'Sheet Users tidak ditemukan.' };
      propSetJson_('users_cache', users);
    }

    var uLower = String(username).trim().toLowerCase();
    for (var i = 0; i < users.length; i++) {
      var u = users[i];
      if (String(u.nama).toLowerCase() === uLower && u.pwd === password) {
        if (!u.active) return { success: false, message: 'Akun tidak aktif. Hubungi admin.' };
        return { success: true, user: { user_id: u.id, nama: u.nama, role: u.role } };
      }
    }
    return { success: false, message: 'Nama atau password salah.' };
  } catch (e) { return { success: false, message: 'Error: ' + e.message }; }
}

function loadUsersFromSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({ id: data[i][0], nama: String(data[i][1]).trim(), pwd: String(data[i][2]).trim(), role: String(data[i][3]).trim(), active: !!data[i][4] });
  }
  return users;
}

/**
 * Call this after modifying Users sheet (e.g., add/remove/change password)
 */
function invalidateUsersCache() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('users_cache');
    invalidateOperatorList_();
  } catch (e) {}
}

function getOperatorList() {
  var cached = getCachedOperatorList_();
  if (cached) return cached;
  try {
    var users = propGetJson_('users_cache');
    if (!users) {
      users = loadUsersFromSheet_();
      if (users) propSetJson_('users_cache', users);
    }
    if (!users) return [];
    var operators = [];
    for (var i = 0; i < users.length; i++) {
      if (users[i].active) operators.push({ user_id: users[i].id, nama: users[i].nama, role: users[i].role });
    }
    setCachedOperatorList_(operators);
    return operators;
  } catch (e) { return []; }
}
