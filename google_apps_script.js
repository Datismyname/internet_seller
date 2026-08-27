function doGet(e) {
  const action = e.parameter.action;

  // ── LOGIN ──
  if (action === 'login') {
    const username = e.parameter.username;
    const password = e.parameter.password;
    const usersSheet = getOrCreateSheet('Users');
    const users = usersSheet.getDataRange().getValues();
    if (users.length <= 1) {
      return jsonResponse({ success: false, error: 'No users found' });
    }
    const headers = users[0];
    const uIdx = headers.indexOf('username');
    const pIdx = headers.indexOf('password');
    const nIdx = headers.indexOf('name');
    const rIdx = headers.indexOf('role');
    const idIdx = headers.indexOf('id');

    for (let i = 1; i < users.length; i++) {
      if (String(users[i][uIdx]) === String(username) && String(users[i][pIdx]) === String(password)) {
        return jsonResponse({
          success: true,
          user: {
            id: users[i][idIdx],
            username: users[i][uIdx],
            name: users[i][nIdx],
            role: users[i][rIdx]
          }
        });
      }
    }
    return jsonResponse({ success: false, error: 'Wrong username or password' });
  }

  // ── GET ALL CUSTOMERS ──
  if (action === 'getAll') {
    const sheet = getOrCreateSheet('Customers');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ success: true, data: [] });
    }
    const headers = data[0];
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        row[String(headers[j])] = data[i][j];
      }
      rows.push(row);
    }
    return jsonResponse({ success: true, data: rows });
  }

  // ── GET BY SELLER ──
  if (action === 'getBySeller') {
    const sellerId = e.parameter.sellerId;
    const sheet = getOrCreateSheet('Customers');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ success: true, data: [] });
    }
    const headers = data[0];
    const sIdx = headers.indexOf('sellerId');
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][sIdx]) === String(sellerId)) {
        const row = {};
        for (let j = 0; j < headers.length; j++) {
          row[String(headers[j])] = data[i][j];
        }
        rows.push(row);
      }
    }
    return jsonResponse({ success: true, data: rows });
  }

  // ── ADD CUSTOMER ──
  if (action === 'add') {
    try {
      const jsonData = e.parameter.data;
      const data = JSON.parse(jsonData);
      const sheet = getOrCreateSheet('Customers');
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const row = [];
      for (let h of headers) {
        row.push(data[String(h)] !== undefined ? data[String(h)] : '');
      }
      sheet.appendRow(row);
      return jsonResponse({ success: true, message: 'Added' });
    } catch (err) {
      return jsonResponse({ success: false, error: String(err) });
    }
  }

  // ── UPDATE CUSTOMER ──
  if (action === 'update') {
    try {
      const jsonData = e.parameter.data;
      const data = JSON.parse(jsonData);
      const sheet = getOrCreateSheet('Customers');
      const allData = sheet.getDataRange().getValues();
      const headers = allData[0];
      const idCol = headers.indexOf('id');
      if (idCol < 0) return jsonResponse({ success: false, error: 'id column not found' });
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][idCol]) === String(data.id)) {
          for (let j = 0; j < headers.length; j++) {
            if (data[String(headers[j])] !== undefined) {
              sheet.getRange(i + 1, j + 1).setValue(data[String(headers[j])]);
            }
          }
          break;
        }
      }
      return jsonResponse({ success: true, message: 'Updated' });
    } catch (err) {
      return jsonResponse({ success: false, error: String(err) });
    }
  }

  // ── DELETE CUSTOMER ──
  if (action === 'delete') {
    try {
      const id = e.parameter.id;
      const sheet = getOrCreateSheet('Customers');
      const allData = sheet.getDataRange().getValues();
      const headers = allData[0];
      const idCol = headers.indexOf('id');
      if (idCol < 0) return jsonResponse({ success: false, error: 'id column not found' });
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][idCol]) === String(id)) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse({ success: true, message: 'Deleted' });
    } catch (err) {
      return jsonResponse({ success: false, error: String(err) });
    }
  }

  return jsonResponse({ success: false, error: 'Unknown action: ' + action });
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Users') {
      sheet.appendRow(['id','username','password','name','role','createdAt']);
      sheet.appendRow(['admin1','admin','admin123','المدير','admin',new Date().toISOString()]);
    }
    if (name === 'Customers') {
      sheet.appendRow(['id','name','phone','company','package','duration','startDate','endDate','notes','rating','lat','lng','sellerId','sellerName','createdAt']);
    }
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
