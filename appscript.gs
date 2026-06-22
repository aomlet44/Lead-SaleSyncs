// ============================================================
// Salesyncs Business Health Check — Google Apps Script
// วางโค้ดนี้ใน Apps Script ที่: https://script.google.com
// แล้ว Deploy → New deployment → Web app
//   Execute as: Me
//   Who has access: Anyone
// ============================================================

const SHEET_ID = '1o61DHyzvP6H272ve3EbhGE6SRd3vhc3ur4L0ycV81bA';

const HEADERS = [
  'lead_id', 'created_at', 'created_at_iso',
  'name', 'company', 'position', 'phone', 'email',
  'lead_type', 'role_interest',
  'interests', 'interests_keys',
  'overall_score', 'health_level',
  'revenue', 'retention', 'partner', 'data_score', 'sales', 'ai',
  'q1', 'q1_keys', 'q2', 'q2_keys', 'q3', 'q3_keys',
  'q4', 'q4_key', 'q5', 'q5_keys',
  'modules', 'follow_up', 'follow_up_secondary',
  'json_data'
];

function doGet(e) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

  // ── READ: admin ดึง Leads ทั้งหมด ──
  if (e.parameter.action === 'read') {
    try {
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        return jsonResp({ ok: true, rows: [] });
      }
      const headers = data[0];
      const jsonIdx = headers.indexOf('json_data');
      const rows = [];
      for (let i = 1; i < data.length; i++) {
        const raw = data[i][jsonIdx];
        if (raw) {
          try { rows.push(JSON.parse(raw)); } catch (_) {}
        }
      }
      return jsonResp({ ok: true, rows: rows });
    } catch (ex) {
      return jsonResp({ ok: false, error: ex.message });
    }
  }

  // ── WRITE: บันทึก Lead ใหม่ ──
  if (e.parameter.data) {
    try {
      const payload = JSON.parse(e.parameter.data);

      // สร้าง Header row ครั้งแรก
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(HEADERS);
        sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      }

      const row = HEADERS.map(h => (payload[h] !== undefined ? payload[h] : ''));
      sheet.appendRow(row);

      return jsonResp({ ok: true });
    } catch (ex) {
      return jsonResp({ ok: false, error: ex.message });
    }
  }

  return jsonResp({ ok: true, message: 'Salesyncs API ready' });
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
