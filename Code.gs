/**
 * JAMIYATUL ABRAR AL ISLAMIA ANWARA COMPLEX
 * Google Sheets <-> Live Web App & APK Connector (Google Apps Script)
 * 
 * সেটআপ নিয়ম:
 * ১. আপনার গুগল শিট (যেমন: Nobember-2026 বা Blane) খুলুন।
 * ২. উপরের মেনু থেকে Extensions -> Apps Script এ যান।
 * ৩. এই সম্পূর্ণ কোডটি Code.gs ফাইলে পেস্ট করে Save (Ctrl+S) দিন।
 * ৪. Deploy -> New deployment -> Web app -> Execute as: Me -> Who has access: Anyone সিলেক্ট করে Deploy দিন।
 * ৫. প্রাপ্ত Web App URL-টি কপি করে আপনার ওয়েব অ্যাপ/APK-এর "অ্যাডমিন প্যানেল -> গুগল শিট লাইভ সিঙ্ক লিংক" বক্সে পেস্ট করুন!
 */

const SHEET_NAME = 'November';
const SYNC_SECRET_KEY = 'abrar2026'; // নিরাপত্তার জন্য সিক্রেট কোড

function doGet(e) {
  try {
    const key = (e && e.parameter && e.parameter.key) ? e.parameter.key : '';
    if (key !== SYNC_SECRET_KEY) {
      return jsonResponse({ status: 'error', message: 'Unauthorized key' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // বক্স ১: রিসিপ্ট (C7:F28) ও পেমেন্ট (H7:L28) পড়া
    const receipts = [];
    const rValues = sheet.getRange('C7:F28').getValues();
    for (let i = 0; i < rValues.length; i++) {
      const part = String(rValues[i][2] || '').trim();
      const amt = Number(rValues[i][3]) || 0;
      if (part || amt > 0) {
        receipts.push({
          id: 'r_' + (i + 1),
          date: formatSheetDate(rValues[i][1]),
          particulars: part,
          amount: amt,
          isBank: part.toLowerCase().indexOf('opening bank') !== -1,
          isOpeningCash: part.toLowerCase().indexOf('opening cash') !== -1
        });
      }
    }

    const payments = [];
    const pValues = sheet.getRange('H7:L28').getValues();
    for (let i = 0; i < pValues.length; i++) {
      const part = String(pValues[i][3] || '').trim();
      const amt = Number(pValues[i][4]) || 0;
      if (part || amt > 0) {
        payments.push({
          id: 'p_' + (i + 1),
          date: formatSheetDate(pValues[i][1]),
          voucher: Number(pValues[i][2]) || 0,
          particulars: part,
          amount: amt
        });
      }
    }

    // বক্স ৩: নগদ ও বিকাশ ব্যালেন্স (Y13:Y14) পড়া
    const nagadRaw = Number(sheet.getRange('Y13').getValue()) || 0;
    const bkashRaw = Number(sheet.getRange('Y14').getValue()) || 0;

    return jsonResponse({
      status: 'ok',
      sheetTitle: ss.getName(),
      updatedAt: new Date().toISOString(),
      receipts: receipts,
      payments: payments,
      nagadRaw: nagadRaw,
      bkashRaw: bkashRaw
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.key !== SYNC_SECRET_KEY) {
      return jsonResponse({ status: 'error', message: 'Unauthorized key' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // অ্যাপ থেকে পাঠানো নতুন রিসিপ্ট ও পেমেন্ট গুগল শিটে বসানো
    if (Array.isArray(payload.receipts)) {
      const rRows = [];
      for (let i = 0; i < 22; i++) {
        const r = payload.receipts[i];
        if (r) rRows.push([i + 1, r.date || '', r.particulars || '', Number(r.amount) || 0]);
        else rRows.push(['', '', '', '']);
      }
      sheet.getRange('C7:F28').setValues(rRows);
    }

    if (Array.isArray(payload.payments)) {
      const pRows = [];
      for (let i = 0; i < 22; i++) {
        const p = payload.payments[i];
        if (p) pRows.push([i + 1, p.date || '', Number(p.voucher) || '', p.particulars || '', Number(p.amount) || 0]);
        else pRows.push(['', '', '', '', '']);
      }
      sheet.getRange('H7:L28').setValues(pRows);
    }

    if (payload.nagadRaw !== undefined) {
      sheet.getRange('Y13').setValue(Number(payload.nagadRaw) || 0);
    }
    if (payload.bkashRaw !== undefined) {
      sheet.getRange('Y14').setValue(Number(payload.bkashRaw) || 0);
    }

    return jsonResponse({
      status: 'ok',
      message: 'Google Sheet synced successfully!',
      syncedAt: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function formatSheetDate(val) {
  if (val instanceof Date) {
    const dd = ('0' + val.getDate()).slice(-2);
    const mm = ('0' + (val.getMonth() + 1)).slice(-2);
    const yy = String(val.getFullYear()).slice(-2);
    return dd + '-' + mm + '-' + yy;
  }
  return String(val || '');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
