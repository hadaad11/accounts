// Live link between the Excel files on your Drive and the dashboard.
// Setup: script.google.com → New project → paste this file → Deploy → New deployment
//   → type: Web app, Execute as: Me, Who has access: Anyone → copy the Web app URL.
// The URL only ever lives inside the encrypted dashboard, never in the repo.
const FILES = {
  receivables: 'حسابات العملاء.xlsx',
  work: 'اسعار الشغل.xlsx',
  suppliers: 'اسعار البنود للموردين.xlsx',
};

function doGet(e) {
  const known = {};
  String((e && e.parameter.known) || '').split(',').forEach(p => {
    const [id, stamp] = p.split(':');
    if (id) known[id] = stamp;
  });
  const files = {};
  for (const [id, name] of Object.entries(FILES)) {
    let file = null;
    const it = DriveApp.getFilesByName(name);
    while (it.hasNext()) {
      const f = it.next();
      if (!f.isTrashed() && (!file || f.getLastUpdated() > file.getLastUpdated())) file = f;
    }
    if (!file) continue;
    const updated = file.getLastUpdated();
    const stamp = String(updated.getTime());
    files[id] = {
      stamp,
      modified: Utilities.formatDate(updated, 'Africa/Cairo', 'yyyy-MM-dd HH:mm'),
      edit_url: 'https://docs.google.com/spreadsheets/d/' + file.getId() + '/edit',
    };
    if (known[id] !== stamp) files[id].b64 = Utilities.base64Encode(file.getBlob().getBytes());
  }
  return ContentService.createTextOutput(JSON.stringify({files})).setMimeType(ContentService.MimeType.JSON);
}
