// นำไฟล์นี้ไปใส่ใน Apps Script ที่ผูกกับ Google Sheet ผลตอบกลับ
const VOTE_SHEET = 'Form_Responses';
const CANDIDATE_SHEET = 'DashboardData';
const VOTE_HEADER = 'นายกในใจคุณ';

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const votes = ss.getSheetByName(VOTE_SHEET);
  const candidates = ss.getSheetByName(CANDIDATE_SHEET);
  if (!votes) return json({ error: 'ไม่พบชีต ' + VOTE_SHEET });
  const rows = votes.getDataRange().getValues();
  const headers = rows.shift().map(String);
  const voteIndex = headers.findIndex(x => x.trim() === VOTE_HEADER);
  if (voteIndex < 0) return json({ error: 'ไม่พบคอลัมน์ ' + VOTE_HEADER });
  const candidateMap = {};
  if (candidates) candidates.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) candidateMap[String(r[0])] = { no: String(r[0]), name: r[1] || 'ผู้สมัคร ' + r[0], team: r[2] || '', image: r[4] || '' }; });
  rows.forEach(r => { const m = String(r[voteIndex]).match(/\d+/); if (m) { const no = m[0]; candidateMap[no] ||= { no, name: 'ผู้สมัครหมายเลข ' + no, team: '', image: '' }; candidateMap[no].votes = (candidateMap[no].votes || 0) + 1; } });
  const list = Object.values(candidateMap).map(c => ({ ...c, votes: c.votes || 0 })).sort((a,b) => b.votes - a.votes);
  return json({ ok: true, lastUpdate: new Date().toISOString(), total: rows.length, candidates: list });
}
function json(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
