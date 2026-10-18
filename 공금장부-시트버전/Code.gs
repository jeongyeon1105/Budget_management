// 공금 장부: 구글 시트(내역·설정)에 저장하고, 영수증 사진은 드라이브 폴더에 보관한다.
const ITEM_HEAD = ['id', '날짜', '목적', '결제한 사람', '금액', '영수증'];
const FOLDER_NAME = '예산사용내역';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('예산사용내역')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 시트가 없으면 만든다. 텍스트 열은 '@' 서식으로 두어 날짜 변환과 수식 실행을 막는다.
function sheets_() {
  const ss = SpreadsheetApp.getActive();
  let items = ss.getSheetByName('내역');
  if (!items) {
    items = ss.insertSheet('내역');
    items.getRange('A:D').setNumberFormat('@');
    items.getRange('F:F').setNumberFormat('@');
    items.appendRow(ITEM_HEAD);
    items.setFrozenRows(1);
  }
  let cfg = ss.getSheetByName('설정');
  if (!cfg) {
    cfg = ss.insertSheet('설정');
    cfg.getRange('B2').setNumberFormat('@');
    cfg.getRange('A1:B2').setValues([['총 예산', 0], ['접속 코드 (비우면 코드 없이 사용)', '']]);
  }
  return {items: items, cfg: cfg};
}

function checkCode_(cfg, code) {
  const need = String(cfg.getRange('B2').getValue()).trim();
  if (need && need !== String(code || '').trim()) throw new Error('BAD_CODE');
}

function read_(s) {
  const tz = Session.getScriptTimeZone();
  const rows = s.items.getDataRange().getValues().slice(1).filter(function (r) { return r[0]; });
  return {
    total: Number(s.cfg.getRange('B1').getValue()) || 0,
    items: rows.map(function (r) {
      return {
        id: r[0],
        date: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
        purpose: String(r[2]),
        payer: String(r[3]),
        amount: Number(r[4]) || 0,
        receipt: r[5] || ''
      };
    })
  };
}

// 모든 호출의 공통 흐름: 코드 확인 → (쓰기면 잠금) → 작업 → 최신 데이터 반환
function run_(code, write, fn) {
  const lock = LockService.getScriptLock();
  if (write) lock.waitLock(20000);
  try {
    const s = sheets_();
    checkCode_(s.cfg, code);
    fn(s);
    return read_(s);
  } finally {
    if (write) lock.releaseLock();
  }
}

function getData(code) {
  return run_(code, false, function () {});
}

function saveBudget(code, total) {
  const n = Math.floor(Number(total));
  if (!(n >= 0)) throw new Error('예산은 0 이상의 숫자여야 합니다.');
  return run_(code, true, function (s) { s.cfg.getRange('B1').setValue(n); });
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function trash_(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* 이미 없으면 무시 */ }
}

// item.id가 있으면 수정, 없으면 추가. image는 JPEG data URL, removeReceipt는 영수증만 지우기.
function saveItem(code, item, image, removeReceipt) {
  const date = String(item.date || '');
  const purpose = String(item.purpose || '').trim().slice(0, 200);
  const payer = String(item.payer || '').trim().slice(0, 100);
  const amount = Math.floor(Number(item.amount));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !purpose || !payer || !(amount > 0)) throw new Error('입력값을 확인해 주세요.');
  if (image && !/^data:image\/jpeg;base64,/.test(image)) throw new Error('JPEG 이미지만 올릴 수 있습니다.');

  return run_(code, true, function (s) {
    const data = s.items.getDataRange().getValues();
    const row = item.id ? data.findIndex(function (r, i) { return i > 0 && r[0] === item.id; }) + 1 : 0;
    if (item.id && !row) throw new Error('이미 삭제된 내역입니다.');
    let receipt = row ? data[row - 1][5] : '';
    if (image || removeReceipt) { trash_(receipt); receipt = ''; }
    if (image) {
      const blob = Utilities.newBlob(Utilities.base64Decode(image.split(',')[1]), 'image/jpeg', 'receipt_' + Date.now() + '.jpg');
      receipt = folder_().createFile(blob).getId();
    }
    const vals = [[item.id || Utilities.getUuid(), date, purpose, payer, amount, receipt, new Date()]];
    s.items.getRange(row || s.items.getLastRow() + 1, 1, 1, 7).setValues(vals);
  });
}

function deleteItem(code, id) {
  return run_(code, true, function (s) {
    const data = s.items.getDataRange().getValues();
    const row = data.findIndex(function (r, i) { return i > 0 && r[0] === id; }) + 1;
    if (!row) return;
    trash_(data[row - 1][5]);
    s.items.deleteRow(row);
  });
}

// 시트에 기록된 영수증만 내려준다(드라이브의 다른 파일은 열 수 없게).
function getReceipt(code, fileId) {
  const s = sheets_();
  checkCode_(s.cfg, code);
  const known = s.items.getDataRange().getValues().some(function (r) { return r[5] && r[5] === fileId; });
  if (!known) throw new Error('영수증을 찾을 수 없습니다.');
  const b = DriveApp.getFileById(fileId).getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}
