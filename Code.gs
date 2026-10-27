// Google Apps Script для приложения «Учёт».
// Вставьте в Расширения → Apps Script внутри вашей Google Таблицы.
const SECRET = 'ЗАМЕНИТЕ_НА_СВОЙ_ПАРОЛЬ';   // придумайте пароль (латиница и цифры, в кавычках) и впишите такой же в приложении
// После любого изменения кода: Развернуть → Управление развёртываниями → карандаш → Версия «Новая версия» → Развернуть.
// Ссылка /exec останется прежней. «Новое развёртывание» создаст НОВУЮ ссылку.
const H = ['id','date','amount','name','account','sub','kind','person'];

// Разбивка по умолчанию. Лист «Настройки» создаётся сам при первом открытии приложения.
// Каждая строка — кнопка при разборе: Группа (1-й шаг) · Категория (2-й шаг) · Счёт (откуда списывается).
// Первая группа считается главной (её реальный расход крупно в «Анализе»).
const DEF_SET = [
  ['Личное','Еда','Личное'],['Личное','Дорога','Личное'],['Личное','Развлечения','Личное'],
  ['Личное','Одежда','Личное'],['Личное','Другое','Личное'],
  ['Бизнес','Доставка','Доставка'],['Бизнес','Бонус','Бонус'],['Бизнес','Штраф','Штраф'],
  ['Бизнес','На сайт','Доставка'],['Бизнес','На офис','Доставка']
];
function settings_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('Настройки');
  if(!s){
    s = ss.insertSheet('Настройки');
    s.getRange(1,1,1,3).setValues([['Группа','Категория','Счёт']]).setFontWeight('bold');
    s.getRange(2,1,DEF_SET.length,3).setValues(DEF_SET);
    s.setFrozenRows(1);
    s.autoResizeColumns(1,3);
  }
  return s.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim() && String(r[1]).trim() && String(r[2]).trim())
    .map(r => ({g:String(r[0]).trim(), c:String(r[1]).trim(), a:String(r[2]).trim()}));
}

function sheet_(name, head){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if(!s){
    s = ss.insertSheet(name);
    s.appendRow(head);
    s.setFrozenRows(1);
    if(name === 'Операции') s.getRange('A:B').setNumberFormat('@');
  }
  return s;
}
function fd_(v){
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
}
function out_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e){
  if(!e || !e.parameter || e.parameter.key !== SECRET) return out_({error:'key'});
  const v = sheet_('Операции', H).getDataRange().getValues().slice(1);
  const tx = v.filter(r => r[0]).map(r => ({
    id:String(r[0]), d:fd_(r[1]), a:Number(r[2]), n:String(r[3]),
    acc:String(r[4]), sub:String(r[5]), k:String(r[6]), p:String(r[7])
  }));
  const acc = sheet_('Счета', ['name','opening']).getDataRange().getValues().slice(1)
    .filter(r => r[0]).map(r => ({n:String(r[0]), o:Number(r[1]) || 0}));
  return out_({acc:acc, tx:tx, cfg:settings_()});
}
function doPost(e){
  let b;
  try { b = JSON.parse(e.postData.contents); } catch(x){ return out_({error:'json'}); }
  if(b.key !== SECRET) return out_({error:'key'});
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return out_({error:'busy'});
  try {
    if(b.action === 'add'){
      const s = sheet_('Операции', H), last = s.getLastRow();
      const have = new Set(last > 1 ? s.getRange(2,1,last-1,1).getValues().map(r => String(r[0])) : []);
      const rows = (b.tx || []).filter(t => !have.has(t.id))
        .map(t => [t.id, t.d, t.a, t.n || '', t.acc || '', t.sub || '', t.k || '', t.p || '']);
      if(rows.length) s.getRange(last+1, 1, rows.length, 8).setValues(rows);
    } else if(b.action === 'del'){
      const s = sheet_('Операции', H), last = s.getLastRow();
      if(last > 1){
        const ids = s.getRange(2,1,last-1,1).getValues();
        for(let i = ids.length - 1; i >= 0; i--){
          if(String(ids[i][0]) === String(b.id)){ s.deleteRow(i + 2); break; }
        }
      }
    } else if(b.action === 'acc'){
      const s = sheet_('Счета', ['name','opening']);
      if(s.getLastRow() > 1) s.getRange(2,1,s.getLastRow()-1,2).clearContent();
      const rows = (b.acc || []).map(a => [a.n, a.o || 0]);
      if(rows.length) s.getRange(2,1,rows.length,2).setValues(rows);
    }
    return out_({ok:true});
  } finally {
    lock.releaseLock();
  }
}
