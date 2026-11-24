/**
 * Дневник тренировок для колена: веб-приложение для синхронизации со страницей.
 * Хранит записи в двух листах этой таблицы:
 *   _data   служебный лист, не редактируйте его вручную;
 *   Дневник читаемая таблица: одна строка на каждое упражнение каждой тренировки.
 * Инструкция по установке: README.md рядом с этим файлом.
 */
const TOKEN = 'ЗАМЕНИТЕ_НА_СВОЙ_ПАРОЛЬ';
const DATA_SHEET = '_data';
const VIEW_SHEET = 'Дневник';
const WD = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function doGet() {
  return out_({ ok: true, info: 'Дневник тренировок: веб-приложение работает' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (TOKEN === 'ЗАМЕНИТЕ_НА_СВОЙ_ПАРОЛЬ' || req.token !== TOKEN) {
      return out_({ ok: false, error: 'неверный пароль' });
    }
    const data = sheet_(DATA_SHEET);
    if (data.getLastRow() === 0) data.appendRow(['ts', 'deleted', 'json']);

    const rows = data.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < rows.length; i++) rowOf[String(rows[i][0])] = i + 1;

    // новые записи со страницы (уже сохранённые и удалённые не трогаем)
    (req.entries || []).forEach(function (en) {
      const k = String(en.ts);
      if (!en.ts || rowOf[k]) return;
      data.appendRow([en.ts, false, JSON.stringify(en)]);
      rowOf[k] = data.getLastRow();
    });

    // удаления: помечаем строку, чтобы запись не вернулась с другого устройства
    (req.deleted || []).forEach(function (ts) {
      const k = String(ts);
      if (rowOf[k]) {
        data.getRange(rowOf[k], 2).setValue(true);
      } else {
        data.appendRow([ts, true, '']);
        rowOf[k] = data.getLastRow();
      }
    });

    const all = data.getDataRange().getValues();
    const entries = [];
    const deleted = [];
    for (let i = 1; i < all.length; i++) {
      const isDeleted = all[i][1] === true || String(all[i][1]).toUpperCase() === 'TRUE';
      if (isDeleted) deleted.push(Number(all[i][0]));
      else if (all[i][2]) entries.push(JSON.parse(all[i][2]));
    }
    renderView_(entries);
    return out_({ ok: true, entries: entries, deleted: deleted });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function renderView_(entries) {
  const sh = sheet_(VIEW_SHEET);
  sh.clearContents();
  const head = ['Дата', 'День', 'Упражнение', 'Подходы', 'Повторения', 'Вес', 'Боль (макс. из 10)', 'Контроль (из 3)', 'Заметка'];
  const rows = [];
  entries.slice().sort(function (a, b) {
    return a.date === b.date ? b.ts - a.ts : (a.date < b.date ? 1 : -1);
  }).forEach(function (en) {
    const p = String(en.date).split('-');
    const d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    const ok = (en.checks || []).filter(Boolean).length;
    (en.items || []).forEach(function (it) {
      rows.push([d, WD[en.wd] || '', it.name || it.id, it.sets, it.reps, it.weight,
        (en.pain === null || en.pain === undefined) ? '' : en.pain, ok, en.note || '']);
    });
  });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('dd.mm.yyyy');
  }
  sh.setFrozenRows(1);
}
