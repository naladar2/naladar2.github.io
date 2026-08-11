/**
 * ╔══════════════════════════════════════════════════════╗
 * ║  СКЛАД — Google Apps Script                          ║
 * ╚══════════════════════════════════════════════════════╝
 *
 * Листы:
 *   Лист1   — рабочий: A:ШК  B:Название  C:Артикул  D:Код материала
 *                       E:Ячейка  F:Взято  H:буфер ввода (сканер/приложение)
 *   Остатки — справочник: A:ШК  B:Название  C:Артикул  D:Код материала
 *
 * Формат ячейки поддерживает оба варианта:
 *   простой:     A1, K12, JD11
 *   со стеллажом: K1/1, K10/76   (буквы+цифры/цифры)
 *
 * ВАЖНО: setValues() из веб-запроса НЕ вызывает onEdit (это триггер только
 * на ручной ввод человеком). Поэтому логика разбора буфера H продублирована
 * здесь и вызывается напрямую из processInputBuffer().
 * Если данные вставляются в H вручную (или сканером напрямую в таблицу) —
 * сработает onEdit(e) ниже, использующий ту же функцию.
 *
 * КАК ПРОВЕРИТЬ: запустите testScript() — НЕ doGet!
 */

const WORK_SHEET = 'Товары';
const REF_SHEET  = 'Остатки';
const INPUT_COL  = 8; // H

// Колонки рабочего листа (1-индексация, как в Google Sheets)
const COL_BARCODE = 1; // A
const COL_NAME     = 2; // B
const COL_CODE     = 3; // C — артикул
const COL_EXTRA    = 4; // D — код материала
const COL_CELL     = 5; // E — ячейка
const COL_TAKEN    = 6; // F — взято

// ── ТЕСТ ────────────────────────────────────────────────
function testScript() {
  Logger.log('=== ping ===');
  Logger.log(JSON.stringify(processAction('ping', {})));
  Logger.log('=== get ===');
  Logger.log(JSON.stringify(processAction('get', {})));
}

// ── Точка входа для веб-запросов ─────────────────────────
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || 'get';
    return jsonResponse(processAction(action, params));
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ── Триггер: срабатывает при РУЧНОМ редактировании ячеек ─
// (вставка через буфер обмена, ввод со сканера напрямую в Sheets)
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== WORK_SHEET) return;

  const range = e.range;
  if (range.getColumn() !== INPUT_COL || range.getLastRow() <= 1) return;

  const rowStart = Math.max(2, range.getRow());
  const rowEnd = range.getLastRow();
  const numRows = rowEnd - rowStart + 1;
  const workingRange = sheet.getRange(rowStart, INPUT_COL, numRows, 1);
  const values = workingRange.getValues().map(r => r[0]);

  const result = processInputBuffer(sheet, values);
  workingRange.clearContent();

  if (result.added > 0) {
    sheet.getParent().toast('Успешно обработано строк: ' + result.added, 'Система ввода', 3);
  } else if (result.cellsOnly) {
    sheet.getParent().toast('Обновлён адрес ячейки склада', 'Система ввода', 3);
  }
}

// ── Логика API ────────────────────────────────────────────
function processAction(action, params) {

  if (action === 'ping') {
    return { ok: true, message: 'Склад API работает' };
  }

  // Загрузить список товаров из Лист1
  if (action === 'get') {
    SpreadsheetApp.flush(); // на случай незакоммиченных изменений от других вызовов
    const sheet = getWorkSheet();
    const data = readWorkSheetData(sheet);
    return { data: data, count: data.length };
  }

  // Загрузить справочник с листа Остатки
  if (action === 'getCatalog') {
    const sheetName = decodeURIComponent(params.sheet || REF_SHEET);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: 'Лист не найден: ' + sheetName };
    const last = sheet.getLastRow();
    if (last < 2) return { data: [], count: 0 };
    const cols = Math.min(sheet.getLastColumn(), 4);
    const vals = sheet.getRange(2, 1, last - 1, cols).getValues();
    const data = vals.filter(r => String(r[0]).trim() !== '').map(r => r.map(String));
    return { data: data, count: data.length };
  }

  // Главный экшен синхронизации:
  // 1) удаляет взятые строки (по barcode+cell)
  // 2) пишет новые строки через ту же логику что и onEdit-макрос (ячейка → штрихкоды)
  if (action === 'sync') {
    const lines       = JSON.parse(params.lines       || '[]');
    const deleteItems = JSON.parse(params.deleteItems || '[]');
    const sheet = getWorkSheet();

    let deleted = 0;
    if (deleteItems.length > 0) {
      deleted = batchDeleteRows(sheet, deleteItems);
    }

    const result = lines.length > 0 ? processInputBuffer(sheet, lines) : { added: 0 };

    SpreadsheetApp.flush();
    const freshData = readWorkSheetData(sheet);
    return { ok: true, deleted: deleted, added: result.added, data: freshData, count: freshData.length };
  }

  return { error: 'Unknown action: ' + action };
}

// ── Пакетное удаление строк ──────────────────────────────
// Читаем всё одним getValues(), фильтруем в памяти, пишем обратно одним
// setValues(). O(1) запросов к API вместо O(N) — работает в разы быстрее.
//
// ВАЖНО: удаляем ровно столько вхождений каждого ключа, сколько передано
// в deleteItems. Если передан 1 товар "4607|A1" — удалится только 1 строка,
// даже если в таблице таких строк несколько (50 линеек в одной ячейке).
function batchDeleteRows(sheet, deleteItems) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;

  // Считаем сколько раз каждый ключ нужно удалить (Map: ключ → счётчик)
  const delCount = new Map();
  deleteItems.forEach(function(it) {
    const key = String(it.b).trim() + '|' + String(it.c).trim().toUpperCase();
    delCount.set(key, (delCount.get(key) || 0) + 1);
  });

  // Читаем ВСЕ данные одним запросом
  const numCols = COL_TAKEN;
  const allRows = sheet.getRange(2, 1, last - 1, numCols).getValues();

  const keepRows = [];
  let deleted = 0;

  for (var i = 0; i < allRows.length; i++) {
    const bc   = String(allRows[i][COL_BARCODE - 1]).trim();
    const cell = String(allRows[i][COL_CELL - 1]).trim().toUpperCase();
    if (!bc) { keepRows.push(allRows[i]); continue; }

    const key = bc + '|' + cell;
    const remaining = delCount.get(key) || 0;

    if (remaining > 0) {
      // Удаляем эту строку и уменьшаем счётчик
      delCount.set(key, remaining - 1);
      deleted++;
    } else {
      keepRows.push(allRows[i]);
    }
  }

  if (deleted === 0) return 0;

  if (keepRows.length > 0) {
    sheet.getRange(2, 1, keepRows.length, numCols).setValues(keepRows);
  }

  // Очищаем хвост
  const clearStart = keepRows.length + 2;
  if (clearStart <= last) {
    sheet.getRange(clearStart, 1, last - clearStart + 1, numCols).clearContent();
  }

  return deleted;
}

// Чтение данных рабочего листа — общая функция для 'get' и 'sync' 
function readWorkSheetData(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const vals = sheet.getRange(2, 1, last - 1, COL_TAKEN).getValues();
  return vals
    .filter(r => String(r[COL_BARCODE-1]).trim() !== '')
    .map(r => [
      String(r[COL_BARCODE-1]).trim(),
      String(r[COL_NAME-1]),
      String(r[COL_CODE-1]),
      String(r[COL_EXTRA-1]),
      String(r[COL_CELL-1]).trim(),
    ]);
}

// ── Общая логика разбора буфера (используется и onEdit, и API) ─────
// values — массив строк: ['A1', '4607...', '4602...', 'K1/1', '4670...']
// Ячейка — буквы + цифры, или буквы+цифры/цифры (K1/1, K10/76)
// Возвращает {added, cellsOnly}
function processInputBuffer(sheet, values) {
  const documentProperties = PropertiesService.getDocumentProperties();
  let savedCell = documentProperties.getProperty('LAST_STORAGE_CELL') || '';

  const barcodesToInsert = [];
  const storageCellsToInsert = [];
  const statusesToInsert = [];
  let hasChanges = false;
  let onlyCells = true;

  for (let i = 0; i < values.length; i++) {
    const inputValue = String(values[i]).trim();
    if (inputValue === '') continue;
    hasChanges = true;

    if (isCellAddress(inputValue)) {
      savedCell = inputValue.toUpperCase();
      documentProperties.setProperty('LAST_STORAGE_CELL', savedCell);
    } else {
      onlyCells = false;
      if (savedCell === '') {
        // Пропускаем штрихкод без указанной ячейки (не блокируем весь импорт)
        continue;
      }
      barcodesToInsert.push([inputValue]);
      storageCellsToInsert.push([savedCell]);
      statusesToInsert.push([false]);
    }
  }

  let added = 0;
  if (barcodesToInsert.length > 0) {
    const lastRowA = sheet.getLastRow();
    let nextRow = 2;
    if (lastRowA >= 2) {
      const aValues = sheet.getRange('A1:A' + lastRowA).getValues();
      for (let j = aValues.length - 1; j >= 0; j--) {
        if (aValues[j][0] !== '') { nextRow = j + 2; break; }
      }
    }
    const count = barcodesToInsert.length;
    sheet.getRange(nextRow, COL_BARCODE, count, 1).setValues(barcodesToInsert);      // A
    sheet.getRange(nextRow, COL_CELL,    count, 1).setValues(storageCellsToInsert);  // E
    sheet.getRange(nextRow, COL_TAKEN,   count, 1).setValues(statusesToInsert);      // F

    // Протягиваем формулы VLOOKUP из строки 2 на новые строки (B, C, D)
    // Это нужно чтобы название/артикул/код подтянулись из листа «Остатки»
    extendFormulas(sheet, nextRow, count);

    added = count;
  }

  return { added: added, cellsOnly: hasChanges && onlyCells };
}

// Протягивает формулы из строки 2 вниз на newRows строк начиная с startRow
// Если в B2, C2, D2 есть формулы — копирует их на новые строки
function extendFormulas(sheet, startRow, count) {
  try {
    // Проверяем есть ли формулы в строке 2 (шаблонная строка)
    const formulaCols = [COL_NAME, COL_CODE, COL_EXTRA]; // B, C, D
    formulaCols.forEach(function(col) {
      const templateCell = sheet.getRange(2, col);
      const formula = templateCell.getFormula();
      if (!formula) return; // нет формулы — пропускаем

      // Копируем формулу на все новые строки
      for (var r = 0; r < count; r++) {
        const targetRow = startRow + r;
        if (targetRow === 2) continue; // не перезаписываем шаблон
        // Формула из строки 2 с заменой номера строки
        const newFormula = formula.replace(/(\$?[A-Z]+)2/g, function(match, colRef) {
          // Заменяем только незакреплённые строки (без $)
          if (match.startsWith('$')) return match;
          return colRef + targetRow;
        });
        sheet.getRange(targetRow, col).setFormula(newFormula);
      }
    });
  } catch(e) {
    Logger.log('extendFormulas error: ' + e);
    // Не критично — данные уже записаны
  }
}

// Адрес ячейки склада. Поддерживает два формата:
//   простой:      буквы(1-3) + цифры(1-4)         напр. A1, K12, JD11
//   со стеллажом:  буквы(1-3) + цифры/цифры        напр. K1/1, K10/76
// Не должен совпадать с обычным штрихкодом (только цифры).
function isCellAddress(value) {
  // Совпадает с логикой onEdit-макроса пользователя:
  // ячейка = не является числом + длина <= 8 символов
  // Дополнительно поддерживаем формат K1/1, K10/76
  if (!value || value.length > 10) return false;
  if (!isNaN(value)) return false; // чистое число — штрихкод
  // Должно начинаться с буквы
  if (!/^[A-Za-zА-Яа-я]/.test(value)) return false;
  // Паттерны: A1, K12, JD11, K1/1, K10/76
  const simple   = /^[A-Za-zА-Яа-я]{1,3}\d{1,4}$/;
  const fraction = /^[A-Za-zА-Яа-я]{1,3}\d{1,4}\/\d{1,4}$/;
  return simple.test(value) || fraction.test(value);
}

// ── Helpers ──────────────────────────────────────────────
function getWorkSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(WORK_SHEET);
  if (!sheet) throw new Error('Лист не найден: ' + WORK_SHEET);
  return sheet;
}

function jsonResponse(obj) {
  const out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

function doPost(e) {
  return jsonResponse({ error: 'Use GET requests' });
}
