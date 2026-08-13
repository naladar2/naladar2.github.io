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

// Структура листа «Товары»:
// A:ШК  B:Название  C:Артикул  D:Код  E:Ячейка  [F:пусто]  G:Взято  H:Буфер
// B/C/D заполняет ARRAYFORMULA автоматически — скрипт их НЕ трогает
// Скрипт пишет только: A (штрихкод), E (ячейка), G (взято)
const COL_BARCODE = 1; // A
const COL_NAME    = 2; // B — ARRAYFORMULA (не перезаписываем)
const COL_CODE    = 3; // C — ARRAYFORMULA
const COL_EXTRA   = 4; // D — ARRAYFORMULA
const COL_CELL    = 5; // E — ячейка склада
const COL_TAKEN   = 6; // F — взято

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

  // Читаем данные ДО очистки
  const values = workingRange.getValues().map(r => r[0]);

  // Очищаем H НЕМЕДЛЕННО — до записи в A/E/G.
  // Это предотвращает расширение «Умной таблицы» вправо:
  // таблица расширяется только если в H остаются данные в момент
  // когда скрипт пишет в соседние столбцы.
  workingRange.clearContent();
  SpreadsheetApp.flush(); // принудительно применяем очистку

  const result = processInputBuffer(sheet, values);

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

  // Список ячеек с количеством товаров в каждой
  if (action === 'getCells') {
    const sheet = getWorkSheet();
    const last = sheet.getLastRow();
    if (last < 3) return { cells: [] };
    const vals = sheet.getRange(2, COL_CELL, last - 1, 1).getValues();
    const counts = {};
    vals.forEach(function(r) {
      const cell = String(r[0]).trim().toUpperCase();
      if (cell) counts[cell] = (counts[cell] || 0) + 1;
    });
    const cells = Object.keys(counts).sort().map(function(k) {
      return { cell: k, count: counts[k] };
    });
    return { cells: cells };
  }

  // Очистить ячейку — удалить все товары из неё
  if (action === 'clearCell') {
    const targetCell = String(params.cell || '').trim().toUpperCase();
    if (!targetCell) return { error: 'Не указана ячейка' };
    const sheet = getWorkSheet();
    const last = sheet.getLastRow();
    if (last < 3) return { ok: true, deleted: 0 };
    // Удаляем все строки где ячейка = targetCell (снизу вверх)
    const cellCol = sheet.getRange(2, COL_CELL, last - 1, 1).getValues();
    let deleted = 0;
    for (var i = cellCol.length - 1; i >= 0; i--) {
      if (String(cellCol[i][0]).trim().toUpperCase() === targetCell) {
        sheet.deleteRow(i + 2);
        deleted++;
      }
    }
    SpreadsheetApp.flush();
    const freshData = readWorkSheetData(sheet);
    return { ok: true, deleted: deleted, data: freshData };
  }

  return { error: 'Unknown action: ' + action };
}

// ── Пакетное удаление строк ──────────────────────────────
// Читаем только A (штрихкод) и F (ячейка) чтобы определить что удалять.
// Строка 2 содержит шаблонные формулы ВПР — её НИКОГДА не перезаписываем.
// Удаляем ровно столько вхождений каждого ключа, сколько передано в deleteItems.
function batchDeleteRows(sheet, deleteItems) {
  const last = sheet.getLastRow();
  if (last < 2) return 0; // данные с строки 2 (ARRAYFORMULA)

  // Считаем сколько раз каждый ключ нужно удалить
  const delCount = new Map();
  deleteItems.forEach(function(it) {
    const key = String(it.b).trim() + '|' + String(it.c).trim().toUpperCase();
    delCount.set(key, (delCount.get(key) || 0) + 1);
  });

  // Читаем строки данных начиная с 3-й (строка 2 = шаблон с формулами)
  // ARRAYFORMULA: данные с строки 2
  const dataStart = 2;
  const dataCount = last - dataStart + 1;
  if (dataCount < 1) return 0;

  // Читаем A (штрихкод) и E (ячейка)
  const barcodeCol = sheet.getRange(dataStart, COL_BARCODE, dataCount, 1).getValues();
  const cellCol    = sheet.getRange(dataStart, COL_CELL,    dataCount, 1).getValues();

  // Определяем какие строки-номера (в таблице) нужно удалить
  // Идём снизу вверх чтобы номера не сдвигались при deleteRow
  const rowsToDelete = [];
  for (var i = dataCount - 1; i >= 0; i--) {
    const bc   = String(barcodeCol[i][0]).trim();
    const cell = String(cellCol[i][0]).trim().toUpperCase();
    if (!bc) continue;

    const key = bc + '|' + cell;
    const remaining = delCount.get(key) || 0;
    if (remaining > 0) {
      delCount.set(key, remaining - 1);
      rowsToDelete.push(dataStart + i); // реальный номер строки в таблице
    }
  }

  // Удаляем строки — deleteRow здесь оправдан т.к. строки уже выбраны,
  // идём снизу вверх (rowsToDelete уже в обратном порядке)
  rowsToDelete.forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });

  return rowsToDelete.length;
}

// Чтение данных рабочего листа — общая функция для 'get' и 'sync' 
function readWorkSheetData(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  // ARRAYFORMULA: данные с строки 2, читаем A..G
  const vals = sheet.getRange(2, 1, last - 1, COL_TAKEN).getValues();
  return vals
    .filter(function(r) {
      // Фильтруем только строки где есть штрихкод в A
      // ARRAYFORMULA может заполнять B/C/D даже для пустых строк — игнорируем их
      return String(r[COL_BARCODE - 1]).trim() !== '';
    })
    .map(function(r) { return [
      String(r[COL_BARCODE - 1]).trim(), // A — штрихкод
      String(r[COL_NAME   - 1]),         // B — название (ARRAYFORMULA)
      String(r[COL_CODE   - 1]),         // C — артикул
      String(r[COL_EXTRA  - 1]),         // D — код материала
      String(r[COL_CELL   - 1]).trim(),  // E — ячейка склада
    ];});
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
    // ARRAYFORMULA в B/C/D заполняется автоматически при добавлении штрихкода в A.
    // Ищем первую пустую строку в A начиная со строки 2.
    let nextRow = 2;
    if (lastRowA >= 2) {
      const aValues = sheet.getRange('A2:A' + lastRowA).getValues();
      for (let j = aValues.length - 1; j >= 0; j--) {
        if (aValues[j][0] !== '') { nextRow = j + 3; break; }
      }
    }
    const count = barcodesToInsert.length;
    // Записываем только штрихкод (A), ячейку (E) и статус (G).
    // B/C/D заполняет ARRAYFORMULA автоматически — не трогаем.
    sheet.getRange(nextRow, COL_BARCODE, count, 1).setValues(barcodesToInsert); // A — штрихкод
    sheet.getRange(nextRow, COL_CELL,    count, 1).setValues(storageCellsToInsert); // E — ячейка
    sheet.getRange(nextRow, COL_TAKEN,   count, 1).setValues(statusesToInsert);     // F — взято

    added = count;
  }

  return { added: added, cellsOnly: hasChanges && onlyCells };
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
