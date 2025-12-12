function coaches(e) {
  // Проверка наличия объекта события
  if (!e || !e.source) {
    Logger.log(`Ошибка: объект события не определен - ${JSON.stringify({ e: !!e, source: !!e?.source })}`);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  if (sheetName !== 'Тренеры и команды') return;

  const colMap = {
    A: 1, B: 2, C: 3, D: 4, G: 7, H: 8, I: 9, J: 10, K: 11, L: 12, M: 13, O: 15
  };

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(...Object.values(colMap));
  if (lastRow < 2) {
    Logger.log('Нет данных для обработки (lastRow < 2)');
    return;
  }

  // Загрузка данных из листа "Силы клубов"
  const clubsSheet = ss.getSheetByName('Силы клубов');
  const clubsData = clubsSheet ? clubsSheet.getRange('A2:C' + clubsSheet.getLastRow()).getValues() : [];
  const clubMap = {}; // A -> C (для D)
  clubsData.forEach(row => {
    if (row[0]) clubMap[row[0].toString().trim()] = row[2] || '';
  });
  Logger.log(`clubMap: ${Object.keys(clubMap).length} записей, пример: ${JSON.stringify(Object.entries(clubMap).slice(0, 5))}`);

  // Загрузка данных из листа "Глосс."
  const glossSheet = ss.getSheetByName('Глосс.');
  const glossData = glossSheet ? glossSheet.getRange('C2:O' + glossSheet.getLastRow()).getValues() : [];
  Logger.log(`Загружено ${glossData.length} строк из "Глосс."`);

  // Загрузка данных из листа "Тренеры и команды"
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  let data = dataRange.getValues();
  Logger.log(`Загружено ${data.length} строк из "Тренеры и команды"`);

  // Создание словарей для ID (A и B) из PropertiesService
  const properties = PropertiesService.getScriptProperties();
  let cToId = JSON.parse(properties.getProperty('cToId') || '{}');
  let gToId = JSON.parse(properties.getProperty('gToId') || '{}');
  let maxCId = Math.max(0, ...Object.values(cToId).map(Number));
  let maxGId = Math.max(0, ...Object.values(gToId).map(Number));

  // Создание индекса для computeH
  const gIndex = {};
  data.forEach((row, idx) => {
    const g_val = row[colMap.G - 1]?.toString().trim() || '';
    if (g_val) {
      if (!gIndex[g_val]) gIndex[g_val] = [];
      gIndex[g_val].push({ i: row[colMap.I - 1], idx: idx });
    }
  });
  Logger.log(`gIndex: ${Object.keys(gIndex).length} записей, пример: ${JSON.stringify(Object.entries(gIndex).slice(0, 5))}`);

  // Оптимизированный BATCH_SIZE
  const BATCH_SIZE = 500;

  // Функции вычислений
  function computeA(rowData, rowIdx) {
    const c_val = rowData[colMap.C - 1]?.toString().trim() || '';
    if (!c_val) {
      Logger.log(`computeA[строка ${rowIdx + 2}]: c_val пустое, возвращается ''`);
      return '';
    }
    if (!cToId[c_val]) {
      maxCId += 1;
      cToId[c_val] = maxCId;
      Logger.log(`computeA[строка ${rowIdx + 2}]: новое c_val=${c_val}, присвоен ID=${maxCId}`);
    }
    const result = cToId[c_val];
    Logger.log(`computeA[строка ${rowIdx + 2}]: c_val=${c_val}, result=${result}`);
    return result;
  }

  function computeB(rowData, rowIdx) {
    const c_val = rowData[colMap.C - 1]?.toString().trim() || '';
    const g_val = rowData[colMap.G - 1]?.toString().trim() || '';
    if (!c_val || !g_val) {
      Logger.log(`computeB[строка ${rowIdx + 2}]: c_val=${c_val}, g_val=${g_val} - пустое, возвращается ''`);
      return '';
    }
    if (!gToId[g_val]) {
      maxGId += 1;
      gToId[g_val] = maxGId;
      Logger.log(`computeB[строка ${rowIdx + 2}]: новое g_val=${g_val}, присвоен ID=${maxGId}`);
    }
    const result = gToId[g_val];
    Logger.log(`computeB[строка ${rowIdx + 2}]: g_val=${g_val}, result=${result}`);
    return result;
  }

  function computeD(rowData, rowIdx) {
    const c_val = rowData[colMap.C - 1]?.toString().trim() || '';
    if (!c_val) {
      Logger.log(`computeD[строка ${rowIdx + 2}]: c_val пустое, возвращается ''`);
      return '';
    }
    const result = clubMap[c_val] || '';
    Logger.log(`computeD[строка ${rowIdx + 2}]: c_val=${c_val}, result=${result}`);
    return result;
  }

  function computeH(rowData, rowIdx) {
    const i_val = rowData[colMap.I - 1];
    const g_val = rowData[colMap.G - 1]?.toString().trim() || '';
    if (!i_val || !g_val || !(i_val instanceof Date) || isNaN(i_val.getTime())) {
      Logger.log(`computeH[строка ${rowIdx + 2}]: i_val=${i_val}, g_val=${g_val} - пропуск (невалидная дата или g_val)`);
      return '';
    }
    const g_rows = gIndex[g_val] || [];
    let count = 0;
    for (const row of g_rows) {
      if (row.i && row.i instanceof Date && !isNaN(row.i.getTime()) && row.i < i_val) {
        count += 1;
      }
    }
    const result = count + 1;
    Logger.log(`computeH[строка ${rowIdx + 2}]: i_val=${i_val}, g_val=${g_val}, g_rows=${g_rows.length}, count=${count}, result=${result}`);
    return result;
  }

  function computeK(rowData, rowIdx) {
    const g_val = rowData[colMap.G - 1]?.toString().trim() || '';
    const d_val = rowData[colMap.D - 1]?.toString().trim() || '';
    if (!g_val) {
      Logger.log(`computeK[строка ${rowIdx + 2}]: g_val пустое, возвращается ''`);
      return '';
    }
    let count = 0;
    for (const row of glossData) {
      const o_val = row[12]?.toString().trim() || ''; // 'Глосс.'!O (индекс 12)
      if (o_val && o_val === d_val) {
        count += 1;
      }
    }
    const result = count > 0;
    Logger.log(`computeK[строка ${rowIdx + 2}]: g_val=${g_val}, d_val=${d_val}, count=${count}, result=${result}`);
    return result;
  }

  function computeL(rowData, rowIdx) {
    const g_val = rowData[colMap.G - 1]?.toString().trim() || '';
    const k_val = rowData[colMap.K - 1];
    const i_val = rowData[colMap.I - 1];
    if (!g_val) {
      Logger.log(`computeL[строка ${rowIdx + 2}]: g_val пустое, возвращается ''`);
      return '';
    }
    if (!(i_val instanceof Date) || isNaN(i_val.getTime())) {
      Logger.log(`computeL[строка ${rowIdx + 2}]: i_val=${i_val} невалидная дата, возвращается ''`);
      return '';
    }
    let sum = 0;
    if (k_val === false) {
      for (const row of glossData) {
        const c_val = row[0]; // 'Глосс.'!C
        const e_val = row[2]; // 'Глосс.'!E
        const f_val = row[3]; // 'Глосс.'!F
        if (e_val instanceof Date && !isNaN(e_val.getTime()) &&
            f_val instanceof Date && !isNaN(f_val.getTime()) &&
            i_val >= e_val && i_val <= f_val && typeof c_val === 'number') {
          sum += c_val;
        }
      }
    } else {
      for (const row of glossData) {
        const h_val = row[5]; // 'Глосс.'!H
        const j_val = row[7]; // 'Глосс.'!J
        const k_val_gloss = row[8]; // 'Глосс.'!K
        if (j_val instanceof Date && !isNaN(j_val.getTime()) &&
            k_val_gloss instanceof Date && !isNaN(k_val_gloss.getTime()) &&
            i_val >= j_val && i_val <= k_val_gloss && typeof h_val === 'number') {
          sum += h_val;
        }
      }
    }
    Logger.log(`computeL[строка ${rowIdx + 2}]: g_val=${g_val}, k_val=${k_val}, i_val=${i_val}, sum=${sum}`);
    return sum;
  }

  function computeM(rowData, rowIdx) {
    const g_val = rowData[colMap.G - 1]?.toString().trim() || '';
    const k_val = rowData[colMap.K - 1];
    const j_val = rowData[colMap.J - 1];
    if (!g_val) {
      Logger.log(`computeM[строка ${rowIdx + 2}]: g_val пустое, возвращается ''`);
      return '';
    }
    if (!(j_val instanceof Date) || isNaN(j_val.getTime())) {
      Logger.log(`computeM[строка ${rowIdx + 2}]: j_val=${j_val} невалидная дата, возвращается ''`);
      return '';
    }
    let sum = 0;
    if (k_val === false) {
      for (const row of glossData) {
        const c_val = row[0]; // 'Глосс.'!C
        const e_val = row[2]; // 'Глосс.'!E
        const f_val = row[3]; // 'Глосс.'!F
        if (e_val instanceof Date && !isNaN(e_val.getTime()) &&
            f_val instanceof Date && !isNaN(f_val.getTime()) &&
            j_val >= e_val && j_val <= f_val && typeof c_val === 'number') {
          sum += c_val;
        }
      }
    } else {
      for (const row of glossData) {
        const h_val = row[5]; // 'Глосс.'!H
        const j_val_gloss = row[7]; // 'Глосс.'!J
        const k_val_gloss = row[8]; // 'Глосс.'!K
        if (j_val_gloss instanceof Date && !isNaN(j_val_gloss.getTime()) &&
            k_val_gloss instanceof Date && !isNaN(k_val_gloss.getTime()) &&
            j_val >= j_val_gloss && j_val <= k_val_gloss && typeof h_val === 'number') {
          sum += h_val;
        }
      }
    }
    Logger.log(`computeM[строка ${rowIdx + 2}]: g_val=${g_val}, k_val=${k_val}, j_val=${j_val}, sum=${sum}`);
    return sum;
  }

  function computeO(rowData, rowIdx) {
    const j_val = rowData[colMap.J - 1];
    const i_val = rowData[colMap.I - 1];
    if (!(j_val instanceof Date) || isNaN(j_val.getTime()) ||
        !(i_val instanceof Date) || isNaN(i_val.getTime())) {
      Logger.log(`computeO[строка ${rowIdx + 2}]: i_val=${i_val}, j_val=${j_val} - невалидные даты, возвращается ''`);
      return '';
    }
    const result = (j_val - i_val) / (1000 * 60 * 60 * 24); // Разница в днях
    Logger.log(`computeO[строка ${rowIdx + 2}]: i_val=${i_val}, j_val=${j_val}, result=${result}`);
    return result;
  }

  const order = ['A', 'B', 'D', 'H', 'K', 'L', 'M', 'O'];

  // Обработка всех строк по пакетам
  const scriptStartTime = Date.now();
  for (let start = 0; start < data.length; start += BATCH_SIZE) {
    const batchStartTime = Date.now();
    const batchData = data.slice(start, Math.min(start + BATCH_SIZE, data.length));
    const columnsToUpdate = {};

    // Собираем данные для каждого столбца
    order.forEach(col => {
      if (!colMap[col]) {
        Logger.log(`Ошибка: col=${col} отсутствует в colMap`);
        return;
      }
      columnsToUpdate[col] = batchData.map((row, idx) => {
        let value = null;
        try {
          switch (col) {
            case 'A': value = computeA(row, start + idx); break;
            case 'B': value = computeB(row, start + idx); break;
            case 'D': value = computeD(row, start + idx); break;
            case 'H': value = computeH(row, start + idx); break;
            case 'K': value = computeK(row, start + idx); break;
            case 'L': value = computeL(row, start + idx); break;
            case 'M': value = computeM(row, start + idx); break;
            case 'O': value = computeO(row, start + idx); break;
          }
        } catch (error) {
          Logger.log(`Ошибка в compute${col}[строка ${start + idx + 2}]: ${error.message}`);
          return [''];
        }
        return [value !== null ? value : ''];
      });
    });

    // Записываем данные для каждого столбца
    order.forEach(col => {
      if (!colMap[col]) {
        Logger.log(`Пропуск записи: col=${col} отсутствует в colMap`);
        return;
      }
      let valuesToWrite = columnsToUpdate[col];
      let startRow = start + 2;
      let numRows = batchData.length;

      // Для столбцов A и B записываем только строки, где C заполнено
      if (col === 'A' || col === 'B') {
        const filteredValues = [];
        const filteredRows = [];
        batchData.forEach((row, idx) => {
          const c_val = row[colMap.C - 1]?.toString().trim() || '';
          if (c_val) {
            filteredValues.push(columnsToUpdate[col][idx]);
            filteredRows.push(start + idx + 2);
          }
        });

        // Записываем только для строк с заполненным C
        filteredValues.forEach((value, idx) => {
          try {
            if (value[0] !== '') {
              Logger.log(`Запись данных для столбца ${col}, строка ${filteredRows[idx]}`);
              sheet.getRange(filteredRows[idx], colMap[col], 1, 1).setValue(value[0]);
            }
          } catch (error) {
            Logger.log(`Ошибка записи для столбца ${col}, строка ${filteredRows[idx]}: ${error.message}`);
          }
        });
      } else {
        // Для остальных столбцов записываем все строки
        try {
          Logger.log(`Запись данных для столбца ${col}, строки ${startRow} - ${startRow + numRows - 1}`);
          sheet.getRange(startRow, colMap[col], numRows, 1).setValues(valuesToWrite);
        } catch (error) {
          Logger.log(`Ошибка записи для столбца ${col}, строки ${startRow}: ${error.message}`);
        }
      }
    });

    const batchTime = (Date.now() - batchStartTime) / 1000;
    Logger.log(`Обработан пакет с ${start} по ${start + batchData.length}, время: ${batchTime} сек`);

    // Проверка времени выполнения
    if (Date.now() - scriptStartTime > 300000) { // 5 минут
      Logger.log('Достигнут лимит времени, скрипт остановлен');
      return;
    }
  }

  // Сохранение словарей cToId и gToId
  try {
    properties.setProperty('cToId', JSON.stringify(cToId));
    properties.setProperty('gToId', JSON.stringify(gToId));
    Logger.log(`Сохранены словари: cToId=${Object.keys(cToId).length} записей, gToId=${Object.keys(gToId).length} записей`);
  } catch (error) {
    Logger.log(`Ошибка сохранения словарей: ${error.message}`);
  }

  Logger.log('Обработка завершена успешно');
}

// Функция для ручного запуска
function resumeProcessing() {
  Logger.log('Ручной запуск обработки всего листа');
  coaches({ source: SpreadsheetApp.getActiveSpreadsheet(), range: SpreadsheetApp.getActiveSheet().getRange('A1') });
}

// Функция для сброса ID (опционально)
function resetIds() {
  PropertiesService.getScriptProperties().deleteProperty('cToId');
  PropertiesService.getScriptProperties().deleteProperty('gToId');
  Logger.log('Словари cToId и gToId сброшены');
}