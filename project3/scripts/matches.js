function onEdit(e) {
  // Проверка наличия объекта события и источника
  if (!e || !e.source) {
    Logger.log(`Ошибка: объект события или источник не определены - ${JSON.stringify({ e: !!e, source: !!e?.source })}`);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // Обрабатываем изменения в "Матчи" или связанных листах // =IF(I2<>"";1;0)
  if (sheetName === 'Матчи') {
    if (!e.range) {
      Logger.log('Объект range не определён, обработка всех строк в "Матчи" с заполненным I');
      processMatches(ss, null);
      return;
    }

    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    if (startRow < 2) {
      Logger.log(`Изменение в заголовке (строка ${startRow}), обработка всех строк с заполненным I`);
      processMatches(ss, null);
      return;
    }

    // Загружаем только изменённые строки // =IF(I2<>"";1;0)
    const lastCol = 27; // Максимальный столбец (AA)
    const data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
    const rowsToProcess = data.map((row, idx) => ({ row, index: startRow + idx - 2 }))
                             .filter(({ row }) => row[8] !== ''); // I не пустой
    Logger.log(`Обработка изменённых строк в "Матчи": ${rowsToProcess.length} строк (с ${startRow} по ${startRow + numRows - 1})`);
    processMatches(ss, rowsToProcess);
  } else if (['Тренеры и команды', 'Глосс.', 'Силы сборных', 'Силы клубов', 'Даты по странам'].includes(sheetName)) {
    Logger.log(`Изменение в листе "${sheetName}", обработка всех строк в "Матчи" с заполненным I`);
    processMatches(ss, null); // null означает обработку всех строк с заполненным I
  }
}

function processMatches(ss, rowsToProcess) {
  const sheet = ss.getSheetByName('Матчи');
  if (!sheet) {
    Logger.log('Ошибка: лист "Матчи" не найден');
    return;
  }

  const colMap = {
    A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7, H: 8, I: 9, K: 11,
    L: 12, M: 13, N: 14, O: 15, P: 16, Q: 17, R: 18, S: 19, T: 20,
    U: 21, V: 22, W: 23, X: 24, Y: 25, Z: 26, AA: 27
  };

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(...Object.values(colMap));
  if (lastRow < 2) {
    Logger.log('Ошибка: недостаточно данных в листе "Матчи" (меньше 2 строк)');
    return;
  }

  // Загрузка всех данных из "Матчи" или только указанных строк // =IF(ROW()>=2;1;0)
  let data, rowIndexes;
  if (rowsToProcess) {
    data = rowsToProcess.map(({ row }) => row);
    rowIndexes = rowsToProcess.map(({ index }) => index);
  } else {
    const fullData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const filtered = fullData.map((row, idx) => ({ row, index: idx }))
                            .filter(({ row }) => row[8] !== ''); // I не пустой
    data = filtered.map(({ row }) => row);
    rowIndexes = filtered.map(({ index }) => index);
  }
  Logger.log(`Обработка ${data.length} строк с заполненным I`);

  // Загрузка словарей cToId и gToId // =IFERROR(INDEX(Список;MATCH(A2;Список;0));0)
  const properties = PropertiesService.getScriptProperties();
  const cToId = JSON.parse(properties.getProperty('cToId') || '{}');
  const gToId = JSON.parse(properties.getProperty('gToId') || '{}');
  Logger.log(`cToId: ${Object.keys(cToId).length} записей, пример: ${JSON.stringify(Object.entries(cToId).slice(0, 5))}`);
  Logger.log(`gToId: ${Object.keys(gToId).length} записей, пример: ${JSON.stringify(Object.entries(gToId).slice(0, 5))}`);

  // Загрузка данных из листа "Глосс." // =IF(COLUMN()>=3;1;0)
  const glossSheet = ss.getSheetByName('Глосс.');
  const glossData = glossSheet ? glossSheet.getRange('C2:O' + glossSheet.getLastRow()).getValues() : [];
  // Загрузка столбца O для computeAA // =IF(O2<>"";O2;"")
  const glossOValues = glossSheet ? glossSheet.getRange('O2:O' + glossSheet.getLastRow()).getValues().map(row => row[0]).filter(val => val !== '') : [];
  Logger.log(`Загружено ${glossData.length} строк из "Глосс.", O содержит ${glossOValues.length} непустых значений`);

  // Загрузка данных из листа "Даты по странам" // =IF(COLUMN()<=6;1;0)
  const datesSheet = ss.getSheetByName('Даты по странам');
  const datesData = datesSheet ? datesSheet.getRange('A2:F' + datesSheet.getLastRow()).getValues() : [];
  const dateI1 = datesSheet ? datesSheet.getRange('I1').getValue() : null;
  const dateJ1 = datesSheet ? datesSheet.getRange('J1').getValue() : null;
  const dateL1 = datesSheet ? datesSheet.getRange('L1').getValue() : null;
  const dateM1 = datesSheet ? datesSheet.getRange('M1').getValue() : null;
  Logger.log(`Загружено ${datesData.length} строк из "Даты по странам", I1=${dateI1}, J1=${dateJ1}, L1=${dateL1}, M1=${dateM1}`);

  // Загрузка данных из листа "Тренеры и команды" // =IF(COLUMN()<=11;1;0)
  const trainersSheet = ss.getSheetByName('Тренеры и команды');
  const trainersData = trainersSheet ? trainersSheet.getRange('A2:K' + trainersSheet.getLastRow()).getValues() : [];
  const p1_val = trainersSheet ? trainersSheet.getRange('P1').getValue() : '';
  Logger.log(`Загружено ${trainersData.length} строк из "Тренеры и команды"`);

  // Создание индексов для оптимизации // =IF(A2<>"";CONCATENATE(A2;"|";B2);"")
  const trainersIndex = {};
  const trainersCIndex = {};
  trainersData.forEach(row => {
    const a_val = row[0];
    const b_val = row[1];
    const c_val = row[2];
    const h_val = Number(row[7]) || 0;
    const i_val = row[8] ? new Date(row[8]) : new Date(0);
    const j_val = row[9] ? new Date(row[9]) : new Date('9999-12-31');
    const k_val = row[10];
    const key = `${a_val}|${b_val}`;
    if (!trainersIndex[key]) trainersIndex[key] = [];
    trainersIndex[key].push({ h: h_val, i: i_val, j: j_val });
    if (c_val) {
      if (!trainersCIndex[c_val]) trainersCIndex[c_val] = [];
      trainersCIndex[c_val].push({ j: j_val, k: k_val });
    }
  });
  Logger.log(`trainersIndex: ${Object.keys(trainersIndex).length} записей`);
  Logger.log(`trainersCIndex: ${Object.keys(trainersCIndex).length} записей`);

  // Загрузка данных из листа "Сезоны и трофеи" // =IF(COLUMN()<=6;1;0)
  const seasonsSheet = ss.getSheetByName('Сезоны и трофеи');
  const seasonsData = seasonsSheet ? seasonsSheet.getRange('B2:G' + seasonsSheet.getLastRow()).getValues() : [];
  Logger.log(`Загружено ${seasonsData.length} строк из "Сезоны и трофеи"`);

  // Загрузка данных из листа "Силы клубов" // =IF(COLUMN()<=13;1;0)
  const clubsSheet = ss.getSheetByName('Силы клубов');
  const clubsData = clubsSheet ? clubsSheet.getRange('A2:M' + clubsSheet.getLastRow()).getValues() : [];
  const clubHeaders = clubsSheet ? clubsSheet.getRange('D1:M1').getValues()[0] : [];
  const clubMap = {};
  clubsData.forEach(row => {
    if (row[0]) clubMap[row[0]] = row.slice(3);
  });
  Logger.log(`clubMap: ${Object.keys(clubMap).length} записей, пример: ${JSON.stringify(Object.entries(clubMap).slice(0, 5))}`);

  // Загрузка данных из листа "Силы сборных" // =IF(COLUMN()<=11;1;0)
  const nationalSheet = ss.getSheetByName('Силы сборных');
  const nationalData = nationalSheet ? nationalSheet.getRange('B2:L' + nationalSheet.getLastRow()).getValues() : [];
  const nationalHeaders = nationalSheet ? nationalSheet.getRange('C1:L1').getValues()[0] : [];
  const nationalMap = {};
  nationalData.forEach(row => {
    if (row[0]) nationalMap[row[0]] = row.slice(1);
  });
  Logger.log(`nationalMap: ${Object.keys(nationalMap).length} записей, пример: ${JSON.stringify(Object.entries(nationalMap).slice(0, 5))}`);

  // Загрузка параметров // =IF(COLUMN()=2;1;0)
  const paramsSheet = ss.getSheetByName('Параметры расчета');
  const param_b1 = paramsSheet ? Number(paramsSheet.getRange('B1').getValue()) || 0 : 0;
  const param_b2 = paramsSheet ? Number(paramsSheet.getRange('B2').getValue()) || 0 : 0;
  const param_b3 = paramsSheet ? Number(paramsSheet.getRange('B3').getValue()) || 0 : 0;
  const param_b4 = paramsSheet ? Number(paramsSheet.getRange('B4').getValue()) || 0 : 0;
  Logger.log(`Параметры: B1=${param_b1}, B2=${param_b2}, B3=${param_b3}, B4=${param_b4}`);

  // Кэш для оптимизации // =IFERROR(MATCH(A2;Кэш;0);0)
  const computeACache = new Map();
  const computeCCache = new Map();
  const computeAACache = new Map();

function computeA(rowData, i_val) {
  if (!i_val) {
    Logger.log(`computeA[строка ${rowData.index + 2}]: i_val=${i_val} - пустое, возвращается ''`);
    return '';
  }
  const i_date = new Date(i_val);
  const k_val = rowData[colMap.K - 1];
  const aa_val = rowData[colMap.AA - 1] === true || rowData[colMap.AA - 1] === 'TRUE' || rowData[colMap.AA - 1] === 1;
  const cacheKey = `${i_val}|${k_val}|${aa_val}`;
  if (computeACache.has(cacheKey)) {
    return computeACache.get(cacheKey);
  }

  let result = '';
  // Выбор диапазона дат в зависимости от AA
  const startDateRef = aa_val ? dateL1 : dateI1;
  const endDateRef = aa_val ? dateM1 : dateJ1;

  // Проверка, попадает ли дата в диапазон
  if (startDateRef && endDateRef && i_date >= new Date(startDateRef) && i_date <= new Date(endDateRef)) {
    // SUMIFS из Даты по странам!C:C с дополнительными условиями по датам (столбцы E и F)
    result = datesData.reduce((sum, row) => {
      const country = row[0]; // A
      const code = Number(row[2]) || 0; // C
      const startDate = row[4] ? new Date(row[4]) : new Date(0); // E
      const endDate = row[5] ? new Date(row[5]) : new Date('9999-12-31'); // F
      if (country === k_val && startDate <= i_date && endDate >= i_date) {
        sum += code;
      }
      return sum;
    }, 0);
  } else {
    // Выбор SUMIFS в зависимости от AA
    if (!aa_val) {
      // SUMIFS из Глосс.!C2:C
      result = glossData.reduce((sum, row) => {
        const code = Number(row[0]) || 0; // C
        const startDate = row[2] ? new Date(row[2]) : new Date(0); // E
        const endDate = row[3] ? new Date(row[3]) : new Date('9999-12-31'); // F
        if (startDate <= i_date && endDate >= i_date) {
          sum += code;
        }
        return sum;
      }, 0);
    } else {
      // SUMIFS из Глосс.!H2:H
      result = glossData.reduce((sum, row) => {
        const code = Number(row[5]) || 0; // H
        const startDate = row[7] ? new Date(row[7]) : new Date(0); // J
        const endDate = row[8] ? new Date(row[8]) : new Date('9999-12-31'); // K
        if (startDate <= i_date && endDate >= i_date) {
          sum += code;
        }
        return sum;
      }, 0);
    }
  }
  
  computeACache.set(cacheKey, result);
  Logger.log(`computeA[строка ${rowData.index + 2}]: i_val=${i_val}, k_val=${k_val}, aa_val=${aa_val}, result=${result}`);
  return result;
}

  // Новая функция computeZ // =IF(A2<>"";IF(AND(Q2=0;R2=0;I2<>"");TRUE;FALSE);"")
  function computeZ(rowData) {
    const a_val = rowData[colMap.A - 1];
    const i_val = rowData[colMap.I - 1];
    const q_val = Number(rowData[colMap.Q - 1]) || 0;
    const r_val = Number(rowData[colMap.R - 1]) || 0;
    
    if (!i_val) {
      Logger.log(`computeZ[строка ${rowData.index + 2}]: i_val=${i_val} - пустое, возвращается ''`);
      return '';
    }
    const result = (q_val === 0 && r_val === 0 && i_val !== '');
    Logger.log(`computeZ[строка ${rowData.index + 2}]: q_val=${q_val}, r_val=${r_val}, i_val=${i_val}, result=${result}`);
    return result;
  }

  // Новая функция computeAA // =IF(I2<>"";IF(COUNTIF(O:O;K2)>0;TRUE;FALSE);"")
  function computeAA(rowData) {
    const i_val = rowData[colMap.I - 1];
    const k_val = rowData[colMap.K - 1];
    
    if (!i_val) {
      Logger.log(`computeAA[строка ${rowData.index + 2}]: i_val=${i_val} - пустое, возвращается ''`);
      return '';
    }
    const cacheKey = `${k_val}`;
    if (computeAACache.has(cacheKey)) {
      return computeAACache.get(cacheKey);
    }
    const result = glossOValues.includes(k_val);
    computeAACache.set(cacheKey, result);
    Logger.log(`computeAA[строка ${rowData.index + 2}]: k_val=${k_val}, result=${result}`);
    return result;
  }

  // Остальные функции compute (без изменений) // =IF(I2<>"";IF(L2<>"";VLOOKUP(L2;Список;2;FALSE);"");"")
  function computeE(rowData, isF) {
    try {
      const i_val = rowData[colMap.I - 1];
      const teamId = isF ? rowData[colMap.M - 1] : rowData[colMap.L - 1];
      if (!i_val || !teamId) {
        Logger.log(`compute${isF ? 'F' : 'E'}[строка ${rowData.index + 2}]: i_val=${i_val}, teamId=${teamId} - пустое, возвращается ''`);
        return '';
      }
      const result = cToId[teamId] || '';
      Logger.log(`compute${isF ? 'F' : 'E'}[строка ${rowData.index + 2}]: teamId=${teamId}, result=${result}`);
      return result;
    } catch (error) {
      Logger.log(`compute${isF ? 'F' : 'E'}[строка ${rowData.index + 2}]: ошибка в L/M - ${error.message}, возвращается ''`);
      return '';
    }
  }

  function computeG(rowData, isH) {
    try {
      const i_val = rowData[colMap.I - 1];
      const club = isH ? rowData[colMap.P - 1] : rowData[colMap.O - 1];
      if (!i_val || !club) {
        Logger.log(`compute${isH ? 'H' : 'G'}[строка ${rowData.index + 2}]: i_val=${i_val}, club=${club} - пустое, возвращается ''`);
        return '';
      }
      const result = gToId[club] || '';
      Logger.log(`compute${isH ? 'H' : 'G'}[строка ${rowData.index + 2}]: club=${club}, result=${result}`);
      return result;
    } catch (error) {
      Logger.log(`compute${isH ? 'H' : 'G'}[строка ${rowData.index + 2}]: ошибка в O/P (зависит от L/M) - ${error.message}, возвращается ''`);
      return '';
    }
  }

  function computeU(rowData) {
    const t_val = rowData[colMap.T - 1];
    if (t_val === '' || t_val == null) return '';
    const t_num = Number(t_val);
    if (isNaN(t_num)) return '';
    return 4 - t_num;
  }

  function computeB(rowData, i_val) {
    const a_val = Number(rowData[colMap.A - 1]);
    if (!i_val || isNaN(a_val)) return '';
    const i_date = new Date(i_val);
    if (a_val === 14) {
      return i_date.getFullYear() === 2019 ? '20S' : '20W';
    } else {
      const year = i_date.getFullYear() % 100;
      const month = i_date.getMonth() + 1;
      return month < 7 ? year.toString().padStart(2, '0') + 'W' : (year + 1).toString().padStart(2, '0') + 'S';
    }
  }

  function computeO(rowData, isP) {
    try {
      const a_val = Number(rowData[colMap.A - 1]);
      const i_val = rowData[colMap.I - 1];
      const l_val = isP ? rowData[colMap.M - 1] : rowData[colMap.L - 1];
      const q_val = Number(isP ? rowData[colMap.R - 1] : rowData[colMap.Q - 1]);
      if (!a_val || !i_val || !l_val || isNaN(q_val)) {
        Logger.log(`compute${isP ? 'P' : 'O'}[строка ${rowData.index + 2}]: a_val=${a_val}, i_val=${i_val}, l_val=${l_val}, q_val=${q_val} - пустое, возвращается ''`);
        return '';
      }
      if (q_val > 0 && q_val < 6000 && !seasonsData.some(row => row[5] === l_val && Number(row[0]) === a_val)) {
        return '?';
      }
      const i_date = new Date(i_val);
      for (const row of trainersData) {
        const team_C = row[2];
        const team_E = row[4];
        const team_F = row[5];
        const team_G = row[6];
        const team_I = row[8] ? new Date(row[8]) : new Date(0);
        const team_J = row[9] ? new Date(row[9]) : new Date('9999-12-31');
        if (
          team_I <= i_date &&
          team_J >= i_date &&
          (team_C === l_val || (team_E && team_E === l_val) || (team_F && team_F === l_val))
        ) {
          return team_G || '';
        }
      }
      return '?';
    } catch (error) {
      Logger.log(`compute${isP ? 'P' : 'O'}[строка ${rowData.index + 2}]: ошибка в L/M - ${error.message}, возвращается ''`);
      return '';
    }
  }

  function computeQ(rowData, isR) {
    try {
      const l_val = isR ? rowData[colMap.M - 1] : rowData[colMap.L - 1];
      const b_val = rowData[colMap.B - 1];
      const k_val = rowData[colMap.K - 1];
      if (!l_val || !b_val) {
        Logger.log(`compute${isR ? 'R' : 'Q'}[строка ${rowData.index + 2}]: l_val=${l_val}, b_val=${b_val} - пустое, возвращается 0`);
        return 0;
      }
      const headers = k_val === 'Сборные' ? nationalHeaders : clubHeaders;
      const dataMap = k_val === 'Сборные' ? nationalMap : clubMap;
      const col_idx = headers.indexOf(b_val);
      if (col_idx === -1) {
        Logger.log(`compute${isR ? 'R' : 'Q'}[строка ${rowData.index + 2}]: b_val=${b_val} не найдено в ${k_val === 'Сборные' ? 'nationalHeaders' : 'clubHeaders'}, возвращается 0`);
        return 0;
      }
      const dataRow = dataMap[l_val] || [];
      const result = Number(dataRow[col_idx]) || 0;
      Logger.log(`compute${isR ? 'R' : 'Q'}[строка ${rowData.index + 2}]: l_val=${l_val}, b_val=${b_val}, k_val=${k_val}, result=${result}`);
      return result;
    } catch (error) {
      Logger.log(`compute${isR ? 'R' : 'Q'}[строка ${rowData.index + 2}]: ошибка в L/M - ${error.message}, возвращается 0`);
      return 0;
    }
  }

  function computeC(rowData, isD) {
    try {
      const o_val = isD ? rowData[colMap.P - 1] : rowData[colMap.O - 1];
      if (o_val === '') return '?';
      if (o_val === '-') return '-';
      const e_val = isD ? rowData[colMap.F - 1] : rowData[colMap.E - 1];
      const g_val = isD ? rowData[colMap.H - 1] : rowData[colMap.G - 1];
      const i_val = rowData[colMap.I - 1];
      if (!e_val || !g_val || !i_val) {
        Logger.log(`compute${isD ? 'D' : 'C'}[строка ${rowData.index + 2}]: e_val=${e_val}, g_val=${g_val}, i_val=${i_val} - пустое, возвращается ''`);
        return '';
      }
      const i_date = new Date(i_val);
      const cacheKey = `${e_val}|${g_val}|${i_val}`;
      if (computeCCache.has(cacheKey)) {
        return computeCCache.get(cacheKey);
      }
      const key = `${e_val}|${g_val}`;
      const trainers = trainersIndex[key] || [];
      let sum = 0;
      for (const trainer of trainers) {
        if (trainer.i <= i_date && trainer.j >= i_date) {
          sum += trainer.h;
        }
      }
      const result = sum || '';
      computeCCache.set(cacheKey, result);
      return result;
    } catch (error) {
      Logger.log(`compute${isD ? 'D' : 'C'}[строка ${rowData.index + 2}]: ошибка в L/M или зависимых - ${error.message}, возвращается ''`);
      return '';
    }
  }

  function computeV(rowData, s_val, t_val, n_val, q_val, r_val) {
    try {
      if (isNaN(s_val) || isNaN(t_val) || n_val === undefined || isNaN(q_val) || isNaN(r_val) || r_val === 0) {
        Logger.log(`computeV[строка ${rowData.index + 2}]: s_val=${s_val}, t_val=${t_val}, n_val=${n_val}, q_val=${q_val}, r_val=${r_val} - некорректные данные, возвращается ''`);
        return '';
      }
      const t_bonus = t_val === 3 ? 1 : t_val === 1 ? 0 : 0.5;
      const ratio = q_val / r_val || 0;
      let adj;
      if (n_val) {
        if (s_val === 2) {
          adj = ratio > param_b3 ? 0 : ratio < param_b4 ? 1 : 0.5;
        } else if (s_val === 1) {
          adj = ratio > param_b3 ? -0.5 : ratio < param_b4 ? 0.5 : 0;
        } else if (s_val === 0) {
          adj = ratio > param_b3 ? -1 : ratio < param_b4 ? 0 : -0.5;
        } else {
          adj = 0;
        }
      } else {
        if (s_val === 2) {
          adj = (q_val >= r_val || (q_val < r_val && ratio > param_b1)) ? 0 :
                (q_val < r_val && ratio < param_b2) ? 1 : 0.5;
        } else if (s_val === 1) {
          adj = (q_val >= r_val || (q_val < r_val && ratio > param_b1)) ? -0.5 :
                (q_val < r_val && ratio < param_b2) ? 0.5 : 0;
        } else if (s_val === 0) {
          adj = (q_val >= r_val || (q_val < r_val && ratio > param_b1)) ? -1 :
                (q_val < r_val && ratio < param_b2) ? 0 : -0.5;
        } else {
          adj = 0;
        }
      }
      return s_val + t_bonus + adj;
    } catch (error) {
      Logger.log(`computeV[строка ${rowData.index + 2}]: ошибка в Q/R (зависит от L/M) - ${error.message}, возвращается ''`);
      return '';
    }
  }

  function computeX(rowData, n_val, q_val, r_val) {
    try {
      if (n_val === undefined || isNaN(q_val) || isNaN(r_val) || r_val === 0) {
        Logger.log(`computeX[строка ${rowData.index + 2}]: n_val=${n_val}, q_val=${q_val}, r_val=${r_val} - некорректные данные, возвращается ''`);
        return '';
      }
      const ratio = q_val / r_val || 0;
      if (n_val) {
        return ratio > param_b3 ? 3 : ratio < param_b4 ? 0 : 1.5;
      } else {
        return (q_val >= r_val || (q_val < r_val && ratio > param_b1)) ? 3 :
               (q_val < r_val && ratio < param_b2) ? 0 : 1.5;
      }
    } catch (error) {
      Logger.log(`computeX[строка ${rowData.index + 2}]: ошибка в Q/R (зависит от L/M) - ${error.message}, возвращается ''`);
      return '';
    }
  }

  function computeW(rowData, v_val) {
    try {
      const i_val = rowData[colMap.I - 1];
      if (!i_val || isNaN(v_val)) {
        Logger.log(`computeW[строка ${rowData.index + 2}]: i_val=${i_val}, v_val=${v_val} - некорректные данные, возвращается ''`);
        return '';
      }
      return 3 - v_val;
    } catch (error) {
      Logger.log(`computeW[строка ${rowData.index + 2}]: ошибка в V (зависит от L/M) - ${error.message}, возвращается ''`);
      return '';
    }
  }

  function computeY(rowData, x_val) {
    try {
      const i_val = rowData[colMap.I - 1];
      if (!i_val || isNaN(x_val)) {
        Logger.log(`computeY[строка ${rowData.index + 2}]: i_val=${i_val}, x_val=${x_val} - некорректные данные, возвращается ''`);
        return '';
      }
      return 3 - x_val;
    } catch (error) {
      Logger.log(`computeY[строка ${rowData.index + 2}]: ошибка в X (зависит от L/M) - ${error.message}, возвращается ''`);
      return '';
    }
  }

  // Порядок обработки столбцов (Z и AA добавлены после Q, R, но перед C, D) // =IF(COLUMN()=MATCH("Z";$A$1:$AA$1;0);1;0)
  const order = ['E', 'F', 'G', 'H', 'A', 'B', 'O', 'P', 'Q', 'R', 'Z', 'AA', 'C', 'D', 'U', 'V', 'X', 'W', 'Y'];

  // Обработка строк // =IF(AND(ROW()>=2;I2<>"");1;0)
  const scriptStartTime = Date.now();
  const results = data.map((row, idx) => {
    const resultRow = new Array(lastCol).fill(null);
    const i_val = row[colMap.I - 1];
    if (!i_val) {
      Logger.log(`[строка ${rowIndexes[idx] + 2}]: i_val=${i_val} - пустое, пропускается`);
      return resultRow;
    }

    row.index = rowIndexes[idx]; // Для логирования
    order.forEach(col => {
      let value = null;
      const s_val = Number(row[colMap.S - 1]);
      const t_val = Number(row[colMap.T - 1]);
      const n_val = row[colMap.N - 1] === 'TRUE' || row[colMap.N - 1] === true || row[colMap.N - 1] === 1;
      const q_val = Number(row[colMap.Q - 1]);
      const r_val = Number(row[colMap.R - 1]);
      switch (col) {
        case 'A': value = computeA(row, i_val); break;
        case 'B': value = computeB(row, i_val); break;
        case 'C': value = computeC(row, false); break;
        case 'D': value = computeC(row, true); break;
        case 'E': value = computeE(row, false); break;
        case 'F': value = computeE(row, true); break;
        case 'G': value = computeG(row, false); break;
        case 'H': value = computeG(row, true); break;
        case 'O': value = computeO(row, false); break;
        case 'P': value = computeO(row, true); break;
        case 'Q': value = computeQ(row, false); break;
        case 'R': value = computeQ(row, true); break;
        case 'U': value = computeU(row); break;
        case 'V': value = computeV(row, s_val, t_val, n_val, q_val, r_val); break;
        case 'W': value = computeW(row, row[colMap.V - 1]); break;
        case 'X': value = computeX(row, n_val, q_val, r_val); break;
        case 'Y': value = computeY(row, row[colMap.X - 1]); break;
        case 'Z': value = computeZ(row); break;
        case 'AA': value = computeAA(row); break;
      }
      if (value !== null) {
        resultRow[colMap[col] - 1] = value;
        row[colMap[col] - 1] = value;
      }
    });
    return resultRow;
  });

  // Запись изменений одним вызовом // =IF(OR(E2<>"";F2<>"";G2<>"";H2<>"";A2<>"";B2<>"";O2<>"";P2<>"";Q2<>"";R2<>"";Z2<>"";AA2<>"";C2<>"";D2<>"";U2<>"";V2<>"";X2<>"";W2<>"";Y2<>"");1;0)
  if (rowsToProcess) {
    if (results.some(row => row.some(val => val !== null))) {
      const ranges = [];
      let currentStart = rowIndexes[0] + 2;
      let currentLength = 1;
      for (let i = 1; i < rowIndexes.length; i++) {
        if (rowIndexes[i] === rowIndexes[i-1] + 1) {
          currentLength++;
        } else {
          ranges.push({ start: currentStart, length: currentLength, data: data.slice(i - currentLength, i) });
          currentStart = rowIndexes[i] + 2;
          currentLength = 1;
        }
      }
      ranges.push({ start: currentStart, length: currentLength, data: data.slice(data.length - currentLength) });

      ranges.forEach(({ start, length, data }) => {
        sheet.getRange(start, 1, length, lastCol).setValues(data);
      });
      Logger.log(`Обновлены изменённые строки с заполненным I: ${rowIndexes.map(i => i + 2).join(', ')}, время: ${(Date.now() - scriptStartTime) / 1000} сек`);
    } else {
      Logger.log(`Нет изменений для строк ${rowIndexes.map(i => i + 2).join(', ')}`);
    }
  } else {
    if (results.some(row => row.some(val => val !== null))) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).setValues(
        sheet.getRange(2, 1, lastRow - 1, lastCol).getValues().map((row, idx) => {
          const resultIdx = rowIndexes.indexOf(idx);
          return resultIdx !== -1 ? data[resultIdx] : row;
        })
      );
      Logger.log(`Обновлены все строки с заполненным I (2 по ${lastRow}), время: ${(Date.now() - scriptStartTime) / 1000} сек`);
    } else {
      Logger.log(`Нет изменений для строк с заполненным I`);
    }
  }

  if (Date.now() - scriptStartTime > 300000) {
    Logger.log('Достигнут лимит времени (300 сек), скрипт остановлен');
    return;
  }

  Logger.log('Обработка завершена успешно');
}

// Функция для ручного запуска обработки с начала // =IF(AND(ROW()>=2;I2<>"");1;0)
function resumeProcessing() {
  Logger.log('Ручной запуск обработки всех строк с заполненным I в листе "Матчи"');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  processMatches(ss, null);
}