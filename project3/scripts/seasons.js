function seasons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Сезоны и трофеи");
  const matchesSheet = ss.getSheetByName("Матчи");
  const ratingSheet = ss.getSheetByName("Рейтинг");
  const trainersSheet = ss.getSheetByName("Тренеры и команды");
  const lastRow = mainSheet.getLastRow();

  // Проверка наличия листов
  if (!mainSheet || !matchesSheet || !ratingSheet || !trainersSheet) {
    Logger.log(`Error: Sheets not found - mainSheet: ${!!mainSheet}, matchesSheet: ${!!matchesSheet}, ratingSheet: ${!!ratingSheet}, trainersSheet: ${!!trainersSheet}`);
    return;
  }

  // Получаем все данные
  const mainData = mainSheet.getRange(2, 1, lastRow - 1, 46).getValues(); // A:AT
  const matchesData = matchesSheet.getDataRange().getValues();
  const ratingData = ratingSheet.getDataRange().getValues();
  const trainersData = trainersSheet.getRange('A2:M' + trainersSheet.getLastRow()).getValues();
  Logger.log(`Loaded: mainData=${mainData.length} rows, matchesData=${matchesData.length} rows, ratingData=${ratingData.length} rows, trainersData=${trainersData.length} rows`);

  // Создаем кэш для условий из Рейтинг
  const ratingConditionCache = {};
  for (let i = 1; i < ratingData.length; i++) {
    const fValue = ratingData[i][2]; // C в Рейтинг (F в Сезоны)
    if (ratingData[i][4] !== "") { // E не пустой
      ratingConditionCache[fValue] = true;
    }
  }
  Logger.log(`ratingConditionCache: ${Object.keys(ratingConditionCache).length} entries, example: ${JSON.stringify(Object.entries(ratingConditionCache).slice(0, 5))}`);

  // Обрабатываем столбец A (автоинкремент)
  const columnA = [];
  for (let i = 0; i < mainData.length; i++) {
    if (i === 0) {
      columnA.push([1]);
    } else {
      const current = mainData[i], prev = mainData[i-1];
      const value = (current[5] === prev[5] && current[6] === prev[6] && current[8] === prev[8]) 
        ? columnA[i-1][0] 
        : columnA[i-1][0] + 1;
      columnA.push([value]);
      Logger.log(`A[row ${i+2}]: fValue=${current[5]}, gValue=${current[6]}, iValue=${current[8]}, prevF=${prev[5]}, prevG=${prev[6]}, prevI=${prev[8]}, result=${value}`);
    }
  }

  // Обрабатываем столбец I
  const columnI = [];
  for (let i = 0; i < mainData.length; i++) {
    const result = computeSeasonI(mainData[i], i);
    columnI.push([result]);
  }

  // Обрабатываем столбец L
  const columnL = [];
  for (let i = 0; i < mainData.length; i++) {
    const fValue = mainData[i][5]; // F
    if (ratingConditionCache[fValue]) {
      const [cValue, bValue, dValue] = [mainData[i][2], mainData[i][1], mainData[i][3]]; // C, B, D
      let vValues = [], wValues = [], xValues = [], yValues = [];
      
      for (let j = 1; j < matchesData.length; j++) {
        const match = matchesData[j];
        if (match[6] == cValue && match[0] == bValue && match[4] == dValue) {
          if (typeof match[21] === 'number' && match[21] !== null) vValues.push(match[21]); // V
          if (typeof match[23] === 'number' && match[23] !== null) xValues.push(match[23]); // X
        }
        if (match[7] == cValue && match[0] == bValue && match[5] == dValue) {
          if (typeof match[22] === 'number' && match[22] !== null) wValues.push(match[22]); // W
          if (typeof match[24] === 'number' && match[24] !== null) yValues.push(match[24]); // Y
        }
      }
      
      const avgV = vValues.length > 0 ? vValues.reduce((sum, val) => sum + val, 0) / vValues.length : 0;
      const avgW = wValues.length > 0 ? wValues.reduce((sum, val) => sum + val, 0) / wValues.length : 0;
      const avgX = xValues.length > 0 ? xValues.reduce((sum, val) => sum + val, 0) / xValues.length : 0;
      const avgY = yValues.length > 0 ? yValues.reduce((sum, val) => sum + val, 0) / yValues.length : 0;
      
      const numerator = (avgV + avgW) / 2;
      const denominator = (avgX + avgY) / 2;
      const result = denominator !== 0 ? (numerator / denominator) * 100 : "";
      Logger.log(`L[row ${i+2}]: fValue=${fValue}, cValue=${cValue}, bValue=${bValue}, dValue=${dValue}, vCount=${vValues.length}, wCount=${wValues.length}, xCount=${xValues.length}, yCount=${yValues.length}, avgV=${avgV}, avgW=${avgW}, avgX=${avgX}, avgY=${avgY}, numerator=${numerator}, denominator=${denominator}, result=${result}`);
      columnL.push([result]);
    } else {
      columnL.push([""]);
      Logger.log(`L[row ${i+2}]: fValue=${fValue} not in ratingConditionCache`);
    }
  }

  // Обрабатываем столбец M
  const columnM = [];
  for (let i = 0; i < mainData.length; i++) {
    const fValue = mainData[i][5]; // F
    if (ratingConditionCache[fValue]) {
      const [cValue, dValue, iValue] = [mainData[i][2], mainData[i][3], mainData[i][8]]; // C, D, I
      const isNextSame = (i < mainData.length - 1) && 
                         (mainData[i+1][8] === iValue && mainData[i+1][5] === fValue);
      
      if (isNextSame) {
        columnM.push([""]);
        Logger.log(`M[row ${i+2}]: isNextSame=true, fValue=${fValue}, iValue=${iValue}, result=""`);
      } else {
        let vValues = [], wValues = [], xValues = [], yValues = [];
        let debugMatches = []; // Для отладки
        for (let j = 1; j < matchesData.length; j++) {
          const match = matchesData[j];
          if (match[6] == cValue && match[4] == dValue && match[2] == iValue) {
            if (typeof match[21] === 'number' && match[21] !== null) vValues.push(match[21]); // V
            if (typeof match[23] === 'number' && match[23] !== null) xValues.push(match[23]); // X
            debugMatches.push(`V/X: row=${j+1}, G=${match[6]}, E=${match[4]}, C=${match[2]}, V=${match[21]}, X=${match[23]}`);
          }
          if (match[7] == cValue && match[5] == dValue && match[3] == iValue) {
            if (typeof match[22] === 'number' && match[22] !== null) wValues.push(match[22]); // W
            if (typeof match[24] === 'number' && match[24] !== null) yValues.push(match[24]); // Y
            debugMatches.push(`W/Y: row=${j+1}, H=${match[7]}, F=${match[5]}, D=${match[3]}, W=${match[22]}, Y=${match[24]}`);
          }
        }
        
        const avgV = vValues.length > 0 ? vValues.reduce((sum, val) => sum + val, 0) / vValues.length : 0;
        const avgW = wValues.length > 0 ? wValues.reduce((sum, val) => sum + val, 0) / wValues.length : 0;
        const avgX = xValues.length > 0 ? xValues.reduce((sum, val) => sum + val, 0) / xValues.length : 0;
        const avgY = yValues.length > 0 ? yValues.reduce((sum, val) => sum + val, 0) / yValues.length : 0;
        
        const numerator = (avgV + avgW) / 2;
        const denominator = (avgX + avgY) / 2;
        let result;
        try {
          result = (numerator / denominator) * 100;
        } catch (error) {
          result = `#DIV/0!`; // Имитация ошибки деления на 0
        }
        Logger.log(`M[row ${i+2}]: fValue=${fValue}, cValue=${cValue}, dValue=${dValue}, iValue=${iValue}, vCount=${vValues.length}, wCount=${wValues.length}, xCount=${xValues.length}, yCount=${yValues.length}, avgV=${avgV}, avgW=${avgW}, avgX=${avgX}, avgY=${avgY}, numerator=${numerator}, denominator=${denominator}, result=${result}, debugMatches=${debugMatches.join("; ")}`);
        columnM.push([result]);
      }
    } else {
      columnM.push([""]);
      Logger.log(`M[row ${i+2}]: fValue=${fValue} not in ratingConditionCache`);
    }
  }

  // Обрабатываем столбец N
  const columnN = [];
  const groups = {};
  
  // Группируем по столбцу F
  for (let i = 0; i < mainData.length; i++) {
    const fValue = mainData[i][5];
    if (!groups[fValue]) groups[fValue] = [];
    groups[fValue].push(i);
  }
  Logger.log(`groups: ${Object.keys(groups).length} groups, example: ${JSON.stringify(Object.entries(groups).slice(0, 5).map(([k, v]) => [k, v.length]))}`);

  // Считаем произведение для групп
  for (let i = 0; i < mainData.length; i++) {
    const fValue = mainData[i][5];
    if (ratingConditionCache[fValue]) {
      const groupIndexes = groups[fValue];
      const isFirstInGroup = groupIndexes[0] === i;
      
      if (isFirstInGroup) {
        let product = 1;
        let hasNonOneValues = false;
        
        for (let col = 30; col <= 45; col++) {
          for (const idx of groupIndexes) {
            const value = mainData[idx][col];
            if (value && value != 1) hasNonOneValues = true;
            product *= (value === "" || value === null) ? 1 : value;
          }
        }
        
        const result = (product === 1 && !hasNonOneValues) ? 1 : (product === 1 ? "" : product);
        for (const idx of groupIndexes) {
          columnN[idx] = [idx === i ? result : ""];
        }
        Logger.log(`N[row ${i+2}]: fValue=${fValue}, groupSize=${groupIndexes.length}, product=${product}, hasNonOneValues=${hasNonOneValues}, result=${result}`);
      } else if (!columnN[i]) {
        columnN[i] = [""];
      }
    } else {
      columnN[i] = [""];
      Logger.log(`N[row ${i+2}]: fValue=${fValue} not in ratingConditionCache`);
    }
  }

  // Функция для столбца I
  function computeSeasonI(rowData, rowIdx) {
    const d_val = rowData[3]; // D
    const c_val = rowData[2]; // C
    const b_val = rowData[1]; // B
    if (!d_val || !c_val || !b_val || typeof b_val !== 'number' || isNaN(b_val)) {
      Logger.log(`computeSeasonI[row ${rowIdx + 2}]: d_val=${d_val}, c_val=${c_val}, b_val=${b_val} - invalid data, returning ''`);
      return '';
    }
    for (let i = 0; i < trainersData.length; i++) {
      const a_val = trainersData[i][0]; // A
      const b_val_trainers = trainersData[i][1]; // B
      const l_val = trainersData[i][11]; // L
      const m_val = trainersData[i][12]; // M
      const h_val = trainersData[i][7]; // H
      if (a_val === d_val &&
          b_val_trainers === c_val &&
          typeof l_val === 'number' && !isNaN(l_val) &&
          typeof m_val === 'number' && !isNaN(m_val) &&
          l_val <= b_val && m_val >= b_val) {
        Logger.log(`computeSeasonI[row ${rowIdx + 2}]: d_val=${d_val}, c_val=${c_val}, b_val=${b_val}, found H=${h_val}`);
        return h_val || '';
      }
    }
    Logger.log(`computeSeasonI[row ${rowIdx + 2}]: d_val=${d_val}, c_val=${c_val}, b_val=${b_val} - no match found`);
    return '';
  }

  // Записываем результаты
  try {
    mainSheet.getRange(2, 1, columnA.length, 1).setValues(columnA); // A
    mainSheet.getRange(2, 9, columnI.length, 1).setValues(columnI); // I
    mainSheet.getRange(2, 12, columnL.length, 1).setValues(columnL); // L
    mainSheet.getRange(2, 13, columnM.length, 1).setValues(columnM); // M
    mainSheet.getRange(2, 14, columnN.length, 1).setValues(columnN); // N
    Logger.log("Data writing completed: columns A, I, L, M, N");
  } catch (error) {
    Logger.log(`Error writing data: ${error.message}`);
  }
}

function resumeProcessingSeasons() {
  Logger.log('Manual processing of all rows starting from row 2 for "Сезоны и трофеи"');
  seasons();
}