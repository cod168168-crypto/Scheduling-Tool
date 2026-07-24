/**
 * 自動觸發函數：監控 A42, A49, A32, A33 核取方塊
 */
function onEdit(e) {
  // 【防錯機制】若在編輯器內手動按「▷ 執行」，e 會是 undefined，直接結束不報錯
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const a1 = range.getA1Notation();

  // 1. 處理 B18:B25 員工團隊 (下半部主排班)
  if (a1 === "A42" && sheet.getRange("A42").getValue() === true) {
    distributeLowerTeam(sheet);
    sheet.getRange("A42").setValue(false); // 執行完後取消勾選
  }

  // 2. 處理 B9:B14 員工團隊 (上半部主排班)
  if (a1 === "A49" && sheet.getRange("A49").getValue() === true) {
    distributeUpperTeam(sheet);
    sheet.getRange("A49").setValue(false); // 執行完後取消勾選
  }

  // 3-1. 勾選 A32：將 B9:B14 員工平均分配填入 D32:M32
  if (a1 === "A32" && sheet.getRange("A32").getValue() === true) {
    distributeRowEqually(sheet, "B9:B14", "D9:M14", "D32:M32");
    sheet.getRange("A32").setValue(false); // 執行完後取消勾選
  }

  // 3-2. 勾選 A33：將 B18:B25 員工平均分配填入 D33:M33
  if (a1 === "A33" && sheet.getRange("A33").getValue() === true) {
    distributeRowEqually(sheet, "B18:B25", "D18:M25", "D33:M33");
    sheet.getRange("A33").setValue(false); // 執行完後取消勾選
  }
}

/**
 * 處理 B18:B25 團隊主排班
 */
function distributeLowerTeam(sheet) {
  const empRangeStr = "B18:B25";
  const dataRange = sheet.getRange("D18:AH25");
  const platformRange = sheet.getRange("A36:A41");
  
  // 優先順序: A36, A37 > A40 > A41 > A38, A39 (對應索引 0, 1, 4, 5, 2, 3)
  const priority = [0, 1, 4, 5, 2, 3];
  const skipIfFiveIdx = 5; // 當只有 5 人上班時排除 A41 (index 5)
  
  autoDistribute(sheet, empRangeStr, dataRange, platformRange, priority, skipIfFiveIdx, null);
}

/**
 * 處理 B9:B14 團隊主排班
 */
function distributeUpperTeam(sheet) {
  const empRangeStr = "B9:B14";
  const dataRange = sheet.getRange("D9:AH14");
  const platformRange = sheet.getRange("A44:A48"); // 共 5 個平台
  
  // 固定優先權分配 (相對索引)：
  // 0(B9)  -> 分配 A44(idx 0)
  // 1(B10) -> 分配 A47(idx 3)
  // 2(B11) -> 分配 A45(idx 1)
  const fixedAssignments = { 
    0: 0, 
    1: 3, 
    2: 1 
  }; 
  
  // 剩餘平台給 B12:B14 輪替
  const priority = [0, 3, 1, 4, 2]; 
  const skipIfFourIdx = 2; // 當只有 4 人上班，排除 A46 (index 2)

  autoDistribute(sheet, empRangeStr, dataRange, platformRange, priority, skipIfFourIdx, fixedAssignments);
}

/**
 * 通用核心：主排班垂直/每日平台分配邏輯
 */
function autoDistribute(sheet, empRangeStr, dataRange, platformRange, priority, skipIdx, fixedMap) {
  const dataValues = dataRange.getValues();
  const platforms = platformRange.getValues().map(r => r[0]);
  const leaveTypes = ["休", "特", "婚", "年", "排"];
  
  const numEmp = dataValues.length;
  const numDays = dataValues[0].length;
  
  // 統計矩陣：stats[員工][平台索引] = 本月領取次數
  let stats = Array.from({ length: numEmp }, () => Array(platforms.length).fill(0));

  for (let col = 0; col < numDays; col++) {
    let workingEmps = []; // 存儲今天上班的員工索引
    for (let row = 0; row < numEmp; row++) {
      let cellVal = dataValues[row][col];
      if (!leaveTypes.includes(cellVal)) {
        dataValues[row][col] = ""; // 清空非休假格子準備填入
        workingEmps.push(row);
      }
    }

    if (workingEmps.length === 0) continue;

    // 處理當日可用平台 (根據人數排除特定平台)
    let currentPriority = [...priority];
    if (empRangeStr === "B18:B25" && workingEmps.length === 5) {
      currentPriority = currentPriority.filter(p => p !== skipIdx);
    } else if (empRangeStr === "B9:B14" && workingEmps.length === 4) {
      currentPriority = currentPriority.filter(p => p !== skipIdx);
    }

    // 1. 優先處理「固定分配」(如 B9, B10, B11)
    if (fixedMap) {
      for (let empIdxStr in fixedMap) {
        let empIdx = parseInt(empIdxStr);
        let targetPIdx = fixedMap[empIdx];
        
        if (workingEmps.includes(empIdx) && currentPriority.includes(targetPIdx)) {
          dataValues[empIdx][col] = platforms[targetPIdx];
          stats[empIdx][targetPIdx]++;
          
          workingEmps = workingEmps.filter(e => e !== empIdx);
          currentPriority = currentPriority.filter(p => p !== targetPIdx);
        }
      }
    }

    // 2. 剩下的平台根據「平均輪替」分配給剩下的員工
    currentPriority.forEach(pIdx => {
      if (workingEmps.length > 0) {
        workingEmps.sort((a, b) => stats[a][pIdx] - stats[b][pIdx]);
        let chosenRow = workingEmps[0];

        dataValues[chosenRow][col] = platforms[pIdx];
        stats[chosenRow][pIdx]++;
        workingEmps = workingEmps.filter(e => e !== chosenRow);
      }
    });
  }

  // 回填資料並美化格式
  dataRange.setValues(dataValues);
  applyStyles(dataRange, platforms, platformRange);
}

/**
 * 功能 (3-1, 3-2)：將指定員工團隊平均輪流填入單一橫向列 (D32:M32 或 D33:M33)
 * 特色：保持純文字、黑色字體、不粗體、無特定框線
 */
function distributeRowEqually(sheet, empRangeStr, attendanceRangeStr, targetRowRangeStr) {
  // 1. 取得員工姓名列表
  const empNames = sheet.getRange(empRangeStr).getValues().map(r => r[0]);
  
  // 2. 取得對應日期範圍的員工出勤表 (判斷當天是否請假)
  const attendanceValues = sheet.getRange(attendanceRangeStr).getValues();
  const leaveTypes = ["休", "特", "婚", "年", "排"];
  
  const numDays = attendanceValues[0].length; // 10 天 (D 到 M 欄)
  const resultRow = [];
  
  // 紀錄每位員工在此列被分配到的次數，用於「平均分配」
  let assignmentCounts = Array(empNames.length).fill(0);

  for (let col = 0; col < numDays; col++) {
    // 找出今天上班（非休假）的員工索引
    let availableEmpIndices = [];
    for (let r = 0; r < empNames.length; r++) {
      let status = attendanceValues[r][col];
      if (!leaveTypes.includes(status)) {
        availableEmpIndices.push(r);
      }
    }

    if (availableEmpIndices.length === 0) {
      resultRow.push(""); // 當天無人上班則留空
      continue;
    }

    // 按「被分配次數最少」排序，若次數相同則隨機打亂保持公平
    availableEmpIndices.sort((a, b) => {
      if (assignmentCounts[a] === assignmentCounts[b]) {
        return Math.random() - 0.5;
      }
      return assignmentCounts[a] - assignmentCounts[b];
    });

    // 選取被分配最少次數的人
    let chosenIdx = availableEmpIndices[0];
    resultRow.push(empNames[chosenIdx]);
    assignmentCounts[chosenIdx]++;
  }

  // 將結果填入目標單列 (D32:M32 或 D33:M33)
  const targetRange = sheet.getRange(targetRowRangeStr);
  targetRange.setValues([resultRow]);
  
  // 【更新】文字格式設定：黑色字體、正常字重(不粗體)、水平垂直對齊居中、並清除特殊邊框
  targetRange.setFontColor("#000000")
             .setFontWeight("normal")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setBorder(false, false, false, false, false, false); // 清除特殊框線，維持基本表格樣式
}

/**
 * 格式套用：主要排班區域的背景、字體、對齊、白色雙實線邊框
 */
function applyStyles(dataRange, platforms, platformRange) {
  const values = dataRange.getValues();
  
  const styles = platforms.map((p, i) => {
    let cell = platformRange.getCell(i + 1, 1);
    return {
      name: p,
      bg: cell.getBackground(),
      fc: cell.getFontColor(),
      fw: cell.getFontWeight(),
      ff: cell.getFontFamily(),
      ha: cell.getHorizontalAlignment(),
      va: cell.getVerticalAlignment()
    };
  });

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      let content = values[r][c];
      let style = styles.find(s => s.name === content);
      
      if (style) {
        let targetCell = dataRange.getCell(r + 1, c + 1);
        targetCell.setBackground(style.bg)
                  .setFontColor(style.fc)
                  .setFontWeight(style.fw)
                  .setFontFamily(style.ff)
                  .setHorizontalAlignment(style.ha)
                  .setVerticalAlignment(style.va)
                  .setBorder(true, true, true, true, null, null, "#ffffff", SpreadsheetApp.BorderStyle.DOUBLE);
      }
    }
  }
}
