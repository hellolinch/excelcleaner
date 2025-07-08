const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const remoteMain = require('@electron/remote/main');
const fs = require('fs');
const xlsx = require('xlsx');
remoteMain.initialize();

function createWindow() {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
    },
  });
  remoteMain.enable(win.webContents);
  win.loadFile(path.join(__dirname, 'index-native.html'));
  win.removeMenu();
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// 发送日志到前端
function sendLog(event, message) {
  if (event && event.sender) {
    event.sender.send('log-update', message);
  }
}

// 处理Excel数据清洗与合并
ipcMain.handle('process-excel', async (event, inputDir, outputDir) => {
  try {
    sendLog(event, `开始扫描目录: ${inputDir}`);
    const files = fs.readdirSync(inputDir).filter(f => f.endsWith('.xlsx'));
    sendLog(event, `找到 ${files.length} 个Excel文件: ${files.join(', ')}`);
    
    let allData = [];
    
    for (const file of files) {
      sendLog(event, `正在处理文件: ${file}`);
      const filePath = path.join(inputDir, file);
      
      try {
        const workbook = xlsx.readFile(filePath);
        sendLog(event, `  - 文件包含 ${workbook.SheetNames.length} 个工作表`);
        
        workbook.SheetNames.forEach(sheetName => {
          sendLog(event, `  - 正在处理工作表: ${sheetName}`);
          const sheet = workbook.Sheets[sheetName];
          
          // 不依赖sheet['!ref']，直接读取更大的范围
          // 假设最多有10000行，10列（A-J）
          const maxRows = 10000;
          const maxCols = 10;
          const rows = [];
          
          for (let rowNum = 0; rowNum < maxRows; rowNum++) {
            const row = {};
            let hasData = false;
            
            for (let colNum = 0; colNum < maxCols; colNum++) {
              const cellAddress = xlsx.utils.encode_cell({r: rowNum, c: colNum});
              const cell = sheet[cellAddress];
              const colName = xlsx.utils.encode_col(colNum);
              const cellValue = cell ? (cell.v || '') : '';
              row[colName] = cellValue;
              
              if (cellValue && String(cellValue).trim() !== '') {
                hasData = true;
              }
            }
            
            // 如果连续10行都没有数据，认为已经到达文件末尾
            if (!hasData) {
              let emptyCount = 0;
              for (let checkRow = rowNum; checkRow < rowNum + 10 && checkRow < maxRows; checkRow++) {
                let checkHasData = false;
                for (let colNum = 0; colNum < maxCols; colNum++) {
                  const cellAddress = xlsx.utils.encode_cell({r: checkRow, c: colNum});
                  const cell = sheet[cellAddress];
                  if (cell && cell.v && String(cell.v).trim() !== '') {
                    checkHasData = true;
                    break;
                  }
                }
                if (!checkHasData) {
                  emptyCount++;
                }
              }
              
              if (emptyCount >= 10) {
                break; // 连续10行空行，结束读取
              }
            }
            
            rows.push(row);
          }
          
          sendLog(event, `  - 实际读取行数: ${rows.length}`);
          
          // 查找所有包含"点号"的行（可能有多个表）
          const tableHeaders = [];
          for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const rowValues = Object.values(row);
            
            for (let j = 0; j < rowValues.length; j++) {
              const cellValue = String(rowValues[j]).trim();
              if (cellValue.includes('点号') || cellValue.includes('点 号')) {
                tableHeaders.push({
                  pointRowIndex: i,
                  pointColIndex: j
                });
                sendLog(event, `  - 找到第${tableHeaders.length}个表的点号列: 第${i + 1}行第${j + 1}列`);
                break;
              }
            }
          }
          
          if (tableHeaders.length === 0) {
            sendLog(event, `  - 警告: 未找到包含"点号"的表头行`);
            return;
          }
          
          sendLog(event, `  - 共找到 ${tableHeaders.length} 个界址点成果表`);
          
          // 处理每个表
          let sheetData = [];
          for (let tableIndex = 0; tableIndex < tableHeaders.length; tableIndex++) {
            const { pointRowIndex, pointColIndex } = tableHeaders[tableIndex];
            sendLog(event, `  - 正在处理第${tableIndex + 1}个表（从第${pointRowIndex + 1}行开始）`);
            
            // 查找x(m)和y(m)列（在点号行的下一行或同一行）
            let xColIndex = -1;
            let yColIndex = -1;
            
            for (let i = pointRowIndex; i <= pointRowIndex + 2 && i < rows.length; i++) {
              const row = rows[i];
              const rowValues = Object.values(row);
              
              for (let j = 0; j < rowValues.length; j++) {
                const cellValue = String(rowValues[j]).trim();
                if (cellValue.includes('x(m)') || cellValue.includes('X(m)')) {
                  xColIndex = j;
                }
                if (cellValue.includes('y(m)') || cellValue.includes('Y(m)')) {
                  yColIndex = j;
                }
              }
            }
            
            if (xColIndex === -1 || yColIndex === -1) {
              sendLog(event, `  - 警告: 第${tableIndex + 1}个表未找到x(m)或y(m)列`);
              continue;
            }
            
            // 确定数据结束行（下一个表的开始行或文件结束）
            const dataStartRow = pointRowIndex + 2;
            let dataEndRow = rows.length - 1;
            if (tableIndex + 1 < tableHeaders.length) {
              dataEndRow = tableHeaders[tableIndex + 1].pointRowIndex - 1;
            }
            
            sendLog(event, `  - 第${tableIndex + 1}个表数据范围: 第${dataStartRow + 1}行到第${dataEndRow + 1}行`);
            
            // 提取该表的数据行
            const tableDataRows = [];
            for (let i = dataStartRow; i <= dataEndRow; i++) {
              const row = rows[i];
              const rowValues = Object.values(row);
              
              const pointValue = String(rowValues[pointColIndex] || '').trim();
              const xValue = String(rowValues[xColIndex] || '').trim();
              const yValue = String(rowValues[yColIndex] || '').trim();
              
              // 跳过空行或无效数据
              if (!pointValue || !xValue || !yValue || 
                  pointValue.includes('点号') || pointValue.includes('点 号') ||
                  xValue.includes('x(m)') || yValue.includes('y(m)') ||
                  pointValue.includes('界址点成果表') || pointValue.includes('宗地号')) {
                continue;
              }
              
              // 检查是否为数字
              const xNum = parseFloat(xValue);
              const yNum = parseFloat(yValue);
              if (isNaN(xNum) || isNaN(yNum)) {
                continue;
              }
              
              tableDataRows.push({
                点号: pointValue,
                'x(m)': xNum,
                'y(m)': yNum
              });
            }
            
            sendLog(event, `  - 第${tableIndex + 1}个表提取到 ${tableDataRows.length} 行有效数据`);
            sheetData = sheetData.concat(tableDataRows);
          }
          
          sendLog(event, `  - 工作表总共提取到 ${sheetData.length} 行数据`);
          
          // 智能去重（保留闭合多边形的首尾点，去重中间重复点）
          if (sheetData.length > 0) {
            sendLog(event, `  - 开始智能去重分析...`);
            
            // 按点号分组统计
            const pointStats = {};
            sheetData.forEach((row, index) => {
              const pointNum = row['点号'];
              if (!pointStats[pointNum]) {
                pointStats[pointNum] = [];
              }
              pointStats[pointNum].push({
                index: index,
                data: row
              });
            });
            
            // 找出重复的点号
            const duplicatePoints = {};
            Object.keys(pointStats).forEach(pointNum => {
              if (pointStats[pointNum].length > 1) {
                duplicatePoints[pointNum] = pointStats[pointNum];
                sendLog(event, `    - 发现重复点号 ${pointNum}：${pointStats[pointNum].length} 次`);
              }
            });
            
            // 构建去重后的数据
            const deduplicatedData = [];
            const removedIndices = new Set();
            
            // 处理重复点号的去重逻辑
            Object.keys(duplicatePoints).forEach(pointNum => {
              const occurrences = duplicatePoints[pointNum];
              
              if (pointNum === 'J1' && occurrences.length >= 2) {
                // J1点特殊处理：保留第一个和最后一个，去重中间的
                sendLog(event, `    - J1点闭合处理：保留首尾，去重中间 ${occurrences.length - 2} 个重复点`);
                
                // 标记中间的J1点为需要移除
                for (let i = 1; i < occurrences.length - 1; i++) {
                  removedIndices.add(occurrences[i].index);
                }
              } else {
                // 其他重复点号：只保留第一个
                sendLog(event, `    - ${pointNum}点去重：保留第1个，移除后面 ${occurrences.length - 1} 个重复点`);
                
                // 标记除第一个外的所有重复点为需要移除
                for (let i = 1; i < occurrences.length; i++) {
                  removedIndices.add(occurrences[i].index);
                }
              }
            });
            
            // 构建最终数据（排除被标记移除的点）
            sheetData.forEach((row, index) => {
              if (!removedIndices.has(index)) {
                deduplicatedData.push(row);
              }
            });
            
            sendLog(event, `  - 智能去重完成：${sheetData.length} → ${deduplicatedData.length} 行（移除 ${removedIndices.size} 个重复点）`);
            
            // 检查是否存在J1闭合
            const hasJ1 = deduplicatedData.some(row => row['点号'] === 'J1');
            const j1Count = deduplicatedData.filter(row => row['点号'] === 'J1').length;
            if (hasJ1) {
              sendLog(event, `  - 检测到 ${j1Count} 个J1点${j1Count >= 2 ? '（闭合多边形）' : ''}`);
            }
            
            allData = allData.concat(deduplicatedData);
          } else {
            allData = allData.concat(sheetData);
          }
        });
        
      } catch (err) {
        sendLog(event, `  - 处理文件失败: ${err.message}`);
      }
    }
    
    sendLog(event, `总共合并 ${allData.length} 行数据`);
    
    if (allData.length === 0) {
      return { success: false, message: '没有找到有效的数据，请检查Excel文件格式' };
    }
    
    // 导出合并后的Excel
    const ws = xlsx.utils.json_to_sheet(allData);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, '合并结果');
    const outPath = path.join(outputDir, '清洗合并结果.xlsx');
    xlsx.writeFile(wb, outPath);
    
    sendLog(event, `处理完成！结果已保存到: ${outPath}`);
    return { success: true, message: `处理完成，共处理 ${allData.length} 行数据，结果已保存到：${outPath}` };
    
  } catch (err) {
    sendLog(event, `处理异常: ${err.message}`);
    return { success: false, message: err.message };
  }
}); 