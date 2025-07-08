const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

console.log('开始测试Excel数据清洗...');

// 简化版的数据清洗逻辑
function processExcelFiles(inputDir, outputDir) {
  console.log(`开始扫描目录: ${inputDir}`);
  
  try {
    const files = fs.readdirSync(inputDir).filter(f => f.endsWith('.xlsx'));
    console.log(`找到 ${files.length} 个Excel文件: ${files.join(', ')}`);
    
    let allData = [];
    
    for (const file of files) {
      console.log(`\n正在处理文件: ${file}`);
      const filePath = path.join(inputDir, file);
      
      try {
        const workbook = xlsx.readFile(filePath);
        console.log(`  - 文件包含 ${workbook.SheetNames.length} 个工作表`);
        
        workbook.SheetNames.forEach(sheetName => {
          console.log(`  - 正在处理工作表: ${sheetName}`);
          const sheet = workbook.Sheets[sheetName];
          
          // 获取工作表的所有数据
          const range = xlsx.utils.decode_range(sheet['!ref']);
          console.log(`  - 工作表范围: ${sheet['!ref']}`);
          
          const rows = [];
          for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            const row = {};
            for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
              const cellAddress = xlsx.utils.encode_cell({r: rowNum, c: colNum});
              const cell = sheet[cellAddress];
              const colName = xlsx.utils.encode_col(colNum);
              row[colName] = cell ? (cell.v || '') : '';
            }
            rows.push(row);
          }
          
          console.log(`  - 总行数: ${rows.length}`);
          
          // 查找包含"点号"的行
          let found = false;
          for (let i = 0; i < Math.min(20, rows.length); i++) {
            const row = rows[i];
            const rowValues = Object.values(row);
            
            console.log(`  - 第${i+1}行前5列:`, rowValues.slice(0, 5));
            
            // 检查是否包含目标关键字
            for (let j = 0; j < rowValues.length; j++) {
              const cellValue = String(rowValues[j]).trim();
              if (cellValue.includes('点号') || cellValue.includes('点 号')) {
                console.log(`  *** 找到点号列: 第${i+1}行第${j+1}列, 内容: "${cellValue}"`);
                found = true;
              }
              if (cellValue.includes('x(m)') || cellValue.includes('X(m)')) {
                console.log(`  *** 找到x列: 第${i+1}行第${j+1}列, 内容: "${cellValue}"`);
              }
              if (cellValue.includes('y(m)') || cellValue.includes('Y(m)')) {
                console.log(`  *** 找到y列: 第${i+1}行第${j+1}列, 内容: "${cellValue}"`);
              }
            }
          }
          
          if (!found) {
            console.log(`  - 警告: 在前20行中未找到包含"点号"的单元格`);
          }
        });
        
      } catch (err) {
        console.log(`  - 处理文件失败: ${err.message}`);
      }
    }
    
  } catch (err) {
    console.log(`扫描目录失败: ${err.message}`);
  }
}

// 测试
processExcelFiles('../00', '../输出'); 