const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// 调试Excel文件结构
function debugExcelFile(filePath) {
  console.log(`\n=== 调试文件: ${filePath} ===`);
  
  try {
    const workbook = xlsx.readFile(filePath);
    console.log(`Sheet数量: ${workbook.SheetNames.length}`);
    console.log(`Sheet名称: ${workbook.SheetNames.join(', ')}`);
    
    workbook.SheetNames.forEach((sheetName, index) => {
      console.log(`\n--- Sheet ${index + 1}: ${sheetName} ---`);
      const sheet = workbook.Sheets[sheetName];
      const json = xlsx.utils.sheet_to_json(sheet, { defval: '' });
      
      console.log(`行数: ${json.length}`);
      if (json.length > 0) {
        console.log('前3行数据:');
        json.slice(0, 3).forEach((row, i) => {
          console.log(`第${i+1}行:`, row);
        });
        
        console.log('列名:', Object.keys(json[0]));
        
        // 检查是否包含目标列
        const pointCol = Object.keys(json[0]).find(k => k.replace(/\s/g, '').includes('点号'));
        const xCol = Object.keys(json[0]).find(k => k.includes('x(m)'));
        const yCol = Object.keys(json[0]).find(k => k.includes('y(m)'));
        
        console.log('找到的列:');
        console.log(`  点号列: ${pointCol || '未找到'}`);
        console.log(`  x(m)列: ${xCol || '未找到'}`);
        console.log(`  y(m)列: ${yCol || '未找到'}`);
      }
    });
  } catch (err) {
    console.error('读取文件出错:', err.message);
  }
}

// 调试指定目录下的所有Excel文件
function debugDirectory(dirPath) {
  console.log(`调试目录: ${dirPath}`);
  
  try {
    const files = fs.readdirSync(dirPath).filter(f => f.endsWith('.xlsx'));
    console.log(`找到 ${files.length} 个Excel文件: ${files.join(', ')}`);
    
    files.forEach(file => {
      const filePath = path.join(dirPath, file);
      debugExcelFile(filePath);
    });
  } catch (err) {
    console.error('读取目录出错:', err.message);
  }
}

// 调试您的Excel文件
debugDirectory('../00'); 