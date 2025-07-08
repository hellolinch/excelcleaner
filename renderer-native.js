const { dialog } = require('@electron/remote');
const { ipcRenderer } = require('electron');

let inputDir = '';
let outputDir = '';
let processing = false;

function updateLog(message) {
  const logArea = document.getElementById('log');
  logArea.value += '\n' + message;
  logArea.scrollTop = logArea.scrollHeight;
}

function updateButtons() {
  const buttons = document.querySelectorAll('button');
  buttons.forEach(btn => {
    btn.disabled = processing;
  });
}

async function handleSelectInput() {
  const result = await dialog.showOpenDialog({ properties: ['openDirectory'] });
  if (!result.canceled && result.filePaths.length > 0) {
    inputDir = result.filePaths[0];
    document.getElementById('input-path').textContent = inputDir;
    updateLog(`已选择输入文件夹：${inputDir}`);
  }
}

async function handleSelectOutput() {
  const result = await dialog.showOpenDialog({ properties: ['openDirectory'] });
  if (!result.canceled && result.filePaths.length > 0) {
    outputDir = result.filePaths[0];
    document.getElementById('output-path').textContent = outputDir;
    updateLog(`已选择输出文件夹：${outputDir}`);
  }
}

async function handleStart() {
  if (!inputDir || !outputDir) {
    updateLog('请先选择输入和输出文件夹！');
    return;
  }
  processing = true;
  updateButtons();
  updateLog('开始处理Excel文件...');
  
  try {
    const result = await ipcRenderer.invoke('process-excel', inputDir, outputDir);
    if (result.success) {
      updateLog('✓ ' + result.message);
    } else {
      updateLog('✗ 处理失败：' + result.message);
    }
  } catch (err) {
    updateLog('✗ 处理异常：' + err.message);
  }
  
  processing = false;
  updateButtons();
}

function handleStop() {
  updateLog('[待实现] 停止处理');
}

function handleClearLog() {
  document.getElementById('log').value = '';
}

// 监听主进程发送的实时日志
ipcRenderer.on('log-update', (event, message) => {
  updateLog(message);
});

// 页面加载完成后绑定事件
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('select-input').addEventListener('click', handleSelectInput);
  document.getElementById('select-output').addEventListener('click', handleSelectOutput);
  document.getElementById('start-btn').addEventListener('click', handleStart);
  document.getElementById('stop-btn').addEventListener('click', handleStop);
  document.getElementById('clear-btn').addEventListener('click', handleClearLog);
}); 