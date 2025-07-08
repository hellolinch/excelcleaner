import React, { useState } from 'react';
const { dialog } = window.require('@electron/remote');
const { ipcRenderer } = window.require('electron');

function App() {
  const [inputDir, setInputDir] = useState('');
  const [outputDir, setOutputDir] = useState('');
  const [log, setLog] = useState('');
  const [processing, setProcessing] = useState(false);

  // 选择文件夹的回调
  const handleSelectInput = async () => {
    const result = await dialog.showOpenDialog({ properties: ['openDirectory'] });
    if (!result.canceled && result.filePaths.length > 0) {
      setInputDir(result.filePaths[0]);
      setLog(log + `\n已选择输入文件夹：${result.filePaths[0]}`);
    }
  };
  const handleSelectOutput = async () => {
    const result = await dialog.showOpenDialog({ properties: ['openDirectory'] });
    if (!result.canceled && result.filePaths.length > 0) {
      setOutputDir(result.filePaths[0]);
      setLog(log + `\n已选择输出文件夹：${result.filePaths[0]}`);
    }
  };
  const handleStart = async () => {
    if (!inputDir || !outputDir) {
      setLog(log + '\n请先选择输入和输出文件夹！');
      return;
    }
    setProcessing(true);
    setLog(log + '\n正在处理，请稍候...');
    try {
      const result = await ipcRenderer.invoke('process-excel', inputDir, outputDir);
      if (result.success) {
        setLog(l => l + `\n${result.message}`);
      } else {
        setLog(l => l + `\n处理失败：${result.message}`);
      }
    } catch (err) {
      setLog(l => l + `\n处理异常：${err.message}`);
    }
    setProcessing(false);
  };
  const handleStop = () => {
    setLog(log + '\n[待实现] 停止处理');
  };
  const handleClearLog = () => {
    setLog('');
  };

  return (
    <div style={{ width: 600, margin: '40px auto', fontFamily: 'sans-serif' }}>
      <h2>Excel数据清洗合并工具</h2>
      <div style={{ marginBottom: 16 }}>
        输入文件夹：
        <button onClick={handleSelectInput} disabled={processing}>选择文件夹</button>
        <span style={{ marginLeft: 8, color: '#888' }}>{inputDir}</span>
      </div>
      <div style={{ marginBottom: 16 }}>
        输出文件夹：
        <button onClick={handleSelectOutput} disabled={processing}>选择文件夹</button>
        <span style={{ marginLeft: 8, color: '#888' }}>{outputDir}</span>
      </div>
      <div style={{ marginBottom: 16 }}>
        <button onClick={handleStart} disabled={processing}>开始处理</button>
        <button onClick={handleStop} style={{ marginLeft: 8 }} disabled={processing}>停止</button>
        <button onClick={handleClearLog} style={{ marginLeft: 8 }} disabled={processing}>清空日志</button>
      </div>
      <div>
        <div>处理日志：</div>
        <textarea
          value={log}
          readOnly
          style={{ width: '100%', height: 200, resize: 'vertical', fontFamily: 'monospace' }}
        />
      </div>
    </div>
  );
}

export default App; 