const { shell, ipcRenderer } = require('electron')
const path = require('path');
const animation = require('./js/animations.js');
const data = require('./js/financialPL.js');
//const srcFilePath = path.join(process.resourcesPath, 'workbooks', 'financial_pl', 'source.xlsx');
//const newFilePath = path.join(process.resourcesPath, 'workbooks', 'financial_pl', 'new.xlsx');
const srcFilePath = path.join(path.dirname(__dirname), 'workbooks', 'financial_pl', 'source.xlsx');
const newFilePath = path.join(path.dirname(__dirname), 'workbooks', 'financial_pl', 'new.xlsx');
//const app = require('electron').remote.app;

prepBtn.addEventListener('click', () => {
  animation.prepToRun();
  shell.openItem(srcFilePath);
});

runBtn.addEventListener('click', () => {
  data.Reformat();
});

openBtn.addEventListener('click', () => {
  shell.openItem(newFilePath);
  ipcRenderer.send('close-me')
});