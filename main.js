const { app, BrowserWindow, ipcMain } = require('electron')
//const path = require('path')

function createWindow () {
  let win = new BrowserWindow({
    width: 362,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      //preload: path.join(__dirname, 'preload.js');
    }
  });
  win.on('will-resize', (e) => {
    e.preventDefault();
  });
  win.center();
  win.removeMenu();
  //win.setResizable(false);
  win.loadFile('renderer/index.html');
  ipcMain.on('close-me', (evt, arg) => {
    app.quit()
  })
};

app.on('ready', function(){
  createWindow();
});
