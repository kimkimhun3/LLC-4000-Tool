const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const fs = require('fs');

let win;

const createWindow = () => {
  win = new BrowserWindow({
    width: 800,
    height: 600,
    icon: './jkj.ico',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
      //webSecurity: false,
      //devTools: false // Disable developer tools
    },
    //autoHideMenuBar: true
  });

  win.loadFile('index.html');
};

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// ipcMain.on('open-file-dialog', (event) => {
//   dialog.showOpenDialog(win, {
//     properties: ['openFile'],
//     filters: [{ name: 'Text Files', extensions: ['txt'] }],
//   }).then((file) => {
//     if (!file.canceled && file.filePaths.length > 0) {
//       const content = fs.readFileSync(file.filePaths[0], 'utf-8');
//       event.reply('file-path', file.filePaths[0]);
//       event.reply('file-content', content);
//     }
//   });
// });

const path = require('path');

ipcMain.on('open-file-dialog', (event) => {
  dialog.showOpenDialog(win, {
    properties: ['openFile'],
    filters: [{ name: 'Text Files', extensions: ['txt'] }],
  }).then((file) => {
    if (!file.canceled && file.filePaths.length > 0) {
      const inputFilePath = file.filePaths[0];
      const inputDir = path.dirname(inputFilePath);
      const baseName = path.basename(inputFilePath, '.txt');

      const content = fs.readFileSync(inputFilePath, 'utf-8');
      event.reply('file-path', inputFilePath);
      event.reply('file-content', content);

      // Automatically create the four files in the same directory
      const filesToCreate = [
        { suffix: '_1.txt', eventName: 'file-created' },
        { suffix: '_2.txt', eventName: 'file-created-2' },
        { suffix: '_3.txt', eventName: 'file-created-3' },
        { suffix: '_4.xlsx', eventName: 'file-created-4' },
      ];

      filesToCreate.forEach(({ suffix, eventName }) => {
        const newFilePath = path.join(inputDir, `${baseName}${suffix}`);
        fs.writeFileSync(newFilePath, '', 'utf-8'); // Creates an empty file
        event.reply(eventName, newFilePath);
      });
    }
  });
});


ipcMain.on('open-save-dialog', (event) => {
  dialog.showSaveDialog(win, {
    defaultPath: 'new-file.txt',
    filters: [{ name: 'Text Files', extensions: ['txt'] }],
  }).then((file) => {
    if (!file.canceled) {
      fs.writeFileSync(file.filePath, '', 'utf-8');
      event.reply('file-created', file.filePath);
    }
  });
});

ipcMain.on('open-save-dialog2', (event) => {
  dialog.showSaveDialog(win, {
    defaultPath: 'new-file.txt',
    filters: [{ name: 'Text Files', extensions: ['txt'] }],
  }).then((file) => {
    if (!file.canceled) {
      fs.writeFileSync(file.filePath, '', 'utf-8');
      event.reply('file-created-2', file.filePath);
    }
  });
});

ipcMain.on('open-save-dialog3', (event) => {
  dialog.showSaveDialog(win, {
    defaultPath: 'new-file.txt',
    filters: [{ name: 'Text Files', extensions: ['txt'] }],
  }).then((file) => {
    if (!file.canceled) {
      fs.writeFileSync(file.filePath, '', 'utf-8');
      event.reply('file-created-3', file.filePath);
    }
  });
});

ipcMain.on('open-save-dialog4', (event) => {
  dialog.showSaveDialog(win, {
    defaultPath: 'new-file.xlsx',
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
  }).then((file) => {
    if (!file.canceled) {
      // createExcelFile(file.filePath);
      fs.writeFileSync(file.filePath, '', 'utf-8');
      event.reply('file-created-4', file.filePath);
    }
  });
});

ipcMain.on('save-filesssss', (event, args) => {
  if (args.filePath) fs.writeFileSync(args.filePath, args.content, 'utf-8');
  if (args.filePath2) fs.writeFileSync(args.filePath2, args.content2, 'utf-8');
  if (args.filePath3) fs.writeFileSync(args.filePath3, args.content3, 'utf-8');
  // if (args.filePath4) saveExcelFile(args.filePath4, args.content4, 'utf-8');
});

