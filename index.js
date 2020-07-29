const { app, BrowserWindow } = require('electron')
const { ipcMain } = require('electron');

let win
let secondWindow;

function createWindow () {
  win = new BrowserWindow({ width: 1400, height: 600, 'minWidth': 1200 })
  win.loadFile('index.html')
  win.webContents.openDevTools()
  win.on('closed', () => {
    win = null
  })
}

app.on('ready', createWindow)
app.on('window-all-closed', ()=>{
  app.quit()
});

ipcMain.on('create-window',(event, arg)=>{
     secondWindow = new BrowserWindow({
         height: 500,
         width: 700,
         parent: win,
		     modal: true
     });
     secondWindow.custom = {
         'file_obj': arg,
     };
     secondWindow.loadFile('./templates/sort.html')
//     secondWindow.webContents.openDevTools()
     secondWindow.on('closed', () => {
       secondWindow = null
     })
     event.returnValue = 'Done'
})

ipcMain.on('send-updated-list',(event, arg)=>{
  console.log("Here it is", arg)
  win.webContents.send('update-stock' , arg);
  secondWindow.close()
})
