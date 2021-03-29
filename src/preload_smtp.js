const { contextBridge, ipcRenderer } = require("electron")
const path = require('path')
const remote = require('electron').remote

contextBridge.exposeInMainWorld(
  'electron',
  {
    updateSMTP: (argsJson) => ipcRenderer.send('update-smtp-json', argsJson),
    SMTPWindowLoad: () => {ipcRenderer.send("SMTPWindowLoaded")},
    openSMTPSettingsWin: () => ipcRenderer.send("showSMTPSetting")
  }
)

ipcRenderer.on("SMTP-resultSent", function(evt, result){
  document.getElementById("inputUsername").value = result[0].username;
  document.getElementById("inputPassword").value = result[0].password;
  document.getElementById("inputHost").value = result[0].host;
  document.getElementById("inputPortNumber").value = result[0].port;
})

ipcRenderer.on("updateResponse", function(evt, res){
  document.getElementById("updateRes").innerHTML = res;
})