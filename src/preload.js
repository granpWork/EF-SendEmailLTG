const { contextBridge, ipcRenderer } = require("electron")
const path = require('path')
const remote = require('electron').remote

contextBridge.exposeInMainWorld(
  'electron',
  {
    excelFileToUpload: (file) => ipcRenderer.send('excel-file', file),
    MainWindowLoad: () => ipcRenderer.send('mainwindow-loaded'),
    verifySmtpConn: (args) => ipcRenderer.send('verify-smtp-connection', args),
    sendEmailTrigger: (args) => ipcRenderer.send('send-mail-trigger',args)
  }
)

ipcRenderer.on("mainwindow-hidden-smtp", function(evt, result){
  document.getElementById("h_host").value = result[0].host;
  document.getElementById("h_username").value = result[0].username;
  document.getElementById("h_password").value = result[0].password;
  document.getElementById("h_port").value = result[0].port;

  // console.log("updated: hidden input_>"+ result[0].host)
  // console.log("updated: hidden input_>"+ result[0].username)
  // console.log("updated: hidden input_>"+ result[0].password)
  // console.log("updated: hidden input_>"+ result[0].port)
})

ipcRenderer.on('res-upload-excel', (event, arg) => {
  let target = document.getElementById("tse");
  let recordCount = document.getElementById("recordCount");

  // recordCount.innerHTML = "<span style='color: green;'>Total Record: "+arg+"</span>"
  
  if(arg != 0 && target.style.display === "none"){
    target.style.display = "block";
  } else if (arg == 0 && target.style.display === "block") {
    target.style.display = "none";
  }

})


ipcRenderer.on('verify-smtp-connection-response', (event, arg) => {
  document.getElementById("spinnerVerfifyconnection").style.display = "none";

  document.getElementById("connectionInfo").innerHTML = "<span style='color: black;'>"+arg+"</span>";
});

ipcRenderer.on("email-sent-response", (event, arg) => {
  console.log(arg)
  document.getElementById("connectionInfo").innerHTML = "<span style='color: black;'>"+arg+"</span>";
  document.getElementById("spinnerVerfifyconnection").style.display = "none";
});

