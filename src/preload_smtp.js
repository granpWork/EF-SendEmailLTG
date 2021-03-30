const { contextBridge, ipcRenderer } = require("electron")
const path = require('path')
const nodemailer = require('nodemailer')

contextBridge.exposeInMainWorld(
  'electron',
  {
    updateSMTP: (argsJson) => ipcRenderer.send('update-smtp-json', argsJson),
    SMTPWindowLoad: () => {ipcRenderer.send("SMTPWindowLoaded")},
    openSMTPSettingsWin: () => ipcRenderer.send("showSMTPSetting"),
    verifySmtpConnInSMTPWin: () => ipcRenderer.send("verify-smtp-in-smtpwin")
  }
)

ipcRenderer.on('smtpwin-row-request', (event, arg) => {
  let verifyConnSpinner = document.getElementById("SMTP_spinnerVerfifyconnection");

  let responseMessage;
  let host = arg[0].host;
  let username = arg[0].username;
  let password = arg[0].password;
  let port = arg[0].usernportame;

  let transporter = nodemailer.createTransport({
    host: host,
    port: port,
    secure: false, // true for 465, false for other ports
    auth: {
      user: username, // generated ethereal user
      pass: password, // generated ethereal password
    },
  })

  transporter.verify(function(error, success) {
    verifyConnSpinner.style.display = "none";
    if (error) {
      responseMessage = error.message;
    } else {
      responseMessage = "Server is ready to take our messages";
    }
    document.getElementById("connectionInfo").innerHTML = "<span style='color: black;'>"+responseMessage+"</span>";
  })

})

ipcRenderer.on("SMTP-resultSent", function(evt, result){
  document.getElementById("inputUsername").value = result[0].username;
  document.getElementById("inputPassword").value = result[0].password;
  document.getElementById("inputHost").value = result[0].host;
  document.getElementById("inputPortNumber").value = result[0].port;
})

ipcRenderer.on("updateResponse", function(evt){

  // document.getElementById("updateRes").innerHTML = "<span style='color: black;'>Successfully Updated.</span>";
})

ipcRenderer.on("updateResponse-reply", function(evt, result){
  // var para = document.createElement("P");
  // var t = document.createTextNode(result);
  // para.appendChild(t);
  // document.getElementById("updateResponse").appendChild(para);
console.log(result)

// if(result !== null){
//   var para = document.createElement("P");
//   var t = document.createTextNode(result);
//   para.appendChild(t);
//   document.getElementById("updateResponse").appendChild(para);

//   document.getElementById("updateResponse").innerHTML = result;
// }
document.getElementById("updateResponse").innerHTML = result;
})

