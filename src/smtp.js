// let $ = require('jquery');
// require('popper.js');
// require('bootstrap');
const submitFormButton = document.querySelector("#smtpForm");
let verifyConnBtn = document.getElementById("SMTP_tse");
let verifyConnSpinner = document.getElementById("SMTP_spinnerVerfifyconnection");
let smtp_verifyConnection = document.getElementById("smtp_verifyConnection");
verifyConnSpinner.style.display = "none";

document.addEventListener("DOMContentLoaded", function(){
    // smtp_verifyConnection.disabled = true;
    window.electron.SMTPWindowLoad();
})

smtp_verifyConnection.addEventListener('click', _ => {

    verifyConnSpinner.style.display = "block";
    document.getElementById("updateResponse").innerHTML = "";
    

    window.electron.verifySmtpConnInSMTPWin()
})

//submit form
submitFormButton.addEventListener("submit", function(event){
    event.preventDefault();

    document.getElementById("connectionInfo").innerHTML = "";
    
    let host = document.getElementById("inputHost").value;
    let username = document.getElementById("inputUsername").value;
    let password = document.getElementById("inputPassword").value;
    let port = document.getElementById("inputPortNumber").value;

    var obj = new Object();
        obj.host = host;
        obj.username  = username;
        obj.password = password;
        obj.port = port;

   var jsonString= JSON.stringify(obj);

//    console.log(jsonString);

   window.electron.updateSMTP(jsonString);
   
});