// let $ = require('jquery');
// require('popper.js');
// require('bootstrap');
const submitFormButton = document.querySelector("#smtpForm");

document.addEventListener("DOMContentLoaded", function(){
    window.electron.SMTPWindowLoad();
})


submitFormButton.addEventListener("submit", function(event){

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

   console.log(jsonString);

   window.electron.updateSMTP(jsonString);
   
});