const submitFormButton = document.querySelector("#fileUploadForm");
let SendEmailDIV = document.getElementById("tse");
let spinner = document.getElementById("spinnerVerfifyconnection");
let btnSendEmail = document.getElementById("sendEmailbtn");
let verifyConnection = document.getElementById("verifyConnection");
let connectionInfo = document.getElementById("connectionInfo");

SendEmailDIV.style.display = "none";
spinner.style.display = "none";

document.addEventListener("DOMContentLoaded", function(){
    window.electron.MainWindowLoad();
})


submitFormButton.addEventListener("submit", function(event){
    event.preventDefault();   // stop the form from submitting

    let inputVal = document.getElementById('fileInput').files[0].path;
    window.electron.excelFileToUpload(inputVal)
})

verifyConnection.addEventListener('click', _ => {

    spinner.style.display = "block";
    let obj = new Object();
    obj.host = document.getElementById('h_host').value;
    obj.username = document.getElementById('h_username').value;
    obj.password = document.getElementById('h_password').value;
    obj.port = document.getElementById('h_port').value;

    window.electron.verifySmtpConn(obj)
})

btnSendEmail.addEventListener('click', _ => {

    spinner.style.display = "block";
    connectionInfo.style.display = "node";
    let obj = new Object();
    obj.host = document.getElementById('h_host').value;
    obj.username = document.getElementById('h_username').value;
    obj.password = document.getElementById('h_password').value;
    obj.port = document.getElementById('h_port').value;

    window.electron.sendEmailTrigger(obj)
})
