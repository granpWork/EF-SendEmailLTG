const submitFormButton = document.querySelector("#fileUploadForm");
let SendEmailDIV = document.getElementById("tse");
let btnSendEmail = document.getElementById("sendEmailbtn");
let verifyConnection = document.getElementById("verifyConnection");

SendEmailDIV.style.display = "none";

document.addEventListener("DOMContentLoaded", function(){
    window.electron.MainWindowLoad();
})


submitFormButton.addEventListener("submit", function(event){
    event.preventDefault();   // stop the form from submitting

    let inputVal = document.getElementById('fileInput').files[0].path;
    window.electron.excelFileToUpload(inputVal)
})

verifyConnection.addEventListener('click', _ => {
    let obj = new Object();
    obj.host = document.getElementById('h_host').value;
    obj.username = document.getElementById('h_username').value;
    obj.password = document.getElementById('h_password').value;
    obj.port = document.getElementById('h_port').value;

    window.electron.verifySmtpConn(obj)
})

btnSendEmail.addEventListener('click', _ => {

    let obj = new Object();
    obj.host = document.getElementById('h_host').value;
    obj.username = document.getElementById('h_username').value;
    obj.password = document.getElementById('h_password').value;
    obj.port = document.getElementById('h_port').value;

    window.electron.sendEmailTrigger(obj)
})
