const submitFormButton = document.querySelector("#fileUploadForm");
let btnSendEmail = document.getElementById("sendEmailReadybtn");
let spinner = document.getElementById("sendEmailSpinner");


document.addEventListener("DOMContentLoaded", function(){
    spinner.style.display = "none";
    btnSendEmail.disabled = true;
    window.electron.MainWindowLoad();
})
submitFormButton.addEventListener("submit", function(event){
    event.preventDefault();   // stop the form from submitting

    let inputVal = document.getElementById('fileInput').files[0].path;
    window.electron.excelFileToUpload(inputVal)
})


btnSendEmail.addEventListener('click', _ => {

    spinner.style.display = "block";
    let obj = new Object();
    obj.host = document.getElementById('h_host').value;
    obj.username = document.getElementById('h_username').value;
    obj.password = document.getElementById('h_password').value;
    obj.port = document.getElementById('h_port').value;

    // window.electron.sendEmailTrigger(obj)
    window.electron.sendEmailTrigger()
})
