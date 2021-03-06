const { contextBridge, ipcRenderer } = require("electron")
const path = require('path')
const nodemailer = require('nodemailer')
const formatCurrency = require('format-currency');

contextBridge.exposeInMainWorld(
  'electron',
  {
    excelFileToUpload: (file) => ipcRenderer.send('excel-file', file),
    MainWindowLoad: () => ipcRenderer.send('mainwindow-loaded'),
    verifySmtpConn: (args) => ipcRenderer.send('verify-smtp-connection', args),
    sendEmailTrigger: () => ipcRenderer.send('send-mail-trigger-SMTPTable_result')
  }
)

ipcRenderer.on('SMTPRow-Response-error', (event,err) => {
  console.log("from SMTPRow-Response-error -> "+err);
  let sendEmailSpinner = document.getElementById("sendEmailSpinner");
  sendEmailSpinner.style.display = "none";
  document.getElementById("emailSentResMsg").innerHTML = "<span style='color: black;'>Cant Connect to SMTP:"+err+"</span>";
})

ipcRenderer.on('SMTPRow-Response', (event, arg, data) => {
  let sendEmailSpinner = document.getElementById("sendEmailSpinner");
  let objSMTP = new Object();
  objSMTP.host = arg[0].host;
  objSMTP.username = arg[0].username;
  objSMTP.password = arg[0].password;
  objSMTP.port = arg[0].usernportame;

  let opts = { format: '%v %c', code: 'Php' }
  let obj = new Object();

    data.forEach(dt => {
      obj.firstName = `${dt["First Name"]}`
      obj.lastName = `${dt["Last Name"]}`
      obj.middleName = `${dt["Middle Name"]}`
      obj.email = `${dt["Email Address"]}`
      obj.companyCode = `${dt["Company Code"]}`
      obj.empNumber = `${dt["Employee Number"]}`
      obj.comapny = getCompanyName(obj.companyCode);
      obj.name = `${dt["First Name"]}`+" "+`${dt["Middle Name"]}`+" "+`${dt["Last Name"]}`;
      obj.modernaOrders = `${dt["For how many people are you reserving Moderna vaccines?"]}`;
      obj.covovaxOrders = `${dt["For how many people are you reserving Covovax (Novavax) vaccines?"]}`;
      obj.modernaControlNumber = `${dt["Moderna Reservation Control Numbers"]}`
      obj.covovaxControlNumber = `${dt["Covovax Reservation Control Numbers"]}`
      obj.modernaTotalCost = formatCurrency(modernaTotalCost(obj.modernaOrders), opts);
      obj.covovaxTotalCost = formatCurrency(covovaxTotalCost(obj.covovaxOrders), opts);
      obj.regLink = registrationURLMapping(`${dt["Company Code"]}`);
      obj.vaccineTotalCost = formatCurrency(vaccineTotalCost(modernaTotalCost(obj.modernaOrders), covovaxTotalCost(obj.covovaxOrders)), opts);
    
      if(`${dt["Email Address"]}` !== "undefined"){
        main(objSMTP, JSON.stringify(obj)).catch(console.error);
        // console.log("json _> "+JSON.stringify(obj));
      }
  })

  sendEmailSpinner.style.display = "none";


  
})

async function verifySMTP(smtp_c){

  let responseMessage;
  let host = smtp_c.host;
  let username = smtp_c.username;
  let password = smtp_c.password;
  let port = smtp_c.port;

  let isError = false;

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
    if (error) {
      console.log(error.message);

      responseMessage = error.message;
      document.getElementById("emailSentResMsg").innerHTML= "<span style='color: black;'>Error: Please check your SMTP Credential.</span>";
      isError = true;
    } else {
      console.log("Server is ready to take our messages");
      // console.log(success);
      responseMessage = "Server is ready to take our messages";

      isError = false;
    }
  })

  return isError;
}

async function main(smtp_c, obj_json) {
  let obj = JSON.parse(obj_json);

  let host = smtp_c.host;
  let username = smtp_c.username;
  let password = smtp_c.password;
  let port = smtp_c.port;

  let email = obj.email;
  let name = obj.name;
  let employeeID = obj.empNumber;
  let company = obj.comapny;
  let modernaOrders = obj.modernaOrders;
  let covovaxOrders = obj.covovaxOrders;
  let modernaControlNumbers = obj.modernaControlNumber;
  let covovaxControlNumbers = obj.covovaxControlNumber;
  let registrationURL = obj.regLink;
  let vaccineTotalCost = obj.vaccineTotalCost;
  let modernaTotalCost = obj.modernaTotalCost;
  let covovaxTotalCost = obj.covovaxTotalCost;


  console.log(name);
  let transporter = nodemailer.createTransport({
    host: host,
    port: port,
    secure: false, // true for 465, false for other ports
    auth: {
      user: username, // generated ethereal user
      pass: password, // generated ethereal password
    },
  })

  // send mail with defined transport object
  let info = await transporter.sendMail({
    from: '"LTG" <'+username+'>', // sender address
    to: email, // list of receivers
    subject: "Vaccine Orders", // Subject line
    text: "Hello world?", // plain text body
    html: `<p style="line-height:18px;">Dear <b>${name},</b><br /> Employee Number: ${employeeID}<br /> ${company}<br/></p> <p>This is to confirm the orders you have placed are as follows:</p> <p style="margin-left: 40px">${modernaOrders} <b>MODERNA</b></p> <p style="margin-left: 60px;line-height:18px;">With control numbers:<br /><span style="margin-left: 0px;"></span><b>${modernaControlNumbers}</b></p> <p style="margin-left: 40px;">And</p> <p style="margin-left: 40px">${covovaxOrders} <b>COVOVAX (NOVAVAX)</b></p> <p style="margin-left: 60px;line-height:18px;">With control numbers:<br /><span style="margin-left: 0px;"><b>${covovaxControlNumbers}</b></span></p> <p>Please ensure that the provided control numbers are used as you register your household / family members' information. Below is the link to the Registration form.</p> <p>CLICK HERE TO REGISTER	 : ${registrationURL}</URL></p> <p>The total cost of your order is <b>${vaccineTotalCost}</b>. Please see the breakdown below:</p> <p style="line-height:18px;"> Moderna = ${modernaOrders} x Php 3,700 = Php ${modernaTotalCost}<br /> Covovax (Novavax) = ${covovaxOrders} x Php 3,000 = Php ${covovaxTotalCost} </p> <p>As you have confirmed and agreed to upon reserving the vaccine orders, the total amount will be fully shouldered by you via the approved payment method/scheme that you had with your company.</p> <p>Please advise your preferred payment scheme by replying to this email. You will be asked to sign/agree to an Authority to Deduct form.</p> <br /> <p>Thank you.</p> <p>${company} HR Department/Management</p>`, // html body
  }).then(function(info){
    var para = document.createElement("P");
    var t = document.createTextNode("Email sent to: "+info.envelope.to[0]);
    para.appendChild(t);
    document.getElementById("emailSentResMsg").appendChild(para);

    document.getElementById('fileInput').value="" ; 
  })   // if successful
  .catch(function(err){console.log('got error'); 
    console.log(err)
    document.getElementById('fileInput').value="" ; 
  });   // if an error occurs;

  // console.log("Message sent: %s", info.envelo);
  // Message sent: <b658f8ca-6296-ccf4-8306-87d57a0b4321@example.com>
}



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
  let btnSendEmail = document.getElementById("sendEmailReadybtn");
  let recordCount = document.getElementById("recordCount");

  if(arg){ 
    btnSendEmail.disabled = false;
  }

  recordCount.innerHTML = "<span style='color: green;'>Total Records: "+arg+"</span>"
 
  console.log(arg);
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

function getCompanyName(companyCode){
  let cm = new Map([
    ["ALL","All Seasons Realty Corp."],
    ["APL","Allianz-PNB Life Insurance, Inc. (APLII)"],
    ["ABI","Asia Brewery, Inc. (ABI) and Subsidiaries"],
    ["BCH","Basic Holdings Corp."],
    ["CPH","Century Park Hotel"],
    ["EPP","Eton Properties Philippines, Inc. (Eton) and Subsidiaries"],
    ["FFI","Foremost Farms, Inc."],
    ["FTC","Fortune Tobacco Corp."],
    ["GDC","Grandspan Development Corp."],
    ["HII","Himmel Industries, Inc."],
    ["LRC","Landcom Realty Corp."],
    ["LTG","LT Group, Inc. (Parent Company)"],
    ["LTGC","LTGC Directors"],
    ["MAC","MacroAsia Corp., Subsidiaries & Affiliates"],
    ["PAL","Philippine Airlines, Inc. (PAL), Subsidiaries and Affiliates"],
    ["PNB","Philippine National Bank (PNB) and Subsidiaries"],
    ["PMI","PMFTC Inc."],
    ["RAP","Rapid Movers & Forwarders, Inc."],
    ["TYK","Tan Yan Kee Foundation, Inc. (TYKFI)"],
    ["TDI","Tanduay Distillers, Inc. (TDI) and subsidiaries"],
    ["CHI","Charter House Inc."],
    ["SPV","SPV Group"],
    ["TMC","Topkick Movers Corporation"],
    ["UNI","University of the East (UE)"],
    ["UER","University of the East Ramon Magsaysay Memorial Medical Center (UERMMMC)"],
    ["VMC","Victorias Milling Company, Inc. (VMC)"],
    ["ZHI","Zebra Holdings, Inc."],
    ["DIR","Directors of LTGC"],
    ["STN","Sabre Travel Network Phils., Inc."]
  ]);

  if(!cm.has(companyCode)){
    return '--'
  }
  return cm.get(companyCode);
}

function generateControlNumbers(company, empid, vaccine, vaccineOrders){
  var controlNumber = [];
  let initVaccine = "";

  if(vaccine == "moderna") {
    initVaccine = "_M";
  }else {
    initVaccine = "_C";
  }

  for(i = 1; i <= vaccineOrders; i++) {
    controlNumber.push(company+"_"+empid+initVaccine+i);
  }

  let revCTRLNum = controlNumber.reverse();
  let modernaControlNumbers = "";
  revCTRLNum.forEach(element => modernaControlNumbers = element + "<br>"+ modernaControlNumbers);

  return modernaControlNumbers;
}

function modernaTotalCost(vaccineOrder){
  return vaccineOrder * 3700;
}

function covovaxTotalCost(vaccineOrder){
  return vaccineOrder * 3000;
}

function vaccineTotalCost(modernaTotalCost, covovaxTotalCost){
  return modernaTotalCost + covovaxTotalCost;
}

function registrationURLMapping(companyCode){
  let cm = new Map([
    ["ALL","<a href='registerallwellness.ltg.com.ph'>registerallwellness.ltg.com.ph</a>"],
    ["APL","<a href='registeraplsignup.ltg.com.ph'>registeraplsignup.ltg.com.ph</a>"],
    ["ABI","<a href='registerabihealth.ltg.com.ph'>registerabihealth.ltg.com.ph</a>"],
    ["BCH","<a href='registerbhccare.ltg.com.ph'>registerbhccare.ltg.com.ph</a>"],
    ["CPH","<a href='registercphregister.ltg.com.ph'>registercphregister.ltg.com.ph</a>"],
    ["CHI","<a href='registerchistronger.ltg.com.ph'>registerchistronger.ltg.com.ph</a>"],
    ["EPP","<a href='registeretoncares.ltg.com.ph'>registeretoncares.ltg.com.ph</a>"],
    ["FFI","<a href='registerffisignature.ltg.com.ph'>registerffisignature.ltg.com.ph</a>"],
    ["FTC","<a href='registerftcsecure.ltg.com.ph'>registerftcsecure.ltg.com.ph</a>"],
    ["GDC","<a href='registergdccure.ltg.com.ph'>registergdccure.ltg.com.ph</a>"],
    ["HII","<a href='registerhiienroll.ltg.com.ph'>registerhiienroll.ltg.com.ph</a>"],
    ["LRC","<a href='registerlrcrecord.ltg.com.ph'>registerlrcrecord.ltg.com.ph</a>"],
    ["LTG","<a href='registerone.ltg.com.ph'>registerone.ltg.com.ph</a>"],
    // ["LTGC","<a href=''></a>"],
    ["MAC","<a href='registermacfile.ltg.com.ph'>registermacfile.ltg.com.ph</a>"],
    ["PAL","<a href='registerpalenlist.ltg.com.ph'>registerpalenlist.ltg.com.ph</a>"],
    ["PNB","<a href='registerpnbprotect.ltg.com.ph'>registerpnbprotect.ltg.com.ph</a>"],
    ["PMI","<a href='registerpmisafe.ltg.com.ph'>registerpmisafe.ltg.com.ph</a>"],
    ["RAP","<a href='registerrapidshield.ltg.com.ph'>registerrapidshield.ltg.com.ph</a>"],
    ["STN","<a href='registerstnsafety.ltg.com.ph'>registerstnsafety.ltg.com.ph</a>"],
    ["TYK","<a href='registertykdefense.ltg.com.ph'>registertykdefense.ltg.com.ph</a>"],
    ["TDI","<a href='registertdilisting.ltg.com.ph'>registertdilisting.ltg.com.ph</a>"],
    ["SPV","<a href='registerspvinvestment.ltg.com.ph'>registerspvinvestment.ltg.com.ph</a>"],
    // ["TMC","<a href=''></a>"],
    ["UNI","<a href='registeruefight.ltg.com.ph'>registeruefight.ltg.com.ph</a>"],
    ["UER","<a href='registeruermimmune.ltg.com.ph'>registeruermimmune.ltg.com.ph</a>"],
    ["VMC","<a href='registervmcregistry.ltg.com.ph'>registervmcregistry.ltg.com.ph</a>"],
    ["ZHI","<a href='registerzhiwin.ltg.com.ph'>registerzhiwin.ltg.com.ph</a>"]
    // ["DIR","<a href=''></a>"]
  ]);

  if(!cm.has(companyCode)){
    return '--'
  }
  return cm.get(companyCode);
}

