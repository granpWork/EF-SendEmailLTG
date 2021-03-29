const {app, Menu, BrowserWindow, ipcMain} = require('electron')
const path = require('path');
let $ = require("jquery") 
const reader = require('xlsx')
const nodemailer = require('nodemailer')
const formatCurrency = require('format-currency')

let excelData;
let smtpResult;
module.exports = {
  getExcelData: getExcelData,
  setExcelData: setExcelData,
  getSmtpResult: getSmtpResult,
  setSmtpResult: setSmtpResult,
}
function getExcelData() { 
  return excelData;
}
function setExcelData(data) { 
  excelData = data;
}

function getSmtpResult() { 
  return smtpResult;
}
function setSmtpResult(data) { 
  smtpResult = data;
}

//creating sqlite3 database
const knex = require('knex')({
  client: 'sqlite3',
  connection: {
     filename: "./database.sqlite"
  },
  useNullAsDefault: true
});

//creating SMTP Table if there is no table and insert data if table is newly created
knex.schema.hasTable('smtp').then(function(exists) {
  if (!exists) {
    // console.log("creating table");
    return knex.schema.createTable('smtp', function(t) {
      t.increments('id').primary();
      t.string('host', 100);
      t.string('username', 100);
      t.string('password', 100);
      t.int('port', 5);
    }).then(function(){
      return knex('smtp').insert([
        {host: "smtp.service.com", username: "appsuport@service.com", password: "1234567", port: "32"}
      ]);
    });
  }
});
function getSMTPTable(){
  return knex.select('host','username', 'password', 'port').from('smtp');
}

const createSMTPWindow = (parentWindow) => {
  // Create the browser window.
  const SMTPWindow = new BrowserWindow({
    width: 400,
    height: 600,
    parent:parentWindow,
    webPreferences: {
      nodeIntegration: true,
      preload: path.join(__dirname, 'preload_smtp.js')
    },
    title: "SMTP Setting"
  });

  // and load the index.html of the app.
  SMTPWindow.loadFile(path.join(__dirname, 'smtp.html'));

  SMTPWindow.on('closed', (e) =>{
    e.preventDefault()

    updateHiddenInput(parentWindow);
  })


  SMTPWindow.setMenu(null)
  // Open the DevTools.
  // SMTPWindow.webContents.openDevTools()

  SMTPWindow.onber

  ipcMain.on("SMTPWindowLoaded", function () {
    let resultSMTP = getSMTPTable();
    resultSMTP.then(function(rows){
        try {
          SMTPWindow.webContents.send("SMTP-resultSent", rows)
        } catch (e) {
          console.log(e);
        }
      })
  })

  

  ipcMain.on('update-smtp-json', (event, args) => {
    let obj = JSON.parse(args);
    knex('smtp')
    .where({ id: 1 })
    .update({ host: obj.host, username: obj.username, password: obj.password, port: obj.port}).then(function(evt, rows){
      let resposeMsg = "Successfully Updated.";

      updateHiddenInput(parentWindow);
  
      SMTPWindow.webContents.send("updateResponse", resposeMsg)
    })

  });
}

function updateHiddenInput(parentWindow){
  let resultSMTP = getSMTPTable();
  resultSMTP.then(function(rows){
      try {
        parentWindow.webContents.send("mainwindow-hidden-smtp", rows)
        // console.log(rows)
      } catch (e) {
        console.log(e);
      }
    })
}

ipcMain.on('excel-file', (event, args) => {

  const file = reader.readFile(args)
  let data = []
  const sheets = file.SheetNames

  for(let i = 0; i < sheets.length; i++)
  {
    const temp = reader.utils.sheet_to_json(
          file.Sheets[file.SheetNames[i]])
    temp.forEach((res) => {
        data.push(res)
    })
  }
  // console.log(data)

  setExcelData(data)

  // data.forEach(dt => {
  //   console.log(`${dt.COMPANY} : ${dt.NAME}`)
  // })
  
  event.reply('res-upload-excel',Object.keys(getExcelData()).length)
})

ipcMain.on('send-mail-trigger', (event, smtpCred) => {
  // console.log(getExcelData());
  // console.log("args_>"+smtpCred.host);
  verifySMTP(smtpCred);
  let opts = { format: '%v %c', code: 'Php' }
  let obj = new Object();

  getExcelData().forEach(dt => {
    obj.email = `${dt["EMAIL ADDRESS"]}`
    obj.comapny = getCompanyName(`${dt.COMPANY}`);
    obj.companyCode = `${dt.COMPANY}`;
    obj.empNumber = `${dt["EMPLOYEE ID"]}`;
    obj.name = `${dt.NAME}`;
    obj.modernaOrders = `${dt["MODERNA ORDERS"]}`;
    obj.covovaxOrders = `${dt["COVOVAX ORDERS"]}`;
    obj.modernaControlNumber = generateControlNumbers(`${dt.COMPANY}`, `${dt["EMPLOYEE ID"]}`, "moderna", `${dt["MODERNA ORDERS"]}`);
    obj.covovaxControlNumber = generateControlNumbers(`${dt.COMPANY}`, `${dt["EMPLOYEE ID"]}`, "covovax", `${dt["COVOVAX ORDERS"]}`);
    obj.modernaTotalCost = formatCurrency(modernaTotalCost(`${dt["MODERNA ORDERS"]}`), opts);
    obj.covovaxTotalCost = formatCurrency(covovaxTotalCost(`${dt["COVOVAX ORDERS"]}`), opts);
    obj.regLink = registrationURLMapping(`${dt.COMPANY}`);
    obj.vaccineTotalCost = formatCurrency(vaccineTotalCost(modernaTotalCost(`${dt["MODERNA ORDERS"]}`), covovaxTotalCost(`${dt["COVOVAX ORDERS"]}`)), opts);

    console.log("=================================")
    main(smtpCred, JSON.stringify(obj)).catch(console.error);
  })
});

async function verifySMTP(smtp_c){

  let responseMessage;
  let host = smtp_c.host;
  let username = smtp_c.username;
  let password = smtp_c.password;
  let port = smtp_c.port;

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
      // console.log(error.message);

      responseMessage = error.message;
    } else {
      // console.log("Server is ready to take our messages");
      // console.log(success);
      responseMessage = "Server is ready to take our messages";
    }
  })

  return responseMessage;
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
  });

  console.log("Message sent: %s", info.messageId);
  // // Message sent: <b658f8ca-6296-ccf4-8306-87d57a0b4321@example.com>
}


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














// Handle creating/removing shortcuts on Windows when installing/uninstalling.
if (require('electron-squirrel-startup')) { // eslint-disable-line global-require
  app.quit();
}

const createWindow = () => {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  });

  // and load the index.html of the app.
  mainWindow.loadFile(path.join(__dirname, 'index.html'));

  // Open the DevTools.
  // mainWindow.webContents.openDevTools();

  ipcMain.on("mainwindow-loaded", function () {
    let resultSMTP = getSMTPTable();
    resultSMTP.then(function(rows){
        try {
          mainWindow.webContents.send("mainwindow-hidden-smtp", rows)
        } catch (e) {
          console.log(e);
        }
      })
  })

  ipcMain.on('verify-smtp-connection', (event, smtpCred) => {
    let responseMessage;
    let host = smtpCred.host;
    let username = smtpCred.username;
    let password = smtpCred.password;
    let port = smtpCred.port;
  
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
        // console.log(error.message);
        responseMessage = error.message;
      } else {
        // console.log("Server is ready to take our messages");
        // console.log(success);
        responseMessage = "Server is ready to take our messages";
      }
      mainWindow.webContents.send("verify-smtp-connection-response", responseMessage)
    })
  
  });

  // Create Menus.
  mainWindow.webContents.openDevTools()
  var menu = Menu.buildFromTemplate([
    {
      label: 'Menu',
      submenu: [
        { label: 'SMTP Settings',
            click(){
              createSMTPWindow(mainWindow)
            }
        },
        { label: 'Exit',
            click() {
              app.quit();
            }  
        }
      ]
    }
  ])
  
  Menu.setApplicationMenu(menu);
};

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', createWindow);

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and import them here.
