// DB settings
var Users;
var PaymentRecords;
var SubsidyRecords;
var Applications;

Tamotsu.onInitialized(function() {
  Users = Tamotsu.Table.define({ sheetName: '會員資料' }, {
    // instanceProperties
    memberStatus: function() {
      var memberStatus = "";
      if (JSON.parse(this['會員資格']).length===0) {
        memberStatus = "新會員";
      } else if (this['永久會員']) {
        memberStatus = "永久會員";
      } else if (this['會員種類']==='一般會員') {
        var now = new Date();
        if (JSON.parse(this['會員資格']).indexOf(now.getFullYear()) !== -1) {
          memberStatus = "常年會員";
        } else {
          // 學生會員
          if (now.getFullYear() - this['生日'].getFullYear() < 20) {
            memberStatus = "常年會員";
          } else {
            memberStatus = "常年會員(待繳費)";
          }
        }
      }
      return memberStatus;
    },
    memberType: function() {
      if (this['會員種類']==='一般會員') {
        var now = new Date();
        if (now.getFullYear() - this['生日'].getFullYear() < 20) {
          return "學生會員";
        }
      }
      return this['會員種類']
    },
  });
  PaymentRecords = Tamotsu.Table.define({ sheetName: '繳費記錄' });
  SubsidyRecords = Tamotsu.Table.define({ sheetName: '補助記錄' });
  Applications = Tamotsu.Table.define({ sheetName: '入會申請' }, {
    // instanceProperties
    memberStatus: function() {
      return '新入會'
    },
    memberType: function() {
      if (this['會員種類']==='一般會員') {
        var now = new Date();
        if (now.getFullYear() - this['生日'].getFullYear() < 20) {
          return "學生會員";
        }
      }
      return this['會員種類']
    },
  });
  Users.first();
  PaymentRecords.first();
  SubsidyRecords.first();
  Applications.first();
  
});



function onOpen(e) {
  Logger.log("onOpen");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("工作人員設定").activate();
}

function onEdit(e) {
  Logger.log("onEdit");
  Logger.log(e.source.getActiveSheet().getName());
  Logger.log(e.source.getActiveRange().getRow());
  Logger.log(e.source.getActiveRange().getColumn());
//  if (e.source.getActiveSheet().getName() == "校友服務" && e.source.getActiveRange().getRow() == 1 && e.source.getActiveRange().getColumn() == 2) {
//    setService();
//  }
//  else if (e.source.getActiveSheet().getName() == "校友服務" && e.source.getActiveRange().getRow() == 4 && e.source.getActiveRange().getColumn() == 2) {
//    showService();
//  }
  if (e.source.getActiveSheet().getName() == "校友服務" && e.source.getActiveRange().getRow() == 4 && e.source.getActiveRange().getColumn() == 2) {
    showService();
  }
}

function setService() {
  Logger.log("setService");

  mainSheet = SpreadsheetApp.getActiveSheet();
  mainSheet.hideRows(7, 30);
  id = mainSheet.getRange("B1").getValue();
  Logger.log(id);
  var memberStatus = lookupMembership(id);
  Logger.log(memberStatus);
  mainSheet.getRange(2, 2, 1, 1).setValue(memberStatus);
  mainSheet.getRange("B4").clear();
  mainSheet.getRange("5:5").clear();
  switch (memberStatus) {
    case "非會員":
      mainSheet.getRange("B5").setValue("入會");
      break;
    case "新入會":
      mainSheet.getRange("B5").setValue("審查");
      break;
    case "常年會員(待繳費)":
    case "新會員":
      mainSheet.getRange("B5").setValue("繳費");
      break;
    case "常年會員":
    case "永久會員":
    case "學生會員":
      mainSheet.getRange("B5").setValue("補助");
      mainSheet.getRange("C5").setValue("紀念品");
      break;
  }
}

function lookupMembership(id) {
  Logger.log("lookupMembership");
  Tamotsu.initialize(SpreadsheetApp.openById("1IynSzWYhWl93xjSqJyuAeldFYV25MXx6AwVDStLSH3c"));
  Logger.log(id);
  var dstSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dstSheet = dstSpreadsheet.getSheetByName("校友服務");
  var name = "無資料";
  var sex = "無資料";
  var memberStatus = "非會員";
  var memberType = "無資料";
  var found = -1;
  var user = Users.where({ '身分證字號': id }).all();
  
  if (user.length > 1) {
    // duplicate records
  } else if (user.length === 0) {
    // not found in Users
    
    var applicant = Applications.where({ '身分證字號': id }).all();
    if(applicant.length > 1) {
      // duplicate records
    } else if (applicant.length === 1) {
      applicant = applicant[0];
      dstSheet.getRange(2, 4, 1, 1).setValue(applicant.memberStatus());
      dstSheet.getRange(2, 5, 1, 1).setValue(applicant.memberType());
      dstSheet.getRange(2, 6, 1, 1).setValue(applicant['姓名']);
      dstSheet.getRange(2, 7, 1, 1).setValue(applicant['性別']);
      for (var j = 2; j < Object.keys(applicant).length; j++) {
        dstSheet.getRange(12, j+2, 1, 1).setValue(applicant[Object.keys(applicant)[j]]);
      }
      
      return applicant.memberStatus();
    }
  } else {
    user = user[0];
    
    dstSheet.getRange(2, 4, 1, 1).setValue(user.memberStatus());
    dstSheet.getRange(2, 5, 1, 1).setValue(user.memberType());
    dstSheet.getRange(2, 6, 1, 1).setValue(user['姓名']);
    dstSheet.getRange(2, 7, 1, 1).setValue(user['性別']);
    return user.memberStatus();
  }
  
//  dstSheet.getRange("A2").setValue(-1); // ???
  
}

function showService() {
  Logger.log("showService");
  
  mainSheet = SpreadsheetApp.getActiveSheet();
  mainSheet.hideRows(7, 30);
  srv = mainSheet.getRange("B4").getValue();
  switch (srv) {
    case "補助":
      mainSheet.showRows(7);
      break;
    case "繳費":
      mainSheet.showRows(9);
      break;
    case "審查":
      mainSheet.showRows(11, 2);
      break;
    case "紀念品":
      mainSheet.showRows(14, 10);
      break;
    case "入會":
      mainSheet.showRows(25, 1);
      break;
  }
}

function submitButton() {
  Logger.log("submitButton");
  Tamotsu.initialize(SpreadsheetApp.openById("1IynSzWYhWl93xjSqJyuAeldFYV25MXx6AwVDStLSH3c"));
  mainSheet = SpreadsheetApp.getActiveSheet();
  srv = mainSheet.getRange("B4").getValue();
  switch (srv) {
    case "補助":
      submitSubsidy();
      break;
    case "審查":
      submitInspect();
      break;
    case "紀念品":
      submitSouvenir();
      break;
  }
  mainSheet.hideRows(7, 30);
  mainSheet.getRange("B1").clear();
}

function submitSubsidy() {
  Logger.log("submitSubsidy");
  mainSheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('Confirmation of subsidy', "補助" + mainSheet.getRange("F2").getValue() + "$" + mainSheet.getRange("B7").getValue() + "?", ui.ButtonSet.OK_CANCEL);
  if (result != ui.Button.OK)
    return;
  
  var newData = [];
  var d = new Date();
  infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("工作人員設定");
  
  SubsidyRecords.create({
    '時間': d,
    '工作人員': Session.getActiveUser().getEmail(),
    '會員身分證字號': mainSheet.getRange("B1").getValue(),
    '金額': mainSheet.getRange("B7").getValue(),
    '類別': infoSheet.getRange("B1").getValue(),
    '活動時間': infoSheet.getRange("B2").getValue(),
    '活動地點': infoSheet.getRange("B3").getValue(),
  });
}

function submitSouvenir() {
  Logger.log("submitSouvenir");
  mainSheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
}

function submitInspect() {
  Logger.log("submitInspect");
  mainSheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('Confirmation of inspect', mainSheet.getRange("B11").getValue() + mainSheet.getRange("F2").getValue() + "的入會申請?", ui.ButtonSet.OK_CANCEL);
  if (result != ui.Button.OK)
    return;
  
  
  var d = new Date();
  var id = mainSheet.getRange("B1").getValue();
  var applicant = Applications.where({ '身分證字號': id }).first();
  if (mainSheet.getRange("B11").getValue() == "批准")
    applicant['審核狀態']="approved";
  else
    applicant['審核狀態']="rejected";
  
  applicant['modifiedAt']=d;
  applicant['modifiedBy']=Session.getActiveUser().getEmail();
  applicant.save();
  
  Users.create({
    '姓名': applicant['姓名'],
    '性別': applicant['性別'],
    '生日': applicant['生日'],
    '身分證字號': applicant['身分證字號'],
    '電子郵件': applicant['電子郵件'],
    '會員種類': applicant['會員種類'],
    '永久會員': false,
    '會員資格': "[]",
    '會員證狀態': "",
    '會員證補發': 0,
    'createdAt': d,
    'modifiedAt': d,
    '戶籍地址': applicant['戶籍地址'],
    '電話號碼': applicant['電話號碼'],
    '最高學歷': applicant['最高學歷'],
    '職業': applicant['職業'],
    '服務單位': applicant['服務單位'],
    '臉書帳號': applicant['臉書帳號'],
    '畢業部別': applicant['畢業部別'],
    '畢業屆別': applicant['畢業屆別'],
    '畢業年份': applicant['畢業年份'],
    '畢業學號': applicant['畢業學號'],
  });
  
//  var newData = [];
//  for (var i = 0; i < 6; i++) {
//    newData[i] = applySheet.getRange(rowNum, i+2, 1, 1).getValue();
//  }
//  newData[6] = false;
//  newData[7] = "[]";
//  newData[8] = "";
//  newData[9] = 0;
//  newData[10] = d.toLocaleTimeString();
//  newData[11] = d.toLocaleTimeString();
//  for (var i = 12; i < 22; i++) {
//    newData[i] = applySheet.getRange(rowNum, i-4, 1, 1).getValue();
//  }
//  SpreadsheetApp.openById("1uMLTcENFjvLSFkcjCr0bb9Wpu3str_hn9gO5V_4gbtY").getSheetByName("會員資料").appendRow(newData);
}