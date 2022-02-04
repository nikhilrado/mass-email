const TO_FIELD = "B2"
const SUBJECT_FIELD = "B3"
const CC_FIELD = "B4"
const BCC_FIELD = "B5"
const BODY_FIELD = "A6"
const DATA_RANGE = "G1:K7"

const SEND_MAIL = false
var sheet = SpreadsheetApp.getActiveSheet();

function myFunction() {
    var emailTemplateTo = sheet.getRange(TO_FIELD).getValue();
    var emailTemplateSubject = sheet.getRange(SUBJECT_FIELD).getValue();
    var emailTemplateCC = sheet.getRange(CC_FIELD).getValue();
    var emailTemplateBCC = sheet.getRange(BCC_FIELD).getValue();
    var emailTemplateBodyRichValue = sheet.getRange(BODY_FIELD).getRichTextValue();
    var inputData = sheet.getRange(DATA_RANGE).getValues();
    var inputDataHeaders = inputData[0];
    inputData.shift() //removes headers from inputData
    Logger.log(inputDataHeaders)
    Logger.log(inputData)
    Logger.log(emailTemplateBodyRichValue)
    //TODO: remove blank cells from data

    dataRangeSplit = extractFirstRow(DATA_RANGE)
    firstRow = parseInt(dataRangeSplit[2]) + 1
    Logger.log(inputData.length)
    for (var i = 0; i < inputData.length; i++){

      var dict = {};
      for (var j = 0; j < inputDataHeaders.length; j++ ){
        dict[inputDataHeaders[j]] = inputData[i][j];

      }
      var emailArgs = {};
      emailArgs["to"] = parseStringTemplate(emailTemplateTo, dict)
      emailArgs["subject"] = parseStringTemplate(emailTemplateSubject, dict)
      emailArgs["cc"] = parseStringTemplate(emailTemplateCC, dict)
      emailArgs["bcc"] = parseStringTemplate(emailTemplateBCC, dict)
      emailArgs["htmlBody"] = parseStringTemplate(richTextToHTML(emailTemplateBodyRichValue),dict)
      Logger.log(emailArgs["to"])
      Logger.log(emailArgs["htmlBody"]);
      
      if(sendEmail({to: emailArgs["to"], name: emailArgs["name"], subject: emailArgs["subject"], htmlBody: emailArgs["htmlBody"], cc: emailArgs["cc"],bcc: emailArgs["bcc"]})){
        var color = "#659160"
      } else {
        var color = "#f26b61"
      }
      sheet.getRange(dataRangeSplit[1]+(firstRow+i)).setBackgroundColor(color);

      
    }
    Logger.log(extractFirstRow(DATA_RANGE))
}

//splits range up into array with length 5. Index 1 contains full a1 cell address. Indexes 1, 2 give column letter. Inexes 3, 4 give row number.
//https://regex101.com/ is really helpful for regex validation
function extractFirstRow(range) {
  var regex = /([A-Za-z]*)(.\d*):([A-Za-z]*)(.\d*)/;
  var arr = regex.exec(range);
  return arr; 
}

function parseCSV (csv, dict){
  csvList = csv.split(",");
  for (i = 0; i < csvList.length; i++){
    i = i;
  }
}

function sendEmail(args){
  var test = true;
  if (SEND_MAIL){
    MailApp.sendEmail({to: args["to"],name: args["name"],subject: args["subject"],htmlBody: args["htmlBody"], cc: args["cc"], bcc: args["bcc"]});
    test = false
  }
  if(test){test = "TEST ";}else{test = "";}
  Logger.log(test+"EMAIL SENT\nname: " + args["name"] + "\nto: " + args["to"])
  return !test;
}

function richTextToHTML(richTextObject){
    //documentation https://developers.google.com/apps-script/reference/spreadsheet/rich-text-value
    richTextList = richTextObject.getRuns();
    HTMLoutput = "";
    for (var i = 0; i < richTextList.length; i++){
      richTextRun = richTextList[i];
      textStyle = richTextRun.getTextStyle();
      text = richTextRun.getText();
      if(richTextRun.getLinkUrl()){text = "<a href='" + richTextRun.getLinkUrl() + "'>" + text + "</a>"}
      if(textStyle.isBold()){text = "<strong>" + text +"</strong>";}
      if(textStyle.isItalic()){text = "<i>" + text +"</i>";}
      if(textStyle.isStrikethrough()){text = "<strike>" + text +"</strike>";}
      if(textStyle.isUnderline()){text = "<u>" + text +"</u>";}
      text = text.replaceAll("\n", "<br>")
      
      if(textStyle.getForegroundColor() && textStyle.getForegroundColor() != "#000000"){
        //gets hex color, can't use "textStyle.getForegroundColor()" cause returns things like "ACCENT2" which don't work outside of the sheet
        hexColor = textStyle.getForegroundColorObject().asRgbColor().asHexString()
        text = "<span style='color:" + hexColor + "'>" + text +"</span>";
      }
      //TODO: remove styling if it is a link
      HTMLoutput += text
    }
    return HTMLoutput;
}

//modified function from pekaaw on stackoverflow: https://stackoverflow.com/a/59084440
function parseStringTemplate(str, obj) {
    //let parts = str.split(/\$\{(?!\d)[\wæøåÆØÅ]*\}/);
    let parts = str.split(/{(?!\d)[\wæøåÆØÅ]*\}/);
    let args = str.match(/[^{\}]+(?=})/g) || [];
    let parameters = args.map(argument => obj[argument] || (obj[argument] === undefined ? "" : obj[argument]));
    return String.raw({ raw: parts }, ...parameters);
}
