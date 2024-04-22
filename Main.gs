/**
 * Adds menu in Sheets
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Send emails')
    .addItem('Create campaign', 'openModal')
    .addToUi();
}

/**
 * Display setup modal
 */
function openModal() {
  // Get ss data
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetData = sheet.getDataRange().getValues();
  const sheetDataHeaders = sheetData[0];
  let recipientsCol = sheetDataHeaders[0];
  let htmlColumns = "<option value=" + recipientsCol + " >" + recipientsCol + "</option > ";
  sheetDataHeaders.slice(1).forEach(function (col) {
    htmlColumns += '<option value=' + col + '>' + col + "</option>";
  })
  // Get user's draft
  let drafts = GmailApp.getDrafts();
  let htmlDrafts = "";
  drafts.forEach(function (draft) {
    htmlDrafts += '<option value=' + draft.getId() + '>' + draft.getMessage().getSubject() + "</option>"
  })
  let modal = HtmlService.createTemplateFromFile("modal");
  // Fill modal's info
  modal.drafts = htmlDrafts;
  modal.htmlColumns = htmlColumns;
  modal.nbOfRows = sheetData.map(x => x[0]).filter(n => n).length - 1;
  modal = modal.evaluate();
  modal.setHeight(400).setWidth(500);
  SpreadsheetApp.getUi().showModalDialog(modal, "Email campaign setup");
}


/**
 * to include css and js code in the HTML files
 * @param {string} filename
 * @returns {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get data from spreadsheet and build a map to ease process
 */
function getSheetsInfo(campaignSelectedInfo) {

  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetData = sheet.getDataRange().getValues();
  const sheetDataHeaders = sheetData[0];
  let sheetsDataObj = {};
  console.log("campaignSelectedInfo", campaignSelectedInfo)
  sheetsDataObj.draftId = campaignSelectedInfo.draftId;
  sheetsDataObj.containsCc = (campaignSelectedInfo.ccCol != null) ? true : false;
  sheetsDataObj.containsBcc = (campaignSelectedInfo.bccCol != null) ? true : false;

  for (let i = 0; i < sheetDataHeaders.length; i++) {
    // Get recipients data from sheets
    if (sheetDataHeaders[i] == "Recipients") {
      sheetsDataObj["recipientList"] = sheetData.map(x => x[i]).slice(1);
    } else if (campaignSelectedInfo.ccCol && sheetDataHeaders[i] == campaignSelectedInfo.ccCol) {
      sheetsDataObj["ccList"] = sheetData.map(x => x[i]).slice(1);
    } else if (campaignSelectedInfo.bccCol && sheetDataHeaders[i] == campaignSelectedInfo.ccCol) {
      sheetsDataObj["bccList"] = sheetData.map(x => x[i]).slice(1);
    } else {
      sheetsDataObj[sheetDataHeaders[i]] = sheetData.map(x => x[i]).slice(1);
    }
  }
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Status");
  console.log(sheetsDataObj);
  return sheetsDataObj;
}

/**
 * Send email with given info
 * @param {array} emailInfo
 */
function sendEmail(emailInfo, progress) {
  // Get selected draft
  const draft = GmailApp.getDraft(emailInfo.draftId);
  let draftMessage = draft.getMessage().getPlainBody();
  let draftSubject = draft.getMessage().getSubject();

  // Check keys and extract tag data
  const emailInfoKeys = Object.keys(emailInfo);
  const fixedCampaignData = ["bccList", "ccList", "draftId", "recipient", "containsCc", "containsBcc"];
  const tags = emailInfoKeys.filter(item => !fixedCampaignData.includes(item));
  console.log(tags)

  // Replace tags by values in subject and email body
  for (let i = 0; i < tags.length; i++) {
    console.log(String(draftMessage).includes("{{" + tags[i] + "}}"))
    draftMessage = draftMessage.replaceAll("{{" + tags[i] + "}}", emailInfo[tags[i]]);
    draftSubject = draftSubject.replaceAll("{{" + tags[i] + "}}", emailInfo[tags[i]]);
  }

  // Prepare email obj
  let email = {
    "subject": draftSubject,
    "htmlBody": draftMessage,
    "to": emailInfo.recipient
  }
  // Check if cc selected by user
  if (emailInfo.containsCc) {
    email["cc"] = emailInfo.cc;
  }
  if (emailInfo.containsBcc) {
    email["bcc"] = emailInfo.bcc;
  }
  const sheet = SpreadsheetApp.getActiveSheet();

  // Send email ad add status/timestamp in sheets
  try {
    const template = HtmlService.createTemplateFromFile("content");
    template.name = "Lidia";
    template.email = emailInfo.recipient;
    template.emailBody = draftMessage;
    const message = template.evaluate().getContent();

    GmailApp.sendEmail(
      emailInfo.recipient, draftSubject, '',
      { htmlBody: message }
    );

    // MailApp.sendEmail(email);
    sheet.getRange(emailInfo.rowIndexInSheets, sheet.getLastColumn()).setValue("SENT").setNote(new Date());
  } catch (e) {
    sheet.getRange(emailInfo.rowIndexInSheets, sheet.getLastColumn()).setValue("FAIL").setNote(new Date() + " " + e.message);
  }
  return progress;
}


// handles the get request to the server
function doGet(e) {
  var method = e.parameter['method'];
  switch (method) {
    case 'track':
      var email = e.parameter['email'];
      updateEmailStatus(email);
    default:
      break;
  }
}

function updateEmailStatus(emailToTrack) {
  // get the active spreadsheet and data
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetData = sheet.getDataRange().getValues();
  const sheetDataHeaders = sheetData[0];

  // iterate through the data, starting at index 1
  for (var i = 1; i < sheetData.length; i++) {
    let email = sheetData[i][sheetDataHeaders.indexOf("Recipients")];

    if (emailToTrack === email) {
      // update the value in sheet
      sheet.getRange(i + 1, sheetDataHeaders.indexOf('Status') + 1).setValue('OPENED');
      break;
    }
  }
}