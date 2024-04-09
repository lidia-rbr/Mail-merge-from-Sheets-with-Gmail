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
  let recipientColumns = "<option value=" + recipientsCol + " selected >" + recipientsCol + "</option > ";
  let htmlColumns = "<option value=" + recipientsCol + " >" + recipientsCol + "</option > ";
  sheetDataHeaders.slice(1).forEach(function (col) {
    recipientColumns += '<option value=' + col + '>' + col + "</option>";
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
  modal.recipientColumns = recipientColumns;
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
  sheetsDataObj.draftId = campaignSelectedInfo.draftId;
  sheetsDataObj.containsCc = (campaignSelectedInfo.ccCol != "blank") ? true : false;
  sheetsDataObj.containsBcc = (campaignSelectedInfo.bccCol!= "blank") ? true : false;

  for (let i = 0; i < sheetDataHeaders.length; i++) {
    // Get recipients data from sheets
    if (sheetDataHeaders[i] == campaignSelectedInfo.recipientsCol) {
      sheetsDataObj["recipientList"] = sheetData.map(x => x[i]).slice(1);
    } else if (campaignSelectedInfo.ccCol && sheetDataHeaders[i] == campaignSelectedInfo.ccCol) {
      sheetsDataObj["ccList"] = sheetData.map(x => x[i]).slice(1);
    } else if (campaignSelectedInfo.bccCol && sheetDataHeaders[i] == campaignSelectedInfo.ccCol) {
      sheetsDataObj["bccList"] = sheetData.map(x => x[i]).slice(1);
    } else {
      sheetsDataObj[sheetDataHeaders[i]] = sheetData.map(x => x[i]).slice(1);
    }
  }
  sheet.getRange(1,sheet.getLastColumn()+1).setValue("Status");
  return sheetsDataObj;
}

/**
 * Send email with given info
 * @param {array} emailInfo
 */
function sendEmail(emailInfo, progress) {
  // Get selected draft
  const draft = GmailApp.getDraft(emailInfo.draftId);
  let draftMessage = draft.getMessage().getBody();
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

  console.log(email)
  
  // Send email ad add status/timestamp in sheets
  try {
    MailApp.sendEmail(email);
    sheet.getRange(emailInfo.rowIndexInSheets,sheet.getLastColumn()).setValue("SENT").setNote(new Date());
  } catch (e) {
    sheet.getRange(emailInfo.rowIndexInSheets,sheet.getLastColumn()).setValue("FAIL").setNote(new Date() + " " + e.message);
  }
  return progress;
}
