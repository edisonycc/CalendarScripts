const RECIPIENT_TechnicalManager = "ccye@hku.hk";
const EMAIL_SUBJECT = "InnoWing Equipment Booking Request";
//const SUPERVISER_EMAIL = "edison.ccye@gmail.com";
const EMAIL_TEMPLATE_DOC_URL = 'https://docs.google.com/document/d/1J_UwpYEO0WAf9Kr7cOhWliuPk0PA1gpVlxPOgcKcjNE/edit?usp=sharing';


/**
 * Installs a trigger on the Spreadsheet for when a Form response is submitted.
 */
function installTrigger() {
    ScriptApp.newTrigger('onFormSubmit')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onFormSubmit()
        .create();
}

/**
 * Sends a customized email for every response on a form.
 *
 * @param {Object} event - Form submit event
 */
function onFormSubmit(e) {
    var responses = e.namedValues;

    // If the question title is a label, it can be accessed as an object field.
    // If it has spaces or other characters, it can be accessed as a dictionary.

    let timestamp = responses.Timestamp[0];
    let equipment = responses.Equipment[0].toLowerCase();
    let applicantName = responses['Member name'][0].trim();
    let applicantEmail = responses['Contact email (Please use HKU email)'][0].trim();
    let applicantPhone = responses['Contact phone number'][0].trim();
    let requestDate = responses['The requested date'][0].trim();
    let requestTime = responses['The requested time'][0].trim();
    let projectType = responses['Project type'][0].trim();
    let supervisor = responses['Academic supervisor (please input "No" if the project has no academic supervisor)'][0].trim();
    let supervisorEmail = responses['Supervisor\'s contact email (please input "No" if the project has no academic supervisor)'][0].trim();
    if (supervisorEmail.toLowerCase() == "no" && !supervisorEmail.toLowerCase().includes('@'))
        supervisorEmail = "";
    let uid = responses['HKU ID (Student ID / Staff ID)'][0].trim();
    let upload = responses['Please upload the CAD file '][0].trim();


    // // Parse topics of interest into a list (since there are multiple items
    // // that are saved in the row as blob of text).
    // var topics = Object.keys(topicUrls).filter(function(topic) {
    //     // indexOf searches for the topic in topicsString and returns a non-negative
    //     // index if the topic is found, or it will return -1 if it's not found.
    //     return topicsString.indexOf(topic.toLowerCase()) != -1;
    // });

    // If the equipment is not 3d printer, send an email to the recipient.
    var status = '';
    if (equipment.includes("waterjet")) {
        MailApp.sendEmail({
            to: RECIPIENT_TechnicalManager + ',' + applicantEmail,
            //to: RECIPIENT,
            subject: EMAIL_SUBJECT,
            // htmlBody: createEmailBody(name, equipment, email, phone),
            htmlBody: createEmailBody(timestamp,
                equipment,
                applicantName,
//                applicantEmail,
                applicantPhone,
                requestDate,
                requestTime,
                projectType,
                supervisor,
                supervisorEmail,
                uid,
                upload
            ),
            cc: supervisorEmail
        });
        status = 'Sent';
    }
    else {
        status = 'No need to send email';
    }

    // Append the status on the spreadsheet to the responses' row.
    var sheet = SpreadsheetApp.getActiveSheet();
    var row = sheet.getActiveRange().getRow();
    var column = e.values.length + 1;
    sheet.getRange(row, column).setValue(status);

    Logger.log("status=" + status + "; responses=" + JSON.stringify(responses));
}

/**
 * Creates email body and includes the links based on topic.
 *
 * @param {string} recipient - The recipient's email address.
 * @param {string[]} topics - List of topics to include in the email body.
 * @return {string} - The email body as an HTML string.
 */
function createEmailBody(timestamp,
                         equipment,
                         applicantName,
//                         applicantEmail,
                         applicantPhone,
                         requestDate,
                         requestTime,
                         projectType,
                         supervisor,
                         supervisorEmail,
                         uid,
                         upload
) {
    // Make sure to update the emailTemplateDocId at the top.
    var docId = DocumentApp.openByUrl(EMAIL_TEMPLATE_DOC_URL).getId();
    var emailBody = docToHtml(docId);
    emailBody = emailBody.replace(/{{timestamp}}/g, timestamp);
    emailBody = emailBody.replace(/{{equipment}}/g, equipment);
    emailBody = emailBody.replace(/{{applicantName}}/g, applicantName);
    emailBody = emailBody.replace(/{{applicantPhone}}/g, applicantPhone);
    emailBody = emailBody.replace(/{{requestDate}}/g, requestDate);
    emailBody = emailBody.replace(/{{requestTime}}/g, requestTime);
    emailBody = emailBody.replace(/{{projectType}}/g, projectType);
    emailBody = emailBody.replace(/{{supervisor}}/g, supervisor);
    emailBody = emailBody.replace(/{{supervisorEmail}}/g, supervisorEmail);
    emailBody = emailBody.replace(/{{uid}}/g, uid);
    emailBody = emailBody.replace(/{{upload}}/g, upload);

    return emailBody;
}

/**
 * Downloads a Google Doc as an HTML string.
 *
 * @param {string} docId - The ID of a Google Doc to fetch content from.
 * @return {string} The Google Doc rendered as an HTML string.
 */
function docToHtml(docId) {

    // Downloads a Google Doc as an HTML string.
    var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" +
        docId + "&exportFormat=html";
    var param = {
        method: "get",
        headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true,
    };
    return UrlFetchApp.fetch(url, param).getContentText();
}