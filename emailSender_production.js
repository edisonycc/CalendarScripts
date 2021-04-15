const RECIPIENT_SETO = "ccye@hku.hk";
const EMAIL_SUBJECT = "InnoWing Equipment Booking Request";
//const SUPERVISER_EMAIL = "edison.ccye@gmail.com";

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
    var sheet = SpreadsheetApp.getActiveSheet();

    // If the question title is a label, it can be accessed as an object field.
    // If it has spaces or other characters, it can be accessed as a dictionary.

    let timestamp = responses.Timestamp[0];
    sheet.getRange('A10').setValue(timestamp);
    let equipment = responses.Equipment[0].toLowerCase();
    sheet.getRange('B10').setValue(equipment);
    let applicantName = responses['Member name'][0].trim();
    sheet.getRange('C10').setValue(applicantName);
    let applicantEmail = responses['Contact email (Please use HKU email)'][0].trim();
    sheet.getRange('D10').setValue(applicantEmail);
    let applicantPhone = responses['Contact phone number'][0].trim();
    sheet.getRange('E10').setValue(applicantPhone);
    let requestDate = responses['The requested date'][0].trim();
    sheet.getRange('F10').setValue(requestDate);
    let requestTime = responses['The requested time'][0].trim();
    sheet.getRange('G10').setValue(requestTime);
    let projectType = responses['Project type'][0].trim();
    sheet.getRange('H10').setValue(projectType);
    let supervisor = responses['Academic supervisor (please input "No" if the project has no academic supervisor)'][0].trim();
    sheet.getRange('I10').setValue(supervisor);
    let superVisorEmail = responses['Supervisor\'s contact email (please input "No" if the project has no academic supervisor)'][0].trim();
    sheet.getRange('J10').setValue(superVisorEmail);
    if (superVisorEmail.toLowerCase() == "no")
        superVisorEmail = "";
    let uid = responses['HKU ID (Student ID / Staff ID)'][0].trim();
    sheet.getRange('K10').setValue(uid);
    let upload = responses['Please upload the CAD file '][0].trim();
    sheet.getRange('L10').setValue(upload);

        let emailBody = `Dear Sir / Madam,

I want to apply equipment for my project part fabrication.
The equipment is ${equipment}.
${upload ? "Cad file: " + upload : 'I do not provide sketch drawing'}.
I want to use the equipment on ${requestDate}.
Booking session: ${requestTime}

My project type is ${projectType}.
${supervisor.toLowerCase() != "no" ? "My project supervisor name is " + supervisor : "The project doesn't have a supervisor." }.
${superVisorEmail != "" ? "My project supervisor email is " + superVisorEmail : ""}.

My name is ${applicantName}.
My HKUID is ${uid}.
My email is ${applicantEmail}
My contact phone number is ${applicantPhone}

${applicantName}
${timestamp}
`;



    // If the equipment is waterjet machine, send an email to the recipient.
    let status = '';
    if (equipment.includes("waterjet")) {
        MailApp.sendEmail({
            to: RECIPIENT_SETO + ',' + applicantEmail,
            subject: EMAIL_SUBJECT,
            body: emailBody,
            cc: superVisorEmail
        });
        //status = 'Sent';
    }
    else {
        status = 'No need to send email';
    }
    //sheet.getRange('M10').setValue(status);



    // // Append the status on the spreadsheet to the responses' row.
    // var sheet = SpreadsheetApp.getActiveSheet();
    // var row = sheet.getActiveRange().getRow();
    // var column = e.values.length + 1;
    // sheet.getRange(row, column).setValue(emailBody);
    //
    // Logger.log("status=" + status + "; responses=" + JSON.stringify(responses));
}

