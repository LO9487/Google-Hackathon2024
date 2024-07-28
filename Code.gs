function formatReport() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let headers = sheet.getRange('A1:I1');
  let table = sheet.getDataRange();

  // Formatting headers
  headers.setFontWeight('bold');
  headers.setFontColor('white');
  headers.setBackground('#52489C');

  // Formatting table
  table.setFontFamily('Roboto');
  table.setHorizontalAlignment('center');
  table.setBorder(true, true, true, true, false, true, '#52489C', SpreadsheetApp.BorderStyle.SOLID);
}

function sendFollowUpEmails() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  const data = sheet.getDataRange().getValues();
  const emailCusColumn = 3; // Column D (0-based index 3)
  const emailRepColumn = 6; // Column G (0-based index 6)
  const followUpColumn = 7; // Column H (0-based index 7)
  const followUpStatus = 8; // Column I (0-based index 8)
  const salesRepColumn = 5; // Column F (0-based index 5)
  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    var row = data[i];
    var followUpDate = new Date(row[followUpColumn]);

    Logger.log('Processing row: ' + (i + 1));
    Logger.log('Follow-up Date: ' + followUpDate);

    if (followUpDate >= today && row[followUpStatus] != 'Followed-up') {
      var emailAddressCus = row[emailCusColumn];
      var emailAddressRep = row[emailRepColumn];
      var salesRep = row[salesRepColumn];
      var subject = `Follow-Up Reminder for Lead: ${row[1]}`;
      var message = `Dear ${salesRep},\n\nPlease follow up with ${row[1]} at ${emailAddressCus} before ${followUpDate}.\n\nBest Regards,\nSales Team`;

      try {
        Logger.log('Sending email to: ' + emailAddressRep);
        GmailApp.sendEmail(emailAddressRep, subject, message);
        sheet.getRange(i + 1, followUpStatus + 1).setValue('Followed-up'); // Update follow-up status
      } catch (error) {
        Logger.log('Error sending email to: ' + emailAddressRep + ' - ' + error.message);
      }
    }
  }
}

// function doGet(e) {
//   return HtmlService.createHtmlOutputFromFile('AddLeadDialog');
// }


function showAddLeadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddLeadDialog')
    .setWidth(400)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Lead');
}

function addLead(name, emailCus, phone, salesRep, emailRep, followUpDate) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  const lastRow = sheet.getLastRow();
  const newLeadId = lastRow;

  sheet.appendRow([newLeadId, name, phone, emailCus, 'new', salesRep, emailRep, followUpDate, 'Pending']);

  const subject = `New Lead Assigned: ${name}`;
  const message = `Dear ${salesRep},\n\nA new lead has been assigned to you.\n\nName: ${name}\nPhone: ${phone}\nEmail: ${emailCus}\nDue Date: ${followUpDate}\nPlease follow up as soon as possible.\n\nBest Regards,\nSales Team`;

  try {
    GmailApp.sendEmail(emailRep, subject, message);
  } catch (error) {
    Logger.log('Error sending email to: ' + emailRep + ' - ' + error.message);
  }
}

function generateSummaryReport() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  const data = sheet.getDataRange().getValues();
  const salesRepColumn = 5; // Column F (0-based index 5)
  const statusColumn = 8; // Column I (0-based index 8)
  let report = {};
  
  for (let i = 1; i < data.length; i++) {
    let salesRep = data[i][salesRepColumn];
    let status = data[i][statusColumn];
    if (!report[salesRep]) {
      report[salesRep] = { total: 0, pending: 0, followedUp: 0 };
    }
    report[salesRep].total++;
    if (status === 'Pending') {
      report[salesRep].pending++;
    } else if (status === 'Followed-up') {
      report[salesRep].followedUp++;
    }
  }

  let summary = 'Sales Representative Lead Summary Report:\n\n';
  for (let salesRep in report) {
    summary += `${salesRep}:\n`;
    summary += `  Total Leads: ${report[salesRep].total}\n`;
    summary += `  Pending Leads: ${report[salesRep].pending}\n`;
    summary += `  Followed-up Leads: ${report[salesRep].followedUp}\n\n`;
  }

  const managerEmail = 'lojkeng@gmail.com';
  const subject = 'Weekly Sales Lead Summary Report';
  try {
    GmailApp.sendEmail(managerEmail, subject, summary);
  } catch (error) {
    Logger.log('Error sending summary report to: ' + managerEmail + ' - ' + error.message);
  }
}

function updateStatusBasedOnDate() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const followUpColumn = 7; // Column H (0-based index 7)
  const followUpStatus = 8; // Column I (0-based index 8)
  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let followUpDate = new Date(row[followUpColumn]);

    Logger.log('Row ' + (i + 1) + ': Follow-Up Date = ' + followUpDate);
    Logger.log('Row ' + (i + 1) + ': Today = ' + today);
    Logger.log('Row ' + (i + 1) + ': Current Status = ' + row[followUpStatus]);

    if (followUpDate < today) {
      Logger.log('Row ' + (i + 1) + ': Follow-Up Date is before today');
      if (row[followUpStatus] !== 'Followed-up' && row[followUpStatus] !== 'Overdue') {
        Logger.log('Row ' + (i + 1) + ': Status is not Followed-up or Overdue');
        try {
          sheet.getRange(i + 1, followUpStatus + 1).setValue('Overdue'); // Update follow-up status
          Logger.log('Row ' + (i + 1) + ': Status updated to Overdue');
        } catch (error) {
          Logger.log('Error updating status for row ' + (i + 1) + ': ' + error.message);
        }
      }
    }
  }
}



function exportLeadsToPDF() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let pdfFile = DriveApp.createFile(sheet.getBlob().getAs('application/pdf')).setName('Leads_Report.pdf');
  
  Logger.log('PDF file created: ' + pdfFile.getUrl());
  
  // Get the email address of the user who is currently accessing the file
  const currentUserEmail = Session.getActiveUser().getEmail();
  const subject = 'Leads Report PDF';
  const message = 'Please find the attached leads report PDF.';
  
  // Send the email with the PDF attachment
  try {
    MailApp.sendEmail({
      to: currentUserEmail,
      subject: subject,
      body: message,
      attachments: [pdfFile]
    });
    Logger.log('Email sent successfully to: ' + currentUserEmail);
  } catch (error) {
    Logger.log('Error sending email to: ' + currentUserEmail + ' - ' + error.message);
  }
}

function showAddSalesRepDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddSalesRepDialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Sales Representative');
}

function showDeleteSalesRepDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DeleteSalesRepDialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Delete Sales Representative');
}

function addSalesRep(name, email) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesReps');
  sheet.appendRow([name, email]);
}

function deleteSalesRep(name) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesReps');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function getSalesReps() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesReps');
  const data = sheet.getDataRange().getValues();
  let salesReps = [];
  
  for (let i = 1; i < data.length; i++) {
    salesReps.push({ name: data[i][0], email: data[i][1] });
  }
  
  return salesReps;
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Formatting')
    .addItem('Format Report', 'formatReport')
    .addItem('Send Follow-Up Emails', 'sendFollowUpEmails')
    .addItem('Add New Customer', 'showAddLeadDialog')
    .addItem('Generate Summary Report', 'generateSummaryReport')
    .addItem('Update Status Based on Date', 'updateStatusBasedOnDate')
    .addItem('Export Leads to PDF', 'exportLeadsToPDF')
    .addItem('Add Sales Representative', 'showAddSalesRepDialog')
    .addItem('Delete Sales Representative', 'showDeleteSalesRepDialog')
    .addToUi();
}
