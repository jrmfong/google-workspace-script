/**
 * This script runs when a Google Form is submitted.
 *
 * It performs the following actions:
 * 1. If the respondent agrees to both terms in the form, the script generates a PDF copy of the Gift Aid Declaration.
 * 2. The PDF is stored in the Finance Team shared drive, emailed to the respondent for their records,
 *    and the Finance Team is also notified by email.
 * 3. If the respondent does not agree to the terms, the script does not generate the Gift Aid Declaration.
 *    Instead, it sends an email notification to both the Finance Team and the respondent.
 */

/**
 * Deployment
 * 1. It is highly recommended to use a dedicated service account to create the Google Form and deploy this script.
 * 2. Create the Google Form with reference to the sample Gift Aid declaration for multiple donations, available here:
 *    https://www.gov.uk/claim-gift-aid/gift-aid-declarations
 * 3. Create the template Google Docs file provided in this repository. Make sure the service account has read access to it.
 * 4. Make sure the target storage folder exists. Create it if necessary. Make sure the service account has read and write access to it.
 * 5. Update the script constants as required, then deploy the script by attaching it to the form.
 */

// Test data for manual execution of testOnFormSubmit
function testOnFormSubmit() {
  const e = {
    namedValues: {
      "Legal last name": ["Chan"],
      "Legal first name": ["Peter"],
      "Home address (line 1)": ["123 High Street"],
      "Your home address (line 2) (Optional)": [""],
      "Town or City": ["Peterborough"],
      "County (Optional)": [""],
      "Full post code": ["AB1 1AA"],
      "Email Address": ["a_testing_email_address"],
      "What is the coverage of this declaration?": [
        "All donations I have made in the past 4 years, today, and in the future"
      ],
      "I want to Gift Aid my donations to our org, including any donations covered by the option selected above.": ["Agree"],
      "I am a UK taxpayer and understand that if I pay less Income Tax and/or Capital Gains Tax than the amount of Gift Aid claimed on all my donations in that tax year, it is my responsibility to pay any difference.": ["Disagree"],
      "Timestamp": ["2026-04-06 21:00:00"]
    }
  };

  onFormSubmit(e);
}

function onFormSubmit(e) {
  if (!e || !e.namedValues) {
    throw new Error(
      "onFormSubmit must be run by a form-submit trigger or via testOnFormSubmit()."
    );
  }

// Important constants, such as the template file ID and the folder ID used to store Gift Aid PDFs
  const CONFIG = {
    templateDocId: "place_the_file_id_of_the_template",
    targetFolderId: "place_the_folder_id_of_the_target_folder",
    orgName: "Your charitable organisation name",
    financeEmail: "the_finance_team_email_address",
  };

  // Define constants from the Google Form responses
  // Make sure they all match the actual question titles in the Google Form
  const answers = e.namedValues;
  const lastName = getResponse_(answers, "Legal last name");
  const firstName = getResponse_(answers, "Legal first name");
  const addressLine1 = getResponse_(answers, "Home address (line 1)");
  const addressLine2 = getResponse_(
    answers,
    "Your home address (line 2) (Optional)"
  );
  const townCity = getResponse_(answers, "Town or City");
  const county = getResponse_(answers, "County (Optional)");
  const postcode = getResponse_(answers, "Full post code");
  const email = getResponse_(answers, "Email Address");
  const declarationScope = getResponse_(
    answers,
    "What is the coverage of this declaration?"
  );
  const giftAidConsent = getResponse_(
    answers,
    "I want to Gift Aid my donations to my org, including any donations covered by the option selected above."
  );
  const taxpayerConfirmation = getResponse_(
    answers,
    "I am a UK taxpayer and understand that if I pay less Income Tax and/or Capital Gains Tax than the amount of Gift Aid claimed on all my donations in that tax year, it is my responsibility to pay any difference."
  );
  const timestamp = getResponse_(answers, "Timestamp");

  const fullAddress = buildAddress_([
    addressLine1,
    addressLine2,
    townCity,
    county,
    postcode,
  ]);

  // Additional constant definition derived from the form responses
  const fullName = `${firstName} ${lastName}`.trim();
  const declarationDate = formatDate_(timestamp);
  const requiredAccepted =
    isAccepted_(giftAidConsent) && isAccepted_(taxpayerConfirmation);

   // If the respondent does not agree to all terms, send a rejection email to the respondent and the Finance Team
  if (!requiredAccepted) {
    Logger.log("Required declaration terms were not accepted. PDF not generated.");

    if (email) {
      Logger.log("Sending rejection email to the respondent...");
      sendGiftAidRejectedEmail_({
        to: email,
        fullName: fullName,
        orgName: CONFIG.orgName,
        declarationDate: declarationDate,
        giftAidConsent: giftAidConsent,
        taxpayerConfirmation: taxpayerConfirmation,
      });
    }

    if (CONFIG.financeEmail) {
      Logger.log("Sending rejection email alert to the Finance Team...");
      sendFinanceRejectionNotification_({
        to: CONFIG.financeEmail,
        fullName: fullName,
        orgName: CONFIG.orgName,
        declarationDate: declarationDate,
        email: email,
        fullAddress: fullAddress,
        declarationScope: declarationScope,
        giftAidConsent: giftAidConsent,
        taxpayerConfirmation: taxpayerConfirmation,
      });
    }
    
    // Exit early because the required declarations were not accepted
    return;
  }

  // Proceed with PDF generation because the respondent agreed to both terms
  Logger.log("Required declaration terms were accepted. PDF is now being generated.");

  // Define constants for file generation
  const templateFile = DriveApp.getFileById(CONFIG.templateDocId);
  const targetFolder = DriveApp.getFolderById(CONFIG.targetFolderId);
  const safeName = sanitizeFileName_(fullName || "Unknown Donor");
  const pdfBaseName = `Gift Aid Declaration - ${safeName} - ${declarationDate}`;

  // Generate the Gift Aid Declaration
  const copiedFile = templateFile.makeCopy(pdfBaseName, targetFolder);
  const doc = DocumentApp.openById(copiedFile.getId());
  const body = doc.getBody();

  // Replace template placeholders with Google Form responses
  body.replaceText("{{FIRST_NAME}}", firstName);
  body.replaceText("{{LAST_NAME}}", lastName);
  body.replaceText("{{FULL_ADDRESS}}", fullAddress);
  body.replaceText("{{POSTCODE}}", postcode);
  body.replaceText("{{EMAIL}}", email);
  body.replaceText("{{DECLARATION_SCOPE}}", declarationScope);
  body.replaceText("{{DECLARATION_DATE}}", declarationDate);
  body.replaceText("{{ORG_NAME}}", CONFIG.orgName);
  body.replaceText("{{GIFT_AID_CONSENT}}", giftAidConsent);
  body.replaceText("{{TAXPAYER_CONFIRM}}", taxpayerConfirmation);

  // Save the completed document
  doc.saveAndClose();

  // Convert the document to PDF
  const pdfBlob = copiedFile
    .getBlob()
    .getAs(MimeType.PDF)
    .setName(pdfBaseName + ".pdf");

  const pdfFile = targetFolder.createFile(pdfBlob);

  Logger.log("PDF created: " + pdfFile.getName());

  // Send the PDF to the respondent
  if (email) {
    Logger.log("Sending email with the PDF to the respondent...");
    sendGiftAidEmail_({
      to: email,
      fullName: fullName,
      orgName: CONFIG.orgName,
      declarationDate: declarationDate,
      pdfBlob: pdfBlob,
    });
  }

  // Send the PDF to the Finance Team
  if (CONFIG.financeEmail) {
    Logger.log("Sending email with the PDF to the Finance Team...");
    sendFinanceNotification_({
      to: CONFIG.financeEmail,
      fullName: fullName,
      orgName: CONFIG.orgName,
      declarationDate: declarationDate,
      pdfBlob: pdfBlob,
      pdfFileUrl: pdfFile.getUrl(),
    });
  }

  copiedFile.setTrashed(true);
}

// This marks the end of the main script and below are the function definitions
function sendGiftAidEmail_(data) {
  const subject = "Your Gift Aid Declaration Copy";

  const bodyText =
    `Dear ${data.fullName},\n\n` +
    `Thank you for submitting your Gift Aid declaration to ${data.orgName}.\n` +
    `Please find attached a copy of your declaration for your records.\n\n` +
    `Declaration date: ${data.declarationDate}\n\n` +
    `If you have any questions or need to update your details, please contact the org.\n\n` +
    `Kind regards,\n` +
    `${data.orgName}`;

  const bodyHtml = `
    <html>
      <body style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
        <p>Dear ${escapeHtml_(data.fullName)},</p>
        <p>
          Thank you for submitting your Gift Aid declaration to
          ${escapeHtml_(data.orgName)}.
        </p>
        <p>
          Please find attached a copy of your declaration for your records.
        </p>
        <p>
          Declaration date: ${escapeHtml_(data.declarationDate)}
        </p>
        <p>
          If you have any questions or need to update your details, please contact us via donation@kccc.org.uk.
        </p>
        <p>
          Kind regards,<br>
          ${escapeHtml_(data.orgName)}
        </p>
      </body>
    </html>
  `;

  GmailApp.sendEmail(data.to, subject, bodyText, {
    htmlBody: bodyHtml,
    attachments: [data.pdfBlob],
    name: data.orgName,
  });
}

function sendGiftAidRejectedEmail_(data) {
  const subject = "Gift Aid Declaration Could Not Be Completed";

  const bodyText =
    `Dear ${data.fullName},\n\n` +
    `Thank you for submitting your response to ${data.orgName}.\n\n` +
    `We were unable to complete your Gift Aid declaration because one or more required confirmations were not agreed to.\n\n` +
    `Gift Aid consent response: ${data.giftAidConsent}\n` +
    `Taxpayer confirmation response: ${data.taxpayerConfirmation}\n` +
    `Submission date: ${data.declarationDate}\n\n` +
    `If this was unintentional, please complete and submit the form again with the required confirmations selected.\n` +
    `If you have any questions, please contact donation@kccc.org.uk.\n\n` +
    `Kind regards,\n` +
    `${data.orgName}`;

  const bodyHtml = `
    <html>
      <body style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
        <p>Dear ${escapeHtml_(data.fullName)},</p>
        <p>
          Thank you for submitting your response to
          ${escapeHtml_(data.orgName)}.
        </p>
        <p>
          We were unable to complete your Gift Aid declaration because one or more
          required confirmations were not agreed to.
        </p>
        <p>
          <strong>Gift Aid consent response:</strong>
          ${escapeHtml_(data.giftAidConsent)}<br>
          <strong>Taxpayer confirmation response:</strong>
          ${escapeHtml_(data.taxpayerConfirmation)}<br>
          <strong>Submission date:</strong>
          ${escapeHtml_(data.declarationDate)}
        </p>
        <p>
          If this was unintentional, please complete and submit the form again with
          the required confirmations selected.
        </p>
        <p>
          If you have any questions, please contact us via donation@kccc.org.uk.
        </p>
        <p>
          Kind regards,<br>
          ${escapeHtml_(data.orgName)}
        </p>
      </body>
    </html>
  `;

  GmailApp.sendEmail(data.to, subject, bodyText, {
    htmlBody: bodyHtml,
    name: data.orgName,
  });
}

function sendFinanceNotification_(data) {
  const subject = "New Gift Aid Declaration Submitted";

  const bodyText =
    `A new Gift Aid declaration has been submitted.\n\n` +
    `Name: ${data.fullName}\n` +
    `Declaration date: ${data.declarationDate}\n` +
    `PDF file: ${data.pdfFileUrl}\n`;

  const bodyHtml = `
    <html>
      <body style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
        <p>A new Gift Aid declaration has been submitted.</p>
        <p><strong>Name:</strong> ${escapeHtml_(data.fullName)}</p>
        <p><strong>Declaration date:</strong> ${escapeHtml_(data.declarationDate)}</p>
        <p><strong>PDF file:</strong> <a href="${escapeHtml_(data.pdfFileUrl)}">Open PDF</a></p>
      </body>
    </html>
  `;

  GmailApp.sendEmail(data.to, subject, bodyText, {
    htmlBody: bodyHtml,
    attachments: [data.pdfBlob],
    name: data.orgName,
  });
}

function sendFinanceRejectionNotification_(data) {
  const subject = "Gift Aid Declaration Not Completed";

  const bodyText =
    `A Gift Aid declaration submission could not be completed.\n\n` +
    `Name: ${data.fullName}\n` +
    `Email: ${data.email}\n` +
    `Address: ${data.fullAddress}\n` +
    `Declaration scope: ${data.declarationScope}\n` +
    `Gift Aid consent response: ${data.giftAidConsent}\n` +
    `Taxpayer confirmation response: ${data.taxpayerConfirmation}\n` +
    `Submission date: ${data.declarationDate}\n`;

  const bodyHtml = `
    <html>
      <body style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
        <p>A Gift Aid declaration submission could not be completed.</p>
        <p><strong>Name:</strong> ${escapeHtml_(data.fullName)}</p>
        <p><strong>Email:</strong> ${escapeHtml_(data.email)}</p>
        <p><strong>Address:</strong> ${escapeHtml_(data.fullAddress)}</p>
        <p><strong>Declaration scope:</strong> ${escapeHtml_(data.declarationScope)}</p>
        <p><strong>Gift Aid consent response:</strong> ${escapeHtml_(data.giftAidConsent)}</p>
        <p><strong>Taxpayer confirmation response:</strong> ${escapeHtml_(data.taxpayerConfirmation)}</p>
        <p><strong>Submission date:</strong> ${escapeHtml_(data.declarationDate)}</p>
      </body>
    </html>
  `;

  GmailApp.sendEmail(data.to, subject, bodyText, {
    htmlBody: bodyHtml,
    name: data.orgName,
  });
}

function getResponse_(namedValues, questionTitle) {
  const value = namedValues[questionTitle];
  if (!value || value.length === 0) {
    return "";
  }
  return String(value[0]).trim();
}

function isAccepted_(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return (
    normalized === "yes" ||
    normalized === "i agree" ||
    normalized === "agreed" ||
    normalized === "agree" ||
    normalized === "true" ||
    normalized === "accepted"
  );
}

function buildAddress_(parts) {
  return parts
    .map(function (part) {
      return String(part || "").trim();
    })
    .filter(function (part) {
      return part !== "";
    })
    .join(", ");
}

function formatDate_(value) {
  if (!value) {
    return Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }

  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return String(value);
  }

  return Utilities.formatDate(
    parsed,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

function sanitizeFileName_(name) {
  return String(name || "")
    .replace(/[\\\/:*?"<>|#]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
