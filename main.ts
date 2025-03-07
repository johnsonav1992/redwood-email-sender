const emailColumnLetter = "A"; // Column for email addresses
const statusColumnLetter = "B"; // Column for email sending status

// Number of recipients per email batch
const BATCH_SIZE = parseInt(getProperty_("BATCH_SIZE") as string);
const TARGET_SEND_TO_EMAIL = getProperty_("TARGET_SEND_TO_EMAIL");
const EMAIL_SUBJECT = getProperty_("EMAIL_SUBJECT");

/**
 * Main sender function for the emails job
 */
function sendEmailsInBatches_() {
  if (!TARGET_SEND_TO_EMAIL || !EMAIL_SUBJECT || !BATCH_SIZE) {
    deleteTrigger_();

    throw new Error(
      "Missing required properties in script properties to run the email job." +
        "Please set these properties in the script properties before running: " +
        "TARGET_SEND_TO_EMAIL (the main email address to send to), EMAIL_SUBJECT (subject line), BATCH_SIZE (how many emails get sent per batch)"
    );
  }

  const remainingSendQuota = MailApp.getRemainingDailyQuota();
  const currentStoredIndex = getProperty_("currentIndex") as string;

  if (!remainingSendQuota) {
    console.log(
      "Send quota is: " + remainingSendQuota + ". Cannot send any more emails"
    );
    deleteTrigger_();
    return deleteProperty_("currentIndex");
  }

  console.log("Starting new batch...");
  console.log("Remaining send quota is: " + remainingSendQuota);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const allSheetEmailData = sheet
    .getRange(`${emailColumnLetter}:${statusColumnLetter}`)
    .getValues() as Array<EmailRow>;

  const emails = allSheetEmailData
    .map((row) => row[0])
    .filter((email) => !!email); // Keep only rows with non-empty email

  const totalEmails = emails.length - 1; // Subtract 1 to account for header row

  // Get the current index from the script properties to know what row to start sending from
  let currentEmailRowToStartOn = parseInt(currentStoredIndex) || 1; // Start from 1 initially to skip headers
  let emailBatch: EmailBatchEntry[] = [];
  let sentCount = 0;

  for (
    let i = currentEmailRowToStartOn;
    i <= totalEmails && sentCount < BATCH_SIZE;
    i++
  ) {
    const email = emails[i];
    const cellNum = i + 1;
    const status = allSheetEmailData[i][1];
    const isSent = status === "Sent" || status === "sent";

    if (!isSent && email) {
      emailBatch.push({ email, rowNum: cellNum });
      sentCount++;
    }
  }

  if (currentEmailRowToStartOn > totalEmails || !emailBatch.length)
    return exitJob_();

  try {
    const emailAddresses = emailBatch
      .map((item) => item.email.trim())
      .join(",");

    const allScriptProperties =
      PropertiesService.getScriptProperties().getProperties();

    const emailImages = Object.fromEntries(
      Object.entries(allScriptProperties)
        .filter(([key]) => key.endsWith("-img"))
        .map(([key, value]) => {
          return [key, getImg_(value)];
        })
    );

    if (Object.keys(emailImages).length === 0) {
      console.log(
        "No images found in script properties - make sure you have them added in if your html template requires images, using this pattern: \n" +
          "Property: imagename-img, Value: fileId from Google Drive. \n" +
          "Be sure your html contains references to these images in the format <img src='cid:imagename-img' />."
      );
    }

    const html = HtmlService.createHtmlOutputFromFile("template").getContent();

    GmailApp.sendEmail(TARGET_SEND_TO_EMAIL, EMAIL_SUBJECT, "body", {
      htmlBody: html,
      bcc: emailAddresses,
      replyTo: TARGET_SEND_TO_EMAIL,
      inlineImages: emailImages,
    });

    emailBatch.forEach((item) => {
      sheet.getRange(`${statusColumnLetter}${item.rowNum}`).setValue("Sent");
    });

    console.log("Sent email batch to: " + emailAddresses);

    if (currentEmailRowToStartOn === totalEmails) return exitJob_();

    currentEmailRowToStartOn += emailBatch.length;
  } catch (e) {
    console.error("Error sending email batch to: " + emails, e);
  }

  setProperty_("currentIndex", currentEmailRowToStartOn.toString());
}

/**
 * Sets up the email job and runs it from the apps script dropdown.
 * This function initializes the necessary trigger for the email job.
 */
function setUpEmailJobAndRun() {
  setupTrigger_();
}

/**
 * Checks and logs the remaining daily email send quota.
 */
function checkQuota() {
  const remainingSendQuota = MailApp.getRemainingDailyQuota();
  console.log("Remaining quota: " + remainingSendQuota);
}
