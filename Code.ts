type EmailBatch = { email: string; rowNum: number };
const emailColumn = "A"; // Column for email addresses
const statusColumn = "B"; // Column for email sending status
const BATCH_SIZE = 2; // Number of recipients per email batch

// Function to send emails in batches
const sendEmailsInBatches_ = () => {
  const remainingSendQuota = MailApp.getRemainingDailyQuota();
  const properties = PropertiesService.getScriptProperties();

  if (!remainingSendQuota) {
    Logger.log(
      "Send quota is: " + remainingSendQuota + ". Cannot send any more emails"
    );
    Logger.log("Deleting trigger...");
    deleteTriggers_();
    properties.deleteProperty("currentIndex");
    return;
  }

  Logger.log("Starting new batch...");
  Logger.log("Remaining send quota is: " + remainingSendQuota);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const emailRange = sheet.getRange(`${emailColumn}:${emailColumn}`);
  const emailData = emailRange.getValues();

  const filteredEmailData = emailData.filter((row) => row[0]); // Keep only rows with non-empty email
  const totalEmails = filteredEmailData.length - 1; // Subtract 1 to account for header row

  // Get the current index from the script properties to know what row to start sending from
  let currentIndex =
    parseInt(properties.getProperty("currentIndex") as string) || 1; // Start from 1 initially to skip headers

  if (currentIndex > totalEmails) {
    Logger.log("All emails sent. Deleting the trigger...");
    deleteTriggers_();
    return properties.deleteProperty("currentIndex");
  }

  let emailBatch: EmailBatch[] = [];
  let sentCount = 0;

  // Loop through emails starting from the currentIndex
  for (let i = currentIndex; i <= totalEmails && sentCount < BATCH_SIZE; i++) {
    const email = filteredEmailData[i][0];
    const cellNum = i + 1;
    const status = sheet.getRange(`${statusColumn}${cellNum}`).getValue();

    if (status !== "Sent" && email) {
      emailBatch.push({ email, rowNum: cellNum });
      sentCount++;
    }
  }

  if (emailBatch.length > 0) {
    try {
      MailApp.sendEmail({
        to: emailBatch.map((item) => item.email).join(","),
        subject,
        htmlBody: html,
        cc: ccEmails,
        bcc: bccEmails,
      });

      emailBatch.forEach((item) => {
        sheet.getRange(`${statusColumn}${item.rowNum}`).setValue("Sent");
      });

      Logger.log(
        "Sent email batch to: " + emailBatch.map((item) => item.email).join(",")
      );

      currentIndex += emailBatch.length;
    } catch (e) {
      Logger.log(
        "Error sending email batch to: " +
          emailBatch.map((item) => item.email).join(",")
      );
    }
  }

  properties.setProperty("currentIndex", currentIndex.toString());

  if (currentIndex > totalEmails) {
    Logger.log("All emails processed. Deleting trigger...");
    deleteTriggers_();
    properties.deleteProperty("currentIndex");
  }
};

const setupTrigger_ = () => {
  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    Logger.log("Creating new trigger...");
    ScriptApp.newTrigger("sendEmailsInBatches_")
      .timeBased()
      .everyMinutes(1) // Triggers can only be created with minimum 1-minute intervals
      .create();
  } else {
    Logger.log("Trigger already exists.");
  }
};

const deleteTriggers_ = () => {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "sendEmailsInBatches_") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
};

// This function is called to start sending emails
const startSendingEmails = () => {
  setupTrigger_();
};

const checkQuota = () => {
  var remainingSendQuota = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining quota: " + remainingSendQuota);
};

const mads =
  "https://drive.google.com/uc?export=view&id=1FB3Kj1juXC1NCkxfq6Hp7OT16hFliENm";
const redwood =
  "https://drive.google.com/uc?export=view&id=1FuKawK5cBa-rShmqrBZFP535_C-c93BU";
const phone =
  "https://drive.google.com/uc?export=view&id=1borma253c_oxk157f0c0Klu4K7RBR799";
const printer =
  "https://drive.google.com/uc?export=view&id=1zYGjC2WV8js24yxC84d38Q_QNprvgcP5";
const email =
  "https://drive.google.com/uc?export=view&id=1t_Kbfy3ew2ArzyRlibcgvFIjSTBcDCkj";
const network =
  "https://drive.google.com/uc?export=view&id=1SZnP57NQsEon5JnaW7864hwJRuOBksRP";

const subject = "TRS & Retirement Planning Seminar 10/29/2024";
const ccEmails = "";
const bccEmails = "";

const html = `<div style="display:none;">TRS & Retirement Planning Seminar 10/29/2024</div><div id=":8f" class="ii gt" jslog="20277; u014N:xr6bB; 1:WyIjdGhyZWFkLWY6MTgxMjkwMjg2MTY5NTYyNDA2OHxtc2ctZjoxODEyOTAyODYxNjk1NjI0MDY4Il0.; 4:WyIjbXNnLWY6MTgxMjkwMjg2MTY5NTYyNDA2OCIsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLDBd"><div id=":8e" class="a3s aiL "><div style="line-break:after-white-space"><div class="adM"><br>
</div><img alt="TRS Invite 10.29.2024.png" src="https://drive.google.com/uc?export=view&id=1hdI689A69Mu-EAOB2ybqkMVarKEewavC" data-image-whitelisted="" class="CToWUd a6T" data-bit="iit" tabindex="0"><div class="a6S" dir="ltr" style="opacity: 0.01; left: 1120px; top: 1567.5px;"><span data-is-tooltip-wrapper="true" class="a5q" jsaction="JIbuQc:.CLIENT"><button class="VYBDae-JX-I VYBDae-JX-I-ql-ay5-ays CgzRE" jscontroller="PIVayb" jsaction="click:h5M12e; clickmod:h5M12e;pointerdown:FEiYhc;pointerup:mF5Elf;pointerenter:EX0mI;pointerleave:vpvbp;pointercancel:xyn4sd;contextmenu:xexox;focus:h06R8; blur:zjh6rb;mlnRJb:fLiPzd;" data-idom-class="CgzRE" jsname="hRZeKc" aria-label="Download attachment TRS Invite 10.29.2024.png" data-tooltip-enabled="true" data-tooltip-id="tt-c146" data-tooltip-classes="AZPksf" id="" jslog="91252; u014N:cOuCgd,Kr2w4b,xr6bB; 4:WyIjbXNnLWY6MTgxMjkwMjg2MTY5NTYyNDA2OCJd; 43:WyJpbWFnZS9wbmciXQ.."><span class="OiePBf-zPjgPe VYBDae-JX-UHGRz"></span><span class="bHC-Q" jscontroller="LBaJxb" jsname="m9ZlFb" soy-skip="" ssk="6:RWVI5c"></span><span class="VYBDae-JX-ank-Rtc0Jf" jsname="S5tZuc" aria-hidden="true"><span class="bzc-ank" aria-hidden="true"><svg viewBox="0 -960 960 960" height="20" width="20" focusable="false" class=" aoH"><path d="M480-336L288-528l51-51L444-474V-816h72v342L621-579l51,51L480-336ZM263.72-192Q234-192 213-213.15T192-264v-72h72v72H696v-72h72v72q0,29.7-21.16,50.85T695.96-192H263.72Z"></path></svg></span></span><div class="VYBDae-JX-ano"></div></button><div class="ne2Ple-oshW8e-J9" id="tt-c146" role="tooltip" aria-hidden="true">Download</div></span><span data-is-tooltip-wrapper="true" class="a5q" jsaction="JIbuQc:.CLIENT"><button class="VYBDae-JX-I VYBDae-JX-I-ql-ay5-ays CgzRE" jscontroller="PIVayb" jsaction="click:h5M12e; clickmod:h5M12e;pointerdown:FEiYhc;pointerup:mF5Elf;pointerenter:EX0mI;pointerleave:vpvbp;pointercancel:xyn4sd;contextmenu:xexox;focus:h06R8; blur:zjh6rb;mlnRJb:fLiPzd;" data-idom-class="CgzRE" jsname="XVusie" aria-label="Add attachment to Drive TRS Invite 10.29.2024.png" data-tooltip-enabled="true" data-tooltip-id="tt-c147" data-tooltip-classes="AZPksf" id="" jslog="54185; u014N:xr6bB; 1:WyIjdGhyZWFkLWY6MTgxMjkwMjg2MTY5NTYyNDA2OHxtc2ctZjoxODEyOTAyODYxNjk1NjI0MDY4Il0.; 4:WyIjbXNnLWY6MTgxMjkwMjg2MTY5NTYyNDA2OCJd; 43:WyJpbWFnZS9wbmciLDM4NzQzN10."><span class="OiePBf-zPjgPe VYBDae-JX-UHGRz"></span><span class="bHC-Q" jscontroller="LBaJxb" jsname="m9ZlFb" soy-skip="" ssk="6:RWVI5c"></span><span class="VYBDae-JX-ank-Rtc0Jf" jsname="S5tZuc" aria-hidden="true"><span class="bzc-ank" aria-hidden="true"><svg viewBox="0 -960 960 960" height="20" width="20" focusable="false" class=" aoH"><path d="M232-120q-17,0-31.5-8.5t-22.29-22.09L80.79-320.41Q73-334 73-351t8-31L329-809q8-14 22.5-22.5t31.06-8.5H577.44q16.56,0 31.06,8.5t22.42,22.37L811-500q-21-5-42-4.5T727-500L571-768H389L146-351l92,159H575q11,21.17 25.5,39.59T634-120H232Zm68-171l-27-48L445.95-641H514L624-449q-14.32,13-26.53,28.5T576-388L480-556L369-362H565q-6,17-9.5,34.7T552-291H300ZM732-144V-252H624v-72H732V-432h72v108H912v72H804v108H732Z"></path></svg></span></span><div class="VYBDae-JX-ano"></div></button><div class="ne2Ple-oshW8e-J9" id="tt-c147" role="tooltip" aria-hidden="true">Add to Drive</div></span><span data-is-tooltip-wrapper="true" class="a5q" jsaction="JIbuQc:.CLIENT"><button class="VYBDae-JX-I VYBDae-JX-I-ql-ay5-ays CgzRE" jscontroller="PIVayb" jsaction="click:h5M12e; clickmod:h5M12e;pointerdown:FEiYhc;pointerup:mF5Elf;pointerenter:EX0mI;pointerleave:vpvbp;pointercancel:xyn4sd;contextmenu:xexox;focus:h06R8; blur:zjh6rb;mlnRJb:fLiPzd;" data-idom-class="CgzRE" jsname="wtaDCf" aria-label="Save to Photos" data-tooltip-enabled="true" data-tooltip-id="tt-c148" data-tooltip-classes="AZPksf" id="" jslog="54186; u014N:xr6bB; 1:WyIjdGhyZWFkLWY6MTgxMjkwMjg2MTY5NTYyNDA2OHxtc2ctZjoxODEyOTAyODYxNjk1NjI0MDY4Il0.; 4:WyIjbXNnLWY6MTgxMjkwMjg2MTY5NTYyNDA2OCJd; 43:WyJpbWFnZS9wbmciLDM4NzQzN10."><span class="OiePBf-zPjgPe VYBDae-JX-UHGRz"></span><span class="bHC-Q" jscontroller="LBaJxb" jsname="m9ZlFb" soy-skip="" ssk="6:RWVI5c"></span><span class="VYBDae-JX-ank-Rtc0Jf" jsname="S5tZuc" aria-hidden="true"><span class="bzc-ank" aria-hidden="true"><svg viewBox="0 -960 960 960" height="20" width="20" focusable="false" class=" aoH"><path d="M516-384h72V-516H720v-72H588V-720H516v132H384v72H516v132ZM312-240q-29.7,0-50.85-21.15T240-312V-792q0-29.7 21.15-50.85T312-864H792q29.7,0 50.85,21.15T864-792v480q0,29.7-21.15,50.85T792-240H312Zm0-72H792V-792H312v480ZM168-96q-29.7,0-50.85-21.15T96-168V-720h72v552H720v72H168ZM312-792v480V-792Z"></path></svg></span></span><div class="VYBDae-JX-ano"></div></button><div class="ne2Ple-oshW8e-J9" id="tt-c148" role="tooltip" aria-hidden="true">Save to Photos</div></span></div><div><div><p style="font-family:&quot;Helvetica Neue&quot;,Helvetica,Arial,sans-serif;margin-bottom:1.5rem">Best regards,</p><table width="640" cellpadding="0" cellspacing="0" style="padding-bottom:10px;font-family:&quot;Helvetica Neue&quot;,Helvetica,Arial,sans-serif"><tbody><tr><td valign="Top" style="width:136px;border-right-width:2px;border-right-style:solid;border-right-color:rgb(122,21,3)"><div style="font-size:14px;line-height:20px;margin:0px 10px 0px 0px;text-align:center"><img height="126" alt="MadisonJ_Head.png" src=${mads} style="width:126px;height:126px" data-image-whitelisted="" class="CToWUd" data-bit="iit">&nbsp;</div><p style="font-size:14px;line-height:20px;margin:5px 10px 0px 0px;text-align:center"><a href="https://redwoodfp.com/" style="word-break:break-word;color:rgb(0,0,0);text-decoration:none" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://redwoodfp.com/&amp;source=gmail&amp;ust=1729005315869000&amp;usg=AOvVaw34sLGb3Gexth1T8oQCHCsz"><img height="81" alt="Logo.jpg" src=${redwood} style="width:126px;height:81px" data-image-whitelisted="" class="CToWUd" data-bit="iit"></a>&nbsp;</p></td><td style="padding-left:10px"><div style="font-size:14px;line-height:20px;margin:0px"><b style="font-size:18px;color:rgb(122,21,3)">Madison Johnson</b><br>Client Services Associate<br></div><p style="font-size:14px;line-height:20px;margin:10px 0px 0px"><img height="12" alt="phone-call.png" src=${phone} style="width:12px;height:12px" data-image-whitelisted="" class="CToWUd" data-bit="iit">&nbsp;<a href="tel:8173327995" style="word-break:break-word;color:rgb(0,0,0);text-decoration:none" target="_blank">(817) 332-7995 (Office)</a><br><img height="12" alt="printer.png" src=${printer} style="width:12px;height:12px" data-image-whitelisted="" class="CToWUd" data-bit="iit">&nbsp;<a href="tel:8173327996" style="word-break:break-word;color:rgb(0,0,0);text-decoration:none" target="_blank">(817) 332-7996 (Fax)</a><br><img height="12" alt="envelope.png" src=${email} style="width:12px;height:12px" data-image-whitelisted="" class="CToWUd" data-bit="iit">&nbsp;<a href="mailto:madisonj@redwoodfp.com" style="word-break:break-word;color:rgb(0,0,0);text-decoration:none" target="_blank">madisonj@redwoodfp.com</a><br><img height="12" alt="web.png" src=${network} style="width:12px;height:12px" data-image-whitelisted="" class="CToWUd" data-bit="iit">&nbsp;<a href="https://www.redwoodfp.com/" style="word-break:break-word;color:rgb(0,0,0);text-decoration:none" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://www.redwoodfp.com/&amp;source=gmail&amp;ust=1729005315869000&amp;usg=AOvVaw1z-yqNb7t_pLtR0uOB3xPu">www.redwoodfp.com</a><br></p><p style="font-size:12px;line-height:18px;margin:10px 0px 0px">Please do not transmit orders or&nbsp;instructions regarding your Redwood&nbsp;Financial or GWN Securities account by&nbsp;e-mail. For your protection, Redwood&nbsp;Financial or GWN Securities does not&nbsp;act on such instructions.</p></td></tr></tbody></table></div><i style="color:rgb(170,170,170)">&nbsp;If you would like to be removed from future email communication, respond REMOVE</i></div></div></div><div class="yj6qo"></div></div>`;
