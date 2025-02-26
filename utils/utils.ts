function setupTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    console.log("Creating new trigger...");
    ScriptApp.newTrigger("sendEmailsInBatches_")
      .timeBased()
      .everyMinutes(1) // Triggers can only be created with minimum 1-minute intervals
      .create();
  } else {
    console.log("Trigger already exists.");
  }
}

function exitJob_() {
  const properties = PropertiesService.getScriptProperties();

  console.log("All emails sent.");

  deleteTrigger_();
  properties.deleteProperty("currentIndex");
}

function deleteTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "sendEmailsInBatches_") {
      console.log("Deleting trigger...");
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function getImg_(fieldId: string) {
  return DriveApp.getFileById(fieldId).getAs("image/png");
}

function getProperty_(scriptPropertyKey: string) {
  const properties = PropertiesService.getScriptProperties();

  return properties.getProperty(scriptPropertyKey);
}

function setProperty_(scriptPropertyKey: string, value: string) {
  const properties = PropertiesService.getScriptProperties();

  properties.setProperty(scriptPropertyKey, value);
}

function deleteProperty_(scriptPropertyKey: string) {
  const properties = PropertiesService.getScriptProperties();

  properties.deleteProperty(scriptPropertyKey);
}
