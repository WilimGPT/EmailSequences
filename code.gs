function runSequences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const contactsSheet = ss.getSheetByName("Contacts");
  const sequencesSheet = ss.getSheetByName("Sequences");
  const settingsSheet = ss.getSheetByName("Settings");

  const contacts = contactsSheet.getDataRange().getValues();
  const sequences = sequencesSheet.getDataRange().getValues();
  const signatureHTML = settingsSheet.getRange("A2").getValue();

  const now = new Date();

  for (let i = 1; i < contacts.length; i++) {
    const row = contacts[i];

    const email = row[0];
