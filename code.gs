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
    const firstName = row[1];
    const sequenceName = row[2];
    const step = row[3];
    const lastSent = row[4];
    const status = row[5];

    if (!email || !sequenceName || !step) continue;
    if (status !== "Active") continue;

    const sequenceRow = sequences.find(
      r => r[0] === sequenceName && r[1] === step
    );

    if (!sequenceRow) continue;

    const delayMin = Number(sequenceRow[3]); // DelayMin (source of truth)
    const subjectTemplate = sequenceRow[4];
    const bodyTemplate = sequenceRow[5];

    if (lastSent) {
      const last = new Date(lastSent);
      const diffMin = (now - last) / (1000 * 60);
      if (diffMin < delayMin) continue;
    }

    if (hasReplied(email)) {
      contactsSheet.getRange(i + 1, 6).setValue("Replied");
      continue;
    }

    const subject = subjectTemplate.replace(/{{name}}/g, firstName);
    const body =
      bodyTemplate.replace(/{{name}}/g, firstName) +
      signatureHTML;

    GmailApp.sendEmail(email, subject, "", {
      htmlBody: body
    });

    contactsSheet.getRange(i + 1, 4).setValue(step + 1);   // Advance step
    contactsSheet.getRange(i + 1, 5).setValue(new Date()); // Update LastSent
  }
}

function hasReplied(email) {
  const threads = GmailApp.search(`from:${email}`);
  return threads.length > 0;
}
