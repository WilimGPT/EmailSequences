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
    const startAfter = row[6]; // ⬅️ Now column G after removing ThreadId

    if (!email || !sequenceName || !step) continue;
    if (status !== "Active") continue;

    // ✅ HARD GATE: do not allow sequence to start before StartAfter
    if (startAfter && new Date(startAfter) > now) {
      continue;
    }

    const sequenceRow = sequences.find(
      r => r[0] === sequenceName && r[1] === step
    );

    // ✅ NO MORE STEPS = COMPLETED (we’ll wire notification later, as discussed)
    if (!sequenceRow) {
      contactsSheet.getRange(i + 1, 6).setValue("Completed");
      continue;
    }

    const delayMin = Number(sequenceRow[3]);
    const subjectTemplate = sequenceRow[4];
    const bodyTemplate = sequenceRow[5];
    const replyInThread = sequenceRow[6] === true;

    // ✅ Delay logic (based only on LastSent)
    if (lastSent) {
      const last = new Date(lastSent);
      const diffMin = (now - last) / (1000 * 60);
      if (diffMin < delayMin) continue;
    }

    // ✅ SAFE REPLY DETECTION: ANY email since LastSent stops the sequence
    if (hasRepliedSinceLastSend(email, lastSent)) {
      contactsSheet.getRange(i + 1, 6).setValue("Replied");
      continue;
    }

    const subject = subjectTemplate.replace(/{{name}}/g, firstName);
    const body =
      bodyTemplate.replace(/{{name}}/g, firstName) +
      signatureHTML;

    // ✅ SEND (threading is now purely cosmetic UX, not logic)
    if (replyInThread) {
      GmailApp.sendEmail(email, subject, "", {
        htmlBody: body
      });
    } else {
      GmailApp.sendEmail(email, subject, "", {
        htmlBody: body
      });
    }

    // ✅ Advance step + update LastSent
    contactsSheet.getRange(i + 1, 4).setValue(step + 1);
    contactsSheet.getRange(i + 1, 5).setValue(new Date());
  }
}

// ✅ REPLY DETECTION: ANY email from them since LastSent
function hasRepliedSinceLastSend(email, lastSent) {
  if (!lastSent) return false;

  const afterDate = Utilities.formatDate(
    new Date(lastSent),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );

  const query = `from:${email} after:${afterDate}`;
  const threads = GmailApp.search(query);

  return threads.length > 0;
}
