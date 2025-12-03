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
    const threadId = row[6]; // ✅ New column

    if (!email || !sequenceName || !step) continue;
    if (status !== "Active") continue;

    const sequenceRow = sequences.find(
      r => r[0] === sequenceName && r[1] === step
    );

    if (!sequenceRow) continue;

    const delayMin = Number(sequenceRow[3]); // DelayMin
    const subjectTemplate = sequenceRow[4];
    const bodyTemplate = sequenceRow[5];
    const replyInThread = sequenceRow[6] === true; // ✅ Checkbox

    if (lastSent) {
      const last = new Date(lastSent);
      const diffMin = (now - last) / (1000 * 60);
      if (diffMin < delayMin) continue;
    }

    // ✅ Check for reply ONLY on the stored most recent thread
    if (threadId && hasRepliedToThread(threadId)) {
      contactsSheet.getRange(i + 1, 6).setValue("Replied");
      continue;
    }

    const subject = subjectTemplate.replace(/{{name}}/g, firstName);
    const body =
      bodyTemplate.replace(/{{name}}/g, firstName) +
      signatureHTML;

    let sentThreadId = null;

    // ✅ Send as reply in-thread
    if (replyInThread && threadId) {
      const thread = GmailApp.getThreadById(threadId);
      thread.reply("", {
        htmlBody: body,
        subject: subject
      });
      sentThreadId = threadId;
    } 
    // ✅ Send as brand new thread
    else {
      const message = GmailApp.sendEmail(email, subject, "", {
        htmlBody: body
      });
      sentThreadId = message.getThreadId(); // ✅ Capture new thread
    }

    // ✅ Persist latest thread
    contactsSheet.getRange(i + 1, 7).setValue(sentThreadId);

    // ✅ Advance step + timestamp
    contactsSheet.getRange(i + 1, 4).setValue(step + 1);
    contactsSheet.getRange(i + 1, 5).setValue(new Date());
  }
}

// ✅ Only checks for replies inside ONE specific thread
function hasRepliedToThread(threadId) {
  const thread = GmailApp.getThreadById(threadId);
  const messages = thread.getMessages();

  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    if (!msg.isDraft() && msg.isInInbox() && msg.getFrom()) {
      return true; // ✅ A reply exists
    }
  }
  return false;
}
