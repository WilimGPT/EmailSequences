function runSequences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const contactsSheet = ss.getSheetByName("Contacts");
  const sequencesSheet = ss.getSheetByName("Sequences");
  const settingsSheet = ss.getSheetByName("Settings");

  const contacts = contactsSheet.getDataRange().getValues();
  const sequences = sequencesSheet.getDataRange().getValues();
  const signatureHTML = settingsSheet.getRange("A2").getValue();

  const now = new Date();
  console.log("=== runSequences at " + now.toISOString() + " ===");

  for (let i = 1; i < contacts.length; i++) {
    const rowIndex = i + 1; // Sheet row number
    const row = contacts[i];

    const email = row[0];
    const firstName = row[1];
    const sequenceName = row[2];
    const step = row[3];
    const lastSent = row[4];
    const status = row[5];
    const startAfter = row[6];

    console.log(
      `--- Row ${rowIndex} --- Email=${email}, FirstName=${firstName}, Sequence=${sequenceName}, Step=${step}, Status=${status}, LastSent=${lastSent}, StartAfter=${startAfter}`
    );

    if (!email || !sequenceName || !step) {
      console.log("Missing email/sequence/step. Skipping row.");
      continue;
    }

    if (status !== "Active") {
      console.log(`Status is '${status}', not 'Active'. Skipping row.`);
      continue;
    }

    // HARD GATE: do not allow sequence to start before StartAfter
    if (startAfter) {
      const startAfterDate = new Date(startAfter);
      console.log(`StartAfterDate=${startAfterDate.toISOString()}`);
      if (startAfterDate > now) {
        console.log("StartAfter is in the future. Skipping row.");
        continue;
      }
    }

    const sequenceRow = sequences.find(
      r => r[0] === sequenceName && r[1] === step
    );

    if (!sequenceRow) {
      console.log(
        `No sequence row for Sequence='${sequenceName}', Step=${step}. Marking as Completed.`
      );
      contactsSheet.getRange(rowIndex, 6).setValue("Completed");
      continue;
    }

    const delayMin = Number(sequenceRow[3]);
    const subjectTemplate = sequenceRow[4];
    const bodyTemplate = sequenceRow[5];

    console.log(
      `Found sequence row. DelayMin=${delayMin}, SubjectTemplate="${subjectTemplate}"`
    );

    // Delay logic (based only on LastSent)
    if (lastSent) {
      const last = new Date(lastSent);
      const diffMin = (now - last) / (1000 * 60);
      console.log(
        `LastSent=${last.toISOString()}, diffMin=${diffMin.toFixed(
          2
        )}, required delay=${delayMin}`
      );
      if (diffMin < delayMin) {
        console.log("Delay not yet satisfied. Skipping row.");
        continue;
      }
    } else {
      console.log("No LastSent value. Treating as first send for this contact.");
    }

    // REPLY DETECTION: any email from them since LastSent
    const replied = hasRepliedSinceLastSend(email, lastSent);
    console.log(`hasRepliedSinceLastSend(${email}) = ${replied}`);

    if (replied) {
      console.log("Reply detected. Marking status as Replied and skipping further sends.");
      contactsSheet.getRange(rowIndex, 6).setValue("Replied");
      continue;
    }

    const subject = subjectTemplate.replace(/{{name}}/g, firstName || "");
    const body =
      bodyTemplate.replace(/{{name}}/g, firstName || "") +
      signatureHTML;

    console.log(`Sending email to ${email} with subject: "${subject}"`);
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: body
    });

    // Advance step + update LastSent
    const newStep = Number(step) + 1;
    contactsSheet.getRange(rowIndex, 4).setValue(newStep);
    contactsSheet.getRange(rowIndex, 5).setValue(new Date());
    console.log(`Advanced to Step=${newStep} and updated LastSent.`);
  }

  console.log("=== runSequences finished ===");
}

// REPLY DETECTION: ANY email from them with date > lastSent
function hasRepliedSinceLastSend(email, lastSent) {
  if (!lastSent) {
    console.log(`No lastSent for ${email}; cannot have replies yet.`);
    return false;
  }

  const last = new Date(lastSent);
  console.log(
    `Checking replies for ${email} since ${last.toISOString()}`
  );

  const query = `from:${email}`;
  const threads = GmailApp.search(query);
  console.log(
    `Found ${threads.length} thread(s) from ${email} for reply check.`
  );

  for (let t = 0; t < threads.length; t++) {
    const thread = threads[t];
    const messages = thread.getMessages();
    console.log(` Thread ${t} has ${messages.length} message(s).`);

    for (let m = 0; m < messages.length; m++) {
      const msg = messages[m];
      const from = msg.getFrom();
      const date = msg.getDate();

      console.log(
        `  Msg ${m}: from="${from}", date=${date.toISOString()}`
      );

      if (from && from.indexOf(email) !== -1 && date > last) {
        console.log(
          `  -> Treating this as a reply (date > lastSent).`
        );
        return true;
      }
    }
  }

  console.log(`No replies found for ${email} since ${last.toISOString()}.`);
  return false;
}

function sendDailySummaryAndArchive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contactsSheet = ss.getSheetByName("Contacts");
  const historySheet = ss.getSheetByName("History");

  const data = contactsSheet.getDataRange().getValues();
  if (data.length <= 1) {
    console.log("No data rows in Contacts. Nothing to archive.");
    return;
  }

  const headers = data[0];
  const now = new Date();
  const timezone = Session.getScriptTimeZone();
  const dateLabel = Utilities.formatDate(now, timezone, "yyyy-MM-dd");

  console.log("=== sendDailySummaryAndArchive for " + dateLabel + " ===");

  // Ensure History has headers (once)
  if (historySheet.getLastRow() === 0) {
    const historyHeaders = headers.concat(["ClosedAt"]);
    historySheet.appendRow(historyHeaders);
    console.log("Initialized History sheet headers.");
  }

  const rowsToArchive = [];
  const rowIndexesToDelete = [];

  // Scan Contacts for Completed / Replied
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[5]; // Status column F

    if (status === "Completed" || status === "Replied") {
      rowsToArchive.push(row.concat([now])); // add ClosedAt
      rowIndexesToDelete.push(i + 1);        // sheet row index
    }
  }

  if (rowsToArchive.length === 0) {
    console.log("No Completed/Replied rows to archive today.");
    return; // nothing to email, nothing to move
  }

  // Append to History
  const startRow = historySheet.getLastRow() + 1;
  historySheet
    .getRange(startRow, 1, rowsToArchive.length, rowsToArchive[0].length)
    .setValues(rowsToArchive);

  console.log("Archived " + rowsToArchive.length + " row(s) to History.");

  // Build summary email
  const userEmail = Session.getActiveUser().getEmail();
  let body = "";
  body += "Daily sequence summary for " + dateLabel + "\n\n";
  body += "The following contacts finished or replied and were moved to the History sheet:\n\n";

  rowsToArchive.forEach((r, idx) => {
    const email = r[0];
    const firstName = r[1];
    const sequenceName = r[2];
    const step = r[3];
    const lastSent = r[4];
    const status = r[5];
    const closedAt = r[7];

    body +=
      (idx + 1) + ". " +
      `Name: ${firstName}\n` +
      `   Email: ${email}\n` +
      `   Sequence: ${sequenceName}\n` +
      `   Last Step: ${step}\n` +
      `   Status: ${status}\n` +
      `   LastSent: ${lastSent}\n` +
      `   ClosedAt: ${closedAt}\n\n`;
  });

  const subject = "Daily sequence summary (" + dateLabel + ")";
  GmailApp.sendEmail(userEmail, subject, body);
  console.log("Sent summary email to " + userEmail);

  // Delete from Contacts (bottom-up to avoid index shift)
  rowIndexesToDelete.sort((a, b) => b - a).forEach(rowIndex => {
    contactsSheet.deleteRow(rowIndex);
  });

  console.log("Deleted " + rowIndexesToDelete.length + " row(s) from Contacts.");
  console.log("=== sendDailySummaryAndArchive finished ===");
}

