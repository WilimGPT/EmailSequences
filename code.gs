function runSequences() {
  const now = new Date();

  // 🚫 HARD WEEKEND BLOCK — do NOTHING on Saturdays & Sundays
  const todayDay = now.getDay(); // 0 = Sunday, 6 = Saturday
  if (todayDay === 0 || todayDay === 6) {
    console.log("Weekend detected. runSequences exited without processing.");
    return;
  }

  let ss, contactsSheet, sequencesSheet, settingsSheet;

  // ✅ HARD FATAL GUARD — missing core sheets
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    contactsSheet = ss.getSheetByName("Contacts");
    sequencesSheet = ss.getSheetByName("Sequences");
    settingsSheet = ss.getSheetByName("Settings");

    if (!contactsSheet || !sequencesSheet || !settingsSheet) {
      throw new Error("One or more required sheets are missing.");
    }
  } catch (err) {
    console.error("FATAL: Core sheets missing.", err);
    return;
  }

  const contacts = contactsSheet.getDataRange().getValues();
  const sequences = sequencesSheet.getDataRange().getValues();
  const signatureHTML = settingsSheet.getRange("A2").getValue();

  console.log("=== runSequences at " + now.toISOString() + " ===");

  for (let i = 1; i < contacts.length; i++) {
    const rowIndex = i + 1;

    try {
      const row = contacts[i];

      const email = row[0];
      const firstName = row[1];
      const sequenceName = row[2];
      const step = Number(row[3]);
      const lastSent = row[4];
      const status = row[5];
      const startAfter = row[6];

      console.log(
        `--- Row ${rowIndex} --- ${email} | ${sequenceName} | Step ${step} | Status ${status}`
      );

      // ✅ BASIC VALIDATION
      if (!email || !sequenceName || !step) {
        throw new Error("Missing required fields: Email, Sequence or Step.");
      }

      if (status !== "Active") continue;

      // ✅ StartAfter gate
      if (startAfter) {
        const startAfterDate = new Date(startAfter);
        if (isNaN(startAfterDate.getTime())) {
          throw new Error("Invalid StartAfter date.");
        }

        if (startAfterDate > now) continue;
      }

      // ✅ Find sequence row safely
      const sequenceRow = sequences.find(
        r => r[0] === sequenceName && Number(r[1]) === step
      );

      if (!sequenceRow) {
        console.log("No more steps → marking Completed");
        contactsSheet.getRange(rowIndex, 6).setValue("Completed");
        continue;
      }

      const delayMin = Number(sequenceRow[3]);
      if (isNaN(delayMin)) {
        throw new Error("DelayMin is not a valid number.");
      }

      const subjectTemplate = sequenceRow[4];
      const bodyTemplate = sequenceRow[5];

      // ✅ OPTIMISED BUSINESS-TIME DELAY LOGIC (SAFE)
      if (lastSent) {
        const lastSentDate = new Date(lastSent);
        if (isNaN(lastSentDate.getTime())) {
          throw new Error("LastSent is not a valid date.");
        }

        const businessMinutes = getBusinessMinutesBetween(
          lastSentDate,
          now
        );

        console.log(
          `Business minutes elapsed: ${businessMinutes} / required: ${delayMin}`
        );

        if (businessMinutes < delayMin) continue;
      }

      // ✅ SAFE REPLY DETECTION
      let replied = false;
      try {
        replied = hasRepliedSinceLastSend(email, lastSent);
      } catch (replyErr) {
        console.warn("Reply detection failed for row " + rowIndex, replyErr);
      }

      if (replied) {
        console.log("Reply detected → marking Replied");
        contactsSheet.getRange(rowIndex, 6).setValue("Replied");
        continue;
      }

      const subject = subjectTemplate.replace(/{{name}}/g, firstName || "");
      const body =
        bodyTemplate.replace(/{{name}}/g, firstName || "") +
        signatureHTML;

      console.log(`✅ Sending: ${email} → "${subject}"`);

      // ✅ SAFE SEND
      try {
        GmailApp.sendEmail(email, subject, "", {
          htmlBody: body
        });
      } catch (mailErr) {
        throw new Error("Gmail send failed: " + mailErr.message);
      }

      contactsSheet.getRange(rowIndex, 4).setValue(step + 1);
      contactsSheet.getRange(rowIndex, 5).setValue(new Date());

    } catch (rowErr) {
      console.error(`❌ Row ${rowIndex} failed:`, rowErr.message);

      // ✅ VISUAL ERROR FLAG (turn entire row red)
      try {
        contactsSheet
          .getRange(rowIndex, 1, 1, contacts[0].length)
          .setFontColor("red");

        // ✅ Optional: mark Status as Error
        contactsSheet.getRange(rowIndex, 6).setValue("Error");
      } catch (flagErr) {
        console.error("Failed to flag row visually:", flagErr.message);
      }
    }
  }

  console.log("=== runSequences finished ===");
}


function getBusinessMinutesBetween(start, end) {
  const totalMinutes = Math.floor((end - start) / 60000);
  const totalDays = Math.floor(totalMinutes / 1440);
  const fullWeeks = Math.floor(totalDays / 7);

  let weekendDays = fullWeeks * 2;

  const remainingDays = totalDays % 7;
  const startDay = start.getDay(); // 0 = Sun, 6 = Sat

  for (let i = 1; i <= remainingDays; i++) {
    const day = (startDay + i) % 7;
    if (day === 0 || day === 6) {
      weekendDays++;
    }
  }

  const weekendMinutes = weekendDays * 1440;
  const businessMinutes = totalMinutes - weekendMinutes;

  return Math.max(0, businessMinutes);
}


function hasRepliedSinceLastSend(email, lastSent) {
  if (!lastSent) return false;

  const last = new Date(lastSent);
  if (isNaN(last.getTime())) {
    throw new Error("Invalid LastSent date in reply detection.");
  }

  const query = `from:${email}`;
  const threads = GmailApp.search(query);

  for (let t = 0; t < threads.length; t++) {
    const messages = threads[t].getMessages();

    for (let m = 0; m < messages.length; m++) {
      const msg = messages[m];
      const from = msg.getFrom();
      const date = msg.getDate();

      if (from && from.indexOf(email) !== -1 && date > last) {
        return true;
      }
    }
  }

  return false;
}


function sendDailySummaryAndArchive() {
  let ss, contactsSheet, historySheet;

  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    contactsSheet = ss.getSheetByName("Contacts");
    historySheet = ss.getSheetByName("History");

    if (!contactsSheet || !historySheet) {
      throw new Error("Missing Contacts or History sheet.");
    }
  } catch (fatalErr) {
    console.error("FATAL SUMMARY FAILURE:", fatalErr.message);
    return;
  }

  const data = contactsSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const now = new Date();
  const timezone = Session.getScriptTimeZone();
  const dateLabel = Utilities.formatDate(now, timezone, "yyyy-MM-dd");

  if (historySheet.getLastRow() === 0) {
    historySheet.appendRow(headers.concat(["ClosedAt"]));
  }

  const rowsToArchive = [];
  const rowIndexesToDelete = [];

  for (let i = 1; i < data.length; i++) {
    try {
      const row = data[i];
      const status = row[5];

      if (status === "Completed" || status === "Replied") {
        rowsToArchive.push(row.concat([now]));
        rowIndexesToDelete.push(i + 1);
      }

    } catch (archErr) {
      console.error("Archive scan failed on row", i + 1, archErr.message);
    }
  }

  if (rowsToArchive.length === 0) return;

  historySheet
    .getRange(
      historySheet.getLastRow() + 1,
      1,
      rowsToArchive.length,
      rowsToArchive[0].length
    )
    .setValues(rowsToArchive);

  const userEmail = Session.getActiveUser().getEmail();
  let body = "Daily sequence summary for " + dateLabel + "\n\n";

  rowsToArchive.forEach((r, idx) => {
    body +=
      `${idx + 1}. ${r[1]} (${r[0]}) | ${r[2]} | Status: ${r[5]}\n`;
  });

  try {
    GmailApp.sendEmail(
      userEmail,
      "Daily sequence summary (" + dateLabel + ")",
      body
    );
  } catch (mailErr) {
    console.error("Summary email failed:", mailErr.message);
  }

  rowIndexesToDelete.sort((a, b) => b - a).forEach(rowIndex => {
    try {
      contactsSheet.deleteRow(rowIndex);
    } catch (delErr) {
      console.error("Delete failed on row", rowIndex, delErr.message);
    }
  });
}
