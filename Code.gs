
const CFG = {
  MASTER: "Task_Master",
  CYCLES: "Task_Cycles",
  LOG: "Completion_Log",
  DASHBOARD: "Manager_Dashboard",
  ALERTS: "Alert_Log",
  MANAGER_EMAIL: "rdey35689@gmail.com",
  LEADERSHIP_EMAIL: "rdey70729@gmail.com",
  ESCALATION_HOURS: 24
};

const HEADERS = {
  MASTER: [
    "Task ID", "Task Name", "Assigned To", "Employee Email",
    "Frequency Days", "Proof Type", "Due Time",
    "Manager Email", "Leadership Email", "Active"
  ],
  CYCLES: [
    "Cycle ID", "Task ID", "Task Name", "Assigned To",
    "Employee Email", "Frequency Days", "Due Date", "Deadline",
    "Status", "Proof", "Completed By", "Completed At",
    "Manager Ack", "Manager Ack At", "Manager Alert Sent At",
    "Leadership Escalated At", "Next Cycle Created"
  ],
  LOG: [
    "Cycle ID", "Task ID", "Task Name", "Assigned To",
    "Completed By", "Completed At", "Proof"
  ],
  DASHBOARD: [
    "Cycle ID", "Task ID", "Task Name", "Assigned To",
    "Employee Email", "Due Date", "Deadline", "Status",
    "Proof", "Manager Ack", "Manager Ack At",
    "Manager Alert Sent At", "Leadership Escalated At"
  ],
  ALERTS: [
    "Alert Key", "Cycle ID", "Alert Type", "Recipient",
    "Sent At", "Details"
  ]
};


/* ---------- INITIAL SETUP ---------- */

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, CFG.MASTER, HEADERS.MASTER);
  ensureSheet_(ss, CFG.CYCLES, HEADERS.CYCLES);
  ensureSheet_(ss, CFG.LOG, HEADERS.LOG);
  ensureSheet_(ss, CFG.DASHBOARD, HEADERS.DASHBOARD);
  ensureSheet_(ss, CFG.ALERTS, HEADERS.ALERTS);

  refreshDashboard_();

  SpreadsheetApp.getUi().alert(
    "Setup complete. Your existing sheet data has been preserved."
  );
}


/* ---------- CREATE INITIAL TASK CYCLES ---------- */

function initializeTaskCycles() {
  const tasks = getActiveTasks_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CFG.CYCLES);

  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  const existing = new Set();

  for (let i = 1; i < data.length; i++) {
    const taskId = String(data[i][h["Task ID"]] || "").trim();
    const dueDate = asDate_(data[i][h["Due Date"]]);

    if (taskId && dueDate) {
      existing.add(taskId + "|" + dateKey_(dueDate));
    }
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  tasks.forEach(task => {
    const key = task.id + "|" + dateKey_(today);

    if (!existing.has(key)) {
      createCycle_(task, today);
      existing.add(key);
    }
  });

  refreshDashboard_();

  SpreadsheetApp.getUi().alert(
    "Initial cycle check complete. Existing task/date combinations were not duplicated."
  );
}


/* ---------- MAIN AUTOMATION ---------- */

function processAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.CYCLES);

  if (!sheet || sheet.getLastRow() < 2) {
    refreshDashboard_();
    return;
  }

  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    const cycleId = String(r[h["Cycle ID"]] || "").trim();
    const taskId = String(r[h["Task ID"]] || "").trim();
    const status = String(r[h["Status"]] || "")
      .trim().toUpperCase();

    if (!cycleId || !taskId || status === "COMPLETED") continue;

    const deadline = asDate_(r[h["Deadline"]]);

    if (!deadline) continue;

    // Mark unfinished cycles overdue once their deadline has passed.
    if (now > deadline && status !== "OVERDUE") {
      sheet.getRange(i + 1, h["Status"] + 1).setValue("OVERDUE");
    }

    const managerEmail = getManagerEmail_(taskId);
    const leadershipEmail = getLeadershipEmail_(taskId);

    const managerAlertSent = asDate_(r[h["Manager Alert Sent At"]]);
    const managerAck = String(r[h["Manager Ack"]] || "")
      .trim().toUpperCase();

    // Send one manager alert for an overdue cycle.
    if (now > deadline && !managerAlertSent && managerEmail) {
      const sent = sendAlertOnce_(
        cycleId,
        "MANAGER_OVERDUE",
        managerEmail,
        buildAlertBody_(r, h, "OVERDUE TASK ALERT")
      );

      if (sent) {
        sheet.getRange(i + 1, h["Manager Alert Sent At"] + 1)
          .setValue(new Date());
      }
    }

    // Escalate if the manager has not acknowledged within 24 hours
    // after the manager alert was sent.
    const latestManagerAlert = asDate_(
      sheet.getRange(i + 1, h["Manager Alert Sent At"] + 1)
        .getValue()
    );

    const leadershipEscalated = asDate_(
      sheet.getRange(i + 1, h["Leadership Escalated At"] + 1)
        .getValue()
    );

    if (
      latestManagerAlert &&
      !leadershipEscalated &&
      managerAck !== "YES" &&
      leadershipEmail &&
      now.getTime() - latestManagerAlert.getTime() >=
        CFG.ESCALATION_HOURS * 60 * 60 * 1000
    ) {
      const sent = sendAlertOnce_(
        cycleId,
        "LEADERSHIP_ESCALATION",
        leadershipEmail,
        buildAlertBody_(r, h, "LEADERSHIP ESCALATION")
      );

      if (sent) {
        sheet.getRange(i + 1, h["Leadership Escalated At"] + 1)
          .setValue(new Date());
      }
    }
  }

  // Repair missing next cycles and prevent duplicates.
  createNextCyclesForCompleted_();

  refreshDashboard_();
}


/* ---------- COMPLETE SELECTED TASK ---------- */

function completeSelectedTask() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.CYCLES);
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== CFG.CYCLES) {
    SpreadsheetApp.getUi().alert(
      "Open Task_Cycles and select the row you want to complete."
    );
    return;
  }

  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert("Select a task row, not the header.");
    return;
  }

  const h = headerMap_(sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]);

  const status = String(
    sheet.getRange(row, h["Status"] + 1).getValue() || ""
  ).trim().toUpperCase();

  if (status === "COMPLETED") {
    SpreadsheetApp.getUi().alert("This task is already completed.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Complete Task",
    "Enter proof of completion, such as a ticket ID or WhatsApp update reference:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const proof = response.getResponseText().trim();

  if (!proof) {
    ui.alert("Proof is required. The task was not completed.");
    return;
  }

  const cycleId = String(
    sheet.getRange(row, h["Cycle ID"] + 1).getValue()
  ).trim();

  const taskId = String(
    sheet.getRange(row, h["Task ID"] + 1).getValue()
  ).trim();

  const taskName = String(
    sheet.getRange(row, h["Task Name"] + 1).getValue()
  ).trim();

  const assignedTo = String(
    sheet.getRange(row, h["Assigned To"] + 1).getValue()
  ).trim();

  const email = Session.getActiveUser().getEmail() ||
    Session.getEffectiveUser().getEmail() || "";

  const completedAt = new Date();

  sheet.getRange(row, h["Status"] + 1).setValue("COMPLETED");
  sheet.getRange(row, h["Proof"] + 1).setValue(proof);
  sheet.getRange(row, h["Completed By"] + 1).setValue(email);
  sheet.getRange(row, h["Completed At"] + 1).setValue(completedAt);

  const logSheet = ss.getSheetByName(CFG.LOG);

  logSheet.appendRow([
    cycleId,
    taskId,
    taskName,
    assignedTo,
    email,
    completedAt,
    proof
  ]);

  createNextCyclesForCompleted_();
  refreshDashboard_();

  ui.alert("Task completed and logged successfully.");
}


/* ---------- MANAGER ACKNOWLEDGMENT ---------- */

function acknowledgeSelectedTask() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.CYCLES);

  if (SpreadsheetApp.getActiveSheet().getName() !== CFG.CYCLES) {
    SpreadsheetApp.getUi().alert(
      "Open Task_Cycles and select the row to acknowledge."
    );
    return;
  }

  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert("Select a task row, not the header.");
    return;
  }

  const h = headerMap_(sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]);

  sheet.getRange(row, h["Manager Ack"] + 1).setValue("YES");
  sheet.getRange(row, h["Manager Ack At"] + 1).setValue(new Date());

  refreshDashboard_();

  SpreadsheetApp.getUi().alert("Manager acknowledgment recorded.");
}


/* ---------- DASHBOARD ---------- */

function refreshDashboard() {
  refreshDashboard_();
}

function refreshDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cycles = ss.getSheetByName(CFG.CYCLES);
  const dashboard = ss.getSheetByName(CFG.DASHBOARD);

  if (!cycles || !dashboard) return;

  const cycleData = cycles.getDataRange().getValues();
  const h = headerMap_(cycleData[0]);

  dashboard.clearContents();
  dashboard.getRange(1, 1, 1, HEADERS.DASHBOARD.length)
    .setValues([HEADERS.DASHBOARD]);

  if (cycleData.length < 2) return;

  const output = [];

  for (let i = 1; i < cycleData.length; i++) {
    const r = cycleData[i];

    output.push([
      r[h["Cycle ID"]],
      r[h["Task ID"]],
      r[h["Task Name"]],
      r[h["Assigned To"]],
      r[h["Employee Email"]],
      r[h["Due Date"]],
      r[h["Deadline"]],
      r[h["Status"]],
      r[h["Proof"]],
      r[h["Manager Ack"]],
      r[h["Manager Ack At"]],
      r[h["Manager Alert Sent At"]],
      r[h["Leadership Escalated At"]]
    ]);
  }

  if (output.length) {
    dashboard.getRange(
      2, 1, output.length, HEADERS.DASHBOARD.length
    ).setValues(output);
  }

  dashboard.setFrozenRows(1);
}


/* ---------- TEST EMAIL ---------- */

function sendTestEmail() {
  const recipient = CFG.MANAGER_EMAIL;

  MailApp.sendEmail({
    to: recipient,
    subject: "Allo Innoware Automation - Test Email",
    body:
      "This is a test email from your Google Apps Script automation.\n\n" +
      "If you received this email, email sending is working."
  });

  SpreadsheetApp.getUi().alert(
    "Test email sent to " + recipient
  );
}


/* ---------- AUTOMATION TRIGGER ---------- */

function installAutomationTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "processAutomation") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("processAutomation")
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getUi().alert(
    "Automation trigger installed. It will run every 5 minutes."
  );
}


/* ---------- CREATE NEXT CYCLES FOR COMPLETED TASKS ---------- */

function createNextCyclesForCompleted_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.CYCLES);

  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  const tasks = getActiveTasks_();
  const taskMap = new Map(tasks.map(t => [t.id, t]));
  const existingTaskDates = new Set();

  // Record every task and due-date combination already present.
  for (let i = 1; i < data.length; i++) {
    const taskId = String(data[i][h["Task ID"]] || "").trim();
    const dueDate = asDate_(data[i][h["Due Date"]]);

    if (taskId && dueDate) {
      existingTaskDates.add(taskId + "|" + dateKey_(dueDate));
    }
  }

  // Check completed cycles, even if the creation timestamp is already set.
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const taskId = String(r[h["Task ID"]] || "").trim();
    const status = String(r[h["Status"]] || "").trim().toUpperCase();
    const proof = String(r[h["Proof"]] || "").trim();
    const completedAt = asDate_(r[h["Completed At"]]);
    const dueDate = asDate_(r[h["Due Date"]]);

    if (
      !taskId ||
      status !== "COMPLETED" ||
      !proof ||
      !completedAt ||
      !dueDate
    ) continue;

    const task = taskMap.get(taskId);
    if (!task) continue;

    // Calculate the next due date from the current cycle's due date.
    const nextDate = new Date(dueDate);
    nextDate.setDate(nextDate.getDate() + task.frequency);

    const key = taskId + "|" + dateKey_(nextDate);

    // Create only if this task/date combination does not already exist.
    if (!existingTaskDates.has(key)) {
      createCycle_(task, nextDate);
      existingTaskDates.add(key);
    }

    // Fill the timestamp only if it is blank.
    if (!r[h["Next Cycle Created"]]) {
      sheet.getRange(i + 1, h["Next Cycle Created"] + 1)
        .setValue(new Date());
    }
  }
}


/* ---------- CREATE ONE CYCLE ---------- */

function createCycle_(task, dueDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.CYCLES);

  const date = new Date(dueDate);
  date.setHours(0, 0, 0, 0);

  const deadline = new Date(date);
  const dueTime = String(task.dueTime || "17:00").trim();
  const parts = dueTime.split(":");

  deadline.setHours(
    Number(parts[0]) || 17,
    Number(parts[1]) || 0,
    0,
    0
  );

  const cycleId = task.id + "-" +
    Utilities.getUuid().substring(0, 8);

  sheet.appendRow([
    cycleId,
    task.id,
    task.name,
    task.assignedTo,
    task.email,
    task.frequency,
    date,
    deadline,
    "PENDING",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ]);

  return cycleId;
}


/* ---------- SEND ALERT ONCE ---------- */

function sendAlertOnce_(cycleId, alertType, recipient, body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CFG.ALERTS);

  if (!recipient) return false;

  const alertKey = cycleId + "|" + alertType;
  const data = logSheet.getDataRange().getValues();
  const h = headerMap_(data[0]);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][h["Alert Key"]] || "") === alertKey) {
      return false;
    }
  }

  MailApp.sendEmail({
    to: recipient,
    subject: alertType === "MANAGER_OVERDUE"
      ? "Overdue Task Alert: " + cycleId
      : "Leadership Escalation: " + cycleId,
    body: body
  });

  logSheet.appendRow([
    alertKey,
    cycleId,
    alertType,
    recipient,
    new Date(),
    body
  ]);

  return true;
}


/* ---------- ALERT MESSAGE ---------- */

function buildAlertBody_(row, h, title) {
  return title + "\n\n" +
    "Cycle ID: " + row[h["Cycle ID"]] + "\n" +
    "Task ID: " + row[h["Task ID"]] + "\n" +
    "Task: " + row[h["Task Name"]] + "\n" +
    "Assigned To: " + row[h["Assigned To"]] + "\n" +
    "Employee Email: " + row[h["Employee Email"]] + "\n" +
    "Due Date: " + formatDate_(row[h["Due Date"]]) + "\n" +
    "Deadline: " + formatDate_(row[h["Deadline"]]) + "\n" +
    "Status: " + row[h["Status"]] + "\n" +
    "Proof: " + (row[h["Proof"]] || "Not provided") + "\n\n" +
    "Please review this task in the Manager_Dashboard sheet.";
}


/* ---------- TASK LOOKUPS ---------- */

function getActiveTasks_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.MASTER);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  const tasks = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    if (!isActive_(r[h["Active"]])) continue;

    const id = String(r[h["Task ID"]] || "").trim();
    if (!id) continue;

    tasks.push({
      id: id,
      name: String(r[h["Task Name"]] || "").trim(),
      assignedTo: String(r[h["Assigned To"]] || "").trim(),
      email: getTaskEmail_(r[h["Employee Email"]]),
      frequency: Math.max(
        1, Number(r[h["Frequency Days"]]) || 1
      ),
      proofType: String(r[h["Proof Type"]] || "").trim(),
      dueTime: String(r[h["Due Time"]] || "17:00").trim(),
      managerEmail: String(r[h["Manager Email"]] || "").trim(),
      leadershipEmail: String(
        r[h["Leadership Email"]] || ""
      ).trim()
    });
  }

  return tasks;
}

function getTaskEmail_(email) {
  return String(email || "").trim();
}

function getManagerEmail_(taskId) {
  const task = getActiveTasks_().find(t => t.id === taskId);
  return task && task.managerEmail
    ? task.managerEmail
    : CFG.MANAGER_EMAIL;
}

function getLeadershipEmail_(taskId) {
  const task = getActiveTasks_().find(t => t.id === taskId);
  return task && task.leadershipEmail
    ? task.leadershipEmail
    : CFG.LEADERSHIP_EMAIL;
}


/* ---------- GENERAL HELPERS ---------- */

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function headerMap_(headers) {
  const map = {};

  headers.forEach((header, index) => {
    map[String(header).trim()] = index;
  });

  return map;
}

function isActive_(value) {
  if (value === true) return true;

  const text = String(value || "").trim().toUpperCase();
  return text === "TRUE" || text === "YES" || text === "1";
}

function asDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value);
  }

  if (value === null || value === undefined || value === "") {
    return null;
  }

  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function dateKey_(value) {
  const date = asDate_(value);
  if (!date) return "";

  return Utilities.formatDate(
    date,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    "yyyy-MM-dd"
  );
}

function formatDate_(value) {
  const date = asDate_(value);
  if (!date) return String(value || "");

  return Utilities.formatDate(
    date,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    "dd/MM/yyyy HH:mm"
  );
}


/* ---------- DIAGNOSTIC CHECK ---------- */

function checkNextCycle() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CFG.CYCLES);

  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    if (String(r[h["Task ID"]] || "").trim() === "T001") {
      results.push({
        cycleId: r[h["Cycle ID"]],
        dueDate: formatDate_(r[h["Due Date"]]),
        status: r[h["Status"]],
        nextCycleCreated: formatDate_(
          r[h["Next Cycle Created"]]
        )
      });
    }
  }

  Logger.log(JSON.stringify(results));
}
