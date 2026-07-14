function exportCandidateApplicationHistoryJson() {
  const SHEET_NAME = "Job_Rejections";
  const OUTPUT_FILE_NAME = "candidate_application_history_export.json";
  const CONTENT_MAX_CHARS = 4000;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Sheet not found: " + SHEET_NAME);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error("No data rows found in sheet: " + SHEET_NAME);
  }

  const headers = values[0].map(h => String(h || "").trim());
  const rows = values.slice(1);

  const records = [];

  rows.forEach((row, index) => {
    const raw = rowToObject_(headers, row);

    const rawContent = raw["Content"] || raw["Snippet"] || "";
    const date = toIsoDateString_(raw["Run Date"] || raw["Date"]);
    const company = cleanText_(raw["Company"]);
    const subject = cleanText_(raw["Subject"]);
    const content = cleanText_(rawContent).substring(0, CONTENT_MAX_CHARS);
    const role = cleanText_(raw["Role"]) || extractRoleForExport_(subject, content);
    const from = cleanText_(raw["From"]);
    const threadId = cleanText_(raw["Thread ID"]);
    const messageId = cleanText_(raw["Message ID"]);

    if (!date && !company && !subject && !content && !messageId) {
      return;
    }

    records.push({
      date: date,
      company: company,
      role: role || null,
      status: "rejection",
      recruiter: from || null,
      source: "gmail_apps_script",
      evidence: content || subject,
      subject: subject,
      message_id: messageId || null,
      thread_id: threadId || null,
      original_row_number: index + 2
    });
  });

  const payload = JSON.stringify(records, null, 2);

  const existingFiles = DriveApp.getFilesByName(OUTPUT_FILE_NAME);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  const file = DriveApp.createFile(
    OUTPUT_FILE_NAME,
    payload,
    "application/json"
  );

  Logger.log("Exported records: " + records.length);
  Logger.log("File name: " + file.getName());
  Logger.log("File URL: " + file.getUrl());
  Logger.log("File ID: " + file.getId());

  return {
    records: records.length,
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    fileId: file.getId()
  };
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((header, i) => {
    obj[header] = row[i];
  });
  return obj;
}

function cleanText_(value) {
  return String(value || "")
    .replace(/\r/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function toIsoDateString_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return value.toISOString();
  }

  const parsed = new Date(value);
  if (!isNaN(parsed)) {
    return parsed.toISOString();
  }

  return String(value || "").trim();
}

function extractRoleForExport_(subject, content) {
  const text = cleanText_([subject, content].filter(Boolean).join("\n"));

  const patterns = [
    /your application to\s+(.+?)\s+at\s+.+/i,
    /Application update for\s+(.+?)\s+at\s+.+/i,
    /application outcome[:\s-]+(.+?)(?:\n|$)/i,
    /application outcome\s*-\s*(.+?)(?:\n|$)/i,
    /unsuccessful job application\s*-\s*.+?\s*-\s*(.+?)(?:\n|$)/i,
    /thank you for (?:your )?(?:recent )?application for (?:the )?(?:role of |position of )?(.+?)(?: position| role| at | with | on |\n|\.|$)/i,
    /thanks for applying for (?:the )?(?:role of |position of )?(.+?)(?: position| role| at | with | on |\n|\.|$)/i,
    /thank you for applying for (?:the )?(?:role of |position of )?(.+?)(?: position| role| at | with | on |\n|\.|$)/i,
    /thank you for your interest in (?:the )?(.+?)(?: position| role| job at| at | with |\n|\.|$)/i,
    /interest in the\s+(.+?)\s+position/i,
    /application for the\s+(.+?)\s+role\s+(?:within|with|at)\s+.+/i,
    /application for the position of\s+(.+?)(?:\.|\n|$)/i,
    /position of\s+(.+?)(?:\n|\.|$)/i,
    /role of\s+(.+?)(?:\n|\.|$)/i,
    /Position:\s*(.+?)(?:\n|$)/i,
    /Role:\s*(.+?)(?:\n|$)/i,
    /RE:\s*(.+?)(?:\n|$)/i
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1]) {
      return cleanRoleForExport_(match[1]);
    }
  }

  return "";
}

function cleanRoleForExport_(value) {
  let role = cleanText_(value)
    .replace(/^your application for\s+/i, "")
    .replace(/^the\s+/i, "")
    .replace(/\s*\(reference[:\s].*$/i, "")
    .replace(/\s*\(req\d+\).*$/i, "")
    .replace(/\s*req\d+.*$/i, "")
    .replace(/\s*job id\s*-.*$/i, "")
    .replace(/\s+advertised by\s+.*$/i, "")
    .replace(/\s+with\s+[A-Z].*$/i, "")
    .replace(/\s+at\s+[A-Z].*$/i, "")
    .replace(/\s+position$/i, "")
    .replace(/\s+role$/i, "")
    .replace(/\s+/g, " ")
    .trim();

  if (/^(this|our|the|your|advertised)$/i.test(role)) {
    return "";
  }

  if (role.length > 140) {
    role = role.substring(0, 140).trim();
  }

  return role;
}
