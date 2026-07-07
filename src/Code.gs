var LABEL_NAME = "Job_Rejections";
var SEARCH_WINDOW = "newer_than:10d";

function extractJobRejections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = getOrCreateSheet(ss, "Job_Rejections");
  const logSheet = getOrCreateSheet(ss, "Run_Log");

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      "Run Date",
      "Gmail Threads",
      "Found",
      "Gmail Labelled",
      "Rows Added",
      "Already In Sheet",
      "Not Rejection"
    ]);
  }

  const lastRow = mainSheet.getLastRow();
  const existingIds = lastRow > 1
    ? mainSheet.getRange(2, 7, lastRow - 1, 1).getValues().flat()
    : [];

  const existingSet = new Set(existingIds.filter(Boolean));

  const threads = GmailApp.search('label:"' + LABEL_NAME + '" ' + SEARCH_WINDOW);

  let gmailThreads = 0;
  let found = 0;
  let gmailLabelled = 0;
  let rowsAdded = 0;
  let alreadyInSheet = 0;
  let notRejection = 0;

  threads.forEach(thread => {
    gmailThreads++;

    const messages = thread.getMessages();
    let rejectionMsg = null;
    let rejectionBody = "";

    for (const msg of messages) {
      const emailText = getFullEmailText(msg);
      const result = classifyJobEmail(msg.getSubject(), emailText);

      if (result.status === "Rejection") {
        rejectionMsg = msg;
        rejectionBody = emailText;
        break;
      }
    }

    if (!rejectionMsg) {
      notRejection++;
      return;
    }

    found++;

    const messageId = rejectionMsg.getId();

    if (existingSet.has(messageId)) {
      alreadyInSheet++;
      return;
    }

    const subject = rejectionMsg.getSubject();
    const from = rejectionMsg.getFrom();

    mainSheet.appendRow([
      rejectionMsg.getDate(),
      extractCompanyName(subject, rejectionBody, from),
      from,
      subject,
      rejectionBody.substring(0, 4000),
      thread.getId(),
      messageId,
      "Rejection"
    ]);

    existingSet.add(messageId);
    rowsAdded++;
  });

  logSheet.appendRow([
    new Date(),
    gmailThreads,
    found,
    gmailLabelled,
    rowsAdded,
    alreadyInSheet,
    notRejection
  ]);
}

function findAndLabelNewRejections() {
  const label = GmailApp.getUserLabelByName(LABEL_NAME) || GmailApp.createLabel(LABEL_NAME);

  const threads = GmailApp.search(SEARCH_WINDOW);

  let checked = 0;
  let found = 0;
  let labelled = 0;
  let alreadyLabelledCount = 0;
  let notRejection = 0;

  const checkedSubjects = [];

  threads.forEach(thread => {
    checked++;

    const alreadyLabelled = thread.getLabels()
      .some(l => l.getName() === LABEL_NAME);

    const messages = thread.getMessages();
    let rejectionFound = false;
    let matchedSubject = "";
    let matchedReason = "";

    for (const msg of messages) {
      const subject = msg.getSubject();
      const emailText = getFullEmailText(msg);
      const result = classifyJobEmail(subject, emailText);

      checkedSubjects.push(subject);

      if (result.status === "Rejection") {
        rejectionFound = true;
        matchedSubject = subject;
        matchedReason = result.reason;
        break;
      }
    }

    if (!rejectionFound) {
      notRejection++;
      return;
    }

    found++;

    if (alreadyLabelled) {
      alreadyLabelledCount++;
      thread.moveToArchive();
      Logger.log("Already labelled and archived: " + matchedSubject + " | " + matchedReason);
      return;
    }

    thread.addLabel(label);
    labelled++;
    thread.moveToArchive();
    Logger.log("Label added and archived: " + matchedSubject + " | " + matchedReason);
  });

  Logger.log("Checked: " + checked);
  Logger.log("Found rejections: " + found);
  Logger.log("New labels added: " + labelled);
  Logger.log("Already labelled: " + alreadyLabelledCount);
  Logger.log("Not rejection: " + notRejection);

  if (checkedSubjects.length > 0) {
    Logger.log("Checked subjects:");
    checkedSubjects.forEach(subject => {
      Logger.log("- " + subject);
    });
  } else {
    Logger.log("Checked subjects: none");
  }
}

function classifyJobEmail(subject, body) {
  const t = norm(subject + " " + body);

  let rejection = 0;
  let positive = 0;

  const strongRejection = [
    /unlikely to progress further/,
    /regret to (inform|advise)/,
    /will not be (moving forward|progressing|proceeding)/,
    /won't be (moving forward|progressing|proceeding)/,
    /not be (moving forward|progressing|proceeding|taking .*next stage)/,
    /not successful/,
    /(application|candidate|interview|role|position).{0,80}unsuccessful/,
    /unsuccessful.{0,80}(application|candidate|role|position|occasion)/,
    /application (has been )?unsuccessful/,
    /job application (has been )?unsuccessful/,
    /unsuccessful on this occasion/,
    /not shortlisted/,
    /not selected/,
    /decided not to (progress|proceed|move forward)/,
    /decided not to.{0,120}(progress|proceed|move forward|continue|advance|take|shortlist|move ahead)/,
    /decided to (move forward|proceed|progress) with other (candidates|applicants)/,
    /decided to move forward with candidates.{0,160}(closely matches|more closely matches|matches).{0,160}(role requirements|requirements)/,
    /(chosen|progressed|proceeding|moving forward) with other (candidates|applicants)/,
    /other candidates.{0,120}(closer|stronger|better|more closely)/,
    /position has been filled/,
    /role.{0,80}already been filled/,
    /unable to offer you an interview/,
    /not able to advance you/,
    /we have reviewed your application.{0,120}(not|unable)/,
    /we will not be taking your application to the next stage/,
    /your application has not been successful/,
    /application has not been successful/,
    /will not be pursuing your application/,
    /not proceeding with your application/,
    /not progress with your application/,
    /not progressing with your application/,
    /unable to progress your application/,
    /unable to progress with your application/,
    /not advance your application/,
    /not be advancing your application/,
    /won’t be progressing your application/,
    /won't be progressing your application/,
    /after careful review.{0,200}(not|unable|unfortunately|decided)/,
    /unfortunately.{0,200}(number of applicants|volume of applications|high number of applications)/,
    /due to.{0,120}(number of applicants|volume of applications|high number of applications)/,

    // PwC-style rejection
    /profile.{0,120}does not meet.{0,120}requirements/,
    /does not meet.{0,120}(requirements|criteria|selection criteria|role requirements)/,
    /determined.{0,120}profile.{0,120}does not meet/
  ];

  const softRejection = [
    /unfortunately/,
    /after careful consideration/,
    /after careful review/,
    /high number of applications/,
    /large number of applications/,
    /strong pool of candidates/,
    /competitive process/
  ];

  const positiveSignals = [
    /schedule a call/,
    /book a time/,
    /availability/,
    /available for an interview/,
    /invite you to interview/,
    /shortlisted for/,
    /proceeding with your application/,
    /progressing your application/,
    /next stage of the interview/,
    /job offer/,
    /offer of employment/,
    /congratulations/
  ];

  for (const r of strongRejection) {
    if (r.test(t)) rejection += 5;
  }

  for (const r of softRejection) {
    if (r.test(t)) rejection += 1;
  }

  for (const r of positiveSignals) {
    if (r.test(t)) positive += 4;
  }

  if (rejection >= 5 && rejection > positive) {
    return {
      status: "Rejection",
      score: rejection,
      reason: "Strong rejection signal"
    };
  }

  return {
    status: "Ignore",
    score: 0,
    reason: "No job rejection signal"
  };
}

function getFullEmailText(msg) {
  return [
    msg.getPlainBody(),
    stripHtml(msg.getBody())
  ].join(" ");
}

function stripHtml(html) {
  return String(html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&middot;/g, " ")
    .replace(/&zwnj;/g, "")
    .replace(/&#8204;/g, "")
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/\s+/g, " ")
    .trim();
}

function extractCompanyName(subject, body, from) {
  const text = subject + "\n" + body;

  let match =
    text.match(/advertised by ([^\n\r.]+)/i) ||
    text.match(/application for .*? at ([^\n\r.]+)/i) ||
    text.match(/position at ([^\n\r.]+)/i) ||
    text.match(/role at ([^\n\r.]+)/i);

  if (match && match[1]) {
    return match[1].trim();
  }

  return from.replace(/<.*?>/g, "").replace(/["']/g, "").trim();
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow([
      "Date",
      "Company",
      "From",
      "Subject",
      "Snippet",
      "Thread ID",
      "Message ID",
      "Status"
    ]);
  }

  return sheet;
}

function norm(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
