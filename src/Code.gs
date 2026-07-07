var LABEL_NAME = "Job_Rejections";
var SEARCH_WINDOW = "newer_than:10d";
var BATCH_SIZE = 100;

// Optional. Leave blank if this script is bound to the Google Sheet.
// If the script is standalone and getActiveSpreadsheet() fails, paste the Sheet ID here.
var SPREADSHEET_ID = "";

function runJobRejectionScan() {
  findAndLabelNewRejections();
  extractJobRejections();
}

function extractJobRejections() {
  const ss = getSpreadsheet_();
  const mainSheet = getOrCreateSheet(ss, "Job_Rejections");
  const logSheet = getOrCreateSheet(ss, "Run_Log");

  ensureRunLogHeader_(logSheet);

  const lastRow = mainSheet.getLastRow();
  const existingIds = lastRow > 1
    ? mainSheet.getRange(2, 7, lastRow - 1, 1).getValues().flat()
    : [];

  const existingSet = new Set(existingIds.filter(Boolean));

  const threads = GmailApp.search('label:"' + LABEL_NAME + '" ' + SEARCH_WINDOW, 0, BATCH_SIZE);

  let gmailThreads = 0;
  let found = 0;
  let rowsAdded = 0;
  let alreadyInSheet = 0;
  let notRejection = 0;

  threads.forEach(thread => {
    gmailThreads++;

    const result = findBestRejectionMessage_(thread.getMessages());

    if (!result.msg) {
      notRejection++;
      return;
    }

    found++;

    const rejectionMsg = result.msg;
    const rejectionBody = result.body;
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
    found,
    rowsAdded,
    alreadyInSheet,
    notRejection
  ]);
}

function findAndLabelNewRejections() {
  const label = GmailApp.getUserLabelByName(LABEL_NAME) || GmailApp.createLabel(LABEL_NAME);
  const threads = GmailApp.search(SEARCH_WINDOW, 0, BATCH_SIZE);

  let checked = 0;
  let found = 0;
  let labelled = 0;
  let alreadyLabelledCount = 0;
  let notRejection = 0;
  let reviewOnly = 0;

  const checkedSubjects = [];

  threads.forEach(thread => {
    checked++;

    const alreadyLabelled = thread.getLabels()
      .some(l => l.getName() === LABEL_NAME);

    const messages = thread.getMessages();
    const result = findBestRejectionMessage_(messages);

    messages.forEach(msg => checkedSubjects.push(msg.getSubject()));

    if (!result.msg) {
      if (result.bestStatus === "Review") {
        reviewOnly++;
      } else {
        notRejection++;
      }
      return;
    }

    found++;

    if (alreadyLabelled) {
      alreadyLabelledCount++;
      Logger.log("Already labelled: " + result.msg.getSubject() + " | " + result.reason);
      return;
    }

    thread.addLabel(label);
    thread.moveToArchive();
    labelled++;

    Logger.log("Label added: " + result.msg.getSubject() + " | " + result.reason);
  });

  Logger.log("Checked threads: " + checked);
  Logger.log("Found rejections: " + found);
  Logger.log("New labels added: " + labelled);
  Logger.log("Already labelled: " + alreadyLabelledCount);
  Logger.log("Review only: " + reviewOnly);
  Logger.log("Not rejection: " + notRejection);

  if (checkedSubjects.length > 0) {
    Logger.log("Checked subjects:");
    checkedSubjects.forEach(subject => Logger.log("- " + subject));
  } else {
    Logger.log("Checked subjects: none");
  }
}

function findBestRejectionMessage_(messages) {
  let bestStatus = "Ignore";
  let bestScore = 0;
  let bestReason = "";

  // Check newest email first. Rejections are often the latest reply in a thread.
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    const emailText = getFullEmailText(msg);
    const result = classifyJobEmail(msg.getSubject(), emailText);

    if (result.score > bestScore) {
      bestScore = result.score;
      bestStatus = result.status;
      bestReason = result.reason;
    }

    if (result.status === "Rejection") {
      return {
        msg: msg,
        body: emailText,
        status: result.status,
        score: result.score,
        reason: result.reason,
        bestStatus: result.status
      };
    }
  }

  return {
    msg: null,
    body: "",
    status: "Ignore",
    score: bestScore,
    reason: bestReason || "No rejection message found",
    bestStatus: bestStatus
  };
}

function classifyJobEmail(subject, body) {
  const t = norm(subject + " " + body);

  let rejection = 0;
  let positive = 0;
  let acknowledgement = 0;
  let jobContext = 0;
  let reasons = [];

  if (/(application|candidate|applicant|position|role|recruitment|hiring|career|vacancy|job application|selection process|talent acquisition|talent team|recruiter|job|interview)/.test(t)) {
    jobContext += 2;
  }

  const strongRejection = [
    [/unlikely to progress further/, "unlikely to progress further"],
    [/regret to (inform|advise)/, "regret wording"],

    // Negated progression
    [/won't be (moving forward|progressing|proceeding|advancing)/, "won't be progressing"],
    [/will not be (moving forward|progressing|proceeding|advancing)/, "will not be progressing"],
    [/we won't be progressing your application/, "won't be progressing your application"],
    [/we will not be progressing your application/, "will not be progressing your application"],
    [/won't be progressing .*application/, "won't be progressing application"],
    [/will not be progressing .*application/, "will not be progressing application"],
    [/not be (moving forward|progressing|proceeding|advancing|taking .*next stage)/, "not moving forward/progressing"],
    [/not taking your application to the next stage/, "not taking application to next stage"],

    // Telstra-style rejection
    [/decided to move forward with candidates/, "move forward with other candidates"],
    [/move forward with candidates whose experience/, "candidates whose experience more closely matches"],
    [/candidates whose experience more closely matches/, "experience more closely matches"],
    [/experience more closely matches .*requirements/, "experience more closely matches requirements"],
    [/(candidate|candidates|applicant|applicants) whose (experience|skills|background|profile|qualifications).{0,120}(closely|better|more closely).{0,50}(match|matches|align|aligns|meet|meets)/, "other candidates better match"],

    // Common rejection wording
    [/not successful/, "not successful"],
    [/(application|candidate|applicant|interview|role|position).{0,80}unsuccessful/, "unsuccessful application/candidate"],
    [/unsuccessful.{0,80}(application|candidate|applicant|role|position|occasion)/, "unsuccessful wording"],
    [/application (has been )?unsuccessful/, "application unsuccessful"],
    [/job application (has been )?unsuccessful/, "job application unsuccessful"],
    [/unsuccessful on this occasion/, "unsuccessful on this occasion"],
    [/not shortlisted/, "not shortlisted"],
    [/not selected/, "not selected"],
    [/not selected to (move forward|progress|proceed|continue)/, "not selected to proceed"],
    [/decided not to (progress|proceed|move forward|continue)/, "decided not to progress"],
    [/decided to (move forward|proceed|progress|continue) with other (candidates|applicants)/, "progressing other candidates"],
    [/(chosen|progressed|proceeding|moving forward|continuing) with other (candidates|applicants)/, "other candidates chosen"],
    [/other candidates.{0,120}(closer|stronger|better|more closely|better suited|more suitable)/, "other candidates stronger"],
    [/another candidate.{0,120}(closer|stronger|better|more closely|better suited|more suitable)/, "another candidate stronger"],

    // Filled / closed / no longer available
    [/position has been filled/, "position filled"],
    [/role.{0,80}already been filled/, "role filled"],
    [/(position|role|vacancy).{0,80}(closed|cancelled|withdrawn|no longer available)/, "role closed/cancelled"],

    // Interview / next stage rejection
    [/unable to offer you an interview/, "unable to offer interview"],
    [/not able to offer you an interview/, "not able to offer interview"],
    [/not able to advance you/, "not able to advance"],
    [/unable to advance you/, "unable to advance"],
    [/we will not be taking your application to the next stage/, "not taking to next stage"],
    [/will not be invited to (interview|the next stage)/, "not invited to next stage"],

    // Application reviewed but rejected
    [/we have reviewed your application.{0,120}(not|unable)/, "reviewed application not/unable"],
    [/your application has not been successful/, "application not successful"],
    [/application has not been successful/, "application not successful"],
    [/will not be pursuing your application/, "not pursuing application"],
    [/not proceeding with your application/, "not proceeding with application"],
    [/not proceed with your application/, "not proceed with application"],
    [/not progress with your application/, "not progress with application"],
    [/not progressing with your application/, "not progressing with application"],
    [/unable to progress your application/, "unable to progress application"],
    [/unable to proceed with your application/, "unable to proceed application"],
    [/not advance your application/, "not advance application"],
    [/not be advancing your application/, "not advancing application"],
    [/no longer considering your application/, "no longer considering application"],
    [/no longer under consideration/, "no longer under consideration"],

    // Requirement / fit rejection
    [/profile.{0,120}does not meet.{0,120}requirements/, "profile does not meet requirements"],
    [/does not meet.{0,120}(requirements|criteria|selection criteria|role requirements)/, "does not meet requirements"],
    [/determined.{0,120}profile.{0,120}does not meet/, "profile does not meet"],
    [/not the right fit/, "not the right fit"],
    [/not a match for (this|the) role/, "not a match"],
    [/not aligned with (this|the) role/, "not aligned"],
    [/more closely aligned with (the|our) requirements/, "others more closely aligned"],
    [/better aligned with (the|our) requirements/, "others better aligned"]
  ];

  const softRejection = [
    [/unfortunately/, "unfortunately"],
    [/after careful consideration/, "after careful consideration"],
    [/after careful review/, "after careful review"],
    [/after reviewing your application/, "after reviewing application"],
    [/high number of applications/, "high number of applications"],
    [/large number of applications/, "large number of applications"],
    [/strong pool of candidates/, "strong pool"],
    [/competitive process/, "competitive process"],
    [/not the outcome you were hoping for/, "not the outcome hoped for"],
    [/wish you (all )?the best/, "wish you the best"],
    [/keep in touch/, "keep in touch"],
    [/future opportunities/, "future opportunities"]
  ];

  const acknowledgementOnly = [
    [/application has been received/, "application received"],
    [/we have received your application/, "received application"],
    [/thank you for applying/, "thank you for applying"],
    [/thank you for your application/, "thank you for your application"],
    [/thank you so much for taking the time to apply/, "thank you for taking time to apply"],
    [/currently reviewing/, "currently reviewing"],
    [/taking a look/, "taking a look"],
    [/will get in touch/, "will get in touch"],
    [/details needed/, "details needed"]
  ];

  const positiveSignals = [
    [/schedule a call/, "schedule a call"],
    [/book a time/, "book a time"],
    [/available for an interview/, "available for interview"],
    [/invite you to interview/, "invite to interview"],
    [/invited to interview/, "invited to interview"],
    [/you have been shortlisted/, "shortlisted"],
    [/shortlisted for/, "shortlisted"],
    [/next stage of the interview/, "next stage interview"],
    [/job offer/, "job offer"],
    [/offer of employment/, "offer of employment"],
    [/congratulations/, "congratulations"],

    // Safer progression positives.
    // Do NOT use plain "progressing your application" because rejection emails say:
    // "we won't be progressing your application".
    [/(pleased|happy|excited|delighted) to .{0,80}(progress|proceed|move forward|advance)/, "positive progression"],
    [/we would like to (progress|proceed|move forward|advance)/, "would like to progress"],
    [/we'd like to (progress|proceed|move forward|advance)/, "would like to progress"]
  ];

  strongRejection.forEach(p => {
    if (p[0].test(t)) {
      rejection += 5;
      reasons.push(p[1]);
    }
  });

  softRejection.forEach(p => {
    if (p[0].test(t)) {
      rejection += 1;
      reasons.push(p[1]);
    }
  });

  acknowledgementOnly.forEach(p => {
    if (p[0].test(t)) {
      acknowledgement += 2;
    }
  });

  positiveSignals.forEach(p => {
    if (p[0].test(t)) {
      positive += 4;
    }
  });

  // Rejection wins on ties because recruiter emails often contain both:
  // "thank you for applying" + "unfortunately" + "won't progress".
  if (rejection >= 5 && rejection >= positive) {
    return {
      status: "Rejection",
      score: rejection,
      reason: reasons.slice(0, 3).join("; ") || "Strong rejection signal"
    };
  }

  if (positive >= 4 && positive > rejection) {
    return {
      status: "Positive",
      score: positive,
      reason: "Positive progression signal"
    };
  }

  if (jobContext >= 2 && rejection >= 2) {
    return {
      status: "Review",
      score: jobContext + rejection,
      reason: "Job context with soft rejection wording"
    };
  }

  if (acknowledgement >= 2 && rejection === 0) {
    return {
      status: "Review",
      score: acknowledgement,
      reason: "Acknowledgement / needs review"
    };
  }

  if (jobContext >= 2) {
    return {
      status: "Review",
      score: jobContext,
      reason: "Job context only"
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
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<\/p>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&middot;/g, " ")
    .replace(/&zwnj;/g, "")
    .replace(/&#8204;/g, "")
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&rsquo;/g, "'")
    .replace(/&lsquo;/g, "'")
    .replace(/&#8217;/g, "'")
    .replace(/&#8216;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&ldquo;/g, '"')
    .replace(/&rdquo;/g, '"')
    .replace(/&ndash;/g, "-")
    .replace(/&mdash;/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

function extractCompanyName(subject, body, from) {
  const text = subject + "\n" + body;

  let match =
    text.match(/advertised by ([^\n\r.]+)/i) ||
    text.match(/application for .*? at ([^\n\r.]+)/i) ||
    text.match(/application for job .*? at ([^\n\r.]+)/i) ||
    text.match(/position at ([^\n\r.]+)/i) ||
    text.match(/role at ([^\n\r.]+)/i) ||
    text.match(/talent team\s*<([^>]+)>/i);

  if (match && match[1]) {
    return cleanCompanyName_(match[1]);
  }

  const emailDomainMatch = String(from || "").match(/@([^>\s]+)/);
  if (emailDomainMatch && emailDomainMatch[1]) {
    const domain = emailDomainMatch[1]
      .replace(/^mail\./, "")
      .replace(/^hris\./, "")
      .replace(/\.(com|com\.au|org|org\.au|net|net\.au|gov\.au)$/i, "");

    return cleanCompanyName_(domain);
  }

  return cleanCompanyName_(from);
}

function cleanCompanyName_(s) {
  return String(s || "")
    .replace(/<.*?>/g, "")
    .replace(/["']/g, "")
    .replace(/\bcareers\b/gi, "")
    .replace(/\brecruitment\b/gi, "")
    .replace(/\btalent team\b/gi, "")
    .replace(/\btalent acquisition\b/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    if (name === "Run_Log") {
      ensureRunLogHeader_(sheet);
    } else {
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
  }

  return sheet;
}

function ensureRunLogHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Run Date",
      "Gmail Threads",
      "Found",
      "Gmail Labelled",
      "Rows Added",
      "Already In Sheet",
      "Not Rejection"
    ]);
  }
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("No active spreadsheet found. Set SPREADSHEET_ID at the top of the script.");
  }

  return ss;
}

function norm(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/[’‘`´]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

function testTelstraCase() {
  const subject = "About your job application for job Senior Business Analyst - Digital Health & Data";

  const body = `
Hi Roberto

Thank you so much for taking the time to apply for our Senior Business Analyst - Digital Health & Data role. We truly value the effort you put into your application.

After careful review, we’ve decided to move forward with candidates whose experience more closely matches the role requirements. Unfortunately, we won’t be progressing your application this time.

While this may not be the outcome you were hoping for, we’d like to keep in touch.

Warm regards,
The Talent Acquisition Team
Telstra Health
`;

  const result = classifyJobEmail(subject, body);
  Logger.log(JSON.stringify(result));

  if (result.status !== "Rejection") {
    throw new Error("Telstra test failed. Expected Rejection, got: " + result.status);
  }

  Logger.log("Telstra test passed.");
}
