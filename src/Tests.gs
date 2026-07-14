function testFalsePositiveCases() {
  const cases = [
    {
      name: "Recruitment Hive acknowledgement",
      subject: "Your application has been received",
      from: "Recruitment Hive <info@recruitmenthive.com.au>",
      body: "Thank you for your interest in our advertised position. Recruitment Hive’s team will assess all applications against our Client's role requirements. If you are shortlisted, Recruitment Hive will contact you to further discuss your application. If you are not shortlisted, we will update your details in our systems.",
      expected: "Positive"
    },
    {
      name: "Bluefin acknowledgement",
      subject: "Application Acknowledgment – Bluefin Resources P/L",
      from: "Lisa Lohan <hello@v3.idibu.com>",
      body: "Thank you for submitting your CV to Bluefin Resources P/L. We want to assure you that if you have not yet heard from us, we will be in touch promptly should you be shortlisted for the position you applied for. In the event that we are unable to proceed with your application at this time, we would like to retain your details in our database for future opportunities.",
      expected: "Positive"
    },
    {
      name: "DEWR Dips in Earnings work email",
      subject: "Re: FW: Dips in Earnings [SEC=OFFICIAL]",
      from: "Roberto.Voto@dewr.gov.au",
      body: "12 Week Outcomes Remove 1 Fortnight. Pay Start Pay End Period Days % Reduction Basic Rate Casual Income Continuous Income JRRR Rate Reduction outcome is treated as FULL.",
      expected: "Ignore"
    },
    {
      name: "Telstra real rejection",
      subject: "About your job application for job Senior Business Analyst - Digital Health & Data",
      from: "Telstra Health Talent Team <recruitment@hris.telstrahealth.com>",
      body: "Thank you so much for taking the time to apply for our Senior Business Analyst - Digital Health & Data role. After careful review, we’ve decided to move forward with candidates whose experience more closely matches the role requirements. Unfortunately, we won’t be progressing your application this time.",
      expected: "Rejection"
    },
    {
      name: "LinkedIn real rejection",
      subject: "Your application to Business Change Analyst at Collaborate Recruitment",
      from: "LinkedIn <jobs-noreply@linkedin.com>",
      body: "Your update from Collaborate Recruitment Business Change Analyst Collaborate Recruitment · Sutherland, New South Wales, Australia Applied on Jun 22 Thank you for your interest in the Business Change Analyst position at Collaborate Recruitment in Sutherland, New South Wales, Australia. Unfortunately, we will not be moving forward with your application, but we appreciate your time and interest in Collaborate Recruitment.",
      expected: "Rejection"
    }
  ];

  cases.forEach(c => {
    const result = classifyJobEmail(c.subject, c.body, c.from);
    Logger.log(c.name + ": " + result.status + " | " + result.reason);

    if (result.status !== c.expected) {
      throw new Error(c.name + " failed. Expected " + c.expected + ", got " + result.status + " | " + result.reason);
    }
  });

  Logger.log("All classifier tests passed.");
}

function testContentCleaner() {
  const subject = "Your application to Business Change Analyst at Collaborate Recruitment";
  const body = "This email was intended for Rob Voto. Learn why we included this: https://example.com Help: https://example.com © 2026 LinkedIn Corporation. Your application to Business Change Analyst at Collaborate Recruitment ͏ ͏ ͏ Your update from Collaborate Recruitment Business Change Analyst Collaborate Recruitment &middot; Sutherland, New South Wales, Australia Applied on Jun 22 Thank you for your interest in the Business Change Analyst position at Collaborate Recruitment in Sutherland, New South Wales, Australia. Unfortunately, we will not be moving forward with your application, but we appreciate your time and interest in Collaborate Recruitment. Regards, Collaborate Recruitment Top jobs looking for your skills Technical Business Analyst Transport for NSW";
  const cleaned = cleanEmailContentForSheet_(subject, body);

  Logger.log(cleaned);

  if (cleaned.indexOf("Top jobs looking for your skills") >= 0) {
    throw new Error("Cleaner failed: still contains LinkedIn recommendations.");
  }

  if (cleaned.indexOf("https://") >= 0) {
    throw new Error("Cleaner failed: still contains long URLs.");
  }

  if (cleaned.indexOf("Unfortunately, we will not be moving forward") < 0) {
    throw new Error("Cleaner failed: removed the useful rejection text.");
  }

  Logger.log("Content cleaner test passed.");
}

function runSafeTests() {
  testFalsePositiveCases();
  testContentCleaner();
  Logger.log("All safe tests passed.");
}
