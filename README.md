# Gmail Tasks

Google Apps Script automation for Rob's Gmail workflows.

Current focus: job-rejection processing and export.

## What it does

- Searches recent Gmail messages.
- Detects job rejection emails.
- Applies the Gmail label `Job_Rejections`.
- Archives labelled threads so they leave Inbox.
- Extracts labelled rejection emails into a Google Sheet.

## Runtime

Runtime is Google Apps Script.

GitHub is the reviewable source of truth. Apps Script is the place where the script actually runs.

Local tooling uses `@google/clasp` to inspect and push the Apps Script project. Use Node 20 or newer.

## Repository map

- `src/Code.gs` main Apps Script code.
- `src/Tests.gs` safe in-project test functions.
- `src/GenerateJson.gs` JSON export logic.
- `src/appsscript.json` Apps Script manifest.
- `docs/INDEX.md` documentation entry point.
- `docs/SETUP.md` setup and trigger notes.
- `docs/TEST_CASES.md` known rejection examples and expected behaviour.
- `docs/CHANGELOG.md` change history.
- `package.json` local clasp tooling only; it is not the application runtime.

## Local setup

```bash
npm ci
npx clasp status
```

The clasp project identity is local-only in `.clasp.json`; credentials and local clasp config must not be committed.

## Validation and deployment

Safe functional checks run inside Apps Script via `runSafeTests()` in `src/Tests.gs`.

Pushing code to the Apps Script project is an explicit deployment action:

```bash
npx clasp push
```

Do not treat a Git push as an Apps Script deployment.

## Drive context doc

Use this context doc before changing the script:

https://docs.google.com/document/d/1tVwZtTlfoaajPm0XKCSfH93ivoS_Y41QWFE5qf4DdqY/edit?usp=drivesdk

Rule: do not guess Apps Script state. Read the current GitHub file or pasted Apps Script code before changing anything.
