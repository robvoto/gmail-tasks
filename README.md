# Gmail Tasks

Source of truth for Rob's Gmail automation tasks.

Current focus: Gmail job rejection automation.

## What it does

- Searches recent Gmail messages.
- Detects job rejection emails.
- Applies the Gmail label `Job_Rejections`.
- Archives labelled threads so they leave Inbox.
- Extracts labelled rejection emails into a Google Sheet.

## Runtime

Runtime is Google Apps Script.

GitHub is the reviewable source of truth. Apps Script is the place where the script actually runs.

## Current source files

- `src/Code.gs` main Apps Script code.
- `docs/SETUP.md` setup and trigger notes.
- `docs/TEST_CASES.md` known rejection examples and expected behaviour.
- `docs/CHANGELOG.md` change history.

## Drive context doc

Use this context doc before changing the script:

https://docs.google.com/document/d/1tVwZtTlfoaajPm0XKCSfH93ivoS_Y41QWFE5qf4DdqY/edit?usp=drivesdk

Rule: do not guess Apps Script state. Read the current GitHub file or pasted Apps Script code before changing anything.
