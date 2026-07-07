# Setup

## Runtime

This project runs in Google Apps Script.

Known Apps Script project URL:

https://script.google.com/u/2/home/projects/1VZAH1tKB11fm1feYFwl8gVuIOrwbWO-93MtuxO-NCpbaxW3H34__I_w_/edit

Known Script ID:

```text
1VZAH1tKB11fm1feYFwl8gVuIOrwbWO-93MtuxO-NCpbaxW3H34__I_w_
```

## Source of truth

GitHub repo:

```text
https://github.com/robvoto/gmail-tasks
```

GitHub stores the reviewable source. Apps Script is the live runtime.

## Current repo files

- `src/Code.gs`
- `src/GenerateJson.gs`
- `src/appsscript.json`

Apps Script compiles all files under `src` together.

## Local WSL setup

Local path used during setup:

```bash
~/projects/gmail-tasks/gmail-tasks
```

The repo uses clasp with:

```json
{
  "scriptId": "1VZAH1tKB11fm1feYFwl8gVuIOrwbWO-93MtuxO-NCpbaxW3H34__I_w_",
  "rootDir": "src"
}
```

The `.clasp.json` file is local config. Do not commit secrets. This file only contains script ID and rootDir.

## Apps Script API

Apps Script API has been enabled for Rob's Google account.

If another chat sees this error:

```text
User has not enabled the Apps Script API.
```

Go here and enable Google Apps Script API:

```text
https://script.google.com/home/usersettings
```

Then wait 1 to 3 minutes and retry:

```bash
npx clasp push
```

## Normal deployment flow

When GitHub has new changes and Apps Script must be updated:

```bash
git pull
npx clasp push
```

If clasp asks:

```text
Manifest file has been updated. Do you want to push and overwrite?
```

Answer:

```text
y
```

## Apps Script files seen so far

Original Apps Script project files:

- `Code.gs`
- `One off.gs`
- `Find and label.gs`
- `Check an Email Body.gs`
- `Generate json.gs`

Current GitHub source files map to:

- `src/Code.gs` -> main Gmail rejection scan and export-to-sheet logic
- `src/GenerateJson.gs` -> JSON export logic
- `src/appsscript.json` -> Apps Script manifest required by clasp

## Important Apps Script rule

Apps Script compiles all `.gs` files together.

Do not duplicate global constants or helper functions across files.

Bad:

```javascript
const LABEL_NAME = "Job_Rejections";
```

in multiple `.gs` files.

Use globals in one place only:

```javascript
var LABEL_NAME = "Job_Rejections";
var SEARCH_WINDOW = "newer_than:10d";
var BATCH_SIZE = 100;
```

## Gmail label

The script uses:

```text
Job_Rejections
```

## Archive behaviour

Gmail archive is not a folder move.

Archive means remove from Inbox while keeping labels.

Correct Apps Script behaviour:

```javascript
thread.addLabel(label);
thread.moveToArchive();
```

Known current code archives newly labelled threads. Already-labelled handling should be checked if labelled messages remain in Inbox.

## Main functions

Run this for the normal combined scan:

```javascript
runJobRejectionScan()
```

It calls:

```javascript
findAndLabelNewRejections();
extractJobRejections();
```

Generate JSON export:

```javascript
exportCandidateApplicationHistoryJson()
```

## Triggers

Preferred trigger:

```text
runJobRejectionScan
```

Suggested cadence: daily.

If Apps Script shows unsaved changes, save first. Triggers run saved code only.

## Debug order when an email is missed

1. Check if the subject appears in `checkedSubjects` log.
2. If not, Gmail search did not scan it. Check `SEARCH_WINDOW`, trigger, or query.
3. If yes, classifier missed the wording. Add the narrowest safe regex.
4. If labelled but still in Inbox, check `thread.moveToArchive()`.

## Current state as of 2026-07-07

- Latest user-provided `Code.gs` was committed to GitHub.
- `GenerateJson.gs` was committed separately.
- `src/appsscript.json` was added because `clasp push` requires a manifest.
- User enabled Apps Script API after clasp reported it was disabled.
- Next step is to run `npx clasp push` again from WSL.
