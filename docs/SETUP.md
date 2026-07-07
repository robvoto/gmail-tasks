# Setup

## Runtime

This project runs in Google Apps Script.

Known Apps Script project URL:

https://script.google.com/u/0/home/projects/1VZAH1tKB11fm1feYFwl8gVuIOrwbWO-93MtuxO-NCpbaxW3H34__I_w_/edit

## Apps Script files seen so far

- `Code.gs`
- `One off.gs`
- `Find and label.gs`
- `Check an Email Body.gs`
- `Generate json.gs`

## Important Apps Script rule

Apps Script compiles all `.gs` files together.

Do not duplicate global constants across files.

Bad:

```javascript
const LABEL_NAME = "Job_Rejections";
```

in multiple `.gs` files.

Use globals in one place only:

```javascript
var LABEL_NAME = "Job_Rejections";
var SEARCH_WINDOW = "newer_than:10d";
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

Also archive already-labelled threads if they are still in Inbox.

## Triggers

Create time-driven triggers in Apps Script:

1. `findAndLabelNewRejections`
2. `extractJobRejections`

Suggested cadence: daily.

If Apps Script shows unsaved changes, save first. Triggers run saved code only.

## Debug order when an email is missed

1. Check if the subject appears in `checkedSubjects` log.
2. If not, Gmail search did not scan it. Check `SEARCH_WINDOW`, trigger, or query.
3. If yes, classifier missed the wording. Add the narrowest safe regex.
4. If labelled but still in Inbox, check `thread.moveToArchive()`.
