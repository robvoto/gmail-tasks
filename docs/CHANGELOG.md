# Changelog

## 2026-07-07

Initial repository setup.

Added:

- README.
- Setup notes.
- Test cases for missed rejection emails.
- Main Apps Script source file.

Known fixes included:

- Use `getFullEmailText(msg)` to combine plain text and stripped HTML body.
- Archive after labelling with `thread.moveToArchive()`.
- Archive already-labelled rejection threads if still in Inbox.
- Add missed patterns for Telstra Health and Recruitment Hive / AFP rejection emails.
