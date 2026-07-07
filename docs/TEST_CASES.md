# Test cases

Use this file when a rejection email is missed.

Rule: capture the exact wording from the email body and add the narrowest safe regex to `strongRejection`.

## Telstra Health

Subject:

```text
About your job application for job Senior Business Analyst - Digital Health & Data
```

Body wording:

```text
After careful review, we’ve decided to move forward with candidates whose experience more closely matches the role requirements. Unfortunately, we won’t be progressing your application this time.
```

Expected:

```text
Rejection
```

Required patterns:

```javascript
/decided to move forward with candidates.{0,160}(closely matches|more closely matches|matches).{0,160}(role requirements|requirements)/,
/won’t be progressing your application/,
/won't be progressing your application/,
```

## Recruitment Hive / AFP

Subject:

```text
Senior Business Analyst contract @ Australian Federal Police
```

Body wording:

```text
Unfortunately due to the number of applicants who have expressed interest for this position, and the limitations on how many I can present to the client, I am unable to progress with your application.
```

Expected:

```text
Rejection
```

Required patterns:

```javascript
/unable to progress with your application/,
/unfortunately.{0,200}(number of applicants|volume of applications|high number of applications)/,
/due to.{0,120}(number of applicants|volume of applications|high number of applications)/,
```

## PwC

Body wording:

```text
profile does not meet the requirements for the role
```

Expected:

```text
Rejection
```

Required patterns:

```javascript
/profile.{0,120}does not meet.{0,120}requirements/,
/does not meet.{0,120}(requirements|criteria|selection criteria|role requirements)/,
/determined.{0,120}profile.{0,120}does not meet/,
```

## HTML body only case

Some rejection text exists in `msg.getBody()` but not `msg.getPlainBody()`.

Expected:

```text
Rejection when using getFullEmailText(msg)
```

Required helper:

```javascript
function getFullEmailText(msg) {
  return [
    msg.getPlainBody(),
    stripHtml(msg.getBody())
  ].join(" ");
}
```
