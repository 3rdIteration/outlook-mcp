# agents.md — Security Requirements for AI Coding Agents

This file defines **mandatory security rules** that any AI agent (GitHub Copilot, Claude Code, Cursor, etc.) **MUST** follow when adding or modifying features in this codebase.

This is an MCP server that bridges an LLM to Microsoft 365 APIs. All data returned to the LLM comes from external, untrusted sources (emails, calendar events, file names, rule conditions, etc.). An attacker who controls any of those sources can attempt **prompt injection** — embedding hidden instructions in metadata fields that the LLM might interpret as commands.

---

## Mandatory Security Rules

### Rule 1: Sanitize ALL External Metadata with `sanitizeMetadata()`

**Every** string that originates from an external source (Microsoft Graph API response, Power Automate API response, or user-provided parameters that are echoed back) **MUST** be passed through `sanitizeMetadata()` before being included in any tool response text.

**Import:**
```javascript
const { sanitizeMetadata } = require('../utils/metadata-sanitizer');
```

**What it protects against:**
- Newline injection (`\n`, `\r`) — attacker breaks out of metadata context
- Invisible Unicode (zero-width spaces, direction overrides, soft hyphens)
- Control characters (null bytes, escape sequences)
- Context overflow (truncates at 500 chars)

**Examples of fields that MUST be sanitized:**
- Email subjects, sender names, recipient names/addresses
- Attachment file names and content types
- Calendar event subjects, locations, body previews
- OneDrive file/folder names
- Power Automate flow names, environment names
- Inbox rule names, conditions (fromAddresses, subjectContains, bodyContains)
- Folder display names
- Any user-provided parameter echoed back in a response (subject, name, query, etc.)

**Pattern:**
```javascript
// ❌ WRONG — raw external data in output
text: `Subject: ${email.subject}`

// ✅ CORRECT — sanitized
text: `Subject: ${sanitizeMetadata(email.subject)}`
```

### Rule 2: Wrap Listed External Content with `wrapWithBoundary()`

When a tool returns a **list or block of external content** (email listings, search results, calendar events, file listings, rules, flow data), the formatted content **MUST** be wrapped with `wrapWithBoundary()`. This generates a unique random 32-character hex token that marks the start and end of untrusted content, making it impossible for an attacker to predict or mimic the boundaries.

**Import:**
```javascript
const { sanitizeMetadata, wrapWithBoundary } = require('../utils/metadata-sanitizer');
```

**Pattern:**
```javascript
// ❌ WRONG — no boundary markers
text: `Found ${count} emails:\n\n${emailList}`

// ✅ CORRECT — wrapped with randomized boundaries
text: `Found ${count} emails:\n\n${wrapWithBoundary(emailList, 'EMAIL LIST')}`
```

**Use descriptive labels:** `'EMAIL LIST'`, `'SEARCH RESULTS'`, `'CALENDAR EVENTS'`, `'ONEDRIVE FILES'`, `'ATTACHMENTS'`, `'INBOX RULES'`, `'FLOWS'`, `'ENVIRONMENTS'`, `'FLOW RUNS'`, `'FOLDER LIST'`, `'EMAIL'`, `'ATTACHMENT CONTENT'`.

**When to use `wrapWithBoundary()`:**
- Any handler that returns a formatted list of items from an API
- Email body content (already handled by `processHtmlEmail` with its own boundary)
- Individual attachment content blocks
- Any large block of external data shown to the LLM

**When NOT needed:**
- Simple confirmation messages with only sanitized scalar values (e.g., "Email sent successfully")
- Error messages
- Static instructional text

### Rule 3: Sanitize HTML Email Bodies with `processHtmlEmail()`

Email body content (HTML or plain text) **MUST** be processed through `processHtmlEmail()` from `utils/html-sanitizer.js`. This:
- Strips hidden CSS content (display:none, visibility:hidden, opacity:0)
- Removes script/style/iframe tags
- Strips invisible Unicode characters
- Removes HTML comments (used for hidden prompt injection)
- Converts safe HTML to markdown
- Wraps the result with randomized boundary markers

**Import:**
```javascript
const { processHtmlEmail } = require('../utils/html-sanitizer');
```

**Pattern:**
```javascript
body = processHtmlEmail(email.body.content, {
  addBoundary: true,
  metadata: { from: senderAddress, subject: email.subject, date: date }
});
```

### Rule 4: Never Expose Raw Error Details from APIs

Error messages from Microsoft Graph API or Flow API may contain sensitive information or user-controlled data. Use generic error messages where possible, and sanitize any dynamic content in error output.

```javascript
// ✅ Acceptable
text: `Error listing emails: ${error.message}`

// ❌ Avoid — don't echo full API response bodies
text: `Error: ${JSON.stringify(apiResponse)}`
```

### Rule 5: Token File Security

The token storage file (`~/.outlook-mcp-tokens.json`) contains OAuth tokens. After writing, permissions are set to `0o600` (owner read/write only) on Unix systems. Do not change this behavior or store tokens in less restrictive locations.

---

## Module Coverage Checklist

Every module that returns external data to the LLM must use the appropriate security functions. Current coverage:

| Module | `sanitizeMetadata` | `wrapWithBoundary` | `processHtmlEmail` |
|--------|:-:|:-:|:-:|
| email/read.js | ✅ | ✅ | ✅ |
| email/list.js | ✅ | ✅ | — |
| email/search.js | ✅ | ✅ | — |
| email/attachments.js | ✅ | ✅ | — |
| email/send.js | ✅ | — | — |
| email/draft.js | ✅ | — | — |
| calendar/list.js | ✅ | ✅ | — |
| calendar/create.js | ✅ | — | — |
| folder/list.js | ✅ | ✅ | — |
| folder/create.js | ✅ | — | — |
| folder/move.js | ✅ | — | — |
| onedrive/list.js | ✅ | ✅ | — |
| onedrive/search.js | ✅ | ✅ | — |
| onedrive/download.js | ✅ | — | — |
| onedrive/share.js | ✅ | — | — |
| onedrive/upload.js | ✅ | — | — |
| onedrive/upload-large.js | ✅ | — | — |
| onedrive/folder.js | ✅ | — | — |
| rules/list.js | ✅ | ✅ | — |
| rules/create.js | ✅ | — | — |
| power-automate/list-flows.js | ✅ | ✅ | — |
| power-automate/list-environments.js | ✅ | ✅ | — |
| power-automate/list-runs.js | — | ✅ | — |
| power-automate/run-flow.js | ✅ | — | — |
| power-automate/toggle-flow.js | ✅ | — | — |

---

## Adding a New Tool Handler

When creating a new tool handler that returns data from Microsoft 365 APIs:

1. **Import** the sanitizer at the top of the file:
   ```javascript
   const { sanitizeMetadata, wrapWithBoundary } = require('../utils/metadata-sanitizer');
   ```

2. **Sanitize** every external string before including it in output text:
   ```javascript
   sanitizeMetadata(apiResponse.name)
   sanitizeMetadata(apiResponse.subject)
   ```

3. **Wrap** any list or block of external data:
   ```javascript
   wrapWithBoundary(formattedList, 'DESCRIPTIVE LABEL')
   ```

4. **Add tests** that verify sanitization is applied (check that the output does not contain raw newlines from test inputs).

5. **Update the coverage table** in this file.

---

## Security Utilities Reference

| Function | Module | Purpose |
|----------|--------|---------|
| `sanitizeMetadata(str, maxLength?)` | `utils/metadata-sanitizer.js` | Strip newlines, control chars, invisible Unicode, truncate |
| `wrapWithBoundary(content, label?)` | `utils/metadata-sanitizer.js` | Wrap with random hex boundary markers |
| `processHtmlEmail(html, options?)` | `utils/html-sanitizer.js` | Full HTML-to-safe-text pipeline with boundary wrapping |
| `sanitizeHtmlToText(html)` | `utils/html-sanitizer.js` | Strip hidden HTML content, convert to safe text |
| `removeInvisibleChars(text)` | `utils/html-sanitizer.js` | Remove zero-width and invisible Unicode characters |
