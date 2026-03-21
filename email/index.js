/**
 * Email module for Outlook MCP server
 */
const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleDraftEmail = require('./draft');
const handleMarkAsRead = require('./mark-as-read');
const { handleListAttachments, handleDownloadAttachment, handleDownloadAttachments } = require('./attachments');

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "List recent emails",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Folder name (default: inbox)"
        },
        count: {
          type: "number",
          description: "Number of emails (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search emails by criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text"
        },
        folder: {
          type: "string",
          description: "Folder (default: inbox)"
        },
        from: {
          type: "string",
          description: "Sender address or name"
        },
        to: {
          type: "string",
          description: "Recipient address or name"
        },
        subject: {
          type: "string",
          description: "Subject filter"
        },
        hasAttachments: {
          type: "boolean",
          description: "Has attachments filter"
        },
        unreadOnly: {
          type: "boolean",
          description: "Unread only filter"
        },
        count: {
          type: "number",
          description: "Max results (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Read email content by ID",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "Email ID"
        },
        includeRawHtml: {
          type: "boolean",
          description: "Include raw HTML (unsafe, debug only)"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Send a new email",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Recipient(s), comma-separated"
        },
        cc: {
          type: "string",
          description: "CC recipient(s), comma-separated"
        },
        bcc: {
          type: "string",
          description: "BCC recipient(s), comma-separated"
        },
        subject: {
          type: "string",
          description: "Subject"
        },
        body: {
          type: "string",
          description: "Body (plain text or HTML)"
        },
        isHtml: {
          type: "boolean",
          description: "Send as HTML (auto-detects if omitted)"
        },
        importance: {
          type: "string",
          description: "Importance level",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Save to sent items"
        }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "draft-email",
    description: "Create an email draft",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Recipient(s), comma-separated"
        },
        cc: {
          type: "string",
          description: "CC recipient(s), comma-separated"
        },
        bcc: {
          type: "string",
          description: "BCC recipient(s), comma-separated"
        },
        subject: {
          type: "string",
          description: "Subject"
        },
        body: {
          type: "string",
          description: "Body (plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Importance level",
          enum: ["normal", "high", "low"]
        }
      },
      required: []
    },
    handler: handleDraftEmail
  },
  {
    name: "mark-as-read",
    description: "Mark email as read/unread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "Email ID"
        },
        isRead: {
          type: "boolean",
          description: "Read (true) or unread (false), default: true"
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "list-email-attachments",
    description: "List email attachments",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "Email ID"
        }
      },
      required: ["id"]
    },
    handler: handleListAttachments
  },
  {
    name: "download-email-attachment",
    description: "Download a specific email attachment",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "Email ID"
        },
        attachmentId: {
          type: "string",
          description: "Attachment ID (from list-email-attachments)"
        }
      },
      required: ["emailId", "attachmentId"]
    },
    handler: handleDownloadAttachment
  },
  {
    name: "download-email-attachments",
    description: "Download all attachments from an email",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "Email ID"
        },
        saveToPath: {
          type: "string",
          description: "Local directory to save files (optional)"
        }
      },
      required: ["emailId"]
    },
    handler: handleDownloadAttachments
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleDraftEmail,
  handleMarkAsRead,
  handleListAttachments,
  handleDownloadAttachment,
  handleDownloadAttachments
};
