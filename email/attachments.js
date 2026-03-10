/**
 * Email attachment functionality
 * 
 * Provides listing and downloading of email attachments
 * via the Microsoft Graph API.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List attachments for a specific email
 * @param {object} args - Tool arguments
 * @param {string} args.id - Email ID (required)
 * @returns {object} - MCP response
 */
async function handleListAttachments(args) {
  const emailId = args.id;

  if (!emailId) {
    return {
      content: [{
        type: "text",
        text: "Email ID is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/messages/${encodeURIComponent(emailId)}/attachments`;
    const queryParams = {
      $select: 'id,name,contentType,size,isInline'
    };

    try {
      const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

      const attachments = response.value || [];

      if (attachments.length === 0) {
        return {
          content: [{
            type: "text",
            text: "This email has no attachments."
          }]
        };
      }

      const formatted = attachments.map((att, index) => {
        const size = formatSize(att.size);
        const inline = att.isInline ? ' (inline)' : '';
        return `${index + 1}. ${att.name} (${att.contentType}, ${size})${inline}\n   ID: ${att.id}`;
      }).join('\n');

      return {
        content: [{
          type: "text",
          text: `Found ${attachments.length} attachment(s):\n\n${formatted}`
        }]
      };
    } catch (error) {
      console.error(`Error listing attachments: ${error.message}`);

      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return {
          content: [{
            type: "text",
            text: "The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID."
          }]
        };
      }

      return {
        content: [{
          type: "text",
          text: `Failed to list attachments: ${error.message}`
        }]
      };
    }
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error accessing attachments: ${error.message}`
      }]
    };
  }
}

/**
 * Download a specific attachment from an email
 * @param {object} args - Tool arguments
 * @param {string} args.emailId - Email ID (required)
 * @param {string} args.attachmentId - Attachment ID (required)
 * @returns {object} - MCP response with attachment content
 */
async function handleDownloadAttachment(args) {
  const emailId = args.emailId;
  const attachmentId = args.attachmentId;

  if (!emailId || !attachmentId) {
    return {
      content: [{
        type: "text",
        text: "Both emailId and attachmentId are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/messages/${encodeURIComponent(emailId)}/attachments/${encodeURIComponent(attachmentId)}`;

    try {
      const attachment = await callGraphAPI(accessToken, 'GET', endpoint);

      if (!attachment) {
        return {
          content: [{
            type: "text",
            text: "Attachment not found."
          }]
        };
      }

      const name = attachment.name || 'unknown';
      const contentType = attachment.contentType || 'application/octet-stream';
      const size = formatSize(attachment.size);

      // Handle different attachment types
      // #microsoft.graph.fileAttachment has contentBytes (base64)
      // #microsoft.graph.itemAttachment is an Outlook item (event, message, etc.)
      // #microsoft.graph.referenceAttachment is a link to a file (e.g., OneDrive)
      const odataType = attachment['@odata.type'];

      if (odataType === '#microsoft.graph.itemAttachment') {
        return {
          content: [{
            type: "text",
            text: `Attachment: ${name} (${contentType}, ${size})\nType: Outlook item attachment (event, message, or contact)\n\nItem attachments cannot be downloaded as files. Use the Microsoft Graph API to access the item directly.`
          }]
        };
      }

      if (odataType === '#microsoft.graph.referenceAttachment') {
        return {
          content: [{
            type: "text",
            text: `Attachment: ${name} (${contentType}, ${size})\nType: Reference attachment (cloud file link)\n\nThis is a link to a file stored in the cloud (e.g., OneDrive or SharePoint). Use the OneDrive tools to access it.`
          }]
        };
      }

      // File attachment - has base64 contentBytes
      if (attachment.contentBytes) {
        const isTextType = contentType.startsWith('text/') ||
          contentType === 'application/json' ||
          contentType === 'application/xml' ||
          contentType === 'application/javascript';

        if (isTextType) {
          // Decode base64 to text for text-based files
          const textContent = Buffer.from(attachment.contentBytes, 'base64').toString('utf-8');
          return {
            content: [{
              type: "text",
              text: `Attachment: ${name} (${contentType}, ${size})\n\n--- Content ---\n${textContent}\n--- End ---`
            }]
          };
        }

        // Binary file - return base64 content
        return {
          content: [{
            type: "text",
            text: `Attachment: ${name} (${contentType}, ${size})\n\nBase64-encoded content (binary file):\n${attachment.contentBytes}`
          }]
        };
      }

      return {
        content: [{
          type: "text",
          text: `Attachment: ${name} (${contentType}, ${size})\n\nNo downloadable content available for this attachment.`
        }]
      };
    } catch (error) {
      console.error(`Error downloading attachment: ${error.message}`);

      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return {
          content: [{
            type: "text",
            text: "The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID."
          }]
        };
      }

      return {
        content: [{
          type: "text",
          text: `Failed to download attachment: ${error.message}`
        }]
      };
    }
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error accessing attachment: ${error.message}`
      }]
    };
  }
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = {
  handleListAttachments,
  handleDownloadAttachment
};
