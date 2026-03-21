/**
 * Email attachment functionality
 * 
 * Provides listing and downloading of email attachments
 * via the Microsoft Graph API.
 */
const fs = require('fs');
const path = require('path');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { sanitizeMetadata, wrapWithBoundary } = require('../utils/metadata-sanitizer');

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
        return `${index + 1}. ${sanitizeMetadata(att.name)} (${sanitizeMetadata(att.contentType)}, ${size})${inline}\n   ID: ${att.id}`;
      }).join('\n');

      return {
        content: [{
          type: "text",
          text: `Found ${attachments.length} attachment(s):\n\n${wrapWithBoundary(formatted, 'ATTACHMENTS')}`
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

      const name = sanitizeMetadata(attachment.name || 'unknown');
      const contentType = sanitizeMetadata(attachment.contentType || 'application/octet-stream');
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
            text: wrapWithBoundary(`Attachment: ${name} (${contentType}, ${size})\nType: Outlook item attachment (event, message, or contact)\n\nItem attachments cannot be downloaded as files. Use the Microsoft Graph API to access the item directly.`, 'ATTACHMENT')
          }]
        };
      }

      if (odataType === '#microsoft.graph.referenceAttachment') {
        return {
          content: [{
            type: "text",
            text: wrapWithBoundary(`Attachment: ${name} (${contentType}, ${size})\nType: Reference attachment (cloud file link)\n\nThis is a link to a file stored in the cloud (e.g., OneDrive or SharePoint). Use the OneDrive tools to access it.`, 'ATTACHMENT')
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
              text: wrapWithBoundary(`Attachment: ${name} (${contentType}, ${size})\n\n${textContent}`, 'ATTACHMENT CONTENT')
            }]
          };
        }

        // Binary file - return base64 content
        return {
          content: [{
            type: "text",
            text: wrapWithBoundary(`Attachment: ${name} (${contentType}, ${size})\n\nBase64-encoded content (binary file):\n${attachment.contentBytes}`, 'ATTACHMENT CONTENT')
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
 * Download all attachments from an email
 * @param {object} args - Tool arguments
 * @param {string} args.emailId - Email ID (required)
 * @param {string} args.saveToPath - Optional local directory path to save attachments to
 * @returns {object} - MCP response with attachment details and content
 */
async function handleDownloadAttachments(args) {
  const emailId = args.emailId;
  const saveToPath = args.saveToPath;

  if (!emailId) {
    return {
      content: [{
        type: "text",
        text: "Email ID is required."
      }]
    };
  }

  // Validate saveToPath if provided
  if (saveToPath) {
    try {
      // Resolve to absolute path and check it exists
      const resolvedPath = path.resolve(saveToPath);
      if (!fs.existsSync(resolvedPath)) {
        return {
          content: [{
            type: "text",
            text: `The directory "${saveToPath}" does not exist. Please create it first or specify a valid path.`
          }]
        };
      }
      const stats = fs.statSync(resolvedPath);
      if (!stats.isDirectory()) {
        return {
          content: [{
            type: "text",
            text: `"${saveToPath}" is not a directory. Please specify a directory path.`
          }]
        };
      }
    } catch (error) {
      return {
        content: [{
          type: "text",
          text: `Error accessing path "${saveToPath}": ${error.message}`
        }]
      };
    }
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/messages/${encodeURIComponent(emailId)}/attachments`;

    try {
      const response = await callGraphAPI(accessToken, 'GET', endpoint);

      const attachments = response.value || [];

      if (attachments.length === 0) {
        return {
          content: [{
            type: "text",
            text: "This email has no attachments."
          }]
        };
      }

      // Filter to file attachments only (those with contentBytes)
      const fileAttachments = attachments.filter(
        att => att['@odata.type'] === '#microsoft.graph.fileAttachment' && att.contentBytes
      );
      const otherAttachments = attachments.filter(
        att => att['@odata.type'] !== '#microsoft.graph.fileAttachment' || !att.contentBytes
      );

      if (fileAttachments.length === 0) {
        const otherInfo = otherAttachments.map(att => {
          const type = att['@odata.type'] === '#microsoft.graph.itemAttachment'
            ? 'Outlook item' : att['@odata.type'] === '#microsoft.graph.referenceAttachment'
            ? 'cloud file link' : 'non-downloadable';
          return `- ${sanitizeMetadata(att.name || 'unknown')} (${type})`;
        }).join('\n');

        return {
          content: [{
            type: "text",
            text: `This email has ${attachments.length} attachment(s), but none are downloadable file attachments:\n${otherInfo}`
          }]
        };
      }

      const results = [];
      const errors = [];
      const savedFiles = [];

      for (const att of fileAttachments) {
        const name = sanitizeMetadata(att.name || 'unknown');
        const contentType = sanitizeMetadata(att.contentType || 'application/octet-stream');
        const size = att.size || 0;

        // Save to disk if saveToPath is specified
        if (saveToPath) {
          try {
            const resolvedDir = path.resolve(saveToPath);
            // Sanitize filename: strip path components, null bytes, and control chars
            let safeName = path.basename(name);
            safeName = safeName.replace(/[\x00-\x1f]/g, ''); // Remove control characters including null bytes
            if (!safeName || safeName === '.' || safeName === '..') {
              safeName = `attachment_${att.id || Date.now()}`;
            }
            const filePath = path.join(resolvedDir, safeName);
            const fileBuffer = Buffer.from(att.contentBytes, 'base64');
            fs.writeFileSync(filePath, fileBuffer);
            savedFiles.push({ name: safeName, path: filePath, size });
          } catch (writeError) {
            errors.push(`- ${name} (${contentType}, ${formatSize(size)}): Failed to save - ${writeError.message}`);
            continue;
          }
        }

        results.push({
          name,
          contentType,
          size,
          contentBytes: att.contentBytes
        });
      }

      // Build response text
      let responseText = '';

      if (saveToPath && savedFiles.length > 0) {
        const resolvedDir = path.resolve(saveToPath);
        responseText += `Saved ${savedFiles.length} attachment(s) to ${resolvedDir}:\n\n`;
        responseText += savedFiles.map((f, i) =>
          `${i + 1}. ${f.name} (${formatSize(f.size)}) → ${f.path}`
        ).join('\n');
      } else {
        responseText += `Downloaded ${results.length} attachment(s):\n\n`;
        responseText += results.map((att, i) =>
          `${i + 1}. ${att.name} (${att.contentType}, ${formatSize(att.size)})\n   Base64 content length: ${att.contentBytes.length} chars`
        ).join('\n');
      }

      // Report any file save errors
      if (errors.length > 0) {
        responseText += `\n\nFailed to save ${errors.length} attachment(s):\n${errors.join('\n')}`;
      }

      // Note about non-file attachments
      if (otherAttachments.length > 0) {
        responseText += `\n\nNote: ${otherAttachments.length} non-downloadable attachment(s) were skipped (item or reference attachments).`;
      }

      // Build content array - include attachment data as structured text when not saving
      const content = [{ type: "text", text: wrapWithBoundary(responseText, 'ATTACHMENTS') }];

      if (!saveToPath) {
        for (const att of results) {
          content.push({
            type: "text",
            text: wrapWithBoundary(`Attachment: ${att.name} (${att.contentType}, ${formatSize(att.size)})\n${att.contentBytes}`, 'ATTACHMENT CONTENT')
          });
        }
      }

      return { content };
    } catch (error) {
      console.error(`Error downloading attachments: ${error.message}`);

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
          text: `Failed to download attachments: ${error.message}`
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
  handleDownloadAttachment,
  handleDownloadAttachments
};
