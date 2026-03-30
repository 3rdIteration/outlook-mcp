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
const { sanitizeMetadata, wrapWithBoundary, wrapField, generateBoundaryToken } = require('../utils/metadata-sanitizer');

/**
 * List attachments for a specific email
 * @param {object} args - Tool arguments
 * @param {string} args.emailId - Email ID (required)
 * @returns {object} - MCP response
 */
async function handleListAttachments(args) {
  const emailId = args.emailId;

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

      const boundaryToken = generateBoundaryToken();

      const formatted = attachments.map((att) => ({
        id: wrapField(att.id, boundaryToken),
        name: wrapField(sanitizeMetadata(att.name), boundaryToken),
        contentType: wrapField(sanitizeMetadata(att.contentType), boundaryToken),
        size: wrapField(formatSize(att.size), boundaryToken),
        isInline: att.isInline || false
      }));

      const payload = {
        _boundary: boundaryToken,
        attachments: formatted
      };

      const attachmentList = JSON.stringify(payload, null, 2);

      return {
        content: [{
          type: "text",
          text: `Found ${attachments.length} attachment(s):\n\n${wrapWithBoundary(attachmentList, 'ATTACHMENTS', boundaryToken)}`
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

      const boundaryToken = generateBoundaryToken();
      const name = wrapField(sanitizeMetadata(attachment.name || 'unknown'), boundaryToken);
      const contentType = wrapField(sanitizeMetadata(attachment.contentType || 'application/octet-stream'), boundaryToken);
      const size = wrapField(formatSize(attachment.size), boundaryToken);

      // Handle different attachment types
      // #microsoft.graph.fileAttachment has contentBytes (base64)
      // #microsoft.graph.itemAttachment is an Outlook item (event, message, etc.)
      // #microsoft.graph.referenceAttachment is a link to a file (e.g., OneDrive)
      const odataType = attachment['@odata.type'];

      if (odataType === '#microsoft.graph.itemAttachment') {
        const payload = {
          _boundary: boundaryToken,
          name,
          contentType,
          size,
          type: 'Outlook item attachment (event, message, or contact)',
          note: 'Item attachments cannot be downloaded as files. Use the Microsoft Graph API to access the item directly.'
        };
        return {
          content: [{
            type: "text",
            text: wrapWithBoundary(JSON.stringify(payload, null, 2), 'ATTACHMENT', boundaryToken)
          }]
        };
      }

      if (odataType === '#microsoft.graph.referenceAttachment') {
        const payload = {
          _boundary: boundaryToken,
          name,
          contentType,
          size,
          type: 'Reference attachment (cloud file link)',
          note: 'This is a link to a file stored in the cloud (e.g., OneDrive or SharePoint). Use the OneDrive tools to access it.'
        };
        return {
          content: [{
            type: "text",
            text: wrapWithBoundary(JSON.stringify(payload, null, 2), 'ATTACHMENT', boundaryToken)
          }]
        };
      }

      // File attachment - has base64 contentBytes
      if (attachment.contentBytes) {
        const rawContentType = sanitizeMetadata(attachment.contentType || 'application/octet-stream');
        const isTextType = rawContentType.startsWith('text/') ||
          rawContentType === 'application/json' ||
          rawContentType === 'application/xml' ||
          rawContentType === 'application/javascript';

        if (isTextType) {
          // Decode base64 to text for text-based files
          const textContent = Buffer.from(attachment.contentBytes, 'base64').toString('utf-8');
          const payload = {
            _boundary: boundaryToken,
            name,
            contentType,
            size,
            content: textContent
          };
          return {
            content: [{
              type: "text",
              text: wrapWithBoundary(JSON.stringify(payload, null, 2), 'ATTACHMENT CONTENT', boundaryToken)
            }]
          };
        }

        // Binary file - return base64 content
        const payload = {
          _boundary: boundaryToken,
          name,
          contentType,
          size,
          encoding: 'base64',
          content: attachment.contentBytes
        };
        return {
          content: [{
            type: "text",
            text: wrapWithBoundary(JSON.stringify(payload, null, 2), 'ATTACHMENT CONTENT', boundaryToken)
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

      const boundaryToken = generateBoundaryToken();
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
            errors.push({ name, contentType, size: formatSize(size), error: writeError.message });
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

      // Build response
      const payload = {
        _boundary: boundaryToken
      };

      if (saveToPath && savedFiles.length > 0) {
        const resolvedDir = path.resolve(saveToPath);
        payload.savedTo = resolvedDir;
        payload.attachments = savedFiles.map(f => ({
          name: wrapField(f.name, boundaryToken),
          size: wrapField(formatSize(f.size), boundaryToken),
          path: wrapField(f.path, boundaryToken)
        }));
      } else {
        payload.attachments = results.map(att => ({
          name: wrapField(att.name, boundaryToken),
          contentType: wrapField(att.contentType, boundaryToken),
          size: wrapField(formatSize(att.size), boundaryToken),
          contentBytesLength: att.contentBytes.length
        }));
      }

      // Report any file save errors
      if (errors.length > 0) {
        payload.errors = errors;
      }

      // Note about non-file attachments
      if (otherAttachments.length > 0) {
        payload.skippedCount = otherAttachments.length;
        payload.skippedNote = `${otherAttachments.length} non-downloadable attachment(s) were skipped (item or reference attachments).`;
      }

      const summaryCount = saveToPath && savedFiles.length > 0 ? savedFiles.length : results.length;
      const summaryAction = saveToPath && savedFiles.length > 0 ? 'Saved' : 'Downloaded';
      const responseJson = JSON.stringify(payload, null, 2);

      // Build content array
      const content = [{ type: "text", text: `${summaryAction} ${summaryCount} attachment(s):\n\n${wrapWithBoundary(responseJson, 'ATTACHMENTS', boundaryToken)}` }];

      if (!saveToPath) {
        for (const att of results) {
          const attBoundary = generateBoundaryToken();
          const attPayload = {
            _boundary: attBoundary,
            name: wrapField(att.name, attBoundary),
            contentType: wrapField(att.contentType, attBoundary),
            size: wrapField(formatSize(att.size), attBoundary),
            content: att.contentBytes
          };
          content.push({
            type: "text",
            text: wrapWithBoundary(JSON.stringify(attPayload, null, 2), 'ATTACHMENT CONTENT', attBoundary)
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
