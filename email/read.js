/**
 * Read email functionality
 *
 * Security: HTML emails are sanitized to remove hidden content that could
 * be used for prompt injection attacks. Only visible text is extracted.
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { processHtmlEmail, sanitizeHtmlToText } = require('../utils/html-sanitizer');
const { sanitizeMetadata, wrapWithBoundary, wrapField, generateBoundaryToken } = require('../utils/metadata-sanitizer');

/**
 * Read email handler
 * @param {object} args - Tool arguments
 * @param {string} args.id - Email ID (required)
 * @param {boolean} args.includeRawHtml - If true, include raw HTML (unsafe, for debugging only)
 * @returns {object} - MCP response
 */
async function handleReadEmail(args) {
  const emailId = args.id;
  const includeRawHtml = args.includeRawHtml === true;

  if (!emailId) {
    return {
      content: [{
        type: "text",
        text: "Email ID is required."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Make API call to get email details
    const endpoint = `me/messages/${encodeURIComponent(emailId)}`;
    const queryParams = {
      $select: config.EMAIL_DETAIL_FIELDS
    };

    try {
      const email = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

      if (!email) {
        return {
          content: [
            {
              type: "text",
              text: `Email with ID ${emailId} not found.`
            }
          ]
        };
      }

      // Generate a shared boundary token for JSON payload and outer markers
      const boundaryToken = generateBoundaryToken();

      // Format sender, recipients, etc. (sanitize metadata to prevent prompt injection)
      const senderName = email.from ? sanitizeMetadata(email.from.emailAddress.name) : 'Unknown';
      const senderAddress = email.from ? sanitizeMetadata(email.from.emailAddress.address) : 'unknown';
      const toRecipients = email.toRecipients ? email.toRecipients.map(r => ({
        name: wrapField(sanitizeMetadata(r.emailAddress.name), boundaryToken),
        address: wrapField(sanitizeMetadata(r.emailAddress.address), boundaryToken)
      })) : [];
      const ccRecipients = email.ccRecipients && email.ccRecipients.length > 0 ? email.ccRecipients.map(r => ({
        name: wrapField(sanitizeMetadata(r.emailAddress.name), boundaryToken),
        address: wrapField(sanitizeMetadata(r.emailAddress.address), boundaryToken)
      })) : [];
      const bccRecipients = email.bccRecipients && email.bccRecipients.length > 0 ? email.bccRecipients.map(r => ({
        name: wrapField(sanitizeMetadata(r.emailAddress.name), boundaryToken),
        address: wrapField(sanitizeMetadata(r.emailAddress.address), boundaryToken)
      })) : [];
      const date = new Date(email.receivedDateTime).toLocaleString();

      // Extract and sanitize body content
      let body = '';
      let bodyNote = '';

      if (email.body) {
        if (email.body.contentType === 'html') {
          // Use secure HTML sanitizer to extract visible text only
          // This prevents prompt injection via hidden HTML content
          body = processHtmlEmail(email.body.content, {
            addBoundary: true,
            metadata: {
              from: email.from?.emailAddress?.address || 'unknown',
              subject: email.subject,
              date: date
            }
          });
          bodyNote = '[HTML email - sanitized for security, hidden content removed]';
        } else {
          // Plain text - still wrap with boundary for safety
          body = processHtmlEmail(email.body.content, {
            addBoundary: true,
            metadata: {
              from: email.from?.emailAddress?.address || 'unknown',
              subject: email.subject,
              date: date
            }
          });
        }
      } else {
        body = email.bodyPreview || 'No content';
      }

      // Build structured JSON response with field-level wrapping
      const emailData = {
        _boundary: boundaryToken,
        id: wrapField(emailId, boundaryToken),
        subject: wrapField(sanitizeMetadata(email.subject), boundaryToken),
        from: {
          name: wrapField(senderName, boundaryToken),
          address: wrapField(senderAddress, boundaryToken)
        },
        to: toRecipients,
        cc: ccRecipients,
        bcc: bccRecipients,
        date: wrapField(date, boundaryToken),
        importance: wrapField(email.importance || 'normal', boundaryToken),
        hasAttachments: email.hasAttachments,
        body: body
      };

      if (bodyNote) {
        emailData.bodyNote = bodyNote;
      }

      const emailJson = JSON.stringify(emailData, null, 2);

      // Optionally include raw HTML for debugging (not recommended for normal use)
      let rawHtmlSection = '';
      if (includeRawHtml && email.body?.contentType === 'html') {
        rawHtmlSection = `\n\n--- RAW HTML (UNSAFE - FOR DEBUGGING ONLY) ---\n${email.body.content}\n--- END RAW HTML ---`;
      }

      return {
        content: [
          {
            type: "text",
            text: wrapWithBoundary(emailJson, 'EMAIL', boundaryToken) + rawHtmlSection
          }
        ]
      };
    } catch (error) {
      console.error(`Error reading email: ${error.message}`);
      
      // Improved error handling with more specific messages
      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return {
          content: [
            {
              type: "text",
              text: `The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID.`
            }
          ]
        };
      } else {
        return {
          content: [
            {
              type: "text",
              text: `Failed to read email: ${error.message}`
            }
          ]
        };
      }
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
        text: `Error accessing email: ${error.message}`
      }]
    };
  }
}

module.exports = handleReadEmail;
