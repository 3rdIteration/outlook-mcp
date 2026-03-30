/**
 * List emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');
const { sanitizeMetadata, wrapWithBoundary, wrapField, generateBoundaryToken } = require('../utils/metadata-sanitizer');

/**
 * List emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Resolve the folder path
    const endpoint = await resolveFolderPath(accessToken, folder);
    
    // Add query parameters
    const queryParams = {
      $top: Math.min(50, requestedCount), // Use 50 per page for efficiency
      $orderby: 'receivedDateTime desc',
      $select: config.EMAIL_SELECT_FIELDS
    };
    
    // Make API call with pagination support
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, queryParams, requestedCount);
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: `No emails found in ${folder}.`
        }]
      };
    }
    
    // Generate a shared boundary token for JSON payload and outer markers
    const boundaryToken = generateBoundaryToken();
    
    // Build structured results with sanitized and field-wrapped values
    const emails = response.value.map((email) => {
      const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
      return {
        emailId: wrapField(email.id, boundaryToken),
        subject: wrapField(sanitizeMetadata(email.subject, config.MAX_SUBJECT_LENGTH), boundaryToken),
        from: {
          name: wrapField(sanitizeMetadata(sender.name, config.MAX_SENDER_LENGTH), boundaryToken),
          address: wrapField(sanitizeMetadata(sender.address, config.MAX_SENDER_LENGTH), boundaryToken)
        },
        to: (email.toRecipients || []).map(r => ({
          name: wrapField(sanitizeMetadata(r.emailAddress?.name || 'Unknown', config.MAX_SENDER_LENGTH), boundaryToken),
          address: wrapField(sanitizeMetadata(r.emailAddress?.address || 'unknown', config.MAX_SENDER_LENGTH), boundaryToken)
        })),
        receivedDateTime: wrapField(email.receivedDateTime, boundaryToken),
        isRead: email.isRead,
        hasAttachments: email.hasAttachments,
        importance: wrapField(email.importance, boundaryToken),
        bodyPreview: wrapField(sanitizeMetadata(email.bodyPreview, config.MAX_BODY_PREVIEW_LENGTH), boundaryToken)
      };
    });
    
    // Wrap emails array in object with boundary token to prevent JSON spoofing
    const payload = {
      _boundary: boundaryToken,
      emails
    };
    
    const emailList = JSON.stringify(payload, null, 2);
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} emails in ${folder}:\n\n${wrapWithBoundary(emailList, 'EMAIL LIST', boundaryToken)}`
      }]
    };
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
        text: `Error listing emails: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEmails;
