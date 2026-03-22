/**
 * List emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');
const { sanitizeMetadata, wrapWithBoundary } = require('../utils/metadata-sanitizer');

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
    
    // Build structured results with sanitized fields
    const emails = response.value.map((email) => {
      const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
      return {
        id: email.id,
        subject: sanitizeMetadata(email.subject),
        from: {
          name: sanitizeMetadata(sender.name),
          address: sanitizeMetadata(sender.address)
        },
        to: (email.toRecipients || []).map(r => ({
          name: sanitizeMetadata(r.emailAddress?.name || ''),
          address: sanitizeMetadata(r.emailAddress?.address || '')
        })),
        receivedDateTime: email.receivedDateTime,
        isRead: email.isRead,
        hasAttachments: email.hasAttachments,
        importance: email.importance,
        bodyPreview: sanitizeMetadata(email.bodyPreview)
      };
    });
    
    const emailList = JSON.stringify(emails, null, 2);
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} emails in ${folder}:\n\n${wrapWithBoundary(emailList, 'EMAIL LIST')}`
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
