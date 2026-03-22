/**
 * List rules functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { sanitizeMetadata, wrapWithBoundary, wrapField, generateBoundaryToken } = require('../utils/metadata-sanitizer');

/**
 * List rules handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListRules(args) {
  const includeDetails = args.includeDetails === true;
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Get all inbox rules
    const rules = await getInboxRules(accessToken);
    
    // Format the rules based on detail level
    const formattedRules = formatRulesList(rules, includeDetails);
    
    return {
      content: [{ 
        type: "text", 
        text: formattedRules
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
        text: `Error listing rules: ${error.message}`
      }]
    };
  }
}

/**
 * Get all inbox rules
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of rule objects
 */
async function getInboxRules(accessToken) {
  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders/inbox/messageRules',
      null
    );
    
    return response.value || [];
  } catch (error) {
    console.error(`Error getting inbox rules: ${error.message}`);
    throw error;
  }
}

/**
 * Format rules list for display
 * @param {Array} rules - Array of rule objects
 * @param {boolean} includeDetails - Whether to include detailed conditions and actions
 * @returns {string} - Formatted rules list
 */
function formatRulesList(rules, includeDetails) {
  if (!rules || rules.length === 0) {
    return "No inbox rules found.\n\nTip: You can create rules using the 'create-rule' tool. Rules are processed in order of their sequence number (lower numbers are processed first).";
  }
  
  // Sort rules by sequence to show execution order
  const sortedRules = [...rules].sort((a, b) => {
    return (a.sequence || 9999) - (b.sequence || 9999);
  });
  
  // Generate a shared boundary token for JSON payload and outer markers
  const boundaryToken = generateBoundaryToken();
  
  // Build structured JSON with field-level wrapping
  const ruleItems = sortedRules.map((rule) => {
    const item = {
      displayName: wrapField(sanitizeMetadata(rule.displayName), boundaryToken),
      isEnabled: rule.isEnabled,
      sequence: rule.sequence || null
    };
    
    if (includeDetails) {
      // Add conditions
      const conditions = {};
      if (rule.conditions?.fromAddresses?.length > 0) {
        conditions.fromAddresses = rule.conditions.fromAddresses.map(addr =>
          wrapField(sanitizeMetadata(addr.emailAddress.address), boundaryToken)
        );
      }
      if (rule.conditions?.subjectContains?.length > 0) {
        conditions.subjectContains = rule.conditions.subjectContains.map(s =>
          wrapField(sanitizeMetadata(s), boundaryToken)
        );
      }
      if (rule.conditions?.bodyContains?.length > 0) {
        conditions.bodyContains = rule.conditions.bodyContains.map(s =>
          wrapField(sanitizeMetadata(s), boundaryToken)
        );
      }
      if (rule.conditions?.hasAttachment === true) {
        conditions.hasAttachment = true;
      }
      if (rule.conditions?.importance) {
        conditions.importance = rule.conditions.importance;
      }
      if (Object.keys(conditions).length > 0) {
        item.conditions = conditions;
      }
      
      // Add actions
      const actions = {};
      if (rule.actions?.moveToFolder) {
        actions.moveToFolder = rule.actions.moveToFolder;
      }
      if (rule.actions?.copyToFolder) {
        actions.copyToFolder = rule.actions.copyToFolder;
      }
      if (rule.actions?.markAsRead === true) {
        actions.markAsRead = true;
      }
      if (rule.actions?.markImportance) {
        actions.markImportance = rule.actions.markImportance;
      }
      if (rule.actions?.forwardTo?.length > 0) {
        actions.forwardTo = rule.actions.forwardTo.map(r =>
          wrapField(sanitizeMetadata(r.emailAddress.address), boundaryToken)
        );
      }
      if (rule.actions?.delete === true) {
        actions.delete = true;
      }
      if (Object.keys(actions).length > 0) {
        item.actions = actions;
      }
    }
    
    return item;
  });
  
  const payload = {
    _boundary: boundaryToken,
    rules: ruleItems
  };
  
  const rulesJson = JSON.stringify(payload, null, 2);
  
  const tip = includeDetails
    ? "\n\nRules are processed in order of their sequence number. You can change rule order using the 'edit-rule-sequence' tool."
    : "\n\nTip: Use 'list-rules with includeDetails=true' to see more information about each rule.";
  
  return `Found ${rules.length} inbox rules (sorted by execution order):\n\n${wrapWithBoundary(rulesJson, 'INBOX RULES', boundaryToken)}${tip}`;
}

module.exports = {
  handleListRules,
  getInboxRules
};
