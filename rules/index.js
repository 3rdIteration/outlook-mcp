/**
 * Email rules management module for Outlook MCP server
 */
const { handleListRules, getInboxRules } = require('./list');
const handleCreateRule = require('./create');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { sanitizeMetadata } = require('../utils/metadata-sanitizer');

/**
 * Edit rule sequence handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleEditRuleSequence(args) {
  const { ruleName, sequence } = args;
  
  if (!ruleName) {
    return {
      content: [{ 
        type: "text", 
        text: "Rule name is required. Please specify the exact name of an existing rule."
      }]
    };
  }
  
  if (!sequence || isNaN(sequence) || sequence < 1) {
    return {
      content: [{ 
        type: "text", 
        text: "A positive sequence number is required. Lower numbers run first (higher priority)."
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Get all rules
    const rules = await getInboxRules(accessToken);
    
    // Find the rule by name
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) {
      return {
        content: [{ 
          type: "text", 
          text: `Rule with name "${sanitizeMetadata(ruleName)}" not found.`
        }]
      };
    }
    
    // Update the rule sequence
    const updateResult = await callGraphAPI(
      accessToken,
      'PATCH',
      `me/mailFolders/inbox/messageRules/${rule.id}`,
      {
        sequence: sequence
      }
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Successfully updated the sequence of rule "${sanitizeMetadata(ruleName)}" to ${sequence}.`
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
        text: `Error updating rule sequence: ${error.message}`
      }]
    };
  }
}

// Rules management tool definitions
const rulesTools = [
  {
    name: "list-rules",
    description: "List inbox rules",
    inputSchema: {
      type: "object",
      properties: {
        includeDetails: {
          type: "boolean",
          description: "Include rule conditions and actions"
        }
      },
      required: []
    },
    handler: handleListRules
  },
  {
    name: "create-rule",
    description: "Create an inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        name: {
          type: "string",
          description: "Rule name"
        },
        fromAddresses: {
          type: "string",
          description: "Sender addresses, comma-separated"
        },
        containsSubject: {
          type: "string",
          description: "Subject contains text"
        },
        hasAttachments: {
          type: "boolean",
          description: "Match emails with attachments"
        },
        moveToFolder: {
          type: "string",
          description: "Target folder name"
        },
        markAsRead: {
          type: "boolean", 
          description: "Mark as read"
        },
        isEnabled: {
          type: "boolean",
          description: "Enable rule (default: true)"
        },
        sequence: {
          type: "number",
          description: "Execution order (default: 100)"
        }
      },
      required: ["name"]
    },
    handler: handleCreateRule
  },
  {
    name: "edit-rule-sequence",
    description: "Change rule execution order",
    inputSchema: {
      type: "object",
      properties: {
        ruleName: {
          type: "string",
          description: "Rule name"
        },
        sequence: {
          type: "number",
          description: "New sequence number"
        }
      },
      required: ["ruleName", "sequence"]
    },
    handler: handleEditRuleSequence
  }
];

// Read-only tools
const rulesReadTools = rulesTools.filter(t =>
  ['list-rules'].includes(t.name)
);

// Write tools
const rulesWriteTools = rulesTools.filter(t =>
  ['create-rule', 'edit-rule-sequence'].includes(t.name)
);

module.exports = {
  rulesTools,
  rulesReadTools,
  rulesWriteTools,
  handleListRules,
  handleCreateRule,
  handleEditRuleSequence
};
