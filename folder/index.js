/**
 * Folder management module for Outlook MCP server
 */
const handleListFolders = require('./list');
const handleCreateFolder = require('./create');
const handleMoveEmails = require('./move');

// Folder management tool definitions
const folderTools = [
  {
    name: "list-folders",
    description: "List mail folders",
    inputSchema: {
      type: "object",
      properties: {
        includeItemCounts: {
          type: "boolean",
          description: "Include total and unread counts"
        },
        includeChildren: {
          type: "boolean",
          description: "Include child folders"
        }
      },
      required: []
    },
    handler: handleListFolders
  },
  {
    name: "create-folder",
    description: "Create a mail folder",
    inputSchema: {
      type: "object",
      properties: {
        name: {
          type: "string",
          description: "Folder name"
        },
        parentFolder: {
          type: "string",
          description: "Parent folder (default: root)"
        }
      },
      required: ["name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "move-emails",
    description: "Move emails between folders",
    inputSchema: {
      type: "object",
      properties: {
        emailIds: {
          type: "string",
          description: "Comma-separated email IDs (from list-emails or search-emails)"
        },
        targetFolder: {
          type: "string",
          description: "Target folder name"
        },
        sourceFolder: {
          type: "string",
          description: "Source folder (default: inbox)"
        }
      },
      required: ["emailIds", "targetFolder"]
    },
    handler: handleMoveEmails
  }
];

// Read-only tools
const folderReadTools = folderTools.filter(t =>
  ['list-folders'].includes(t.name)
);

// Write tools
const folderWriteTools = folderTools.filter(t =>
  ['create-folder', 'move-emails'].includes(t.name)
);

module.exports = {
  folderTools,
  folderReadTools,
  folderWriteTools,
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails
};
