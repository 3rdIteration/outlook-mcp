/**
 * OneDrive module for Outlook MCP server
 */
const handleListFiles = require('./list');
const handleSearchFiles = require('./search');
const handleDownload = require('./download');
const handleUpload = require('./upload');
const handleUploadLarge = require('./upload-large');
const handleShare = require('./share');
const { handleCreateFolder, handleDeleteItem } = require('./folder');

// OneDrive tool definitions
const onedriveTools = [
  {
    name: "onedrive-list",
    description: "List OneDrive files and folders",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Path (default: root)"
        },
        count: {
          type: "number",
          description: "Max items (default: 25, max: 50)"
        }
      },
      required: []
    },
    handler: handleListFiles
  },
  {
    name: "onedrive-search",
    description: "Search OneDrive files",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query"
        },
        count: {
          type: "number",
          description: "Max results (default: 25, max: 50)"
        }
      },
      required: ["query"]
    },
    handler: handleSearchFiles
  },
  {
    name: "onedrive-download",
    description: "Get file download URL (provide itemId or path)",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "Item ID"
        },
        path: {
          type: "string",
          description: "File path (alternative to itemId)"
        }
      },
      required: []
    },
    handler: handleDownload
  },
  {
    name: "onedrive-upload",
    description: "Upload file to OneDrive (< 4MB)",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Destination path with filename"
        },
        content: {
          type: "string",
          description: "File content"
        },
        conflictBehavior: {
          type: "string",
          description: "If exists: rename, replace, or fail",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path", "content"]
    },
    handler: handleUpload
  },
  {
    name: "onedrive-upload-large",
    description: "Upload large file to OneDrive (> 4MB, chunked)",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Destination path with filename"
        },
        content: {
          type: "string",
          description: "File content"
        },
        conflictBehavior: {
          type: "string",
          description: "If exists: rename, replace, or fail",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path", "content"]
    },
    handler: handleUploadLarge
  },
  {
    name: "onedrive-share",
    description: "Create a sharing link for OneDrive item",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "Item ID"
        },
        path: {
          type: "string",
          description: "Item path (alternative to itemId)"
        },
        type: {
          type: "string",
          description: "Link type",
          enum: ["view", "edit", "embed"]
        },
        scope: {
          type: "string",
          description: "Link scope",
          enum: ["anonymous", "organization"]
        }
      },
      required: []
    },
    handler: handleShare
  },
  {
    name: "onedrive-create-folder",
    description: "Create a OneDrive folder",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Parent path (default: root)"
        },
        name: {
          type: "string",
          description: "Folder name"
        }
      },
      required: ["name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "onedrive-delete",
    description: "Delete a OneDrive item",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "Item ID"
        },
        path: {
          type: "string",
          description: "Item path (alternative to itemId)"
        }
      },
      required: []
    },
    handler: handleDeleteItem
  }
];

module.exports = {
  onedriveTools,
  handleListFiles,
  handleSearchFiles,
  handleDownload,
  handleUpload,
  handleUploadLarge,
  handleShare,
  handleCreateFolder,
  handleDeleteItem
};
