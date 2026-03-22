/**
 * OneDrive search files functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { sanitizeMetadata, wrapWithBoundary, wrapField, generateBoundaryToken } = require('../utils/metadata-sanitizer');

/**
 * Search files handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchFiles(args) {
  const query = args.query;
  const count = args.count || 25;

  if (!query) {
    return {
      content: [{
        type: "text",
        text: "Search query is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Use the search endpoint
    const endpoint = `me/drive/search(q='${encodeURIComponent(query)}')`;

    const queryParams = {
      $top: Math.min(50, count),
      $select: config.ONEDRIVE_SELECT_FIELDS
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No files found matching "${query}".`
        }]
      };
    }

    // Generate a shared boundary token for JSON payload and outer markers
    const boundaryToken = generateBoundaryToken();
    
    // Format results as structured JSON with field-level wrapping
    const items = response.value.map((item) => ({
      id: wrapField(item.id, boundaryToken),
      name: wrapField(sanitizeMetadata(item.name), boundaryToken),
      type: item.folder ? 'folder' : 'file',
      size: item.size ? formatSize(item.size) : '',
      path: item.parentReference?.path?.replace('/drive/root:', '') || '/',
      lastModified: new Date(item.lastModifiedDateTime).toLocaleString()
    }));
    
    const payload = {
      _boundary: boundaryToken,
      items
    };
    
    const fileList = JSON.stringify(payload, null, 2);

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} items matching "${sanitizeMetadata(query)}":\n\n${wrapWithBoundary(fileList, 'SEARCH RESULTS', boundaryToken)}`
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
        text: `Error searching files: ${error.message}`
      }]
    };
  }
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = handleSearchFiles;
