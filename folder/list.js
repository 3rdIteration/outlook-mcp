/**
 * List folders functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { sanitizeMetadata, wrapWithBoundary, wrapField, generateBoundaryToken } = require('../utils/metadata-sanitizer');

/**
 * List folders handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFolders(args) {
  const includeItemCounts = args.includeItemCounts === true;
  const includeChildren = args.includeChildren === true;
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Get all mail folders
    const folders = await getAllFoldersHierarchy(accessToken, includeItemCounts);
    
    // Generate a shared boundary token for JSON payload and outer markers
    const boundaryToken = generateBoundaryToken();
    
    // If including children, format as hierarchy
    if (includeChildren) {
      const hierarchyData = buildFolderHierarchy(folders, includeItemCounts, boundaryToken);
      const payload = {
        _boundary: boundaryToken,
        folders: hierarchyData
      };
      const folderJson = JSON.stringify(payload, null, 2);
      return {
        content: [{ 
          type: "text", 
          text: `Folder Hierarchy:\n\n${wrapWithBoundary(folderJson, 'FOLDER LIST', boundaryToken)}`
        }]
      };
    } else {
      // Otherwise, format as flat list
      const flatData = buildFolderFlatList(folders, includeItemCounts, boundaryToken);
      const payload = {
        _boundary: boundaryToken,
        folders: flatData
      };
      const folderJson = JSON.stringify(payload, null, 2);
      return {
        content: [{ 
          type: "text", 
          text: `Found ${folders.length} folders:\n\n${wrapWithBoundary(folderJson, 'FOLDER LIST', boundaryToken)}`
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
        text: `Error listing folders: ${error.message}`
      }]
    };
  }
}

/**
 * Get all mail folders with hierarchy information
 * @param {string} accessToken - Access token
 * @param {boolean} includeItemCounts - Include item counts in response
 * @returns {Promise<Array>} - Array of folder objects with hierarchy
 */
async function getAllFoldersHierarchy(accessToken, includeItemCounts) {
  try {
    // Determine select fields based on whether to include counts
    const selectFields = includeItemCounts
      ? 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      : 'id,displayName,parentFolderId,childFolderCount';
    
    // Get all mail folders
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { 
        $top: 100,
        $select: selectFields
      }
    );
    
    if (!response.value) {
      return [];
    }
    
    // Get child folders for folders with children
    const foldersWithChildren = response.value.filter(f => f.childFolderCount > 0);
    
    const childFolderPromises = foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI(
          accessToken,
          'GET',
          `me/mailFolders/${folder.id}/childFolders`,
          null,
          { $select: selectFields }
        );
        
        // Add parent folder info to each child
        const childFolders = childResponse.value || [];
        childFolders.forEach(child => {
          child.parentFolder = folder.displayName;
        });
        
        return childFolders;
      } catch (error) {
        console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
        return [];
      }
    });
    
    const childFolders = await Promise.all(childFolderPromises);
    const allChildFolders = childFolders.flat();
    
    // Add top-level flag to parent folders
    const topLevelFolders = response.value.map(folder => ({
      ...folder,
      isTopLevel: true
    }));
    
    // Combine all folders
    return [...topLevelFolders, ...allChildFolders];
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    throw error;
  }
}

/**
 * Build flat list of folder objects with field wrapping
 * @param {Array} folders - Array of folder objects
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @param {string} boundaryToken - Boundary token for field wrapping
 * @returns {Array} - Array of folder data objects
 */
function buildFolderFlatList(folders, includeItemCounts, boundaryToken) {
  if (!folders || folders.length === 0) {
    return [];
  }
  
  // Sort folders alphabetically, with well-known folders first
  const wellKnownFolderNames = ['Inbox', 'Drafts', 'Sent Items', 'Deleted Items', 'Junk Email', 'Archive'];
  
  const sortedFolders = [...folders].sort((a, b) => {
    const aIsWellKnown = wellKnownFolderNames.includes(a.displayName);
    const bIsWellKnown = wellKnownFolderNames.includes(b.displayName);
    
    if (aIsWellKnown && !bIsWellKnown) return -1;
    if (!aIsWellKnown && bIsWellKnown) return 1;
    
    if (aIsWellKnown && bIsWellKnown) {
      return wellKnownFolderNames.indexOf(a.displayName) - wellKnownFolderNames.indexOf(b.displayName);
    }
    
    return a.displayName.localeCompare(b.displayName);
  });
  
  return sortedFolders.map(folder => {
    const item = {
      id: wrapField(folder.id, boundaryToken),
      displayName: wrapField(sanitizeMetadata(folder.displayName), boundaryToken)
    };
    
    if (folder.parentFolder) {
      item.parentFolder = wrapField(sanitizeMetadata(folder.parentFolder), boundaryToken);
    }
    
    if (includeItemCounts) {
      item.totalItemCount = folder.totalItemCount || 0;
      item.unreadItemCount = folder.unreadItemCount || 0;
    }
    
    return item;
  });
}

/**
 * Build folder hierarchy with field wrapping
 * @param {Array} folders - Array of folder objects
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @param {string} boundaryToken - Boundary token for field wrapping
 * @returns {Array} - Array of folder data objects with children
 */
function buildFolderHierarchy(folders, includeItemCounts, boundaryToken) {
  if (!folders || folders.length === 0) {
    return [];
  }
  
  // Build folder hierarchy
  const folderMap = new Map();
  const rootFolderIds = [];
  
  // First pass: create map of all folders
  folders.forEach(folder => {
    folderMap.set(folder.id, {
      ...folder,
      childIds: []
    });
    
    if (folder.isTopLevel) {
      rootFolderIds.push(folder.id);
    }
  });
  
  // Second pass: build hierarchy
  folders.forEach(folder => {
    if (!folder.isTopLevel && folder.parentFolderId) {
      const parent = folderMap.get(folder.parentFolderId);
      if (parent) {
        parent.childIds.push(folder.id);
      } else {
        rootFolderIds.push(folder.id);
      }
    }
  });
  
  // Build tree recursively
  function buildSubtree(folderId) {
    const folder = folderMap.get(folderId);
    if (!folder) return null;
    
    const item = {
      id: wrapField(folder.id, boundaryToken),
      displayName: wrapField(sanitizeMetadata(folder.displayName), boundaryToken)
    };
    
    if (includeItemCounts) {
      item.totalItemCount = folder.totalItemCount || 0;
      item.unreadItemCount = folder.unreadItemCount || 0;
    }
    
    if (folder.childIds.length > 0) {
      item.children = folder.childIds
        .map(childId => buildSubtree(childId))
        .filter(Boolean);
    }
    
    return item;
  }
  
  return rootFolderIds.map(id => buildSubtree(id)).filter(Boolean);
}

module.exports = handleListFolders;
