const config = require('../../config');

// Build the combined TOOLS array the same way index.js does
const { authTools } = require('../../auth');
const { calendarTools } = require('../../calendar');
const { emailTools } = require('../../email');
const { folderTools } = require('../../folder');
const { rulesTools } = require('../../rules');
const { onedriveTools } = require('../../onedrive');
const { powerAutomateTools } = require('../../power-automate');

const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...onedriveTools,
  ...powerAutomateTools
];

/**
 * Simulates the tools/list handler logic from index.js.
 * Kept in sync with the actual implementation to test pagination behavior.
 */
function simulateToolsList(params) {
  const pageSize = config.TOOLS_PAGE_SIZE;
  const cursor = params?.cursor;
  let startIndex = 0;

  if (cursor) {
    try {
      startIndex = parseInt(Buffer.from(cursor, 'base64').toString('utf8'), 10);
    } catch {
      startIndex = 0;
    }
    if (isNaN(startIndex) || startIndex < 0 || startIndex >= TOOLS.length) {
      startIndex = 0;
    }
  }

  const endIndex = Math.min(startIndex + pageSize, TOOLS.length);
  const pageTools = TOOLS.slice(startIndex, endIndex);

  const result = {
    tools: pageTools.map(tool => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema
    }))
  };

  if (endIndex < TOOLS.length) {
    result.nextCursor = Buffer.from(String(endIndex)).toString('base64');
  }

  return result;
}

describe('tools/list pagination', () => {
  test('first page returns TOOLS_PAGE_SIZE tools with a nextCursor', () => {
    const result = simulateToolsList({});
    expect(result.tools.length).toBe(config.TOOLS_PAGE_SIZE);
    expect(result.nextCursor).toBeDefined();
  });

  test('following cursors returns all tools across pages', () => {
    const allToolNames = [];
    let cursor = undefined;
    let pages = 0;

    while (true) {
      pages++;
      const result = simulateToolsList({ cursor });
      expect(result.tools.length).toBeGreaterThan(0);
      expect(result.tools.length).toBeLessThanOrEqual(config.TOOLS_PAGE_SIZE);

      for (const tool of result.tools) {
        expect(tool).toHaveProperty('name');
        expect(tool).toHaveProperty('description');
        expect(tool).toHaveProperty('inputSchema');
        allToolNames.push(tool.name);
      }

      if (!result.nextCursor) break;
      cursor = result.nextCursor;
    }

    expect(allToolNames.length).toBe(TOOLS.length);
    expect(pages).toBe(Math.ceil(TOOLS.length / config.TOOLS_PAGE_SIZE));
    // Verify no duplicates
    expect(new Set(allToolNames).size).toBe(allToolNames.length);
  });

  test('last page has no nextCursor', () => {
    let cursor = undefined;
    let result;

    // Walk to the last page
    while (true) {
      result = simulateToolsList({ cursor });
      if (!result.nextCursor) break;
      cursor = result.nextCursor;
    }

    expect(result.nextCursor).toBeUndefined();
    expect(result.tools.length).toBeGreaterThan(0);
  });

  test('invalid cursor falls back to first page', () => {
    const result = simulateToolsList({ cursor: 'not-valid-base64!!!' });
    expect(result.tools.length).toBe(config.TOOLS_PAGE_SIZE);
    // Should return the same tools as the first page
    const firstPage = simulateToolsList({});
    expect(result.tools.map(t => t.name)).toEqual(firstPage.tools.map(t => t.name));
  });

  test('out-of-range cursor falls back to first page', () => {
    const badCursor = Buffer.from('9999').toString('base64');
    const result = simulateToolsList({ cursor: badCursor });
    const firstPage = simulateToolsList({});
    expect(result.tools.map(t => t.name)).toEqual(firstPage.tools.map(t => t.name));
  });

  test('negative index cursor falls back to first page', () => {
    const badCursor = Buffer.from('-5').toString('base64');
    const result = simulateToolsList({ cursor: badCursor });
    const firstPage = simulateToolsList({});
    expect(result.tools.map(t => t.name)).toEqual(firstPage.tools.map(t => t.name));
  });

  test('each page response is smaller than the full unpaginated response', () => {
    const fullSize = JSON.stringify({
      tools: TOOLS.map(t => ({ name: t.name, description: t.description, inputSchema: t.inputSchema }))
    }).length;

    let cursor = undefined;
    while (true) {
      const result = simulateToolsList({ cursor });
      const pageSize = JSON.stringify(result).length;
      expect(pageSize).toBeLessThan(fullSize);
      if (!result.nextCursor) break;
      cursor = result.nextCursor;
    }
  });
});
