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
 * @param {object} params - request params (may include cursor)
 * @param {number} pageSize - override page size for testing (defaults to config)
 */
function simulateToolsList(params, pageSize) {
  if (pageSize === undefined) pageSize = config.TOOLS_PAGE_SIZE;

  // No pagination when pageSize <= 0
  if (pageSize <= 0) {
    return {
      tools: TOOLS.map(tool => ({
        name: tool.name,
        description: tool.description,
        inputSchema: tool.inputSchema
      }))
    };
  }

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

describe('tools/list — default (no pagination, TOOLS_PAGE_SIZE=0)', () => {
  test('returns all tools in one response with no nextCursor', () => {
    const result = simulateToolsList({}, 0);
    expect(result.tools.length).toBe(TOOLS.length);
    expect(result.nextCursor).toBeUndefined();
  });

  test('every tool has name, description, and inputSchema', () => {
    const result = simulateToolsList({}, 0);
    for (const tool of result.tools) {
      expect(tool).toHaveProperty('name');
      expect(tool).toHaveProperty('description');
      expect(tool).toHaveProperty('inputSchema');
    }
  });

  test('no duplicate tool names', () => {
    const result = simulateToolsList({}, 0);
    const names = result.tools.map(t => t.name);
    expect(new Set(names).size).toBe(names.length);
  });
});

describe('tools/list — paginated mode (TOOLS_PAGE_SIZE=10)', () => {
  const PAGE_SIZE = 10;

  test('first page returns PAGE_SIZE tools with a nextCursor', () => {
    const result = simulateToolsList({}, PAGE_SIZE);
    expect(result.tools.length).toBe(PAGE_SIZE);
    expect(result.nextCursor).toBeDefined();
  });

  test('following cursors returns all tools across pages', () => {
    const allToolNames = [];
    let cursor = undefined;
    let pages = 0;

    while (true) {
      pages++;
      const result = simulateToolsList({ cursor }, PAGE_SIZE);
      expect(result.tools.length).toBeGreaterThan(0);
      expect(result.tools.length).toBeLessThanOrEqual(PAGE_SIZE);

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
    expect(pages).toBe(Math.ceil(TOOLS.length / PAGE_SIZE));
    // Verify no duplicates
    expect(new Set(allToolNames).size).toBe(allToolNames.length);
  });

  test('last page has no nextCursor', () => {
    let cursor = undefined;
    let result;

    while (true) {
      result = simulateToolsList({ cursor }, PAGE_SIZE);
      if (!result.nextCursor) break;
      cursor = result.nextCursor;
    }

    expect(result.nextCursor).toBeUndefined();
    expect(result.tools.length).toBeGreaterThan(0);
  });

  test('invalid cursor falls back to first page', () => {
    const result = simulateToolsList({ cursor: 'not-valid-base64!!!' }, PAGE_SIZE);
    expect(result.tools.length).toBe(PAGE_SIZE);
    const firstPage = simulateToolsList({}, PAGE_SIZE);
    expect(result.tools.map(t => t.name)).toEqual(firstPage.tools.map(t => t.name));
  });

  test('out-of-range cursor falls back to first page', () => {
    const badCursor = Buffer.from('9999').toString('base64');
    const result = simulateToolsList({ cursor: badCursor }, PAGE_SIZE);
    const firstPage = simulateToolsList({}, PAGE_SIZE);
    expect(result.tools.map(t => t.name)).toEqual(firstPage.tools.map(t => t.name));
  });

  test('negative index cursor falls back to first page', () => {
    const badCursor = Buffer.from('-5').toString('base64');
    const result = simulateToolsList({ cursor: badCursor }, PAGE_SIZE);
    const firstPage = simulateToolsList({}, PAGE_SIZE);
    expect(result.tools.map(t => t.name)).toEqual(firstPage.tools.map(t => t.name));
  });

  test('each page response is smaller than the full unpaginated response', () => {
    const fullSize = JSON.stringify({
      tools: TOOLS.map(t => ({ name: t.name, description: t.description, inputSchema: t.inputSchema }))
    }).length;

    let cursor = undefined;
    while (true) {
      const result = simulateToolsList({ cursor }, PAGE_SIZE);
      const pageBytes = JSON.stringify(result).length;
      expect(pageBytes).toBeLessThan(fullSize);
      if (!result.nextCursor) break;
      cursor = result.nextCursor;
    }
  });
});

describe('tools/list — large page size returns all tools', () => {
  test('page size >= tool count returns everything with no nextCursor', () => {
    const result = simulateToolsList({}, 50);
    expect(result.tools.length).toBe(TOOLS.length);
    expect(result.nextCursor).toBeUndefined();
  });
});

describe('tool definition budget', () => {
  test('total tool definitions fit within small context window budget', () => {
    // Tool definitions should stay under ~12KB / ~3000 tokens to remain
    // functional on clients with 4096-token context windows.
    // At ~4 chars/token, 12000 bytes ≈ 3000 tokens.
    const fullResponse = JSON.stringify({
      tools: TOOLS.map(t => ({ name: t.name, description: t.description, inputSchema: t.inputSchema }))
    });
    expect(fullResponse.length).toBeLessThan(12000);
  });
});
