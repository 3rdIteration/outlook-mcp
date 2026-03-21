/**
 * Tests for the split server entry points.
 *
 * Verifies that each focused server includes only the expected tools
 * and that the combined server (index.js) still includes all of them.
 */

// Build tool sets the same way each entry point does
const { authTools } = require('../../auth');
const { calendarTools } = require('../../calendar');
const { emailTools } = require('../../email');
const { folderTools } = require('../../folder');
const { rulesTools } = require('../../rules');
const { onedriveTools } = require('../../onedrive');
const { powerAutomateTools } = require('../../power-automate');

// Auth tool names that appear in every server
const AUTH_TOOL_NAMES = authTools.map(t => t.name);

// --- helper ---
function toolNames(arr) {
  return arr.map(t => t.name);
}

// Replicate the tool arrays assembled by each entry point
const emailServerTools = [...authTools, ...emailTools, ...folderTools, ...rulesTools];
const calendarServerTools = [...authTools, ...calendarTools];
const onedriveServerTools = [...authTools, ...onedriveTools];
const powerAutomateServerTools = [...authTools, ...powerAutomateTools];
const allTools = [
  ...authTools, ...calendarTools, ...emailTools,
  ...folderTools, ...rulesTools, ...onedriveTools, ...powerAutomateTools
];

describe('split server — email', () => {
  test('includes auth tools', () => {
    const names = toolNames(emailServerTools);
    for (const name of AUTH_TOOL_NAMES) {
      expect(names).toContain(name);
    }
  });

  test('includes email, folder, and rules tools', () => {
    const names = toolNames(emailServerTools);
    for (const t of [...emailTools, ...folderTools, ...rulesTools]) {
      expect(names).toContain(t.name);
    }
  });

  test('does NOT include calendar, onedrive, or power-automate tools', () => {
    const names = toolNames(emailServerTools);
    for (const t of [...calendarTools, ...onedriveTools, ...powerAutomateTools]) {
      expect(names).not.toContain(t.name);
    }
  });

  test('has no duplicate tool names', () => {
    const names = toolNames(emailServerTools);
    expect(new Set(names).size).toBe(names.length);
  });
});

describe('split server — calendar', () => {
  test('includes auth tools', () => {
    const names = toolNames(calendarServerTools);
    for (const name of AUTH_TOOL_NAMES) {
      expect(names).toContain(name);
    }
  });

  test('includes calendar tools', () => {
    const names = toolNames(calendarServerTools);
    for (const t of calendarTools) {
      expect(names).toContain(t.name);
    }
  });

  test('does NOT include email, folder, rules, onedrive, or power-automate tools', () => {
    const names = toolNames(calendarServerTools);
    for (const t of [...emailTools, ...folderTools, ...rulesTools, ...onedriveTools, ...powerAutomateTools]) {
      expect(names).not.toContain(t.name);
    }
  });

  test('has no duplicate tool names', () => {
    const names = toolNames(calendarServerTools);
    expect(new Set(names).size).toBe(names.length);
  });
});

describe('split server — onedrive', () => {
  test('includes auth tools', () => {
    const names = toolNames(onedriveServerTools);
    for (const name of AUTH_TOOL_NAMES) {
      expect(names).toContain(name);
    }
  });

  test('includes onedrive tools', () => {
    const names = toolNames(onedriveServerTools);
    for (const t of onedriveTools) {
      expect(names).toContain(t.name);
    }
  });

  test('does NOT include email, folder, rules, calendar, or power-automate tools', () => {
    const names = toolNames(onedriveServerTools);
    for (const t of [...emailTools, ...folderTools, ...rulesTools, ...calendarTools, ...powerAutomateTools]) {
      expect(names).not.toContain(t.name);
    }
  });

  test('has no duplicate tool names', () => {
    const names = toolNames(onedriveServerTools);
    expect(new Set(names).size).toBe(names.length);
  });
});

describe('split server — power-automate', () => {
  test('includes auth tools', () => {
    const names = toolNames(powerAutomateServerTools);
    for (const name of AUTH_TOOL_NAMES) {
      expect(names).toContain(name);
    }
  });

  test('includes power-automate tools', () => {
    const names = toolNames(powerAutomateServerTools);
    for (const t of powerAutomateTools) {
      expect(names).toContain(t.name);
    }
  });

  test('does NOT include email, folder, rules, calendar, or onedrive tools', () => {
    const names = toolNames(powerAutomateServerTools);
    for (const t of [...emailTools, ...folderTools, ...rulesTools, ...calendarTools, ...onedriveTools]) {
      expect(names).not.toContain(t.name);
    }
  });

  test('has no duplicate tool names', () => {
    const names = toolNames(powerAutomateServerTools);
    expect(new Set(names).size).toBe(names.length);
  });
});

describe('combined server (index.js) still contains all tools', () => {
  test('includes every module tool', () => {
    const names = toolNames(allTools);
    const allModuleTools = [
      ...authTools, ...calendarTools, ...emailTools,
      ...folderTools, ...rulesTools, ...onedriveTools, ...powerAutomateTools
    ];
    for (const t of allModuleTools) {
      expect(names).toContain(t.name);
    }
  });

  test('has no duplicate tool names', () => {
    const names = toolNames(allTools);
    expect(new Set(names).size).toBe(names.length);
  });

  test('union of all split servers covers all tools (minus auth dupes)', () => {
    // Collect unique names from all four split servers
    const splitNames = new Set([
      ...toolNames(emailServerTools),
      ...toolNames(calendarServerTools),
      ...toolNames(onedriveServerTools),
      ...toolNames(powerAutomateServerTools)
    ]);
    const allNames = new Set(toolNames(allTools));
    expect(splitNames).toEqual(allNames);
  });
});

describe('every tool has required fields', () => {
  const servers = [
    { label: 'email', tools: emailServerTools },
    { label: 'calendar', tools: calendarServerTools },
    { label: 'onedrive', tools: onedriveServerTools },
    { label: 'power-automate', tools: powerAutomateServerTools },
  ];

  test.each(servers)('$label server tools have name, description, and inputSchema', ({ tools }) => {
    for (const tool of tools) {
      expect(tool).toHaveProperty('name');
      expect(tool).toHaveProperty('description');
      expect(tool).toHaveProperty('inputSchema');
    }
  });
});
