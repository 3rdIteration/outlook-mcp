/**
 * Tests for the split server entry points.
 *
 * Verifies that each focused server includes only the expected tools
 * and that the combined server (index.js) still includes all of them.
 * Also verifies read/write/safewrite sub-splits.
 */

// Build tool sets the same way each entry point does
const { authTools } = require('../../auth');
const { calendarTools, calendarReadTools, calendarWriteTools } = require('../../calendar');
const { emailTools, emailReadTools, emailWriteTools, emailSafeWriteTools } = require('../../email');
const { folderTools, folderReadTools, folderWriteTools } = require('../../folder');
const { rulesTools, rulesReadTools, rulesWriteTools } = require('../../rules');
const { onedriveTools, onedriveReadTools, onedriveWriteTools } = require('../../onedrive');
const { powerAutomateTools, powerAutomateReadTools, powerAutomateWriteTools } = require('../../power-automate');

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

// Read/Write/SafeWrite sub-split server tool arrays
const emailReadServerTools = [...authTools, ...emailReadTools, ...folderReadTools, ...rulesReadTools];
const emailWriteServerTools = [...authTools, ...emailWriteTools, ...folderWriteTools, ...rulesWriteTools];
const emailSafeWriteServerTools = [...authTools, ...emailSafeWriteTools];
const calendarReadServerTools = [...authTools, ...calendarReadTools];
const calendarWriteServerTools = [...authTools, ...calendarWriteTools];
const onedriveReadServerTools = [...authTools, ...onedriveReadTools];
const onedriveWriteServerTools = [...authTools, ...onedriveWriteTools];
const powerAutomateReadServerTools = [...authTools, ...powerAutomateReadTools];
const powerAutomateWriteServerTools = [...authTools, ...powerAutomateWriteTools];

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
    { label: 'email-read', tools: emailReadServerTools },
    { label: 'email-write', tools: emailWriteServerTools },
    { label: 'email-safewrite', tools: emailSafeWriteServerTools },
    { label: 'calendar-read', tools: calendarReadServerTools },
    { label: 'calendar-write', tools: calendarWriteServerTools },
    { label: 'onedrive-read', tools: onedriveReadServerTools },
    { label: 'onedrive-write', tools: onedriveWriteServerTools },
    { label: 'power-automate-read', tools: powerAutomateReadServerTools },
    { label: 'power-automate-write', tools: powerAutomateWriteServerTools },
  ];

  test.each(servers)('$label server tools have name, description, and inputSchema', ({ tools }) => {
    for (const tool of tools) {
      expect(tool).toHaveProperty('name');
      expect(tool).toHaveProperty('description');
      expect(tool).toHaveProperty('inputSchema');
    }
  });

  test.each(servers)('$label server tools all have callable handler functions', ({ tools }) => {
    for (const tool of tools) {
      expect(typeof tool.handler).toBe('function');
    }
  });
});

// ===== Read/Write sub-split tests =====

describe('read/write sub-splits — email', () => {
  test('read + write covers all email tools (no gaps)', () => {
    const readNames = toolNames(emailReadTools);
    const writeNames = toolNames(emailWriteTools);
    const combined = new Set([...readNames, ...writeNames]);
    const allNames = new Set(toolNames(emailTools));
    expect(combined).toEqual(allNames);
  });

  test('read and write do not overlap', () => {
    const readNames = new Set(toolNames(emailReadTools));
    const writeNames = toolNames(emailWriteTools);
    for (const name of writeNames) {
      expect(readNames.has(name)).toBe(false);
    }
  });

  test('safewrite is a subset of write', () => {
    const writeNames = new Set(toolNames(emailWriteTools));
    for (const name of toolNames(emailSafeWriteTools)) {
      expect(writeNames.has(name)).toBe(true);
    }
  });

  test('safewrite contains only mark-as-read', () => {
    expect(toolNames(emailSafeWriteTools)).toEqual(['mark-as-read']);
  });

  test('read tools are correct', () => {
    const expected = [
      'list-emails', 'search-emails', 'read-email',
      'list-email-attachments', 'download-email-attachment', 'download-email-attachments'
    ];
    expect(toolNames(emailReadTools).sort()).toEqual(expected.sort());
  });

  test('write tools are correct', () => {
    const expected = ['send-email', 'draft-email', 'mark-as-read'];
    expect(toolNames(emailWriteTools).sort()).toEqual(expected.sort());
  });
});

describe('read/write sub-splits — calendar', () => {
  test('read + write covers all calendar tools', () => {
    const combined = new Set([...toolNames(calendarReadTools), ...toolNames(calendarWriteTools)]);
    expect(combined).toEqual(new Set(toolNames(calendarTools)));
  });

  test('read and write do not overlap', () => {
    const readNames = new Set(toolNames(calendarReadTools));
    for (const name of toolNames(calendarWriteTools)) {
      expect(readNames.has(name)).toBe(false);
    }
  });

  test('read tools are correct', () => {
    expect(toolNames(calendarReadTools)).toEqual(['list-events']);
  });

  test('write tools are correct', () => {
    const expected = ['accept-event', 'decline-event', 'create-event', 'cancel-event', 'delete-event'];
    expect(toolNames(calendarWriteTools).sort()).toEqual(expected.sort());
  });
});

describe('read/write sub-splits — folder', () => {
  test('read + write covers all folder tools', () => {
    const combined = new Set([...toolNames(folderReadTools), ...toolNames(folderWriteTools)]);
    expect(combined).toEqual(new Set(toolNames(folderTools)));
  });

  test('read and write do not overlap', () => {
    const readNames = new Set(toolNames(folderReadTools));
    for (const name of toolNames(folderWriteTools)) {
      expect(readNames.has(name)).toBe(false);
    }
  });
});

describe('read/write sub-splits — rules', () => {
  test('read + write covers all rules tools', () => {
    const combined = new Set([...toolNames(rulesReadTools), ...toolNames(rulesWriteTools)]);
    expect(combined).toEqual(new Set(toolNames(rulesTools)));
  });

  test('read and write do not overlap', () => {
    const readNames = new Set(toolNames(rulesReadTools));
    for (const name of toolNames(rulesWriteTools)) {
      expect(readNames.has(name)).toBe(false);
    }
  });
});

describe('read/write sub-splits — onedrive', () => {
  test('read + write covers all onedrive tools', () => {
    const combined = new Set([...toolNames(onedriveReadTools), ...toolNames(onedriveWriteTools)]);
    expect(combined).toEqual(new Set(toolNames(onedriveTools)));
  });

  test('read and write do not overlap', () => {
    const readNames = new Set(toolNames(onedriveReadTools));
    for (const name of toolNames(onedriveWriteTools)) {
      expect(readNames.has(name)).toBe(false);
    }
  });

  test('read tools are correct', () => {
    const expected = ['onedrive-list', 'onedrive-search', 'onedrive-download'];
    expect(toolNames(onedriveReadTools).sort()).toEqual(expected.sort());
  });

  test('write tools are correct', () => {
    const expected = ['onedrive-upload', 'onedrive-upload-large', 'onedrive-share', 'onedrive-create-folder', 'onedrive-delete'];
    expect(toolNames(onedriveWriteTools).sort()).toEqual(expected.sort());
  });
});

describe('read/write sub-splits — power-automate', () => {
  test('read + write covers all power-automate tools', () => {
    const combined = new Set([...toolNames(powerAutomateReadTools), ...toolNames(powerAutomateWriteTools)]);
    expect(combined).toEqual(new Set(toolNames(powerAutomateTools)));
  });

  test('read and write do not overlap', () => {
    const readNames = new Set(toolNames(powerAutomateReadTools));
    for (const name of toolNames(powerAutomateWriteTools)) {
      expect(readNames.has(name)).toBe(false);
    }
  });

  test('read tools are correct', () => {
    const expected = ['flow-list-environments', 'flow-list', 'flow-list-runs'];
    expect(toolNames(powerAutomateReadTools).sort()).toEqual(expected.sort());
  });

  test('write tools are correct', () => {
    const expected = ['flow-run', 'flow-toggle'];
    expect(toolNames(powerAutomateWriteTools).sort()).toEqual(expected.sort());
  });
});

describe('read/write sub-split servers — no duplicates', () => {
  const subSplitServers = [
    { label: 'email-read', tools: emailReadServerTools },
    { label: 'email-write', tools: emailWriteServerTools },
    { label: 'email-safewrite', tools: emailSafeWriteServerTools },
    { label: 'calendar-read', tools: calendarReadServerTools },
    { label: 'calendar-write', tools: calendarWriteServerTools },
    { label: 'onedrive-read', tools: onedriveReadServerTools },
    { label: 'onedrive-write', tools: onedriveWriteServerTools },
    { label: 'power-automate-read', tools: powerAutomateReadServerTools },
    { label: 'power-automate-write', tools: powerAutomateWriteServerTools },
  ];

  test.each(subSplitServers)('$label server has no duplicate tool names', ({ tools }) => {
    const names = toolNames(tools);
    expect(new Set(names).size).toBe(names.length);
  });

  test.each(subSplitServers)('$label server includes auth tools', ({ tools }) => {
    const names = toolNames(tools);
    for (const name of AUTH_TOOL_NAMES) {
      expect(names).toContain(name);
    }
  });
});
