/**
 * Power Automate module for M365 MCP server
 */
const handleListEnvironments = require('./list-environments');
const handleListFlows = require('./list-flows');
const handleRunFlow = require('./run-flow');
const handleListRuns = require('./list-runs');
const handleToggleFlow = require('./toggle-flow');

// Power Automate tool definitions
const powerAutomateTools = [
  {
    name: "flow-list-environments",
    description: "List Power Platform environments",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleListEnvironments
  },
  {
    name: "flow-list",
    description: "List flows in an environment",
    inputSchema: {
      type: "object",
      properties: {
        environmentId: {
          type: "string",
          description: "Environment ID"
        }
      },
      required: ["environmentId"]
    },
    handler: handleListFlows
  },
  {
    name: "flow-run",
    description: "Trigger a manual flow",
    inputSchema: {
      type: "object",
      properties: {
        environmentId: {
          type: "string",
          description: "Environment ID"
        },
        flowId: {
          type: "string",
          description: "Flow ID"
        },
        inputs: {
          type: "string",
          description: "JSON input parameters (optional)"
        }
      },
      required: ["environmentId", "flowId"]
    },
    handler: handleRunFlow
  },
  {
    name: "flow-list-runs",
    description: "Get flow run history",
    inputSchema: {
      type: "object",
      properties: {
        environmentId: {
          type: "string",
          description: "Environment ID"
        },
        flowId: {
          type: "string",
          description: "Flow ID"
        },
        count: {
          type: "number",
          description: "Max runs (default: 10)"
        }
      },
      required: ["environmentId", "flowId"]
    },
    handler: handleListRuns
  },
  {
    name: "flow-toggle",
    description: "Enable or disable a flow",
    inputSchema: {
      type: "object",
      properties: {
        environmentId: {
          type: "string",
          description: "Environment ID"
        },
        flowId: {
          type: "string",
          description: "Flow ID"
        },
        enable: {
          type: "boolean",
          description: "Enable (true) or disable (false)"
        }
      },
      required: ["environmentId", "flowId"]
    },
    handler: handleToggleFlow
  }
];

module.exports = {
  powerAutomateTools,
  handleListEnvironments,
  handleListFlows,
  handleRunFlow,
  handleListRuns,
  handleToggleFlow
};
