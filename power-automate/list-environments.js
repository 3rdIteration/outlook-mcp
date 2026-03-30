/**
 * Power Automate list environments functionality
 */
const { callFlowAPI } = require('./flow-api');
const { getFlowAccessToken } = require('../auth/token-manager');
const { sanitizeMetadata, wrapWithBoundary, wrapField, generateBoundaryToken } = require('../utils/metadata-sanitizer');

/**
 * List environments handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEnvironments(args) {
  try {
    const accessToken = getFlowAccessToken();

    if (!accessToken) {
      return {
        content: [{
          type: "text",
          text: "Power Automate authentication required. Please authenticate with Flow scope first.\n\nNote: Power Automate requires additional Azure AD configuration with the Flow API scope."
        }]
      };
    }

    const response = await callFlowAPI(accessToken, 'GET', '/environments');

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No Power Platform environments found."
        }]
      };
    }

    // Generate a shared boundary token for JSON payload and outer markers
    const boundaryToken = generateBoundaryToken();

    const environments = response.value.map((env) => {
      const props = env.properties || {};
      return {
        id: wrapField(env.name, boundaryToken),
        displayName: wrapField(sanitizeMetadata(props.displayName || env.name), boundaryToken),
        isDefault: props.isDefault || false,
        region: wrapField(props.azureRegionHint || 'Unknown region', boundaryToken)
      };
    });

    const payload = {
      _boundary: boundaryToken,
      environments
    };

    const envList = JSON.stringify(payload, null, 2);

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} environment(s):\n\n${wrapWithBoundary(envList, 'ENVIRONMENTS', boundaryToken)}\n\nUse the environment ID (e.g., 'Default-12345') with other flow commands.`
      }]
    };
  } catch (error) {
    if (error.message === 'FLOW_UNAUTHORIZED') {
      return {
        content: [{
          type: "text",
          text: "Power Automate authentication expired. Please re-authenticate with Flow scope."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error listing environments: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEnvironments;
