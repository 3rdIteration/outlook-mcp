/**
 * Power Automate list flows functionality
 */
const { callFlowAPI } = require('./flow-api');
const { getFlowAccessToken } = require('../auth/token-manager');
const { sanitizeMetadata, wrapWithBoundary, wrapField, generateBoundaryToken } = require('../utils/metadata-sanitizer');

/**
 * List flows handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFlows(args) {
  const environmentId = args.environmentId;

  if (!environmentId) {
    return {
      content: [{
        type: "text",
        text: "Environment ID is required. Use 'flow-list-environments' to get available environments."
      }]
    };
  }

  try {
    const accessToken = getFlowAccessToken();

    if (!accessToken) {
      return {
        content: [{
          type: "text",
          text: "Power Automate authentication required. Please authenticate with Flow scope first."
        }]
      };
    }

    const path = `/environments/${environmentId}/flows`;
    const response = await callFlowAPI(accessToken, 'GET', path);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No flows found in environment ${environmentId}.\n\nNote: Only solution-aware flows are accessible via the API.`
        }]
      };
    }

    // Generate a shared boundary token for JSON payload and outer markers
    const boundaryToken = generateBoundaryToken();

    const flows = response.value.map((flow) => {
      const props = flow.properties || {};
      const state = props.state || 'Unknown';
      const triggerType = props.definition?.triggers ? Object.keys(props.definition.triggers)[0] : 'Unknown';
      const created = props.createdTime ? new Date(props.createdTime).toLocaleDateString() : 'Unknown';

      return {
        id: wrapField(flow.name, boundaryToken),
        displayName: wrapField(sanitizeMetadata(props.displayName || flow.name), boundaryToken),
        state,
        trigger: triggerType,
        created
      };
    });

    const payload = {
      _boundary: boundaryToken,
      flows
    };

    const flowList = JSON.stringify(payload, null, 2);

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} flow(s) in environment:\n\n${wrapWithBoundary(flowList, 'FLOWS', boundaryToken)}`
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
        text: `Error listing flows: ${error.message}`
      }]
    };
  }
}

module.exports = handleListFlows;
