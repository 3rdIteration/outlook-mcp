/**
 * Shared MCP server factory
 *
 * Creates and starts an MCP server with the given tools.
 * Used by the combined index.js and by the per-service entry points
 * (server-email.js, server-calendar.js, server-onedrive.js, server-power-automate.js).
 */
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const config = require('./config');

/**
 * Create and start an MCP server.
 * @param {string} serverName - Display name for this server instance
 * @param {Array} tools - Array of tool definitions (name, description, inputSchema, handler)
 * @returns {Server} The MCP server instance
 */
function createServer(serverName, tools) {
  // Log startup information
  console.error(`STARTING ${serverName.toUpperCase()} MCP SERVER`);
  console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

  const TOOLS = tools;

  // Create server with tools capabilities
  const server = new Server(
    { name: serverName, version: config.SERVER_VERSION },
    {
      capabilities: {
        tools: {}
      }
    }
  );

  // Handle all requests
  server.fallbackRequestHandler = async (request) => {
    try {
      const { method, params, id } = request;
      console.error(`REQUEST: ${method} [${id}]`);

      // Initialize handler
      if (method === "initialize") {
        console.error(`INITIALIZE REQUEST: ID [${id}]`);
        return {
          protocolVersion: "2025-11-25",
          capabilities: {
            tools: {}
          },
          serverInfo: { name: serverName, version: config.SERVER_VERSION }
        };
      }

      // Tools list handler with optional cursor-based pagination
      if (method === "tools/list") {
        console.error(`TOOLS LIST REQUEST: ID [${id}]`);
        console.error(`TOOLS COUNT: ${TOOLS.length}`);
        console.error(`TOOLS NAMES: ${TOOLS.map(t => t.name).join(', ')}`);

        const pageSize = config.TOOLS_PAGE_SIZE;

        // When pageSize is 0 or negative, return all tools without pagination
        if (pageSize <= 0) {
          console.error(`TOOLS PAGE: 0-${TOOLS.length - 1} of ${TOOLS.length} (no pagination)`);
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

        console.error(`TOOLS PAGE: ${startIndex}-${endIndex - 1} of ${TOOLS.length}`);

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

      // Required empty responses for other capabilities
      if (method === "resources/list") return { resources: [] };
      if (method === "prompts/list") return { prompts: [] };

      // Tool call handler
      if (method === "tools/call") {
        try {
          const { name, arguments: args = {} } = params || {};

          console.error(`TOOL CALL: ${name}`);

          // Find the tool handler
          const tool = TOOLS.find(t => t.name === name);

          if (tool && tool.handler) {
            return await tool.handler(args);
          }

          // Tool not found
          return {
            error: {
              code: -32601,
              message: `Tool not found: ${name}`
            }
          };
        } catch (error) {
          console.error(`Error in tools/call:`, error);
          return {
            error: {
              code: -32603,
              message: `Error processing tool call: ${error.message}`
            }
          };
        }
      }

      // For any other method, return method not found
      return {
        error: {
          code: -32601,
          message: `Method not found: ${method}`
        }
      };
    } catch (error) {
      console.error(`Error in fallbackRequestHandler:`, error);
      return {
        error: {
          code: -32603,
          message: `Error processing request: ${error.message}`
        }
      };
    }
  };

  // Make the script executable
  process.on('SIGTERM', () => {
    console.error('SIGTERM received but staying alive');
  });

  // Start the server
  const transport = new StdioServerTransport();
  server.connect(transport)
    .then(() => console.error(`${serverName} connected and listening`))
    .catch(error => {
      console.error(`Connection error: ${error.message}`);
      process.exit(1);
    });

  return server;
}

module.exports = { createServer };
