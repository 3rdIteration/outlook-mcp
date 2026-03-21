/**
 * Calendar module for Outlook MCP server
 */
const handleListEvents = require('./list');
const handleDeclineEvent = require('./decline');
const handleCreateEvent = require('./create');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');

// Calendar tool definitions
const calendarTools = [
  {
    name: "list-events",
    description: "List upcoming calendar events",
    inputSchema: {
      type: "object",
      properties: {
        count: {
          type: "number",
          description: "Number of events (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEvents
  },
  {
    name: "decline-event",
    description: "Decline a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "Event ID"
        },
        comment: {
          type: "string",
          description: "Optional decline comment"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeclineEvent
  },
  {
    name: "create-event",
    description: "Create a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        subject: {
          type: "string",
          description: "Event subject"
        },
        start: {
          type: "string",
          description: "Start time (ISO 8601)"
        },
        end: {
          type: "string",
          description: "End time (ISO 8601)"
        },
        attendees: {
          type: "array",
          items: {
            type: "string"
          },
          description: "Attendee email addresses"
        },
        body: {
          type: "string",
          description: "Event body content"
        }
      },
      required: ["subject", "start", "end"]
    },
    handler: handleCreateEvent
  },
  {
    name: "cancel-event",
    description: "Cancel a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "Event ID"
        },
        comment: {
          type: "string",
          description: "Optional cancel comment"
        }
      },
      required: ["eventId"]
    },
    handler: handleCancelEvent
  },
  {
    name: "delete-event",
    description: "Delete a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "Event ID"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeleteEvent
  }
];

module.exports = {
  calendarTools,
  handleListEvents,
  handleDeclineEvent,
  handleCreateEvent,
  handleCancelEvent,
  handleDeleteEvent
};
