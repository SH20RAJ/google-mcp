// Unified Google MCP backend for Google Apps Script.
// Deploy this script as a Web App. Store the API key in Script Properties as MCP_API_KEY.

var APP_VERSION = '1.0.0';
var SCRIPT_PROPERTY_API_KEY = 'MCP_API_KEY';
var DEFAULT_LIMIT = 10;
var MAX_LIMIT = 50;
var MAX_THREAD_MESSAGES = 25;
var MAX_TEXT_LENGTH = 50000;
var MAX_BODY_EXCERPT = 2000;
var MAX_FILE_CONTENT_BYTES = 200000;

// ---------------------------------------------------------------------------
// Tool definitions
// ---------------------------------------------------------------------------

const gmailTools = {
  listThreads: defineTool_(
    'List Gmail threads from the inbox or from a specific label.',
    {
      label: {
        type: 'string',
        description: 'Optional Gmail label name to filter threads. If omitted, inbox threads are returned.'
      },
      start: {
        type: 'number',
        description: 'Zero-based offset for pagination.'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of threads to return. Defaults to 10.'
      }
    },
    [],
    gmailListThreads_
  ),
  getThread: defineTool_(
    'Get a Gmail thread with message-level details.',
    {
      threadId: {
        type: 'string',
        description: 'The Gmail thread ID.'
      },
      includeBodies: {
        type: 'boolean',
        description: 'When true, includes truncated plain-text bodies for each message.'
      }
    },
    ['threadId'],
    gmailGetThread_
  ),
  search: defineTool_(
    'Search Gmail using Gmail search syntax.',
    {
      query: {
        type: 'string',
        description: 'Gmail search query.'
      },
      start: {
        type: 'number',
        description: 'Zero-based offset for pagination.'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of threads to return. Defaults to 10.'
      }
    },
    ['query'],
    gmailSearch_
  ),
  sendEmail: defineTool_(
    'Send a Gmail email.',
    {
      to: {
        type: 'string',
        description: 'Recipient email address or comma-separated addresses.'
      },
      subject: {
        type: 'string',
        description: 'Email subject.'
      },
      body: {
        type: 'string',
        description: 'Plain text email body.'
      },
      cc: {
        type: 'string',
        description: 'Optional CC recipients.'
      },
      bcc: {
        type: 'string',
        description: 'Optional BCC recipients.'
      },
      htmlBody: {
        type: 'string',
        description: 'Optional HTML body.'
      },
      replyTo: {
        type: 'string',
        description: 'Optional reply-to address.'
      },
      name: {
        type: 'string',
        description: 'Optional sender display name.'
      }
    },
    ['to', 'subject', 'body'],
    gmailSendEmail_
  ),
  reply: defineTool_(
    'Reply to an existing Gmail thread.',
    {
      threadId: {
        type: 'string',
        description: 'The Gmail thread ID.'
      },
      body: {
        type: 'string',
        description: 'Plain text reply body.'
      },
      htmlBody: {
        type: 'string',
        description: 'Optional HTML reply body.'
      },
      cc: {
        type: 'string',
        description: 'Optional CC recipients.'
      },
      bcc: {
        type: 'string',
        description: 'Optional BCC recipients.'
      },
      replyAll: {
        type: 'boolean',
        description: 'When true, sends a reply-all.'
      }
    },
    ['threadId', 'body'],
    gmailReply_
  ),
  archiveThread: defineTool_(
    'Archive a Gmail thread.',
    {
      threadId: {
        type: 'string',
        description: 'The Gmail thread ID.'
      }
    },
    ['threadId'],
    gmailArchiveThread_
  ),
  deleteThread: defineTool_(
    'Move a Gmail thread to trash.',
    {
      threadId: {
        type: 'string',
        description: 'The Gmail thread ID.'
      }
    },
    ['threadId'],
    gmailDeleteThread_
  ),
  markRead: defineTool_(
    'Mark a Gmail thread as read.',
    {
      threadId: {
        type: 'string',
        description: 'The Gmail thread ID.'
      }
    },
    ['threadId'],
    gmailMarkRead_
  ),
  markUnread: defineTool_(
    'Mark a Gmail thread as unread.',
    {
      threadId: {
        type: 'string',
        description: 'The Gmail thread ID.'
      }
    },
    ['threadId'],
    gmailMarkUnread_
  )
};

const driveTools = {
  listFiles: defineTool_(
    'List Drive files, optionally inside a specific folder.',
    {
      folderId: {
        type: 'string',
        description: 'Optional Drive folder ID. If omitted, files are listed across Drive.'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of files to return. Defaults to 10.'
      }
    },
    [],
    driveListFiles_
  ),
  searchFiles: defineTool_(
    'Search Drive files using Drive query syntax.',
    {
      query: {
        type: 'string',
        description: 'Drive file query.'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of files to return. Defaults to 10.'
      }
    },
    ['query'],
    driveSearchFiles_
  ),
  getFile: defineTool_(
    'Get Drive file metadata and optional text content for supported file types.',
    {
      fileId: {
        type: 'string',
        description: 'The Drive file ID.'
      },
      includeContent: {
        type: 'boolean',
        description: 'When true, attempts to return text content for non-Google-native files.'
      }
    },
    ['fileId'],
    driveGetFile_
  ),
  createFile: defineTool_(
    'Create a Drive file.',
    {
      name: {
        type: 'string',
        description: 'File name.'
      },
      content: {
        type: 'string',
        description: 'Text file content.'
      },
      mimeType: {
        type: 'string',
        description: 'Optional MIME type, for example text/plain or application/json.'
      },
      folderId: {
        type: 'string',
        description: 'Optional parent folder ID.'
      }
    },
    ['name', 'content'],
    driveCreateFile_
  ),
  updateFile: defineTool_(
    'Update Drive file metadata or text content.',
    {
      fileId: {
        type: 'string',
        description: 'The Drive file ID.'
      },
      name: {
        type: 'string',
        description: 'Optional new file name.'
      },
      content: {
        type: 'string',
        description: 'Optional new text content for supported file types.'
      },
      description: {
        type: 'string',
        description: 'Optional file description.'
      },
      starred: {
        type: 'boolean',
        description: 'Optional starred state.'
      },
      trashed: {
        type: 'boolean',
        description: 'Optional trashed state.'
      }
    },
    ['fileId'],
    driveUpdateFile_
  ),
  deleteFile: defineTool_(
    'Move a Drive file to trash.',
    {
      fileId: {
        type: 'string',
        description: 'The Drive file ID.'
      }
    },
    ['fileId'],
    driveDeleteFile_
  ),
  createFolder: defineTool_(
    'Create a Drive folder.',
    {
      name: {
        type: 'string',
        description: 'Folder name.'
      },
      parentId: {
        type: 'string',
        description: 'Optional parent folder ID.'
      }
    },
    ['name'],
    driveCreateFolder_
  ),
  shareFile: defineTool_(
    'Share a Drive file with a user.',
    {
      fileId: {
        type: 'string',
        description: 'The Drive file ID.'
      },
      email: {
        type: 'string',
        description: 'User email address.'
      },
      role: {
        type: 'string',
        description: 'Sharing role: viewer, commenter, or editor.'
      }
    },
    ['fileId', 'email', 'role'],
    driveShareFile_
  )
};

const calendarTools = {
  listEvents: defineTool_(
    'List Calendar events in a time range.',
    {
      calendarId: {
        type: 'string',
        description: 'Optional calendar ID. Uses the default calendar when omitted.'
      },
      startTime: {
        type: 'string',
        description: 'Optional ISO datetime for the start of the window. Defaults to now.'
      },
      endTime: {
        type: 'string',
        description: 'Optional ISO datetime for the end of the window. Defaults to 30 days after startTime.'
      },
      query: {
        type: 'string',
        description: 'Optional search text.'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of events to return. Defaults to 10.'
      }
    },
    [],
    calendarListEvents_
  ),
  getEvent: defineTool_(
    'Get a Calendar event by ID.',
    {
      calendarId: {
        type: 'string',
        description: 'Optional calendar ID. Uses the default calendar when omitted.'
      },
      eventId: {
        type: 'string',
        description: 'The Calendar event ID.'
      }
    },
    ['eventId'],
    calendarGetEvent_
  ),
  createEvent: defineTool_(
    'Create a Calendar event.',
    {
      calendarId: {
        type: 'string',
        description: 'Optional calendar ID. Uses the default calendar when omitted.'
      },
      title: {
        type: 'string',
        description: 'Event title.'
      },
      startTime: {
        type: 'string',
        description: 'Event start time as an ISO datetime.'
      },
      endTime: {
        type: 'string',
        description: 'Event end time as an ISO datetime.'
      },
      description: {
        type: 'string',
        description: 'Optional event description.'
      },
      location: {
        type: 'string',
        description: 'Optional event location.'
      },
      guests: {
        type: 'array',
        description: 'Optional array of guest email addresses.',
        items: {
          type: 'string'
        }
      },
      sendInvites: {
        type: 'boolean',
        description: 'When true, sends invitations to guests.'
      }
    },
    ['title', 'startTime', 'endTime'],
    calendarCreateEvent_
  ),
  updateEvent: defineTool_(
    'Update a Calendar event.',
    {
      calendarId: {
        type: 'string',
        description: 'Optional calendar ID. Uses the default calendar when omitted.'
      },
      eventId: {
        type: 'string',
        description: 'The Calendar event ID.'
      },
      title: {
        type: 'string',
        description: 'Optional new title.'
      },
      startTime: {
        type: 'string',
        description: 'Optional new start time as an ISO datetime.'
      },
      endTime: {
        type: 'string',
        description: 'Optional new end time as an ISO datetime.'
      },
      description: {
        type: 'string',
        description: 'Optional new description.'
      },
      location: {
        type: 'string',
        description: 'Optional new location.'
      },
      guestsToAdd: {
        type: 'array',
        description: 'Optional array of guest email addresses to add.',
        items: {
          type: 'string'
        }
      },
      color: {
        type: 'string',
        description: 'Optional event color.'
      },
      visibility: {
        type: 'string',
        description: 'Optional event visibility.'
      }
    },
    ['eventId'],
    calendarUpdateEvent_
  ),
  deleteEvent: defineTool_(
    'Delete a Calendar event.',
    {
      calendarId: {
        type: 'string',
        description: 'Optional calendar ID. Uses the default calendar when omitted.'
      },
      eventId: {
        type: 'string',
        description: 'The Calendar event ID.'
      }
    },
    ['eventId'],
    calendarDeleteEvent_
  )
};

const tasksTools = {
  list: defineTool_(
    'List Google Tasklists or tasks inside a specific tasklist.',
    {
      tasklistId: {
        type: 'string',
        description: 'Optional tasklist ID. When omitted, tasklists are returned.'
      },
      showCompleted: {
        type: 'boolean',
        description: 'When true, includes completed tasks.'
      },
      showHidden: {
        type: 'boolean',
        description: 'When true, includes hidden tasks.'
      },
      showDeleted: {
        type: 'boolean',
        description: 'When true, includes deleted tasks.'
      },
      dueMin: {
        type: 'string',
        description: 'Optional minimum due date as an ISO datetime.'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of items to return. Defaults to 10.'
      }
    },
    [],
    tasksList_
  ),
  create: defineTool_(
    'Create a Google Task.',
    {
      tasklistId: {
        type: 'string',
        description: 'Optional tasklist ID. Uses the first available tasklist when omitted.'
      },
      title: {
        type: 'string',
        description: 'Task title.'
      },
      notes: {
        type: 'string',
        description: 'Optional task notes.'
      },
      due: {
        type: 'string',
        description: 'Optional due date as an ISO datetime.'
      },
      parent: {
        type: 'string',
        description: 'Optional parent task ID.'
      },
      previous: {
        type: 'string',
        description: 'Optional sibling task ID to insert after.'
      }
    },
    ['title'],
    tasksCreate_
  ),
  update: defineTool_(
    'Update a Google Task.',
    {
      tasklistId: {
        type: 'string',
        description: 'Optional tasklist ID. Uses the first available tasklist when omitted.'
      },
      taskId: {
        type: 'string',
        description: 'The Google Task ID.'
      },
      title: {
        type: 'string',
        description: 'Optional new task title.'
      },
      notes: {
        type: 'string',
        description: 'Optional new task notes.'
      },
      due: {
        type: 'string',
        description: 'Optional new due date as an ISO datetime.'
      },
      status: {
        type: 'string',
        description: 'Optional task status, typically needsAction or completed.'
      },
      completed: {
        type: 'string',
        description: 'Optional completion time as an ISO datetime.'
      },
      deleted: {
        type: 'boolean',
        description: 'Optional deleted flag.'
      },
      hidden: {
        type: 'boolean',
        description: 'Optional hidden flag.'
      }
    },
    ['taskId'],
    tasksUpdate_
  ),
  delete: defineTool_(
    'Delete a Google Task.',
    {
      tasklistId: {
        type: 'string',
        description: 'Optional tasklist ID. Uses the first available tasklist when omitted.'
      },
      taskId: {
        type: 'string',
        description: 'The Google Task ID.'
      }
    },
    ['taskId'],
    tasksDelete_
  )
};

const sheetsTools = {
  read: defineTool_(
    'Read values from a Google Sheet.',
    {
      spreadsheetId: {
        type: 'string',
        description: 'The Spreadsheet ID.'
      },
      sheetName: {
        type: 'string',
        description: 'Optional sheet name. Uses the first sheet when omitted.'
      },
      range: {
        type: 'string',
        description: 'Optional A1 range. Uses a bounded data range when omitted.'
      },
      maxRows: {
        type: 'number',
        description: 'Optional safety cap for rows when range is omitted.'
      },
      maxColumns: {
        type: 'number',
        description: 'Optional safety cap for columns when range is omitted.'
      }
    },
    ['spreadsheetId'],
    sheetsRead_
  ),
  write: defineTool_(
    'Write values to a Google Sheet range.',
    {
      spreadsheetId: {
        type: 'string',
        description: 'The Spreadsheet ID.'
      },
      sheetName: {
        type: 'string',
        description: 'Optional sheet name. Uses the first sheet when omitted.'
      },
      range: {
        type: 'string',
        description: 'Target A1 range.'
      },
      values: {
        type: 'array',
        description: 'A 2D array of values. A 1D array is treated as a single row.',
        items: {}
      }
    },
    ['spreadsheetId', 'range', 'values'],
    sheetsWrite_
  ),
  append: defineTool_(
    'Append rows to a Google Sheet.',
    {
      spreadsheetId: {
        type: 'string',
        description: 'The Spreadsheet ID.'
      },
      sheetName: {
        type: 'string',
        description: 'Optional sheet name. Uses the first sheet when omitted.'
      },
      values: {
        type: 'array',
        description: 'A 2D array of rows to append. A 1D array is treated as one row.',
        items: {}
      }
    },
    ['spreadsheetId', 'values'],
    sheetsAppend_
  ),
  clear: defineTool_(
    'Clear values from a Google Sheet range or sheet.',
    {
      spreadsheetId: {
        type: 'string',
        description: 'The Spreadsheet ID.'
      },
      sheetName: {
        type: 'string',
        description: 'Optional sheet name. Uses the first sheet when omitted.'
      },
      range: {
        type: 'string',
        description: 'Optional A1 range. When omitted, clears the whole sheet.'
      }
    },
    ['spreadsheetId'],
    sheetsClear_
  )
};

const docsTools = {
  create: defineTool_(
    'Create a Google Doc.',
    {
      title: {
        type: 'string',
        description: 'Document title.'
      },
      bodyText: {
        type: 'string',
        description: 'Optional initial body text.'
      }
    },
    ['title'],
    docsCreate_
  ),
  read: defineTool_(
    'Read a Google Doc as plain text.',
    {
      documentId: {
        type: 'string',
        description: 'The Document ID.'
      }
    },
    ['documentId'],
    docsRead_
  ),
  appendText: defineTool_(
    'Append plain text to the end of a Google Doc.',
    {
      documentId: {
        type: 'string',
        description: 'The Document ID.'
      },
      text: {
        type: 'string',
        description: 'Text to append as a new paragraph.'
      }
    },
    ['documentId', 'text'],
    docsAppendText_
  )
};

const registry = {
  gmail: gmailTools,
  drive: driveTools,
  calendar: calendarTools,
  tasks: tasksTools,
  sheets: sheetsTools,
  docs: docsTools
};

// ---------------------------------------------------------------------------
// Web app entry points
// ---------------------------------------------------------------------------

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};

    if (isTruthyParam_(params.schema)) {
      return jsonResponse_({
        tools: getOpenAiToolSchemas_()
      });
    }

    if (isTruthyParam_(params.tools)) {
      return jsonResponse_({
        success: true,
        data: {
          tools: getToolDiscovery_()
        }
      });
    }

    return jsonResponse_({
      success: true,
      data: {
        status: 'ok',
        message: 'Google MCP backend is running.',
        version: APP_VERSION,
        apiKeyConfigured: isApiKeyConfigured_(),
        services: Object.keys(registry),
        toolCount: getToolDiscovery_().length,
        schemaHint: 'Use ?schema=1 to fetch OpenAI tool definitions.',
        discoveryHint: 'Use ?tools=1 to list available tools.',
        timestamp: new Date().toISOString()
      }
    });
  } catch (err) {
    return errorResponse_(err, 'GET_ERROR');
  }
}

function doPost(e) {
  try {
    var payload = parseJsonBody_(e);
    var apiKey = payload.apiKey || (e.parameter && e.parameter.key) || (e.postData.type === 'application/json' && payload.apiKey);
    
    // In GAS, headers are not always available in doPost(e) directly unless using a specific wrapper
    // but we can try to extract the key from the payload or parameters.
    validateApiKey_(apiKey);

    // MCP JSON-RPC protocol detection
    if (payload.method) {
      return handleMcpRequest_(payload);
    }

    // Legacy direct execution (if needed)
    if (!payload.tool) {
      throw createError_('Missing required field: tool (or method for MCP)', 'INVALID_REQUEST');
    }

    var args = hasOwn_(payload, 'arguments') ? payload.arguments : {};
    var result = executeTool_(payload.tool, args);

    return jsonResponse_({
      success: true,
      data: result
    });
  } catch (err) {
    return mcpErrorResponse_(err, payload ? payload.id : null);
  }
}

/**
 * Handles standard MCP JSON-RPC requests.
 */
function handleMcpRequest_(payload) {
  var method = payload.method;
  var params = payload.params || {};
  var id = payload.id;
  var result;

  switch (method) {
    case 'initialize':
      result = {
        protocolVersion: '2024-11-05',
        capabilities: {
          tools: {
            listChanged: false
          }
        },
        serverInfo: {
          name: 'google-mcp-backend',
          version: APP_VERSION
        }
      };
      break;

    case 'tools/list':
      result = {
        tools: getOpenAiToolSchemas_().map(function(tool) {
          return {
            name: tool.function.name,
            description: tool.function.description,
            inputSchema: tool.function.parameters
          };
        })
      };
      break;

    case 'tools/call':
      if (!params.name) {
        throw createError_('Missing tool name in tools/call.', 'INVALID_PARAMS');
      }
      var toolName = params.name;
      var args = params.arguments || {};
      var toolResult = executeTool_(toolName, args);
      
      result = {
        content: [
          {
            type: 'text',
            text: JSON.stringify(toolResult, null, 2)
          }
        ],
        isError: false
      };
      break;

    case 'notifications/initialized':
      // Notification doesn't expect a response, but we return a success for the proxy
      return jsonResponse_({ jsonrpc: '2.0', result: null });

    case 'resources/list':
    case 'prompts/list':
      result = { items: [] }; // Not implemented yet
      break;

    default:
      throw createError_('Unsupported MCP method: ' + method, 'METHOD_NOT_FOUND');
  }

  return jsonResponse_({
    jsonrpc: '2.0',
    id: id,
    result: result
  });
}

function mcpErrorResponse_(err, id) {
  var code = -32603; // Internal error
  if (err.code === 'UNAUTHORIZED') code = -32001;
  if (err.code === 'METHOD_NOT_FOUND') code = -32601;
  if (err.code === 'INVALID_PARAMS') code = -32602;

  return jsonResponse_({
    jsonrpc: '2.0',
    id: id,
    error: {
      code: code,
      message: err.message || 'Common MCP error'
    }
  });
}

// ---------------------------------------------------------------------------
// Core execution helpers
// ---------------------------------------------------------------------------

function defineTool_(description, properties, required, handler) {
  return {
    description: description,
    properties: properties || {},
    required: required || [],
    handler: handler
  };
}

function createToolSchema_(name, description, properties, required) {
  return {
    type: 'function',
    function: {
      name: name,
      description: description,
      parameters: {
        type: 'object',
        properties: properties || {},
        required: required || [],
        additionalProperties: false
      }
    }
  };
}

function getOpenAiToolSchemas_() {
  var tools = [];
  var services = Object.keys(registry);

  for (var i = 0; i < services.length; i++) {
    var service = services[i];
    var actions = Object.keys(registry[service]);

    for (var j = 0; j < actions.length; j++) {
      var action = actions[j];
      var toolDef = registry[service][action];
      tools.push(
        createToolSchema_(
          service + '.' + action,
          toolDef.description,
          toolDef.properties,
          toolDef.required
        )
      );
    }
  }

  return tools;
}

function getToolDiscovery_() {
  var tools = [];
  var services = Object.keys(registry);

  for (var i = 0; i < services.length; i++) {
    var service = services[i];
    var actions = Object.keys(registry[service]);

    for (var j = 0; j < actions.length; j++) {
      var action = actions[j];
      var toolDef = registry[service][action];
      tools.push({
        name: service + '.' + action,
        service: service,
        action: action,
        description: toolDef.description
      });
    }
  }

  return tools;
}

function executeTool_(toolName, args) {
  var route = resolveTool_(toolName);
  var normalizedArgs = normalizeArguments_(args);

  validateArguments_(toolName, route.tool, normalizedArgs);

  return route.tool.handler(normalizedArgs);
}

function resolveTool_(toolName) {
  if (!toolName || typeof toolName !== 'string') {
    throw createError_('Tool name must be a non-empty string.', 'INVALID_TOOL');
  }

  var parts = toolName.split('.');
  if (parts.length !== 2) {
    throw createError_('Tool name must use the format service.action.', 'INVALID_TOOL');
  }

  var service = parts[0];
  var action = parts[1];
  var serviceTools = registry[service];

  if (!serviceTools) {
    throw createError_('Unknown service: ' + service, 'UNKNOWN_SERVICE');
  }

  if (!serviceTools[action]) {
    throw createError_('Unknown action for service ' + service + ': ' + action, 'UNKNOWN_ACTION');
  }

  return {
    service: service,
    action: action,
    tool: serviceTools[action]
  };
}

function validateArguments_(toolName, toolDef, args) {
  if (args === null || typeof args !== 'object' || Array.isArray(args)) {
    throw createError_('Tool arguments must be a JSON object.', 'INVALID_ARGUMENTS');
  }

  var required = toolDef.required || [];
  for (var i = 0; i < required.length; i++) {
    var key = required[i];
    if (!hasOwn_(args, key) || args[key] === null || args[key] === '') {
      throw createError_('Missing required argument for ' + toolName + ': ' + key, 'MISSING_ARGUMENT');
    }
  }

  var allowed = toolDef.properties || {};
  var argKeys = Object.keys(args);

  for (var j = 0; j < argKeys.length; j++) {
    var argKey = argKeys[j];
    if (!hasOwn_(allowed, argKey)) {
      throw createError_('Unknown argument for ' + toolName + ': ' + argKey, 'UNKNOWN_ARGUMENT');
    }

    if (args[argKey] === null || args[argKey] === undefined) {
      continue;
    }

    validateArgumentType_(toolName, argKey, allowed[argKey], args[argKey]);
  }
}

function validateArgumentType_(toolName, argKey, schema, value) {
  if (!schema || !schema.type) {
    return;
  }

  var valid = true;
  switch (schema.type) {
    case 'string':
      valid = typeof value === 'string';
      break;
    case 'number':
      valid = typeof value === 'number' && !isNaN(value);
      break;
    case 'boolean':
      valid = typeof value === 'boolean';
      break;
    case 'array':
      valid = Array.isArray(value);
      break;
    case 'object':
      valid = value && typeof value === 'object' && !Array.isArray(value);
      break;
    default:
      valid = true;
  }

  if (!valid) {
    throw createError_(
      'Invalid type for ' + toolName + '.' + argKey + '. Expected ' + schema.type + '.',
      'INVALID_ARGUMENT_TYPE'
    );
  }
}

function normalizeArguments_(args) {
  if (args === undefined || args === null) {
    return {};
  }
  return args;
}

// ---------------------------------------------------------------------------
// Response helpers
// ---------------------------------------------------------------------------

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse_(err, fallbackCode) {
  return jsonResponse_({
    success: false,
    error: {
      code: err && err.code ? err.code : (fallbackCode || 'INTERNAL_ERROR'),
      message: err && err.message ? err.message : 'Unexpected error.'
    }
  });
}

function createError_(message, code) {
  var err = new Error(message);
  err.code = code || 'ERROR';
  return err;
}

// ---------------------------------------------------------------------------
// Security helpers
// ---------------------------------------------------------------------------

function getConfiguredApiKey_() {
  return PropertiesService.getScriptProperties().getProperty(SCRIPT_PROPERTY_API_KEY);
}

function isApiKeyConfigured_() {
  return !!getConfiguredApiKey_();
}

function validateApiKey_(apiKey) {
  var configuredKey = getConfiguredApiKey_();

  if (!configuredKey) {
    throw createError_(
      'Server is not configured. Set the MCP_API_KEY script property before using POST /exec.',
      'CONFIGURATION_ERROR'
    );
  }

  if (!apiKey || apiKey !== configuredKey) {
    throw createError_('Unauthorized request: invalid API key.', 'UNAUTHORIZED');
  }
}

// ---------------------------------------------------------------------------
// Request parsing helpers
// ---------------------------------------------------------------------------

function parseJsonBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw createError_('Missing request body.', 'INVALID_REQUEST');
  }

  try {
    var payload = JSON.parse(e.postData.contents);
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      throw new Error('JSON body must be an object.');
    }
    return payload;
  } catch (err) {
    throw createError_('Invalid JSON body: ' + err.message, 'INVALID_JSON');
  }
}

function isTruthyParam_(value) {
  if (value === true) {
    return true;
  }
  if (value === false || value === undefined || value === null) {
    return false;
  }
  var normalized = String(value).toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'yes';
}

// ---------------------------------------------------------------------------
// Gmail tools
// ---------------------------------------------------------------------------

function gmailListThreads_(args) {
  var start = toNonNegativeInt_(args.start, 0);
  var limit = clampLimit_(args.limit, DEFAULT_LIMIT, MAX_LIMIT);
  var threads;

  if (args.label) {
    threads = GmailApp.search('label:"' + escapeSearchValue_(args.label) + '"', start, limit);
  } else {
    threads = GmailApp.getInboxThreads(start, limit);
  }

  return {
    count: threads.length,
    threads: mapArray_(threads, function(thread) {
      return serializeThreadSummary_(thread);
    })
  };
}

function gmailGetThread_(args) {
  var thread = GmailApp.getThreadById(args.threadId);
  if (!thread) {
    throw createError_('Gmail thread not found: ' + args.threadId, 'NOT_FOUND');
  }

  return {
    thread: serializeThreadDetail_(thread, args.includeBodies !== false)
  };
}

function gmailSearch_(args) {
  var start = toNonNegativeInt_(args.start, 0);
  var limit = clampLimit_(args.limit, DEFAULT_LIMIT, MAX_LIMIT);
  var threads = GmailApp.search(args.query, start, limit);

  return {
    query: args.query,
    count: threads.length,
    threads: mapArray_(threads, function(thread) {
      return serializeThreadSummary_(thread);
    })
  };
}

function gmailSendEmail_(args) {
  var options = {};

  copyOptionalArg_(args, options, 'cc');
  copyOptionalArg_(args, options, 'bcc');
  copyOptionalArg_(args, options, 'htmlBody');
  copyOptionalArg_(args, options, 'replyTo');
  copyOptionalArg_(args, options, 'name');

  GmailApp.sendEmail(args.to, args.subject, args.body, options);

  return {
    sent: true,
    to: args.to,
    subject: args.subject
  };
}

function gmailReply_(args) {
  var thread = GmailApp.getThreadById(args.threadId);
  if (!thread) {
    throw createError_('Gmail thread not found: ' + args.threadId, 'NOT_FOUND');
  }

  var options = {};
  copyOptionalArg_(args, options, 'cc');
  copyOptionalArg_(args, options, 'bcc');
  copyOptionalArg_(args, options, 'htmlBody');

  if (args.replyAll) {
    thread.replyAll(args.body, options);
  } else {
    thread.reply(args.body, options);
  }

  return {
    replied: true,
    thread: serializeThreadSummary_(thread)
  };
}

function gmailArchiveThread_(args) {
  var thread = GmailApp.getThreadById(args.threadId);
  if (!thread) {
    throw createError_('Gmail thread not found: ' + args.threadId, 'NOT_FOUND');
  }

  thread.moveToArchive();

  return {
    archived: true,
    threadId: args.threadId
  };
}

function gmailDeleteThread_(args) {
  var thread = GmailApp.getThreadById(args.threadId);
  if (!thread) {
    throw createError_('Gmail thread not found: ' + args.threadId, 'NOT_FOUND');
  }

  thread.moveToTrash();

  return {
    deleted: true,
    threadId: args.threadId
  };
}

function gmailMarkRead_(args) {
  var thread = GmailApp.getThreadById(args.threadId);
  if (!thread) {
    throw createError_('Gmail thread not found: ' + args.threadId, 'NOT_FOUND');
  }

  thread.markRead();

  return {
    markedRead: true,
    threadId: args.threadId
  };
}

function gmailMarkUnread_(args) {
  var thread = GmailApp.getThreadById(args.threadId);
  if (!thread) {
    throw createError_('Gmail thread not found: ' + args.threadId, 'NOT_FOUND');
  }

  thread.markUnread();

  return {
    markedUnread: true,
    threadId: args.threadId
  };
}

// ---------------------------------------------------------------------------
// Drive tools
// ---------------------------------------------------------------------------

function driveListFiles_(args) {
  var limit = clampLimit_(args.limit, DEFAULT_LIMIT, MAX_LIMIT);
  var iterator = args.folderId
    ? DriveApp.getFolderById(args.folderId).getFiles()
    : DriveApp.getFiles();
  var files = collectIterator_(iterator, limit, serializeFileSummary_);

  return {
    count: files.length,
    files: files
  };
}

function driveSearchFiles_(args) {
  var limit = clampLimit_(args.limit, DEFAULT_LIMIT, MAX_LIMIT);
  var iterator = DriveApp.searchFiles(args.query);
  var files = collectIterator_(iterator, limit, serializeFileSummary_);

  return {
    query: args.query,
    count: files.length,
    files: files
  };
}

function driveGetFile_(args) {
  var file = DriveApp.getFileById(args.fileId);
  if (!file) {
    throw createError_('Drive file not found: ' + args.fileId, 'NOT_FOUND');
  }

  var data = serializeFileDetail_(file);
  if (args.includeContent) {
    data.content = getFileContentResult_(file);
  }

  return {
    file: data
  };
}

function driveCreateFile_(args) {
  var file;
  var mimeType = args.mimeType || 'text/plain';

  if (args.folderId) {
    var folder = DriveApp.getFolderById(args.folderId);
    file = args.mimeType
      ? folder.createFile(args.name, args.content, mimeType)
      : folder.createFile(args.name, args.content);
  } else {
    file = args.mimeType
      ? DriveApp.createFile(args.name, args.content, mimeType)
      : DriveApp.createFile(args.name, args.content);
  }

  return {
    file: serializeFileDetail_(file)
  };
}

function driveUpdateFile_(args) {
  var file = DriveApp.getFileById(args.fileId);
  if (!file) {
    throw createError_('Drive file not found: ' + args.fileId, 'NOT_FOUND');
  }

  if (hasOwn_(args, 'name')) {
    file.setName(args.name);
  }
  if (hasOwn_(args, 'description')) {
    file.setDescription(args.description);
  }
  if (hasOwn_(args, 'starred')) {
    file.setStarred(!!args.starred);
  }
  if (hasOwn_(args, 'trashed')) {
    file.setTrashed(!!args.trashed);
  }
  if (hasOwn_(args, 'content')) {
    var mimeType = file.getMimeType();
    if (mimeType.indexOf('application/vnd.google-apps') === 0) {
      throw createError_(
        'Native Google files cannot be updated through drive.updateFile content. Use docs.* or sheets.* tools instead.',
        'UNSUPPORTED_OPERATION'
      );
    }
    file.setContent(args.content);
  }

  return {
    file: serializeFileDetail_(file)
  };
}

function driveDeleteFile_(args) {
  var file = DriveApp.getFileById(args.fileId);
  if (!file) {
    throw createError_('Drive file not found: ' + args.fileId, 'NOT_FOUND');
  }

  file.setTrashed(true);

  return {
    deleted: true,
    fileId: args.fileId
  };
}

function driveCreateFolder_(args) {
  var folder = args.parentId
    ? DriveApp.getFolderById(args.parentId).createFolder(args.name)
    : DriveApp.createFolder(args.name);

  return {
    folder: {
      id: folder.getId(),
      name: folder.getName(),
      url: folder.getUrl(),
      dateCreated: folder.getDateCreated().toISOString(),
      lastUpdated: folder.getLastUpdated().toISOString()
    }
  };
}

function driveShareFile_(args) {
  var file = DriveApp.getFileById(args.fileId);
  if (!file) {
    throw createError_('Drive file not found: ' + args.fileId, 'NOT_FOUND');
  }

  var role = String(args.role).toLowerCase();
  if (role === 'viewer') {
    file.addViewer(args.email);
  } else if (role === 'editor') {
    file.addEditor(args.email);
  } else if (role === 'commenter') {
    if (typeof file.addCommenter === 'function') {
      file.addCommenter(args.email);
    } else {
      throw createError_('Commenter sharing is not supported in this Apps Script runtime.', 'UNSUPPORTED_OPERATION');
    }
  } else {
    throw createError_('Invalid share role. Use viewer, commenter, or editor.', 'INVALID_ARGUMENT');
  }

  return {
    shared: true,
    email: args.email,
    role: role,
    file: serializeFileSummary_(file)
  };
}

// ---------------------------------------------------------------------------
// Calendar tools
// ---------------------------------------------------------------------------

function calendarListEvents_(args) {
  var calendar = getCalendar_(args.calendarId);
  var startTime = args.startTime ? parseDate_(args.startTime, 'startTime') : new Date();
  var endTime = args.endTime
    ? parseDate_(args.endTime, 'endTime')
    : new Date(startTime.getTime() + 30 * 24 * 60 * 60 * 1000);

  if (endTime.getTime() <= startTime.getTime()) {
    throw createError_('endTime must be later than startTime.', 'INVALID_ARGUMENT');
  }

  var options = {
    max: clampLimit_(args.limit, DEFAULT_LIMIT, MAX_LIMIT)
  };
  if (args.query) {
    options.search = args.query;
  }

  var events = calendar.getEvents(startTime, endTime, options);

  return {
    calendarId: calendar.getId(),
    count: events.length,
    events: mapArray_(events, function(event) {
      return serializeEvent_(event, false);
    })
  };
}

function calendarGetEvent_(args) {
  var calendar = getCalendar_(args.calendarId);
  var event = resolveCalendarEvent_(calendar, args.eventId);

  return {
    event: serializeEvent_(event, true)
  };
}

function calendarCreateEvent_(args) {
  var calendar = getCalendar_(args.calendarId);
  var startTime = parseDate_(args.startTime, 'startTime');
  var endTime = parseDate_(args.endTime, 'endTime');

  if (endTime.getTime() <= startTime.getTime()) {
    throw createError_('endTime must be later than startTime.', 'INVALID_ARGUMENT');
  }

  var options = {};
  copyOptionalArg_(args, options, 'description');
  copyOptionalArg_(args, options, 'location');
  if (args.guests && args.guests.length) {
    options.guests = args.guests.join(',');
  }
  if (hasOwn_(args, 'sendInvites')) {
    options.sendInvites = !!args.sendInvites;
  }

  var event = calendar.createEvent(args.title, startTime, endTime, options);

  return {
    event: serializeEvent_(event, true)
  };
}

function calendarUpdateEvent_(args) {
  var calendar = getCalendar_(args.calendarId);
  var event = resolveCalendarEvent_(calendar, args.eventId);

  if (hasOwn_(args, 'title')) {
    event.setTitle(args.title);
  }

  if (hasOwn_(args, 'startTime') || hasOwn_(args, 'endTime')) {
    var startTime = hasOwn_(args, 'startTime') ? parseDate_(args.startTime, 'startTime') : event.getStartTime();
    var endTime = hasOwn_(args, 'endTime') ? parseDate_(args.endTime, 'endTime') : event.getEndTime();

    if (endTime.getTime() <= startTime.getTime()) {
      throw createError_('endTime must be later than startTime.', 'INVALID_ARGUMENT');
    }

    event.setTime(startTime, endTime);
  }

  if (hasOwn_(args, 'description')) {
    event.setDescription(args.description);
  }
  if (hasOwn_(args, 'location')) {
    event.setLocation(args.location);
  }
  if (args.guestsToAdd && args.guestsToAdd.length) {
    for (var i = 0; i < args.guestsToAdd.length; i++) {
      event.addGuest(args.guestsToAdd[i]);
    }
  }
  if (hasOwn_(args, 'color') && typeof event.setColor === 'function') {
    event.setColor(args.color);
  }
  if (hasOwn_(args, 'visibility') && typeof event.setVisibility === 'function') {
    event.setVisibility(args.visibility);
  }

  return {
    event: serializeEvent_(event, true)
  };
}

function calendarDeleteEvent_(args) {
  var calendar = getCalendar_(args.calendarId);
  var event = resolveCalendarEvent_(calendar, args.eventId);
  event.deleteEvent();

  return {
    deleted: true,
    eventId: args.eventId
  };
}

// ---------------------------------------------------------------------------
// Tasks tools
// ---------------------------------------------------------------------------

function tasksList_(args) {
  ensureTasksServiceEnabled_();
  var limit = clampLimit_(args.limit, DEFAULT_LIMIT, MAX_LIMIT);

  if (!args.tasklistId) {
    var tasklistsResult = Tasks.Tasklists.list({
      maxResults: limit
    }) || {};
    var tasklists = tasklistsResult.items || [];

    return {
      count: tasklists.length,
      tasklists: mapArray_(tasklists, serializeTasklist_)
    };
  }

  var params = {
    maxResults: limit,
    showCompleted: !!args.showCompleted,
    showDeleted: !!args.showDeleted,
    showHidden: !!args.showHidden
  };

  if (args.dueMin) {
    params.dueMin = parseDate_(args.dueMin, 'dueMin').toISOString();
  }

  var tasksResult = Tasks.Tasks.list(args.tasklistId, params) || {};
  var tasks = tasksResult.items || [];

  return {
    tasklistId: args.tasklistId,
    count: tasks.length,
    tasks: mapArray_(tasks, serializeTask_)
  };
}

function tasksCreate_(args) {
  ensureTasksServiceEnabled_();
  var tasklistId = resolveTasklistId_(args.tasklistId);
  var resource = {
    title: args.title
  };

  copyOptionalArg_(args, resource, 'notes');
  if (hasOwn_(args, 'due')) {
    resource.due = parseDate_(args.due, 'due').toISOString();
  }
  copyOptionalArg_(args, resource, 'parent');
  copyOptionalArg_(args, resource, 'previous');

  var created = Tasks.Tasks.insert(resource, tasklistId);

  return {
    tasklistId: tasklistId,
    task: serializeTask_(created)
  };
}

function tasksUpdate_(args) {
  ensureTasksServiceEnabled_();
  var tasklistId = resolveTasklistId_(args.tasklistId);
  var patch = {};

  copyOptionalArg_(args, patch, 'title');
  copyOptionalArg_(args, patch, 'notes');
  copyOptionalArg_(args, patch, 'status');

  if (hasOwn_(args, 'due')) {
    patch.due = parseDate_(args.due, 'due').toISOString();
  }
  if (hasOwn_(args, 'completed')) {
    patch.completed = parseDate_(args.completed, 'completed').toISOString();
  }
  if (hasOwn_(args, 'deleted')) {
    patch.deleted = !!args.deleted;
  }
  if (hasOwn_(args, 'hidden')) {
    patch.hidden = !!args.hidden;
  }

  if (!Object.keys(patch).length) {
    throw createError_('No fields provided to update.', 'INVALID_ARGUMENT');
  }

  var updated = Tasks.Tasks.patch(patch, tasklistId, args.taskId);

  return {
    tasklistId: tasklistId,
    task: serializeTask_(updated)
  };
}

function tasksDelete_(args) {
  ensureTasksServiceEnabled_();
  var tasklistId = resolveTasklistId_(args.tasklistId);
  Tasks.Tasks.remove(tasklistId, args.taskId);

  return {
    deleted: true,
    tasklistId: tasklistId,
    taskId: args.taskId
  };
}

// ---------------------------------------------------------------------------
// Sheets tools
// ---------------------------------------------------------------------------

function sheetsRead_(args) {
  var spreadsheet = SpreadsheetApp.openById(args.spreadsheetId);
  var sheet = getSheet_(spreadsheet, args.sheetName);
  var range;

  if (args.range) {
    range = sheet.getRange(args.range);
  } else {
    var lastRow = Math.max(sheet.getLastRow(), 1);
    var lastColumn = Math.max(sheet.getLastColumn(), 1);
    var maxRows = clampLimit_(args.maxRows, Math.min(lastRow, 200), 1000);
    var maxColumns = clampLimit_(args.maxColumns, Math.min(lastColumn, 50), 100);
    range = sheet.getRange(1, 1, Math.min(lastRow, maxRows), Math.min(lastColumn, maxColumns));
  }

  var values = serializeTableValues_(range.getValues());

  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName(),
    sheetName: sheet.getName(),
    range: range.getA1Notation(),
    rowCount: values.length,
    columnCount: values.length ? values[0].length : 0,
    values: values
  };
}

function sheetsWrite_(args) {
  var spreadsheet = SpreadsheetApp.openById(args.spreadsheetId);
  var sheet = getSheet_(spreadsheet, args.sheetName);
  var values = normalize2DValues_(args.values);
  var range = sheet.getRange(args.range);

  if (range.getNumRows() !== values.length || range.getNumColumns() !== values[0].length) {
    throw createError_(
      'Range size does not match values dimensions. Expected ' +
        range.getNumRows() +
        'x' +
        range.getNumColumns() +
        ' but received ' +
        values.length +
        'x' +
        values[0].length +
        '.',
      'DIMENSION_MISMATCH'
    );
  }

  range.setValues(values);

  return {
    written: true,
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheet.getName(),
    range: range.getA1Notation(),
    rowCount: values.length,
    columnCount: values[0].length
  };
}

function sheetsAppend_(args) {
  var spreadsheet = SpreadsheetApp.openById(args.spreadsheetId);
  var sheet = getSheet_(spreadsheet, args.sheetName);
  var values = normalize2DValues_(args.values);
  var startRow = sheet.getLastRow() + 1;
  if (startRow < 1) {
    startRow = 1;
  }
  var range = sheet.getRange(startRow, 1, values.length, values[0].length);
  range.setValues(values);

  return {
    appended: true,
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheet.getName(),
    range: range.getA1Notation(),
    rowCount: values.length,
    columnCount: values[0].length
  };
}

function sheetsClear_(args) {
  var spreadsheet = SpreadsheetApp.openById(args.spreadsheetId);
  var sheet = getSheet_(spreadsheet, args.sheetName);

  if (args.range) {
    var range = sheet.getRange(args.range);
    range.clearContent();
    return {
      cleared: true,
      spreadsheetId: spreadsheet.getId(),
      sheetName: sheet.getName(),
      range: range.getA1Notation()
    };
  }

  sheet.clearContents();

  return {
    cleared: true,
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheet.getName(),
    range: sheet.getDataRange().getA1Notation()
  };
}

// ---------------------------------------------------------------------------
// Docs tools
// ---------------------------------------------------------------------------

function docsCreate_(args) {
  var doc = DocumentApp.create(args.title);

  if (args.bodyText) {
    doc.getBody().appendParagraph(args.bodyText);
  }

  doc.saveAndClose();

  return {
    document: {
      id: doc.getId(),
      title: doc.getName(),
      url: doc.getUrl()
    }
  };
}

function docsRead_(args) {
  var doc = DocumentApp.openById(args.documentId);
  var text = doc.getBody().getText();
  var truncated = text.length > MAX_TEXT_LENGTH;

  return {
    document: {
      id: doc.getId(),
      title: doc.getName(),
      url: doc.getUrl(),
      text: truncateText_(text, MAX_TEXT_LENGTH),
      truncated: truncated
    }
  };
}

function docsAppendText_(args) {
  var doc = DocumentApp.openById(args.documentId);
  doc.getBody().appendParagraph(args.text);
  doc.saveAndClose();

  return {
    appended: true,
    document: {
      id: doc.getId(),
      title: doc.getName(),
      url: doc.getUrl()
    }
  };
}

// ---------------------------------------------------------------------------
// Gmail serialization helpers
// ---------------------------------------------------------------------------

function serializeThreadSummary_(thread) {
  var messages = thread.getMessages();
  var lastMessage = messages.length ? messages[messages.length - 1] : null;

  return {
    id: thread.getId(),
    subject: thread.getFirstMessageSubject(),
    messageCount: thread.getMessageCount(),
    unread: thread.isUnread(),
    labels: mapArray_(thread.getLabels(), function(label) {
      return label.getName();
    }),
    lastMessageDate: thread.getLastMessageDate().toISOString(),
    snippet: lastMessage ? truncateText_(lastMessage.getPlainBody(), 240) : ''
  };
}

function serializeThreadDetail_(thread, includeBodies) {
  var messages = thread.getMessages();
  var limitedMessages = [];
  var maxMessages = Math.min(messages.length, MAX_THREAD_MESSAGES);

  for (var i = 0; i < maxMessages; i++) {
    limitedMessages.push(serializeMessage_(messages[i], includeBodies));
  }

  return {
    id: thread.getId(),
    subject: thread.getFirstMessageSubject(),
    messageCount: thread.getMessageCount(),
    unread: thread.isUnread(),
    labels: mapArray_(thread.getLabels(), function(label) {
      return label.getName();
    }),
    messagesTruncated: messages.length > MAX_THREAD_MESSAGES,
    messages: limitedMessages
  };
}

function serializeMessage_(message, includeBody) {
  var data = {
    id: message.getId(),
    from: message.getFrom(),
    to: message.getTo(),
    cc: message.getCc(),
    bcc: message.getBcc(),
    replyTo: message.getReplyTo(),
    subject: message.getSubject(),
    date: message.getDate().toISOString()
  };

  if (includeBody) {
    data.body = truncateText_(message.getPlainBody(), MAX_BODY_EXCERPT);
  }

  return data;
}

// ---------------------------------------------------------------------------
// Drive serialization helpers
// ---------------------------------------------------------------------------

function serializeFileSummary_(file) {
  return {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    size: file.getSize(),
    url: file.getUrl(),
    dateCreated: file.getDateCreated().toISOString(),
    lastUpdated: file.getLastUpdated().toISOString(),
    ownerEmail: safeUserEmail_(file.getOwner()),
    trashed: file.isTrashed()
  };
}

function serializeFileDetail_(file) {
  return {
    id: file.getId(),
    name: file.getName(),
    description: file.getDescription(),
    mimeType: file.getMimeType(),
    size: file.getSize(),
    url: file.getUrl(),
    dateCreated: file.getDateCreated().toISOString(),
    lastUpdated: file.getLastUpdated().toISOString(),
    ownerEmail: safeUserEmail_(file.getOwner()),
    editorEmails: mapArray_(file.getEditors(), safeUserEmail_),
    viewerEmails: mapArray_(file.getViewers(), safeUserEmail_),
    parentIds: getParentIds_(file),
    starred: file.isStarred(),
    trashed: file.isTrashed()
  };
}

function getFileContentResult_(file) {
  var mimeType = file.getMimeType();

  if (mimeType.indexOf('application/vnd.google-apps') === 0) {
    return {
      available: false,
      reason: 'Use the service-specific tools for native Google files.'
    };
  }

  if (file.getSize() > MAX_FILE_CONTENT_BYTES) {
    return {
      available: false,
      reason: 'File content is too large to return safely.'
    };
  }

  try {
    var text = file.getBlob().getDataAsString();
    return {
      available: true,
      text: truncateText_(text, MAX_TEXT_LENGTH),
      truncated: text.length > MAX_TEXT_LENGTH
    };
  } catch (err) {
    return {
      available: false,
      reason: 'Failed to read file as text: ' + err.message
    };
  }
}

function getParentIds_(file) {
  var parents = [];
  var iterator = file.getParents();
  while (iterator.hasNext()) {
    parents.push(iterator.next().getId());
  }
  return parents;
}

// ---------------------------------------------------------------------------
// Calendar helpers
// ---------------------------------------------------------------------------

function getCalendar_(calendarId) {
  var calendar = calendarId ? CalendarApp.getCalendarById(calendarId) : CalendarApp.getDefaultCalendar();

  if (!calendar) {
    throw createError_('Calendar not found: ' + calendarId, 'NOT_FOUND');
  }

  return calendar;
}

function resolveCalendarEvent_(calendar, eventId) {
  var event = calendar.getEventById(eventId);

  if (!event && typeof calendar.getEventSeriesById === 'function') {
    event = calendar.getEventSeriesById(eventId);
  }

  if (!event) {
    throw createError_('Calendar event not found: ' + eventId, 'NOT_FOUND');
  }

  return event;
}

function serializeEvent_(event, includeGuests) {
  var data = {
    id: event.getId(),
    title: event.getTitle(),
    startTime: event.getStartTime().toISOString(),
    endTime: event.getEndTime().toISOString(),
    description: event.getDescription(),
    location: event.getLocation(),
    isAllDayEvent: event.isAllDayEvent(),
    myStatus: String(event.getMyStatus()),
    guestCount: event.getGuestList().length
  };

  if (includeGuests) {
    data.guests = mapArray_(event.getGuestList(), function(guest) {
      return {
        email: guest.getEmail(),
        status: String(guest.getGuestStatus())
      };
    });
  }

  return data;
}

// ---------------------------------------------------------------------------
// Tasks helpers
// ---------------------------------------------------------------------------

function ensureTasksServiceEnabled_() {
  if (typeof Tasks === 'undefined') {
    throw createError_(
      'Advanced Google Service "Tasks API" is not enabled. Enable it before using tasks.* tools.',
      'SERVICE_NOT_ENABLED'
    );
  }
}

function resolveTasklistId_(tasklistId) {
  if (tasklistId) {
    return tasklistId;
  }

  var result = Tasks.Tasklists.list({
    maxResults: 1
  }) || {};
  var tasklists = result.items || [];

  if (!tasklists.length) {
    throw createError_('No Google Tasklists are available for this account.', 'NOT_FOUND');
  }

  return tasklists[0].id;
}

function serializeTasklist_(tasklist) {
  return {
    id: tasklist.id,
    title: tasklist.title,
    updated: tasklist.updated || null,
    selfLink: tasklist.selfLink || null
  };
}

function serializeTask_(task) {
  return {
    id: task.id,
    title: task.title,
    notes: task.notes || '',
    status: task.status || 'needsAction',
    due: task.due || null,
    completed: task.completed || null,
    updated: task.updated || null,
    deleted: !!task.deleted,
    hidden: !!task.hidden,
    parent: task.parent || null,
    position: task.position || null,
    selfLink: task.selfLink || null
  };
}

// ---------------------------------------------------------------------------
// Sheets helpers
// ---------------------------------------------------------------------------

function getSheet_(spreadsheet, sheetName) {
  if (!sheetName) {
    return spreadsheet.getSheets()[0];
  }

  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw createError_('Sheet not found: ' + sheetName, 'NOT_FOUND');
  }

  return sheet;
}

function normalize2DValues_(values) {
  if (!Array.isArray(values) || !values.length) {
    throw createError_('values must be a non-empty array.', 'INVALID_ARGUMENT');
  }

  var normalized = Array.isArray(values[0]) ? values : [values];
  var width = normalized[0].length;

  if (!width) {
    throw createError_('values must contain at least one column.', 'INVALID_ARGUMENT');
  }

  for (var i = 0; i < normalized.length; i++) {
    if (!Array.isArray(normalized[i])) {
      throw createError_('values must be a rectangular 2D array.', 'INVALID_ARGUMENT');
    }
    if (normalized[i].length !== width) {
      throw createError_('All rows in values must have the same number of columns.', 'INVALID_ARGUMENT');
    }
  }

  return normalized;
}

function serializeTableValues_(values) {
  return mapArray_(values, function(row) {
    return mapArray_(row, serializeValue_);
  });
}

function serializeValue_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (value === undefined) {
    return null;
  }
  return value;
}

// ---------------------------------------------------------------------------
// Generic utility helpers
// ---------------------------------------------------------------------------

function mapArray_(items, mapper) {
  var results = [];
  for (var i = 0; i < items.length; i++) {
    results.push(mapper(items[i], i));
  }
  return results;
}

function collectIterator_(iterator, limit, mapper) {
  var items = [];
  while (iterator.hasNext() && items.length < limit) {
    items.push(mapper(iterator.next()));
  }
  return items;
}

function copyOptionalArg_(source, target, key) {
  if (hasOwn_(source, key) && source[key] !== undefined && source[key] !== null) {
    target[key] = source[key];
  }
}

function hasOwn_(obj, key) {
  return Object.prototype.hasOwnProperty.call(obj, key);
}

function toNonNegativeInt_(value, fallback) {
  if (value === undefined || value === null || value === '') {
    return fallback;
  }
  var parsed = Number(value);
  if (isNaN(parsed) || parsed < 0) {
    throw createError_('Expected a non-negative number.', 'INVALID_ARGUMENT');
  }
  return Math.floor(parsed);
}

function clampLimit_(value, fallback, max) {
  var parsed = value;
  if (parsed === undefined || parsed === null || parsed === '') {
    parsed = fallback;
  }
  parsed = Number(parsed);
  if (isNaN(parsed) || parsed <= 0) {
    throw createError_('Limit must be a positive number.', 'INVALID_ARGUMENT');
  }
  return Math.min(Math.floor(parsed), max || MAX_LIMIT);
}

function parseDate_(value, fieldName) {
  var date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) {
    throw createError_('Invalid date for ' + fieldName + ': ' + value, 'INVALID_ARGUMENT');
  }
  return date;
}

function truncateText_(text, maxLength) {
  if (text === null || text === undefined) {
    return '';
  }

  var normalized = String(text);
  if (normalized.length <= maxLength) {
    return normalized;
  }

  return normalized.substring(0, maxLength) + '\n...[truncated]';
}

function escapeSearchValue_(value) {
  return String(value).replace(/"/g, '\\"');
}

function safeUserEmail_(user) {
  try {
    return user ? user.getEmail() : null;
  } catch (err) {
    return null;
  }
}
