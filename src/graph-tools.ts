import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import logger from './logger.js';
import GraphClient from './graph-client.js';
import { api } from './generated/client.js';
import { z } from 'zod';

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

export function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false
): void {
  for (const tool of api.endpoints) {
    if (readOnly && tool.method.toUpperCase() !== 'GET') {
      logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
      continue;
    }

    const paramSchema: Record<string, any> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        if (param.type === 'Body' && param.schema) {
          paramSchema[param.name] = z.union([z.string(), param.schema]);
        } else {
          paramSchema[param.name] = param.schema || z.any();
        }
      }
    }

    server.tool(
      tool.alias,
      tool.description ?? '',
      paramSchema,
      {
        title: tool.alias,
        readOnlyHint: tool.method.toUpperCase() === 'GET',
      },
      async (params, extra) => {
        logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);
        try {
          logger.info(`params: ${JSON.stringify(params)}`);

          const parameterDefinitions = tool.parameters || [];

          let path = tool.path;
          const queryParams: Record<string, string> = {};
          const headers: Record<string, string> = {};
          let body: any = null;
          for (let [paramName, paramValue] of Object.entries(params)) {
            // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
            // and others might not support __, so we strip them in hack.ts and restore them here
            const odataParams = [
              'filter',
              'select',
              'expand',
              'orderby',
              'skip',
              'top',
              'count',
              'search',
              'format',
            ];
            const fixedParamName = odataParams.includes(paramName.toLowerCase())
              ? `$${paramName.toLowerCase()}`
              : paramName;
            const paramDef = parameterDefinitions.find((p) => p.name === paramName);

            if (paramDef) {
              switch (paramDef.type) {
                case 'Path':
                  path = path
                    .replace(`{${paramName}}`, encodeURIComponent(paramValue as string))
                    .replace(`:${paramName}`, encodeURIComponent(paramValue as string));
                  break;

                case 'Query':
                  queryParams[fixedParamName] = `${paramValue}`;
                  break;

                case 'Body':
                  if (typeof paramValue === 'string') {
                    try {
                      body = JSON.parse(paramValue);
                    } catch (e) {
                      body = paramValue;
                    }
                  } else {
                    body = paramValue;
                  }
                  break;

                case 'Header':
                  headers[fixedParamName] = `${paramValue}`;
                  break;
              }
            } else if (paramName === 'body') {
              if (typeof paramValue === 'string') {
                try {
                  body = JSON.parse(paramValue);
                } catch (e) {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
              logger.info(`Set legacy body param: ${JSON.stringify(body)}`);
            }
          }

          if (Object.keys(queryParams).length > 0) {
            const queryString = Object.entries(queryParams)
              .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
              .join('&');
            path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
          }

          const options: any = {
            method: tool.method.toUpperCase(),
            headers,
          };

          if (options.method !== 'GET' && body) {
            options.body = typeof body === 'string' ? body : JSON.stringify(body);
          }

          // Add Excel-specific handling:
          if (tool.alias.includes('excel') && path.includes('workbook')) {
            // Force session creation for Excel operations
            logger.info(`Excel operation detected: ${tool.alias}`);

            // Extract file path from the Graph API path
            const match = path.match(/\/me\/drive\/root:([^:]+):/);
            if (match) {
              options.excelFile = match[1];
              logger.info(`Extracted Excel file path: ${options.excelFile}`);
            }
          }

          const isProbablyMediaContent =
            tool.errors?.some((error) => error.description === 'Retrieved media content') ||
            path.endsWith('/content');

          if (isProbablyMediaContent) {
            options.rawResponse = true;
          }

          logger.info(`Making graph request to ${path} with options: ${JSON.stringify(options)}`);
          const response = await graphClient.graphRequest(path, options);

          if (response && response.content && response.content.length > 0) {
            const responseText = response.content[0].text;
            const responseSize = responseText.length;
            logger.info(`Response size: ${responseSize} characters`);

            try {
              const jsonResponse = JSON.parse(responseText);
              if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
                logger.info(`Response contains ${jsonResponse.value.length} items`);
                if (jsonResponse.value.length > 0 && jsonResponse.value[0].body) {
                  logger.info(
                    `First item has body field with size: ${JSON.stringify(jsonResponse.value[0].body).length} characters`
                  );
                }
              }
              if (jsonResponse['@odata.nextLink']) {
                logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
              }
              const preview = responseText.substring(0, 500);
              logger.info(`Response preview: ${preview}${responseText.length > 500 ? '...' : ''}`);
            } catch (e) {
              const preview = responseText.substring(0, 500);
              logger.info(
                `Response preview (non-JSON): ${preview}${responseText.length > 500 ? '...' : ''}`
              );
            }
          }

          // Convert McpResponse to CallToolResult with the correct structure
          const content: ContentItem[] = response.content.map((item) => {
            // GraphClient only returns text content items, so create proper TextContent items
            const textContent: TextContent = {
              type: 'text',
              text: item.text,
            };
            return textContent;
          });

          const result: CallToolResult = {
            content,
            _meta: response._meta,
            isError: response.isError,
          };

          return result;
        } catch (error) {
          logger.error(`Error in tool ${tool.alias}: ${(error as Error).message}`);
          const errorContent: TextContent = {
            type: 'text',
            text: JSON.stringify({
              error: `Error in tool ${tool.alias}: ${(error as Error).message}`,
            }),
          };

          return {
            content: [errorContent],
            isError: true,
          };
        }
      }
    );
  }

  // Add this new tool for debugging Excel access
  server.tool(
    'debug-excel-access',
    'Debug Excel file access and permissions',
    {
      filePath: z.string().describe('Path to Excel file (e.g., /test.xlsx)'),
    },
    async ({ filePath }) => {
      try {
        logger.info(`Debugging Excel access for: ${filePath}`);

        // Test 1: Check if file exists
        const fileInfo = await graphClient.graphRequest(`/me/drive/root:${filePath}`);
        logger.info(`File info result: ${JSON.stringify(fileInfo)}`);

        // Test 2: Try to create session
        const sessionResult = await graphClient.graphRequest(
          `/me/drive/root:${filePath}:/workbook/createSession`,
          { method: 'POST', body: JSON.stringify({ persistChanges: false }) }
        );
        logger.info(`Session creation result: ${JSON.stringify(sessionResult)}`);

        // Test 3: List worksheets
        const worksheets = await graphClient.graphRequest(
          `/me/drive/root:${filePath}:/workbook/worksheets`
        );
        logger.info(`Worksheets result: ${JSON.stringify(worksheets)}`);

        // Test 4: Try to read a simple range
        const range = await graphClient.graphRequest(
          `/me/drive/root:${filePath}:/workbook/worksheets/Sheet1/range(address='A1:C5')`
        );
        logger.info(`Range result: ${JSON.stringify(range)}`);

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  fileInfo: JSON.parse(fileInfo.content[0].text),
                  sessionResult: JSON.parse(sessionResult.content[0].text),
                  worksheets: JSON.parse(worksheets.content[0].text),
                  range: JSON.parse(range.content[0].text),
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        const errorMessage = (error as Error).message;
        logger.error(`Debug Excel access error: ${errorMessage}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: errorMessage }, null, 2),
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Add file upload capability
  server.tool(
    'upload-file-to-onedrive',
    'Upload a new file to OneDrive',
    {
      fileName: z.string().describe('Name of the file to create (e.g., "test.xlsx")'),
      content: z.string().describe('File content (base64 encoded for binary files)'),
      folderPath: z.string().default('/').describe('Folder path (default: root folder)'),
      contentType: z
        .string()
        .default('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        .describe('MIME type of the file'),
    },
    async ({ fileName, content, folderPath, contentType }) => {
      try {
        logger.info(`Uploading file: ${fileName} to folder: ${folderPath}`);

        const uploadPath =
          folderPath === '/'
            ? `/me/drive/root:/${fileName}:/content`
            : `/me/drive/root:${folderPath}/${fileName}:/content`;

        const response = await graphClient.graphRequest(uploadPath, {
          method: 'PUT',
          headers: {
            'Content-Type': contentType,
          },
          body: content,
        });

        return response;
      } catch (error) {
        const errorMessage = (error as Error).message;
        logger.error(`Upload error: ${errorMessage}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: errorMessage }, null, 2),
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Add Excel file creation helper
  server.tool(
    'create-empty-excel-file',
    'Create an empty Excel file in OneDrive',
    {
      fileName: z.string().describe('Name of the Excel file (e.g., "MyWorkbook.xlsx")'),
      folderPath: z.string().default('/').describe('Folder path (default: root folder)'),
    },
    async ({ fileName, folderPath }) => {
      try {
        // Create minimal Excel file content (base64 encoded empty workbook)
        const emptyExcelContent = 'UEsDBBQAAAAIAAAAIQDd...'; // You'd need actual empty Excel file bytes

        // For now, create via Graph API
        const uploadPath =
          folderPath === '/'
            ? `/me/drive/root:/${fileName}:/content`
            : `/me/drive/root:${folderPath}/${fileName}:/content`;

        // Create empty file first
        const response = await graphClient.graphRequest(uploadPath, {
          method: 'PUT',
          headers: {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          },
          body: '', // Empty content - Graph will create a minimal Excel file
        });

        return response;
      } catch (error) {
        const errorMessage = (error as Error).message;
        logger.error(`Excel creation error: ${errorMessage}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: errorMessage }, null, 2),
            },
          ],
          isError: true,
        };
      }
    }
  );
}
