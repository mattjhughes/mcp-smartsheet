/**
 * Smartsheet MCP Server
 * 
 * A Model Context Protocol server implementation for Smartsheet API integration.
 * 
 * SECURITY NOTE:
 * - Never commit API tokens to the repository
 * - This server should be run locally and only connected to trusted MCP clients
 * - Ensure proper access controls are set up on your Smartsheet account
 */
import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

// Smartsheet API client
import * as smartsheet from "smartsheet";
import axios from "axios";

// TypeScript interfaces for Smartsheet API
interface SmartsheetCell {
  columnId: number;
  value: string | number | boolean;
  displayValue?: string;
}

interface SmartsheetRow {
  id?: number;
  sheetId?: string;
  rowNumber?: number;
  cells: SmartsheetCell[];
  toBottom?: boolean;
}

interface SmartsheetColumn {
  id: number;
  title: string;
  type: string;
  index: number;
  primary?: boolean;
}

interface SmartsheetSheet {
  id: string;
  name: string;
  columns: SmartsheetColumn[];
  rows: SmartsheetRow[];
}

// Schema definitions for Smartsheet operations (for Zod validation)
const SheetSchema = z.object({
  sheetId: z.string().describe("The ID of the sheet to read or update"),
});

const RowSchema = z.object({
  sheetId: z.string().describe("The ID of the sheet to add rows to"),
  cells: z.array(
    z.object({
      columnId: z.number().describe("The ID of the column"),
      value: z.union([z.string(), z.number(), z.boolean()]).describe("The value to set in the cell"),
    })
  ).describe("The cells to include in the new row"),
});

/**
 * Smartsheet MCP Server
 * 
 * This server provides Model Context Protocol access to Smartsheet data
 * with tools for listing, reading, and updating sheets.
 */
// Interface for our API client
interface SmartsheetApiClient {
  sheets: {
    listSheets: () => Promise<any>;
    getSheet: (options: {id: string, include?: string}) => Promise<any>;
    addRows: (options: {sheetId: string, body: SmartsheetRow[]}) => Promise<any>;
  };
  search: {
    searchAll: (options: {query: string}) => Promise<any>;
  };
}

class SmartsheetMcpServer {
  private server: McpServer;
  public client: SmartsheetApiClient;
  private apiToken: string = ''; // Store API token for direct calls
  
  constructor() {
    // Initialize MCP server
    this.server = new McpServer({
      name: "Smartsheet MCP Server",
      version: "1.0.0"
    });
    
    // Create a placeholder client - will be properly initialized in setApiToken
    this.client = {
      sheets: {
        listSheets: async () => { throw new Error("Client not initialized"); },
        getSheet: async () => { throw new Error("Client not initialized"); },
        addRows: async () => { throw new Error("Client not initialized"); }
      },
      search: {
        searchAll: async () => { throw new Error("Client not initialized"); }
      }
    };

    this.setupResources();
    this.setupTools();
  }

  /**
   * Initialize the Smartsheet client with the provided API key
   * @param token The Smartsheet API token to use for authentication
   * @throws Error if token cannot be used to initialize the client
   */
  public setApiToken(token: string) {
    try {
      console.error("Initializing Smartsheet client...");
      
      // Store the token for direct API calls
      this.apiToken = token;
      
      // Create a wrapper using axios for direct API calls
      this.client = this.createApiClient();
      
      console.error("Smartsheet direct API client successfully initialized");
    } catch (error) {
      console.error("Error initializing Smartsheet client:", error);
      throw error;
    }
  }
  
  /**
   * Create a Smartsheet API client using axios for direct API calls
   * This separates the client creation logic for better code organization
   * @returns A fully configured SmartsheetApiClient
   */
  private createApiClient(): SmartsheetApiClient {
    return {
      sheets: {
        listSheets: async () => {
          console.error("Calling listSheets via direct API");
          const apiUrl = "https://api.smartsheet.com/2.0/sheets";
          console.error(`API call: GET ${apiUrl}`);
          
          const response = await this.apiGet(apiUrl);
          this.logResponse("listSheets", response);
          return response.data;
        },
        
        getSheet: async (options: any) => {
          const sheetId = options.id;
          console.error(`Calling getSheet for sheet ${sheetId} via direct API`);
          
          // Add include parameters if needed
          let includeParam = '';
          if (options.include) {
            includeParam = `?include=${options.include}`;
          }
          const apiUrl = `https://api.smartsheet.com/2.0/sheets/${sheetId}${includeParam}`;
          console.error(`API call: GET ${apiUrl}`);
          
          const response = await this.apiGet(apiUrl);
          this.logResponse("getSheet", response);
          return response.data;
        },
        
        addRows: async (options: any) => {
          const sheetId = options.sheetId;
          console.error(`Calling addRows for sheet ${sheetId} via direct API`);
          const apiUrl = `https://api.smartsheet.com/2.0/sheets/${sheetId}/rows`;
          console.error(`API call: POST ${apiUrl}`);
          console.error(`Adding ${options.body.length} row(s)`);
          
          const response = await this.apiPost(apiUrl, options.body);
          this.logResponse("addRows", response);
          return response.data;
        }
      },
      
      search: {
        searchAll: async (options: any) => {
          console.error(`Calling searchAll with query ${options.query} via direct API`);
          const apiUrl = `https://api.smartsheet.com/2.0/search?query=${encodeURIComponent(options.query)}`;
          console.error(`API call: GET ${apiUrl}`);
          
          const response = await this.apiGet(apiUrl);
          this.logResponse("searchAll", response);
          return response.data;
        }
      }
    };
  }
  
  /**
   * Helper method for GET requests
   */
  private async apiGet(url: string) {
    return await axios.get<any>(url, { headers: this.getHeaders() });
  }
  
  /**
   * Helper method for POST requests
   */
  private async apiPost<T = any>(url: string, data: any) {
    return await axios.post<T>(url, data, { headers: this.getHeaders() });
  }
  
  /**
   * Helper method for PUT requests
   */
  private async apiPut<T = any>(url: string, data: any) {
    return await axios.put<T>(url, data, { headers: this.getHeaders() });
  }
  
  /**
   * Helper to get authorization headers
   */
  private getHeaders() {
    return {
      'Authorization': `Bearer ${this.apiToken}`,
      'Content-Type': 'application/json'
    };
  }
  
  /**
   * Helper to log API responses
   */
  private logResponse(operation: string, response: any) {
    console.error(`${operation} response status: ${response.status}`);
    // Log only the count of data items, not the actual data, to prevent accidental data exposure
    if (response.data) {
      if (Array.isArray(response.data)) {
        console.error(`${operation} response contains ${response.data.length} items`);
      } else if (response.data.results && Array.isArray(response.data.results)) {
        console.error(`${operation} response contains ${response.data.results.length} results`);
      } else if (response.data.data && Array.isArray(response.data.data)) {
        console.error(`${operation} response contains ${response.data.data.length} data items`);
      } else {
        console.error(`${operation} response contains data (details omitted for security)`);
      }
    }
  }

  /**
   * Set up MCP resources (read-only access points)
   * These allow AI to read data from Smartsheet
   */
  private setupResources() {
    // Resource to list all available sheets
    this.server.resource(
      "sheets",
      "smartsheet://sheets",
      async (uri) => {
        try {
          this.ensureClientInitialized();
          const response = await this.client.sheets.listSheets();
          
          return {
            contents: [{
              uri: uri.href,
              text: JSON.stringify(response.data.map((sheet: any) => ({
                id: sheet.id,
                name: sheet.name,
                permalink: sheet.permalink
              })))
            }]
          };
        } catch (error: unknown) {
          console.error("Error listing sheets:", error);
          return {
            contents: [{
              uri: uri.href,
              text: JSON.stringify({ error: "Failed to list sheets" })
            }]
          };
        }
      }
    );

    // Resource to read a specific sheet
    this.server.resource(
      "sheet",
      new ResourceTemplate("smartsheet://sheets/{sheetId}", { list: undefined }),
      async (uri, { sheetId }) => {
        try {
          this.ensureClientInitialized();
          // Ensure sheetId is a string
          const response = await this.client.sheets.getSheet({
            id: String(sheetId),
            include: "rowPermalinks"  // Use string format for consistency
          });
          
          return {
            contents: [{
              uri: uri.href,
              text: JSON.stringify({
                id: response.id,
                name: response.name,
                columns: response.columns.map((col: any) => ({
                  id: col.id,
                  title: col.title,
                  type: col.type
                })),
                rows: response.rows.map((row: any) => ({
                  id: row.id,
                  cells: row.cells.map((cell: any) => ({
                    columnId: cell.columnId,
                    value: cell.value,
                    displayValue: cell.displayValue
                  }))
                }))
              })
            }]
          };
        } catch (error: unknown) {
          console.error(`Error fetching sheet ${sheetId}:`, error);
          return {
            contents: [{
              uri: uri.href,
              text: JSON.stringify({ error: `Failed to fetch sheet ${sheetId}` })
            }]
          };
        }
      }
    );
  }

  /**
   * Set up MCP tools (actions that modify data)
   * These allow AI to perform actions on Smartsheet
   */
  private setupTools() {
    // Tool to list all sheets
    this.server.tool(
      "listSheets",
      {},
      async () => {
        try {
          this.ensureClientInitialized();
          const response = await this.client.sheets.listSheets();
          
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                sheets: response.data ? response.data.map((sheet: any) => ({
                  id: sheet.id,
                  name: sheet.name,
                  accessLevel: sheet.accessLevel,
                  permalink: sheet.permalink
                })) : []
              })
            }]
          };
        } catch (error: unknown) {
          console.error("Error listing sheets:", error);
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                error: "Failed to list sheets",
                details: error instanceof Error ? error.message : String(error)
              })
            }]
          };
        }
      }
    );
    
    // Tool to get a sheet with its contents
    this.server.tool(
      "getSheet",
      {
        sheetId: z.string().describe("The ID of the sheet to read"),
      },
      async ({ sheetId }) => {
        try {
          this.ensureClientInitialized();
          // Ensure sheetId is a string
          const response = await this.client.sheets.getSheet({
            id: String(sheetId),
            include: "rowPermalinks"  // Use string format for consistency
          });
          
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                id: response.id,
                name: response.name,
                columns: response.columns.map((col: any) => ({
                  id: col.id,
                  title: col.title,
                  type: col.type,
                  index: col.index,
                  primary: col.primary || false
                })),
                rows: response.rows.map((row: any) => ({
                  id: row.id,
                  rowNumber: row.rowNumber,
                  cells: row.cells.map((cell: any) => ({
                    columnId: cell.columnId,
                    value: cell.value,
                    displayValue: cell.displayValue
                  }))
                }))
              })
            }]
          };
        } catch (error: unknown) {
          console.error(`Error fetching sheet ${sheetId}:`, error);
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                error: `Failed to fetch sheet ${sheetId}`,
                details: error instanceof Error ? error.message : String(error)
              })
            }]
          };
        }
      }
    );

    // Tool to add a row to a sheet
    this.server.tool(
      "addRow",
      {
        sheetId: z.string().describe("The ID of the sheet to add a row to"),
        cells: z.array(
          z.object({
            columnId: z.number().describe("The ID of the column"),
            value: z.union([z.string(), z.number(), z.boolean()]).describe("The value to set in the cell"),
          })
        ).describe("The cells to include in the new row"),
      },
      async ({ sheetId, cells }) => {
        try {
          this.ensureClientInitialized();
          const rowToAdd = {
            toBottom: true,
            cells: cells.map(cell => ({
              columnId: cell.columnId,
              value: cell.value
            }))
          };

          const response = await this.client.sheets.addRows({
            sheetId: sheetId,
            body: [rowToAdd]
          });

          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                message: "Row added successfully",
                rowId: response.result ? response.result[0].id : null
              })
            }]
          };
        } catch (error: unknown) {
          console.error("Error adding row:", error);
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                error: "Failed to add row",
                details: error instanceof Error ? error.message : String(error)
              })
            }]
          };
        }
      }
    );

    // Tool to update a row in a sheet
    this.server.tool(
      "updateRow",
      {
        sheetId: z.string().describe("The ID of the sheet containing the row"),
        rowId: z.number().describe("The ID of the row to update"),
        cells: z.array(
          z.object({
            columnId: z.number().describe("The ID of the column"),
            value: z.union([z.string(), z.number(), z.boolean()]).describe("The new value to set in the cell"),
          })
        ).describe("The cells to update in the row"),
      },
      async ({ sheetId, rowId, cells }) => {
        try {
          this.ensureClientInitialized();
          const rowToUpdate = {
            id: rowId,
            cells: cells.map(cell => ({
              columnId: cell.columnId,
              value: cell.value
            }))
          };

          // Direct API call to update a row
          const apiUrl = `https://api.smartsheet.com/2.0/sheets/${sheetId}/rows`;
          console.error(`API call: PUT ${apiUrl}`);
          console.error(`Updating row ${rowId} with ${cells.length} cell changes`);
          
          const response = await this.apiPut(apiUrl, [rowToUpdate]);
          this.logResponse("updateRow", response);

          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                message: "Row updated successfully",
                result: response.data
              })
            }]
          };
        } catch (error: unknown) {
          console.error(`Error updating row ${rowId}:`, error);
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                error: `Failed to update row ${rowId}`,
                details: error instanceof Error ? error.message : String(error)
              })
            }]
          };
        }
      }
    );

    // Tool to list columns in a sheet
    this.server.tool(
      "listColumns",
      {
        sheetId: z.string().describe("The ID of the sheet to get columns from"),
      },
      async ({ sheetId }) => {
        try {
          this.ensureClientInitialized();
          // Ensure sheetId is a string
          const response = await this.client.sheets.getSheet({
            id: String(sheetId),
            include: "rowPermalinks"  // Use string format for consistency
          });
          
          const columns = response.columns.map((col: any) => ({
            id: col.id,
            title: col.title,
            type: col.type,
            index: col.index,
            primary: col.primary || false
          }));

          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                sheetName: response.name,
                columns: columns
              })
            }]
          };
        } catch (error: unknown) {
          console.error(`Error fetching columns for sheet ${sheetId}:`, error);
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                error: `Failed to fetch columns for sheet ${sheetId}`,
                details: error instanceof Error ? error.message : String(error)
              })
            }]
          };
        }
      }
    );

    // Tool to search for data in sheets
    this.server.tool(
      "searchSheets",
      {
        query: z.string().describe("The text to search for in sheets"),
      },
      async ({ query }) => {
        try {
          this.ensureClientInitialized();
          
          // Use our client's searchAll method which now uses direct API calls
          const response = await this.client.search.searchAll({ query });
          
          // The results are in response.results for the Smartsheet API
          const searchResults = response && response.results ? response.results : [];
          console.error(`Got ${searchResults.length} search results`);
          
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                results: searchResults.map((result: any) => ({
                  objectType: result.objectType,
                  text: result.text || result.name || result.title,
                  sheetId: result.parentObjectId || result.parentId || null,
                  objectId: result.objectId || result.id
                }))
              })
            }]
          };
        } catch (error: unknown) {
          console.error(`Error searching sheets for "${query}":`, error);
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                error: `Failed to search sheets for "${query}"`,
                details: error instanceof Error ? error.message : String(error)
              })
            }]
          };
        }
      }
    );
  }

  /**
   * Start the MCP server using the stdio transport
   * This method connects the server to the MCP client using standard I/O
   * 
   * @returns A promise that resolves when the server is started
   * @throws Error if the server fails to start
   */
  public async start() {
    // Check if client is initialized
    if (!this.client) {
      console.error("Warning: Smartsheet client not initialized. Waiting for API token...");
    }

    const transport = new StdioServerTransport();
    console.log("Starting Smartsheet MCP server...");
    await this.server.connect(transport);
    console.log("Smartsheet MCP server started");
  }
  
  /**
   * Check if the client is initialized
   */
  private ensureClientInitialized() {
    if (!this.client) {
      console.error("Smartsheet client not initialized. API token must be provided in the MCP config.");
      throw new Error("Smartsheet client not initialized. API token must be provided in the MCP config.");
    }
  }
}

// Main entry point
async function main() {
  try {
    // Redirect console.log to stderr so it doesn't interfere with MCP protocol
    const originalConsoleLog = console.log;
    console.log = function(...args: any[]) {
      console.error(...args);
    };
    
    // Read the API token from the environment variable
    // Claude's MCP implementation passes config as environment variables
    const apiToken = process.env.API_TOKEN || process.env.SMARTSHEET_API_TOKEN;
    
    if (apiToken) {
      console.error("Found API token in environment variables");
    } else {
      console.error("No API token found in environment variables");
      // Security note: We don't log environment variable names for security
    }
    
    const server = new SmartsheetMcpServer();
    
    // If we have an API token from environment variables, use it
    if (apiToken) {
      server.setApiToken(apiToken);
    } else {
      // Otherwise, try to get it from stdin messages
      let isConfigured = false;
      
      // Set up a promise that will resolve when we receive the API token
      const configPromise = new Promise<void>((resolve, reject) => {
        // Handle initialization timeouts
        const timeoutId = setTimeout(() => {
          if (!isConfigured) {
            reject(new Error("Timed out waiting for API token configuration"));
          }
        }, 30000); // 30 second timeout
        
        // Read configuration from MCP
        process.stdin.on('data', (data) => {
          try {
            console.error("Received data:", data.toString().substring(0, 100) + "...");
            const message = JSON.parse(data.toString());
            
            console.error("Parsed message:", JSON.stringify(message).substring(0, 100) + "...");
            
            // Check if this is the init message from the MCP host
            if (message.method === 'initialize' && message.params) {
              console.error("Received initialize message");
              // Don't resolve yet - just acknowledge we got the init message
              // But look for config in the init message
              if (message.params.config) {
                console.error("Found config in initialize params:", JSON.stringify(message.params.config));
                const initApiToken = message.params.config.apiToken;
                if (initApiToken) {
                  server.setApiToken(initApiToken);
                  isConfigured = true;
                  clearTimeout(timeoutId);
                  resolve();
                  return;
                }
              }
            }
            
            // Check for config in regular jsonrpc message
            if (message.jsonrpc && message.method === "init" && message.params) {
              console.error("Found MCP init message with params:", JSON.stringify(message.params));
              
              // Try various config locations based on MCP implementations
              const config = message.params.config || message.params;
              const msgApiToken = config.apiToken || (config.smartsheet && config.smartsheet.apiToken);
              
              if (msgApiToken) {
                console.error("Found API token in config");
                // Initialize Smartsheet client with the token
                server.setApiToken(msgApiToken);
                isConfigured = true;
                clearTimeout(timeoutId);
                resolve();
                return;
              }
            }
            
            // Handle config coming directly 
            if (message.config && message.config.apiToken) {
              console.error("Found direct config with API token");
              server.setApiToken(message.config.apiToken);
              isConfigured = true;
              clearTimeout(timeoutId);
              resolve();
              return;
            }
          } catch (error) {
            console.error("Error parsing MCP message:", error);
          }
        });
      });
      
      // Start handling standard input immediately
      process.stdin.resume();
      
      try {
        // Wait for configuration before starting server
        await Promise.race([
          configPromise,
          // Start with a delay to allow time for config to arrive
          new Promise(resolve => setTimeout(resolve, 5000))
        ]);
      } catch (error) {
        console.error("Configuration error:", error);
        // Continue anyway and let the client checks handle it
      }
    }
    
    // For debug purposes, set a fake client if none exists yet
    if (!server.client) {
      console.error("WARNING: No API token received. Creating mock client for debugging");
      server.setApiToken("DEBUG_TOKEN");
    }
    
    // Start the server
    await server.start();
  } catch (error) {
    console.error("Failed to start server:", error);
    process.exit(1);
  }
}

main().catch(error => {
  console.error("Unhandled error:", error);
  process.exit(1);
});