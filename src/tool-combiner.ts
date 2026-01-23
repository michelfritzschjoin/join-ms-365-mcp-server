/**
 * Tool Combination Engine for parallel/sequential tool execution with error handling
 */

import GraphClient from './graph-client.js';
import logger from './logger.js';

export interface ToolCall {
  name: string;
  params: Record<string, unknown>;
  dependsOn?: string[]; // Names of tools that must complete first
}

export interface ToolResult {
  toolName: string;
  success: boolean;
  data: unknown;
  error?: string;
  executionTime: number;
}

export interface ToolCombinationResult {
  results: ToolResult[];
  totalTime: number;
  successCount: number;
  failureCount: number;
  aggregatedData: unknown[];
}

export class ToolCombiner {
  private graphClient: GraphClient;
  private readonly maxConcurrent: number;

  constructor(graphClient: GraphClient) {
    this.graphClient = graphClient;
    this.maxConcurrent = parseInt(process.env.MS365_MCP_MAX_CONCURRENT_TOOLS || '5', 10);
  }

  /**
   * Execute tools in parallel/sequential order based on dependencies
   */
  async executeTools(tools: ToolCall[]): Promise<ToolCombinationResult> {
    const startTime = Date.now();
    const results: ToolResult[] = [];
    const executed = new Set<string>();
    const toolMap = new Map<string, ToolCall>();

    // Build tool map
    for (const tool of tools) {
      toolMap.set(tool.name, tool);
    }

    // Execute tools respecting dependencies
    while (executed.size < tools.length) {
      // Find tools that can be executed (no dependencies or all dependencies executed)
      const readyTools = tools.filter(
        (tool) =>
          !executed.has(tool.name) &&
          (!tool.dependsOn || tool.dependsOn.every((dep) => executed.has(dep)))
      );

      if (readyTools.length === 0) {
        // Circular dependency or missing dependency
        logger.warn('No ready tools found, possible circular dependency');
        break;
      }

      // Execute ready tools in parallel (limited by maxConcurrent)
      const batch = readyTools.slice(0, this.maxConcurrent);
      const batchResults = await Promise.allSettled(batch.map((tool) => this.executeTool(tool)));

      // Process batch results
      for (let i = 0; i < batch.length; i++) {
        const result = batchResults[i];
        if (result.status === 'fulfilled') {
          results.push(result.value);
          executed.add(batch[i].name);
        } else {
          // Tool execution failed
          const toolResult: ToolResult = {
            toolName: batch[i].name,
            success: false,
            data: null,
            error: result.reason?.message || 'Unknown error',
            executionTime: 0,
          };
          results.push(toolResult);
          executed.add(batch[i].name);
        }
      }
    }

    const totalTime = Date.now() - startTime;
    const successCount = results.filter((r) => r.success).length;
    const failureCount = results.filter((r) => !r.success).length;
    const aggregatedData = results.filter((r) => r.success).map((r) => r.data);

    return {
      results,
      totalTime,
      successCount,
      failureCount,
      aggregatedData,
    };
  }

  /**
   * Execute a single tool
   */
  private async executeTool(tool: ToolCall): Promise<ToolResult> {
    const startTime = Date.now();
    try {
      logger.info(`Executing tool: ${tool.name}`);

      // Map tool name to Graph API endpoint
      const endpoint = this.mapToolToEndpoint(tool.name, tool.params);
      if (!endpoint) {
        throw new Error(`Unknown tool: ${tool.name}`);
      }

      // Execute Graph API request
      const result = await this.graphClient.makeRequest(endpoint.path, endpoint.options);

      const executionTime = Date.now() - startTime;
      logger.info(`Tool ${tool.name} completed in ${executionTime}ms`);

      return {
        toolName: tool.name,
        success: true,
        data: result,
        executionTime,
      };
    } catch (error) {
      const executionTime = Date.now() - startTime;
      logger.error(`Tool ${tool.name} failed: ${error}`);

      return {
        toolName: tool.name,
        success: false,
        data: null,
        error: (error as Error).message,
        executionTime,
      };
    }
  }

  /**
   * Map tool name to Graph API endpoint
   */
  private mapToolToEndpoint(
    toolName: string,
    params: Record<string, unknown>
  ): { path: string; options: Record<string, unknown> } | null {
    // This is a simplified mapping - in reality, you'd use the tool registry
    // For now, we'll handle common patterns

    if (toolName.startsWith('list-')) {
      const entity = toolName.replace('list-', '');
      return {
        path: `/${entity}`,
        options: {
          method: 'GET',
          queryParams: params,
        },
      };
    }

    if (toolName.startsWith('get-')) {
      const entity = toolName.replace('get-', '');
      const id = params['id'] || params['entityId'];
      if (!id) {
        return null;
      }
      return {
        path: `/${entity}/${id}`,
        options: {
          method: 'GET',
          queryParams: params,
        },
      };
    }

    if (toolName.startsWith('search-')) {
      return {
        path: '/search/query',
        options: {
          method: 'POST',
          body: JSON.stringify(params),
        },
      };
    }

    // Default: try to use tool name as path
    return {
      path: `/${toolName}`,
      options: {
        method: 'GET',
        queryParams: params,
      },
    };
  }

  /**
   * Execute tools sequentially
   */
  async executeSequential(tools: ToolCall[]): Promise<ToolCombinationResult> {
    const startTime = Date.now();
    const results: ToolResult[] = [];

    for (const tool of tools) {
      const result = await this.executeTool(tool);
      results.push(result);

      // Stop on first failure if configured
      if (!result.success && process.env.MS365_MCP_STOP_ON_ERROR === 'true') {
        break;
      }
    }

    const totalTime = Date.now() - startTime;
    const successCount = results.filter((r) => r.success).length;
    const failureCount = results.filter((r) => !r.success).length;
    const aggregatedData = results.filter((r) => r.success).map((r) => r.data);

    return {
      results,
      totalTime,
      successCount,
      failureCount,
      aggregatedData,
    };
  }

  /**
   * Execute tools in parallel (no dependencies)
   */
  async executeParallel(tools: ToolCall[]): Promise<ToolCombinationResult> {
    const startTime = Date.now();

    // Execute all tools in parallel (limited by maxConcurrent)
    const batches: ToolCall[][] = [];
    for (let i = 0; i < tools.length; i += this.maxConcurrent) {
      batches.push(tools.slice(i, i + this.maxConcurrent));
    }

    const allResults: ToolResult[] = [];
    for (const batch of batches) {
      const batchResults = await Promise.allSettled(batch.map((tool) => this.executeTool(tool)));

      for (let i = 0; i < batch.length; i++) {
        const result = batchResults[i];
        if (result.status === 'fulfilled') {
          allResults.push(result.value);
        } else {
          allResults.push({
            toolName: batch[i].name,
            success: false,
            data: null,
            error: result.reason?.message || 'Unknown error',
            executionTime: 0,
          });
        }
      }
    }

    const totalTime = Date.now() - startTime;
    const successCount = allResults.filter((r) => r.success).length;
    const failureCount = allResults.filter((r) => !r.success).length;
    const aggregatedData = allResults.filter((r) => r.success).map((r) => r.data);

    return {
      results: allResults,
      totalTime,
      successCount,
      failureCount,
      aggregatedData,
    };
  }
}

export default ToolCombiner;
