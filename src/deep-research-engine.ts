/**
 * Deep Research Engine for complete questions with multi-step reasoning and iterative refinement
 */

import GraphClient from './graph-client.js';
import SearchFirstStrategy from './intelligent-search.js';
import EntityExtractor from './entity-extractor.js';
import SynonymExpander from './synonym-expander.js';
import QueryRefiner from './query-refiner.js';
import LearningSystem from './learning-system.js';
import ToolCombiner from './tool-combiner.js';
import DataAggregator from './data-aggregator.js';
import DownloadLinkGenerator from './download-link-generator.js';
import logger from './logger.js';
import type { AppSecrets } from './secrets.js';
import { getMaxAggregateItems } from './perf-config.js';

export interface ResearchQuestion {
  question: string;
  context?: string;
  maxDepth?: number;
  maxIterations?: number;
  includeDownloadLinks?: boolean;
}

export interface ResearchStep {
  step: number;
  query: string;
  results: unknown[];
  sources: string[];
  reasoning: string;
  nextSteps?: string[];
}

export interface ResearchResult {
  question: string;
  steps: ResearchStep[];
  finalAnswer: string;
  sources: string[];
  totalItems: number;
  downloadLinks?: Array<{ fileName: string; downloadUrl: string }>;
  confidence: number;
  executionTime: number;
}

export class DeepResearchEngine {
  private graphClient: GraphClient;
  private searchStrategy: SearchFirstStrategy;
  private entityExtractor: EntityExtractor;
  private toolCombiner: ToolCombiner;
  private dataAggregator: DataAggregator;
  private downloadLinkGenerator: DownloadLinkGenerator;
  private readonly maxDepth: number;
  private readonly maxIterations: number;

  constructor(
    graphClient: GraphClient,
    searchStrategy: SearchFirstStrategy,
    entityExtractor: EntityExtractor,
    toolCombiner: ToolCombiner,
    dataAggregator: DataAggregator,
    downloadLinkGenerator: DownloadLinkGenerator,
    secrets: AppSecrets
  ) {
    this.graphClient = graphClient;
    this.searchStrategy = searchStrategy;
    this.entityExtractor = entityExtractor;
    this.toolCombiner = toolCombiner;
    this.dataAggregator = dataAggregator;
    this.downloadLinkGenerator = downloadLinkGenerator;
    this.maxDepth = parseInt(process.env.MS365_MCP_DEEP_RESEARCH_MAX_DEPTH || '5', 10);
    this.maxIterations = parseInt(process.env.MS365_MCP_MAX_RESEARCH_ITERATIONS || '5', 10);
  }

  /**
   * Conduct deep research on a question
   */
  async research(question: ResearchQuestion): Promise<ResearchResult> {
    const startTime = Date.now();
    const steps: ResearchStep[] = [];
    const allSources = new Set<string>();
    const allItems: unknown[] = [];
    let currentQuery = question.question;
    let depth = 0;
    let iteration = 0;

    logger.info(`Starting deep research for: "${question.question}"`);

    // Iterative research loop
    while (
      iteration < (question.maxIterations || this.maxIterations) &&
      depth < (question.maxDepth || this.maxDepth)
    ) {
      iteration++;
      logger.info(
        `Research iteration ${iteration}/${question.maxIterations || this.maxIterations}`
      );

      // 1. Execute search-first strategy
      const itemsPerIteration = parseInt(
        process.env.MS365_MCP_DEEP_RESEARCH_ITEMS_PER_ITERATION || '100',
        10
      );

      // Use learning system recommendations for better entity types
      const searchStrategy = this.searchStrategy as unknown as {
        learningSystem?: LearningSystem;
      };
      let recommendedEntityTypes: string[] | undefined;
      if (searchStrategy.learningSystem) {
        recommendedEntityTypes = searchStrategy.learningSystem.getRecommendedEntityTypes(
          currentQuery,
          `deep-research-${iteration}`
        );
      }

      const searchResult = await this.searchStrategy.execute(currentQuery, {
        maxResults: itemsPerIteration,
        entityTypes: recommendedEntityTypes,
      });

      // 2. Extract entities and information
      const extractedInfo = this.entityExtractor.extractFromResults(
        searchResult.searchResults.items
      );

      // 3. Determine next steps based on results
      const reasoning = this.generateReasoning(searchResult, extractedInfo, question.question);
      const nextSteps = this.determineNextSteps(searchResult, extractedInfo, question.question);

      // 4. Record step
      steps.push({
        step: iteration,
        query: currentQuery,
        results: searchResult.searchResults.items,
        sources: searchResult.searchResults.sources,
        reasoning,
        nextSteps,
      });

      // 5. Collect results
      for (const source of searchResult.searchResults.sources) {
        allSources.add(source);
      }
      allItems.push(...searchResult.searchResults.items);
      allItems.push(...Object.values(searchResult.specificResults).flat());

      // 6. If we have enough information or no next steps, break
      const maxAggregateItems = getMaxAggregateItems();
      if (nextSteps.length === 0 || allItems.length >= maxAggregateItems) {
        logger.info('Research complete: sufficient information gathered or no more steps');
        break;
      }

      // 7. Refine query for next iteration using learned patterns
      if (nextSteps.length > 0) {
        // Use learning system recommendations if available for better query refinement
        // The searchStrategy already learns from each search, but we can use that knowledge
        // for better query refinement in the next iteration
        currentQuery = await this.refineQueryForNextStep(currentQuery, nextSteps[0], extractedInfo);
        depth++;
      } else {
        break;
      }
    }

    // Aggregate all results
    const maxAggregateItems = getMaxAggregateItems();
    const aggregated = this.dataAggregator.aggregate(
      [
        {
          source: 'search',
          items: allItems,
        },
      ],
      {
        sortBy: 'relevance',
        sortOrder: 'desc',
        maxItems: maxAggregateItems,
        deduplicate: true,
        formatForLLM: true,
      }
    );

    // Generate download links if requested
    let downloadLinks: Array<{ fileName: string; downloadUrl: string }> | undefined;
    if (question.includeDownloadLinks) {
      const links = await this.downloadLinkGenerator.addDownloadLinksToResults(
        aggregated.items.map((i) => i.data)
      );
      downloadLinks = links
        .filter((l) => l.downloadLink)
        .map((l) => ({
          fileName: l.downloadLink!.fileName,
          downloadUrl: l.downloadLink!.downloadUrl,
        }));
    }

    // Generate final answer
    const finalAnswer = this.generateFinalAnswer(question.question, steps, aggregated);

    // Calculate confidence
    const confidence = this.calculateConfidence(steps, aggregated, allItems.length);

    const executionTime = Date.now() - startTime;

    logger.info(
      `Deep research complete: ${steps.length} steps, ${allItems.length} items, ${executionTime}ms`
    );

    return {
      question: question.question,
      steps,
      finalAnswer,
      sources: Array.from(allSources),
      totalItems: allItems.length,
      downloadLinks,
      confidence,
      executionTime,
    };
  }

  /**
   * Generate reasoning for a research step
   */
  private generateReasoning(
    searchResult: Awaited<ReturnType<SearchFirstStrategy['execute']>>,
    extractedInfo: ReturnType<EntityExtractor['extractFromResults']>,
    originalQuestion: string
  ): string {
    const itemCount = searchResult.searchResults.items.length;
    const sourceCount = searchResult.searchResults.sources.length;

    if (itemCount === 0) {
      return `No results found. May need to refine search query or try different sources.`;
    }

    const reasoning: string[] = [];
    reasoning.push(`Found ${itemCount} items from ${sourceCount} sources.`);

    if (extractedInfo.sites.length > 0) {
      reasoning.push(`Identified ${extractedInfo.sites.length} SharePoint sites.`);
    }
    if (extractedInfo.files.length > 0) {
      reasoning.push(`Identified ${extractedInfo.files.length} files.`);
    }
    if (extractedInfo.teams.length > 0) {
      reasoning.push(`Identified ${extractedInfo.teams.length} Teams.`);
    }
    if (extractedInfo.users.length > 0) {
      reasoning.push(`Identified ${extractedInfo.users.length} users.`);
    }

    return reasoning.join(' ');
  }

  /**
   * Determine next steps based on results
   */
  private determineNextSteps(
    searchResult: Awaited<ReturnType<SearchFirstStrategy['execute']>>,
    extractedInfo: ReturnType<EntityExtractor['extractFromResults']>,
    originalQuestion: string
  ): string[] {
    const nextSteps: string[] = [];

    // If we found sites but haven't explored them deeply
    if (extractedInfo.sites.length > 0) {
      nextSteps.push(`Explore ${extractedInfo.sites.length} SharePoint sites in detail`);
    }

    // If we found files but haven't examined them
    if (extractedInfo.files.length > 0) {
      nextSteps.push(`Examine ${extractedInfo.files.length} files for relevant content`);
    }

    // If we found teams but haven't checked their activity
    if (extractedInfo.teams.length > 0) {
      nextSteps.push(`Check activity in ${extractedInfo.teams.length} Teams`);
    }

    // If we found users but haven't looked at their contributions
    if (extractedInfo.users.length > 0) {
      nextSteps.push(`Review contributions from ${extractedInfo.users.length} users`);
    }

    // If results are sparse, suggest broader search
    if (searchResult.searchResults.items.length < 5) {
      nextSteps.push('Try broader search terms or different entity types');
    }

    return nextSteps;
  }

  /**
   * Refine query for next research step using learned patterns
   */
  private async refineQueryForNextStep(
    currentQuery: string,
    nextStep: string,
    extractedInfo: ReturnType<EntityExtractor['extractFromResults']>
  ): Promise<string> {
    // Use keywords from extracted info to refine query
    let refinedQuery = currentQuery;
    if (extractedInfo.keywords.length > 0) {
      const topKeywords = extractedInfo.keywords.slice(0, 3).join(' ');
      refinedQuery = `${currentQuery} ${topKeywords}`;
    } else {
      // Use next step description
      refinedQuery = `${currentQuery} ${nextStep}`;
    }

    // Try to get query variants from learning system via searchStrategy
    const searchStrategy = this.searchStrategy as unknown as {
      learningSystem?: LearningSystem;
      queryRefiner?: QueryRefiner;
    };

    // If we have a query refiner, use it to get better variants
    if (searchStrategy.queryRefiner) {
      try {
        const variants = await searchStrategy.queryRefiner.refineQuery(refinedQuery, false);
        if (variants.length > 0 && variants[0] !== refinedQuery) {
          // Use the best variant
          refinedQuery = variants[0];
          logger.debug(`Refined query using query refiner: "${refinedQuery}"`);
        }
      } catch (error) {
        logger.warn(`Failed to refine query: ${error}`);
      }
    }

    return refinedQuery;
  }

  /**
   * Generate final answer from research steps
   */
  private generateFinalAnswer(
    question: string,
    steps: ResearchStep[],
    aggregated: ReturnType<DataAggregator['aggregate']>
  ): string {
    const answerParts: string[] = [];

    answerParts.push(`## Research Summary for: "${question}"`);
    answerParts.push('');
    answerParts.push(`**Total Items Found:** ${aggregated.totalItems}`);
    answerParts.push(`**Unique Items:** ${aggregated.uniqueItems}`);
    answerParts.push(`**Sources:** ${aggregated.sources.join(', ')}`);
    answerParts.push('');

    // Add step summaries
    answerParts.push('### Research Steps:');
    for (const step of steps) {
      answerParts.push(`**Step ${step.step}:** ${step.query}`);
      answerParts.push(`- Found ${step.results.length} items from ${step.sources.length} sources`);
      answerParts.push(`- Reasoning: ${step.reasoning}`);
      if (step.nextSteps && step.nextSteps.length > 0) {
        answerParts.push(`- Next steps: ${step.nextSteps.join('; ')}`);
      }
      answerParts.push('');
    }

    // Add formatted data if available
    if (aggregated.formattedForLLM) {
      answerParts.push('### Key Findings:');
      answerParts.push(aggregated.formattedForLLM);
    }

    return answerParts.join('\n');
  }

  /**
   * Calculate confidence score
   */
  private calculateConfidence(
    steps: ResearchStep[],
    aggregated: ReturnType<DataAggregator['aggregate']>,
    totalItems: number
  ): number {
    let confidence = 0.5; // Base confidence

    // More items = higher confidence
    if (totalItems > 50) {
      confidence += 0.2;
    } else if (totalItems > 20) {
      confidence += 0.1;
    }

    // More sources = higher confidence
    if (aggregated.sources.length > 3) {
      confidence += 0.1;
    } else if (aggregated.sources.length > 1) {
      confidence += 0.05;
    }

    // More steps with results = higher confidence
    const stepsWithResults = steps.filter((s) => s.results.length > 0).length;
    if (stepsWithResults >= 3) {
      confidence += 0.1;
    } else if (stepsWithResults >= 2) {
      confidence += 0.05;
    }

    return Math.min(confidence, 1.0);
  }
}

export default DeepResearchEngine;
