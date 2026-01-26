/**
 * Chat Memory Module - Per-chat memory for OpenWebUI sessions
 *
 * Stores conversation history, mentioned entities, and user preferences
 * for each chat session, enabling contextual awareness across tool calls.
 *
 * Memory is held for a maximum of 72 hours (3 days) to prevent unbounded growth.
 */

import logger from './logger.js';

// Timer type that works across Node.js environments
type TimerHandle = ReturnType<typeof setInterval>;

/**
 * Single conversation entry in chat history
 */
export interface ConversationEntry {
  question: string;
  answer: string;
  timestamp: Date;
  toolUsed?: string;
  resultCount?: number;
  sources?: string[];
}

/**
 * Entity types that can be tracked in conversation
 */
export type EntityType = 'people' | 'files' | 'events' | 'topics';

/**
 * Mentioned entities tracked per chat
 */
export interface MentionedEntities {
  people: Set<string>;
  files: Set<string>;
  events: Set<string>;
  topics: Set<string>;
}

/**
 * User preferences for a chat session
 */
export interface ChatPreferences {
  language?: 'en' | 'de' | 'auto';
  resultLimit?: number;
  preferredSources?: string[];
}

/**
 * Complete chat memory for a single chat session
 */
export interface ChatMemory {
  chatId: string;
  userId?: string;
  conversationHistory: ConversationEntry[];
  mentionedEntities: MentionedEntities;
  preferences: ChatPreferences;
  createdAt: Date;
  lastActivity: Date;
}

/**
 * Serializable version of ChatMemory for JSON export
 */
export interface SerializedChatMemory {
  chatId: string;
  userId?: string;
  conversationHistory: ConversationEntry[];
  mentionedEntities: {
    people: string[];
    files: string[];
    events: string[];
    topics: string[];
  };
  preferences: ChatPreferences;
  createdAt: string;
  lastActivity: string;
}

/**
 * Statistics about the memory store
 */
export interface MemoryStoreStats {
  activeSessions: number;
  totalConversations: number;
  totalEntities: number;
  oldestSession: Date | null;
  newestSession: Date | null;
  memoryUsageEstimate: string;
}

/**
 * Configuration for ChatMemoryStore
 */
interface ChatMemoryStoreConfig {
  ttlHours: number;
  maxHistoryPerChat: number;
  cleanupIntervalMinutes: number;
}

/**
 * Default configuration
 */
const DEFAULT_CONFIG: ChatMemoryStoreConfig = {
  ttlHours: 72, // 72 hours max
  maxHistoryPerChat: 50,
  cleanupIntervalMinutes: 30,
};

/**
 * ChatMemoryStore - Singleton managing all chat memories
 *
 * Provides per-chat memory storage with automatic cleanup of stale sessions.
 */
export class ChatMemoryStore {
  private static instance: ChatMemoryStore | null = null;
  private memories: Map<string, ChatMemory> = new Map();
  private config: ChatMemoryStoreConfig;
  private cleanupInterval: TimerHandle | null = null;

  private constructor(config?: Partial<ChatMemoryStoreConfig>) {
    this.config = {
      ...DEFAULT_CONFIG,
      ...config,
    };

    // Enforce max 72 hours TTL
    if (this.config.ttlHours > 72) {
      this.config.ttlHours = 72;
      logger.warn('Chat memory TTL capped at 72 hours');
    }

    // Load config from environment
    const envTtl = process.env.MS365_MCP_CHAT_MEMORY_TTL;
    if (envTtl) {
      const parsedTtl = parseInt(envTtl, 10);
      if (!isNaN(parsedTtl) && parsedTtl > 0) {
        this.config.ttlHours = Math.min(parsedTtl, 72);
      }
    }

    const envMaxHistory = process.env.MS365_MCP_CHAT_MEMORY_MAX_HISTORY;
    if (envMaxHistory) {
      const parsedMax = parseInt(envMaxHistory, 10);
      if (!isNaN(parsedMax) && parsedMax > 0) {
        this.config.maxHistoryPerChat = parsedMax;
      }
    }

    // Start periodic cleanup
    this.startCleanupInterval();

    logger.info('ChatMemoryStore initialized', {
      ttlHours: this.config.ttlHours,
      maxHistoryPerChat: this.config.maxHistoryPerChat,
      cleanupIntervalMinutes: this.config.cleanupIntervalMinutes,
    });
  }

  /**
   * Get singleton instance of ChatMemoryStore
   */
  static getInstance(config?: Partial<ChatMemoryStoreConfig>): ChatMemoryStore {
    if (!ChatMemoryStore.instance) {
      ChatMemoryStore.instance = new ChatMemoryStore(config);
    }
    return ChatMemoryStore.instance;
  }

  /**
   * Reset singleton instance (for testing)
   */
  static resetInstance(): void {
    if (ChatMemoryStore.instance) {
      ChatMemoryStore.instance.destroy();
      ChatMemoryStore.instance = null;
    }
  }

  /**
   * Start periodic cleanup of stale sessions
   */
  private startCleanupInterval(): void {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
    }

    this.cleanupInterval = setInterval(
      () => {
        this.cleanup();
      },
      this.config.cleanupIntervalMinutes * 60 * 1000
    );

    // Don't prevent Node from exiting
    this.cleanupInterval.unref();
  }

  /**
   * Clean up resources
   */
  destroy(): void {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
      this.cleanupInterval = null;
    }
    this.memories.clear();
  }

  /**
   * Create a new empty ChatMemory
   */
  private createEmptyMemory(chatId: string, userId?: string): ChatMemory {
    return {
      chatId,
      userId,
      conversationHistory: [],
      mentionedEntities: {
        people: new Set(),
        files: new Set(),
        events: new Set(),
        topics: new Set(),
      },
      preferences: {},
      createdAt: new Date(),
      lastActivity: new Date(),
    };
  }

  /**
   * Get or create memory for a chat session
   */
  getMemory(chatId: string, userId?: string): ChatMemory {
    let memory = this.memories.get(chatId);

    if (!memory) {
      memory = this.createEmptyMemory(chatId, userId);
      this.memories.set(chatId, memory);
      logger.debug(`Created new chat memory for ${chatId}`);
    } else {
      // Update last activity
      memory.lastActivity = new Date();
      // Update userId if provided and not set
      if (userId && !memory.userId) {
        memory.userId = userId;
      }
    }

    return memory;
  }

  /**
   * Check if memory exists for a chat
   */
  hasMemory(chatId: string): boolean {
    return this.memories.has(chatId);
  }

  /**
   * Add a conversation entry to chat memory
   */
  addConversation(
    chatId: string,
    question: string,
    answer: string,
    options?: {
      toolUsed?: string;
      resultCount?: number;
      sources?: string[];
      userId?: string;
    }
  ): void {
    const memory = this.getMemory(chatId, options?.userId);

    const entry: ConversationEntry = {
      question,
      answer,
      timestamp: new Date(),
      toolUsed: options?.toolUsed,
      resultCount: options?.resultCount,
      sources: options?.sources,
    };

    memory.conversationHistory.push(entry);

    // Trim history if exceeds max
    if (memory.conversationHistory.length > this.config.maxHistoryPerChat) {
      memory.conversationHistory = memory.conversationHistory.slice(-this.config.maxHistoryPerChat);
    }

    memory.lastActivity = new Date();
    logger.debug(`Added conversation to chat ${chatId}`, {
      questionLength: question.length,
      answerLength: answer.length,
      historySize: memory.conversationHistory.length,
    });
  }

  /**
   * Add an entity to the mentioned entities
   */
  addEntity(chatId: string, type: EntityType, entity: string, userId?: string): void {
    const memory = this.getMemory(chatId, userId);
    memory.mentionedEntities[type].add(entity);
    memory.lastActivity = new Date();
  }

  /**
   * Add multiple entities at once
   */
  addEntities(
    chatId: string,
    entities: Partial<Record<EntityType, string[]>>,
    userId?: string
  ): void {
    const memory = this.getMemory(chatId, userId);

    for (const [type, values] of Object.entries(entities)) {
      if (values && Array.isArray(values)) {
        for (const value of values) {
          memory.mentionedEntities[type as EntityType].add(value);
        }
      }
    }

    memory.lastActivity = new Date();
  }

  /**
   * Get recent conversation context
   */
  getRecentContext(chatId: string, limit: number = 5): ConversationEntry[] {
    const memory = this.memories.get(chatId);
    if (!memory) {
      return [];
    }

    return memory.conversationHistory.slice(-limit);
  }

  /**
   * Get all mentioned entities for a chat
   */
  getMentionedEntities(chatId: string): MentionedEntities | null {
    const memory = this.memories.get(chatId);
    if (!memory) {
      return null;
    }

    return memory.mentionedEntities;
  }

  /**
   * Set a preference for a chat
   */
  setPreference<K extends keyof ChatPreferences>(
    chatId: string,
    key: K,
    value: ChatPreferences[K],
    userId?: string
  ): void {
    const memory = this.getMemory(chatId, userId);
    memory.preferences[key] = value;
    memory.lastActivity = new Date();
  }

  /**
   * Get preferences for a chat
   */
  getPreferences(chatId: string): ChatPreferences | null {
    const memory = this.memories.get(chatId);
    if (!memory) {
      return null;
    }

    return memory.preferences;
  }

  /**
   * Clear memory for a specific chat
   */
  clearMemory(chatId: string): boolean {
    const existed = this.memories.has(chatId);
    this.memories.delete(chatId);
    if (existed) {
      logger.info(`Cleared chat memory for ${chatId}`);
    }
    return existed;
  }

  /**
   * Remove stale sessions (inactive for > TTL hours)
   */
  cleanup(): number {
    const now = new Date();
    const ttlMs = this.config.ttlHours * 60 * 60 * 1000;
    let removedCount = 0;

    for (const [chatId, memory] of this.memories) {
      const age = now.getTime() - memory.lastActivity.getTime();
      if (age > ttlMs) {
        this.memories.delete(chatId);
        removedCount++;
        logger.debug(`Cleaned up stale chat memory: ${chatId}`, {
          ageHours: Math.round(age / (60 * 60 * 1000)),
        });
      }
    }

    if (removedCount > 0) {
      logger.info(`Chat memory cleanup: removed ${removedCount} stale sessions`, {
        remainingSessions: this.memories.size,
      });
    }

    return removedCount;
  }

  /**
   * Get statistics about the memory store
   */
  getStats(): MemoryStoreStats {
    let totalConversations = 0;
    let totalEntities = 0;
    let oldestSession: Date | null = null;
    let newestSession: Date | null = null;

    for (const memory of this.memories.values()) {
      totalConversations += memory.conversationHistory.length;
      totalEntities +=
        memory.mentionedEntities.people.size +
        memory.mentionedEntities.files.size +
        memory.mentionedEntities.events.size +
        memory.mentionedEntities.topics.size;

      if (!oldestSession || memory.createdAt < oldestSession) {
        oldestSession = memory.createdAt;
      }
      if (!newestSession || memory.createdAt > newestSession) {
        newestSession = memory.createdAt;
      }
    }

    // Rough memory estimate
    const memoryBytes = JSON.stringify(this.serializeAll()).length;
    const memoryUsageEstimate =
      memoryBytes < 1024
        ? `${memoryBytes} B`
        : memoryBytes < 1024 * 1024
          ? `${(memoryBytes / 1024).toFixed(2)} KB`
          : `${(memoryBytes / (1024 * 1024)).toFixed(2)} MB`;

    return {
      activeSessions: this.memories.size,
      totalConversations,
      totalEntities,
      oldestSession,
      newestSession,
      memoryUsageEstimate,
    };
  }

  /**
   * Serialize a single memory to JSON-compatible format
   */
  serializeMemory(chatId: string): SerializedChatMemory | null {
    const memory = this.memories.get(chatId);
    if (!memory) {
      return null;
    }

    return {
      chatId: memory.chatId,
      userId: memory.userId,
      conversationHistory: memory.conversationHistory,
      mentionedEntities: {
        people: Array.from(memory.mentionedEntities.people),
        files: Array.from(memory.mentionedEntities.files),
        events: Array.from(memory.mentionedEntities.events),
        topics: Array.from(memory.mentionedEntities.topics),
      },
      preferences: memory.preferences,
      createdAt: memory.createdAt.toISOString(),
      lastActivity: memory.lastActivity.toISOString(),
    };
  }

  /**
   * Serialize all memories
   */
  serializeAll(): SerializedChatMemory[] {
    const serialized: SerializedChatMemory[] = [];
    for (const chatId of this.memories.keys()) {
      const mem = this.serializeMemory(chatId);
      if (mem) {
        serialized.push(mem);
      }
    }
    return serialized;
  }

  /**
   * Get a formatted summary of chat memory for display
   */
  getMemorySummary(chatId: string): string {
    const memory = this.memories.get(chatId);
    if (!memory) {
      return 'No memory found for this chat session.';
    }

    const lines: string[] = [];
    lines.push(`📝 **Chat Memory Summary**`);
    lines.push(`- Chat ID: \`${chatId.substring(0, 8)}...\``);
    if (memory.userId) {
      lines.push(`- User ID: \`${memory.userId.substring(0, 8)}...\``);
    }
    lines.push(`- Created: ${memory.createdAt.toISOString()}`);
    lines.push(`- Last Activity: ${memory.lastActivity.toISOString()}`);
    lines.push('');

    // Conversation history
    lines.push(`**📜 Conversation History** (${memory.conversationHistory.length} entries)`);
    if (memory.conversationHistory.length > 0) {
      const recent = memory.conversationHistory.slice(-3);
      for (const entry of recent) {
        const time = new Date(entry.timestamp).toLocaleTimeString();
        lines.push(`- [${time}] Q: "${entry.question.substring(0, 50)}..."`);
      }
      if (memory.conversationHistory.length > 3) {
        lines.push(`  _(and ${memory.conversationHistory.length - 3} more...)_`);
      }
    } else {
      lines.push('  _No conversations yet_');
    }
    lines.push('');

    // Mentioned entities
    lines.push('**🏷️ Mentioned Entities**');
    const entityCounts = {
      '👤 People': memory.mentionedEntities.people.size,
      '📁 Files': memory.mentionedEntities.files.size,
      '📅 Events': memory.mentionedEntities.events.size,
      '💡 Topics': memory.mentionedEntities.topics.size,
    };

    for (const [label, count] of Object.entries(entityCounts)) {
      if (count > 0) {
        lines.push(`- ${label}: ${count}`);
      }
    }

    if (Object.values(entityCounts).every((c) => c === 0)) {
      lines.push('  _No entities tracked yet_');
    }
    lines.push('');

    // Preferences
    lines.push('**⚙️ Preferences**');
    const prefs = memory.preferences;
    if (Object.keys(prefs).length > 0) {
      if (prefs.language) lines.push(`- Language: ${prefs.language}`);
      if (prefs.resultLimit) lines.push(`- Result Limit: ${prefs.resultLimit}`);
      if (prefs.preferredSources?.length) {
        lines.push(`- Preferred Sources: ${prefs.preferredSources.join(', ')}`);
      }
    } else {
      lines.push('  _No preferences set_');
    }

    return lines.join('\n');
  }

  /**
   * Build context string for query enhancement
   */
  buildContextForQuery(chatId: string): string | null {
    const memory = this.memories.get(chatId);
    if (!memory || memory.conversationHistory.length === 0) {
      return null;
    }

    const parts: string[] = [];

    // Add recent conversation context
    const recent = memory.conversationHistory.slice(-3);
    if (recent.length > 0) {
      parts.push('Recent conversation context:');
      for (const entry of recent) {
        parts.push(`- Q: "${entry.question}"`);
        if (entry.resultCount !== undefined) {
          parts.push(`  (Found ${entry.resultCount} results)`);
        }
      }
    }

    // Add relevant mentioned entities
    const entities = memory.mentionedEntities;
    const mentionedItems: string[] = [];

    if (entities.people.size > 0) {
      mentionedItems.push(`People: ${Array.from(entities.people).slice(0, 5).join(', ')}`);
    }
    if (entities.topics.size > 0) {
      mentionedItems.push(`Topics: ${Array.from(entities.topics).slice(0, 5).join(', ')}`);
    }

    if (mentionedItems.length > 0) {
      parts.push('');
      parts.push('Previously mentioned:');
      parts.push(mentionedItems.join('; '));
    }

    return parts.length > 0 ? parts.join('\n') : null;
  }
}

/**
 * Check if chat memory feature is enabled
 */
export function isChatMemoryEnabled(): boolean {
  const envEnabled = process.env.MS365_MCP_CHAT_MEMORY_ENABLED;
  // Enabled by default unless explicitly disabled
  return envEnabled !== 'false' && envEnabled !== '0';
}

/**
 * Get the singleton ChatMemoryStore instance
 */
export function getChatMemoryStore(): ChatMemoryStore {
  return ChatMemoryStore.getInstance();
}

export default ChatMemoryStore;
