/**
 * Entity Graph Builder
 *
 * Builds a graph of relationships between entities found in M365 data.
 * Enables cross-source analysis and relationship discovery.
 */

// Entity graph types - language-agnostic

/**
 * Entity types in the graph
 */
export type GraphEntityType =
  | 'person'
  | 'project'
  | 'document'
  | 'event'
  | 'email'
  | 'task'
  | 'team'
  | 'location';

/**
 * Relationship types between entities
 */
export type RelationshipType =
  | 'mentions' // Entity A mentions Entity B
  | 'attends' // Person attends Event
  | 'owns' // Person owns Document
  | 'shares' // Person shares with Person
  | 'collaborates' // Person collaborates with Person
  | 'assignedTo' // Task assigned to Person
  | 'relatedTo' // Generic relationship
  | 'partOf'; // Entity is part of another

/**
 * Entity node in the graph
 */
export interface EntityNode {
  id: string;
  type: GraphEntityType;
  name: string;
  sources: string[];
  frequency: number;
  lastSeen: Date;
  properties: Record<string, unknown>;
}

/**
 * Edge between entities
 */
export interface EntityEdge {
  id: string;
  source: string;
  target: string;
  relationship: RelationshipType;
  strength: number;
  context: string[];
  timestamp?: Date;
}

/**
 * Entity cluster (group of related entities)
 */
export interface EntityCluster {
  id: string;
  name: string;
  entities: string[];
  centralEntity?: string;
  theme?: string;
}

/**
 * Complete entity graph
 */
export interface EntityGraph {
  nodes: Map<string, EntityNode>;
  edges: Map<string, EntityEdge[]>;
  clusters: EntityCluster[];
  stats: {
    nodeCount: number;
    edgeCount: number;
    clusterCount: number;
  };
}

/**
 * EntityGraphBuilder - Builds entity relationship graphs
 */
export class EntityGraphBuilder {
  private nodes: Map<string, EntityNode> = new Map();
  private edges: Map<string, EntityEdge[]> = new Map();
  private clusters: EntityCluster[] = [];
  private edgeIdCounter = 0;

  /**
   * Build graph from M365 results
   */
  build(results: unknown[]): EntityGraph {
    // Clear previous graph
    this.clear();

    // Extract entities from results
    for (const item of results) {
      if (typeof item === 'object' && item !== null) {
        this.extractEntitiesFromItem(item as Record<string, unknown>);
      }
    }

    // Build relationships
    this.buildRelationships(results);

    // Identify clusters
    this.identifyClusters();

    return this.getGraph();
  }

  /**
   * Extract entities from a single item
   */
  private extractEntitiesFromItem(item: Record<string, unknown>): void {
    const source = this.identifySource(item);

    // Extract person entities
    this.extractPersonEntities(item, source);

    // Extract document entities
    this.extractDocumentEntities(item, source);

    // Extract event entities
    this.extractEventEntities(item, source);

    // Extract task entities
    this.extractTaskEntities(item, source);
  }

  /**
   * Extract person entities
   */
  private extractPersonEntities(item: Record<string, unknown>, source: string): void {
    // From email: sender, recipients
    if (item.from) {
      this.addPersonFromEmailAddress(item.from, source);
    }
    if (item.toRecipients && Array.isArray(item.toRecipients)) {
      for (const recipient of item.toRecipients) {
        this.addPersonFromEmailAddress(recipient, source);
      }
    }

    // From event: organizer, attendees
    if (item.organizer) {
      this.addPersonFromEmailAddress(item.organizer, source);
    }
    if (item.attendees && Array.isArray(item.attendees)) {
      for (const attendee of item.attendees) {
        this.addPersonFromEmailAddress(attendee, source);
      }
    }

    // From file: createdBy, lastModifiedBy
    if (item.createdBy) {
      this.addPersonFromUser(item.createdBy, source);
    }
    if (item.lastModifiedBy) {
      this.addPersonFromUser(item.lastModifiedBy, source);
    }
  }

  /**
   * Add person from email address object
   */
  private addPersonFromEmailAddress(obj: unknown, source: string): void {
    if (typeof obj !== 'object' || obj === null) return;

    const o = obj as Record<string, unknown>;
    let name: string | undefined;
    let email: string | undefined;

    if (o.emailAddress && typeof o.emailAddress === 'object') {
      const emailObj = o.emailAddress as Record<string, unknown>;
      name = emailObj.name as string;
      email = emailObj.address as string;
    } else {
      name = (o.name as string) || (o.displayName as string);
      email = (o.address as string) || (o.mail as string);
    }

    if (email) {
      this.addOrUpdateNode({
        id: email.toLowerCase(),
        type: 'person',
        name: name || email,
        sources: [source],
        frequency: 1,
        lastSeen: new Date(),
        properties: { email },
      });
    }
  }

  /**
   * Add person from user object
   */
  private addPersonFromUser(obj: unknown, source: string): void {
    if (typeof obj !== 'object' || obj === null) return;

    const o = obj as Record<string, unknown>;
    if (o.user && typeof o.user === 'object') {
      const user = o.user as Record<string, unknown>;
      const email = (user.email as string) || (user.id as string);
      const name = user.displayName as string;

      if (email) {
        this.addOrUpdateNode({
          id: email.toLowerCase(),
          type: 'person',
          name: name || email,
          sources: [source],
          frequency: 1,
          lastSeen: new Date(),
          properties: { email },
        });
      }
    }
  }

  /**
   * Extract document entities
   */
  private extractDocumentEntities(item: Record<string, unknown>, source: string): void {
    if (item.name && item.webUrl) {
      const id = (item.id as string) || (item.webUrl as string);
      this.addOrUpdateNode({
        id,
        type: 'document',
        name: item.name as string,
        sources: [source],
        frequency: 1,
        lastSeen: new Date(),
        properties: {
          webUrl: item.webUrl,
          mimeType: item.file ? (item.file as Record<string, unknown>).mimeType : undefined,
        },
      });
    }
  }

  /**
   * Extract event entities
   */
  private extractEventEntities(item: Record<string, unknown>, source: string): void {
    if (item.subject && item.start) {
      const id = (item.id as string) || `event-${item.subject}-${item.start}`;
      this.addOrUpdateNode({
        id,
        type: 'event',
        name: item.subject as string,
        sources: [source],
        frequency: 1,
        lastSeen: new Date(),
        properties: {
          start: item.start,
          end: item.end,
          location: item.location,
        },
      });
    }
  }

  /**
   * Extract task entities
   */
  private extractTaskEntities(item: Record<string, unknown>, source: string): void {
    if (item.title && (item.status || item.percentComplete !== undefined)) {
      const id = (item.id as string) || `task-${item.title}`;
      this.addOrUpdateNode({
        id,
        type: 'task',
        name: item.title as string,
        sources: [source],
        frequency: 1,
        lastSeen: new Date(),
        properties: {
          status: item.status,
          dueDateTime: item.dueDateTime,
        },
      });
    }
  }

  /**
   * Add or update a node
   */
  private addOrUpdateNode(node: EntityNode): void {
    const existing = this.nodes.get(node.id);
    if (existing) {
      // Update existing node
      existing.frequency++;
      existing.lastSeen = new Date();
      for (const source of node.sources) {
        if (!existing.sources.includes(source)) {
          existing.sources.push(source);
        }
      }
      // Merge properties
      existing.properties = { ...existing.properties, ...node.properties };
    } else {
      this.nodes.set(node.id, node);
    }
  }

  /**
   * Build relationships between entities
   */
  private buildRelationships(results: unknown[]): void {
    for (const item of results) {
      if (typeof item !== 'object' || item === null) continue;
      const obj = item as Record<string, unknown>;

      // Email relationships
      if (obj.from && obj.toRecipients) {
        this.buildEmailRelationships(obj);
      }

      // Event relationships
      if (obj.organizer && obj.attendees) {
        this.buildEventRelationships(obj);
      }

      // Document relationships
      if (obj.createdBy || obj.lastModifiedBy) {
        this.buildDocumentRelationships(obj);
      }
    }
  }

  /**
   * Build relationships from email
   */
  private buildEmailRelationships(item: Record<string, unknown>): void {
    const fromEmail = this.extractEmailAddress(item.from);
    if (!fromEmail) return;

    const recipients = item.toRecipients as unknown[];
    if (!Array.isArray(recipients)) return;

    for (const recipient of recipients) {
      const toEmail = this.extractEmailAddress(recipient);
      if (toEmail && toEmail !== fromEmail) {
        this.addEdge(fromEmail, toEmail, 'mentions', `Email: ${item.subject || 'Untitled'}`);
      }
    }
  }

  /**
   * Build relationships from event
   */
  private buildEventRelationships(item: Record<string, unknown>): void {
    const organizerEmail = this.extractEmailAddress(item.organizer);
    const eventId = item.id as string;

    if (organizerEmail && eventId) {
      this.addEdge(organizerEmail, eventId, 'owns', `Organizes: ${item.subject || 'Event'}`);
    }

    const attendees = item.attendees as unknown[];
    if (!Array.isArray(attendees)) return;

    for (const attendee of attendees) {
      const attendeeEmail = this.extractEmailAddress(attendee);
      if (attendeeEmail && eventId) {
        this.addEdge(attendeeEmail, eventId, 'attends', `Attends: ${item.subject || 'Event'}`);
      }
      // Also add collaboration relationship between attendees
      if (organizerEmail && attendeeEmail && organizerEmail !== attendeeEmail) {
        this.addEdge(
          organizerEmail,
          attendeeEmail,
          'collaborates',
          `Meeting: ${item.subject || 'Event'}`
        );
      }
    }
  }

  /**
   * Build relationships from document
   */
  private buildDocumentRelationships(item: Record<string, unknown>): void {
    const docId = (item.id as string) || (item.webUrl as string);
    if (!docId) return;

    if (item.createdBy) {
      const creatorEmail = this.extractUserEmail(item.createdBy);
      if (creatorEmail) {
        this.addEdge(creatorEmail, docId, 'owns', `Created: ${item.name || 'Document'}`);
      }
    }

    if (item.lastModifiedBy) {
      const modifierEmail = this.extractUserEmail(item.lastModifiedBy);
      if (modifierEmail) {
        this.addEdge(modifierEmail, docId, 'mentions', `Modified: ${item.name || 'Document'}`);
      }
    }
  }

  /**
   * Extract email address from object
   */
  private extractEmailAddress(obj: unknown): string | null {
    if (typeof obj !== 'object' || obj === null) return null;

    const o = obj as Record<string, unknown>;
    if (o.emailAddress && typeof o.emailAddress === 'object') {
      return ((o.emailAddress as Record<string, unknown>).address as string)?.toLowerCase() || null;
    }
    return (o.address as string)?.toLowerCase() || null;
  }

  /**
   * Extract email from user object
   */
  private extractUserEmail(obj: unknown): string | null {
    if (typeof obj !== 'object' || obj === null) return null;

    const o = obj as Record<string, unknown>;
    if (o.user && typeof o.user === 'object') {
      const user = o.user as Record<string, unknown>;
      return ((user.email as string) || (user.id as string))?.toLowerCase() || null;
    }
    return null;
  }

  /**
   * Add edge between nodes
   */
  private addEdge(
    sourceId: string,
    targetId: string,
    relationship: RelationshipType,
    context: string
  ): void {
    const source = sourceId.toLowerCase();
    const target = targetId.toLowerCase();

    if (!this.edges.has(source)) {
      this.edges.set(source, []);
    }

    const existing = this.edges
      .get(source)!
      .find((e) => e.target === target && e.relationship === relationship);
    if (existing) {
      existing.strength++;
      if (!existing.context.includes(context)) {
        existing.context.push(context);
      }
    } else {
      this.edges.get(source)!.push({
        id: `edge-${this.edgeIdCounter++}`,
        source,
        target,
        relationship,
        strength: 1,
        context: [context],
        timestamp: new Date(),
      });
    }
  }

  /**
   * Identify clusters of related entities
   */
  private identifyClusters(): void {
    // Simple clustering based on connected components
    const visited = new Set<string>();
    let clusterId = 0;

    for (const [nodeId] of this.nodes) {
      if (visited.has(nodeId)) continue;

      const cluster: string[] = [];
      this.dfs(nodeId, visited, cluster);

      if (cluster.length > 1) {
        // Find central entity (highest frequency)
        let central = cluster[0];
        let maxFreq = 0;
        for (const id of cluster) {
          const node = this.nodes.get(id);
          if (node && node.frequency > maxFreq) {
            maxFreq = node.frequency;
            central = id;
          }
        }

        this.clusters.push({
          id: `cluster-${clusterId++}`,
          name: this.nodes.get(central)?.name || 'Unnamed Cluster',
          entities: cluster,
          centralEntity: central,
        });
      }
    }
  }

  /**
   * DFS for cluster identification
   */
  private dfs(nodeId: string, visited: Set<string>, cluster: string[]): void {
    if (visited.has(nodeId)) return;
    visited.add(nodeId);
    cluster.push(nodeId);

    const edges = this.edges.get(nodeId) || [];
    for (const edge of edges) {
      this.dfs(edge.target, visited, cluster);
    }
  }

  /**
   * Identify source type from item
   */
  private identifySource(item: Record<string, unknown>): string {
    if (item.from && item.toRecipients) return 'email';
    if (item.start && item.end) return 'calendar';
    if (item.webUrl && item.name) return 'files';
    if (item.title && item.status) return 'tasks';
    return 'unknown';
  }

  /**
   * Get the complete graph
   */
  getGraph(): EntityGraph {
    let edgeCount = 0;
    for (const edges of this.edges.values()) {
      edgeCount += edges.length;
    }

    return {
      nodes: new Map(this.nodes),
      edges: new Map(this.edges),
      clusters: [...this.clusters],
      stats: {
        nodeCount: this.nodes.size,
        edgeCount,
        clusterCount: this.clusters.length,
      },
    };
  }

  /**
   * Get node by ID
   */
  getNode(id: string): EntityNode | undefined {
    return this.nodes.get(id.toLowerCase());
  }

  /**
   * Get edges for a node
   */
  getEdges(nodeId: string): EntityEdge[] {
    return this.edges.get(nodeId.toLowerCase()) || [];
  }

  /**
   * Get related entities
   */
  getRelatedEntities(nodeId: string, maxDepth: number = 2): EntityNode[] {
    const related = new Set<string>();
    const queue: Array<{ id: string; depth: number }> = [{ id: nodeId.toLowerCase(), depth: 0 }];

    while (queue.length > 0) {
      const { id, depth } = queue.shift()!;
      if (depth > maxDepth) continue;

      const edges = this.getEdges(id);
      for (const edge of edges) {
        if (!related.has(edge.target)) {
          related.add(edge.target);
          queue.push({ id: edge.target, depth: depth + 1 });
        }
      }
    }

    return Array.from(related)
      .map((id) => this.nodes.get(id))
      .filter((n): n is EntityNode => n !== undefined);
  }

  /**
   * Clear the graph
   */
  clear(): void {
    this.nodes.clear();
    this.edges.clear();
    this.clusters = [];
    this.edgeIdCounter = 0;
  }
}

export default EntityGraphBuilder;
