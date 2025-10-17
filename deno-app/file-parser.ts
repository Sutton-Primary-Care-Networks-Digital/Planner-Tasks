/**
 * File Parser Module for Deno
 * Handles CSV and Excel file parsing for task data
 */

import { parse } from "@std/csv";
import type { PlannerMember } from "./auth.ts";

export interface ParsedTask {
  title: string;
  description?: string;
  startDate?: string;
  dueDate?: string;
  assignee?: string;
  assignees?: string[];
  bucketName?: string;
  status?: string;
}

export interface ProcessedTask extends ParsedTask {
  assigneeUsers?: (PlannerMember & { originalName: string; source: string })[];
  assigneeLookupFailed?: boolean;
  assigneeLookupFailedList?: string[];
  assigneeCandidates?: (PlannerMember & { originalName: string; source: string })[];
  assigneeNeedsReview?: boolean;
  bucketInfo?: {
    id: string;
    name: string;
    originalName: string;
    exactMatch: boolean;
  };
  bucketLookupFailed?: boolean;
}

export class FileParser {
  /**
   * Parse CSV file content
   */
  async parseCSV(content: string): Promise<Record<string, string>[]> {
    try {
      const records = parse(content, {
        skipFirstRow: true, // Skip header row
        strip: true,
      });

      // Convert array of arrays to array of objects with headers
      const lines = content.split('\n');
      const headers = this.parseCSVLine(lines[0]);
      
      const result: Record<string, string>[] = [];
      
      for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        
        const values = this.parseCSVLine(line);
        const record: Record<string, string> = {};
        
        for (let j = 0; j < Math.min(headers.length, values.length); j++) {
          record[headers[j]] = values[j];
        }
        
        result.push(record);
      }
      
      return result;
    } catch (error) {
      console.error("Error parsing CSV:", error);
      throw error;
    }
  }

  /**
   * Parse a single CSV line handling quotes and commas
   */
  private parseCSVLine(line: string): string[] {
    const result: string[] = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const nextChar = line[i + 1];
      
      if (char === '"') {
        if (inQuotes && nextChar === '"') {
          // Escaped quote
          current += '"';
          i++; // Skip next quote
        } else {
          // Toggle quote state
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        // End of field
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    
    // Add the last field
    result.push(current.trim());
    
    return result;
  }

  /**
   * Parse Excel file content (basic CSV format for now)
   * Note: For full Excel support, you'd need additional libraries
   */
  async parseExcel(content: string): Promise<Record<string, string>[]> {
    // For now, treat as CSV. In a full implementation, you'd use a library like xlsx
    return this.parseCSV(content);
  }

  /**
   * Normalize date string to ISO 8601 format
   */
  normalizeDate(dateStr: string): string | undefined {
    if (!dateStr || dateStr.trim() === "") {
      return undefined;
    }

    try {
      const date = new Date(dateStr.trim());
      
      if (isNaN(date.getTime())) {
        console.warn(`Could not parse date '${dateStr}'`);
        return undefined;
      }

      // Convert to ISO string with Z timezone
      return date.toISOString();
    } catch (error) {
      console.warn(`Error parsing date '${dateStr}':`, error);
      return undefined;
    }
  }

  /**
   * Map status string to percentage
   */
  mapStatusToPercentage(status: string): number {
    if (!status) return 0;
    
    const statusLower = status.toLowerCase().replace(/\s+/g, "_");
    const statusMapping: Record<string, number> = {
      "not_started": 0,
      "not started": 0,
      "in_progress": 50,
      "in progress": 50,
      "completed": 100,
      "complete": 100,
      "done": 100,
    };

    return statusMapping[statusLower] ?? 0;
  }

  /**
   * Process raw CSV/Excel data into task objects
   */
  processData(
    data: Record<string, string>[],
    columnMapping: {
      title: string;
      description?: string;
      startDate?: string;
      dueDate?: string;
      assignee?: string;
      bucketName?: string;
      status?: string;
    }
  ): ParsedTask[] {
    const tasks: ParsedTask[] = [];

    for (const row of data) {
      if (!row[columnMapping.title]?.trim()) {
        continue; // Skip rows without titles
      }

      const task: ParsedTask = {
        title: row[columnMapping.title].trim(),
      };

      // Add optional fields if columns are mapped
      if (columnMapping.description && row[columnMapping.description]) {
        task.description = row[columnMapping.description].trim();
      }

      if (columnMapping.startDate && row[columnMapping.startDate]) {
        task.startDate = this.normalizeDate(row[columnMapping.startDate]);
      }

      if (columnMapping.dueDate && row[columnMapping.dueDate]) {
        task.dueDate = this.normalizeDate(row[columnMapping.dueDate]);
      }

      if (columnMapping.assignee && row[columnMapping.assignee]) {
        const assigneeText = row[columnMapping.assignee].trim();
        if (assigneeText) {
          task.assignee = assigneeText;
          // Handle multiple assignees separated by commas or semicolons
          task.assignees = assigneeText
            .split(/[;,]/)
            .map(a => a.trim())
            .filter(a => a);
        }
      }

      if (columnMapping.bucketName && row[columnMapping.bucketName]) {
        task.bucketName = row[columnMapping.bucketName].trim();
      }

      if (columnMapping.status && row[columnMapping.status]) {
        task.status = row[columnMapping.status].trim();
      }

      tasks.push(task);
    }

    return tasks;
  }

  /**
   * Validate tasks have required fields
   */
  validateTasks(tasks: ParsedTask[]): { valid: boolean; errors: string[] } {
    const errors: string[] = [];

    for (let i = 0; i < tasks.length; i++) {
      const task = tasks[i];
      
      if (!task.title?.trim()) {
        errors.push(`Row ${i + 1}: Missing title`);
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }

  /**
   * Get available columns from parsed data
   */
  getAvailableColumns(data: Record<string, string>[]): string[] {
    if (data.length === 0) return [];
    return Object.keys(data[0]);
  }

  /**
   * Get assignee statistics
   */
  getAssigneeStatistics(tasks: ParsedTask[]): Record<string, number> {
    const stats: Record<string, number> = {};
    let unassigned = 0;

    for (const task of tasks) {
      if (task.assignee) {
        stats[task.assignee] = (stats[task.assignee] || 0) + 1;
      } else {
        unassigned++;
      }
    }

    if (unassigned > 0) {
      stats["[Unassigned]"] = unassigned;
    }

    return stats;
  }

  /**
   * Get bucket statistics
   */
  getBucketStatistics(tasks: ParsedTask[]): Record<string, number> {
    const stats: Record<string, number> = {};
    let noBucket = 0;

    for (const task of tasks) {
      if (task.bucketName) {
        stats[task.bucketName] = (stats[task.bucketName] || 0) + 1;
      } else {
        noBucket++;
      }
    }

    if (noBucket > 0) {
      stats["[No Bucket]"] = noBucket;
    }

    return stats;
  }

  private normalizeKey(value: string): string {
    if (!value) {
      return "";
    }
    return value
      .normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim()
      .replace(/\s+/g, " ");
  }

  private generateMemberKeys(member: PlannerMember): string[] {
    const keys = new Set<string>();
    if (member.displayName) {
      keys.add(this.normalizeKey(member.displayName));
    }
    if (member.givenName && member.surname) {
      keys.add(this.normalizeKey(`${member.givenName} ${member.surname}`));
      keys.add(this.normalizeKey(`${member.surname} ${member.givenName}`));
    }
    if (member.companyName && member.givenName && member.surname) {
      keys.add(this.normalizeKey(`${member.givenName} ${member.surname} ${member.companyName}`));
      keys.add(this.normalizeKey(`${member.surname} ${member.givenName} ${member.companyName}`));
    }
    if (member.mail) {
      keys.add(member.mail.toLowerCase());
    }
    if (member.userPrincipalName) {
      keys.add(member.userPrincipalName.toLowerCase());
    }
    return Array.from(keys).filter((key) => key.length > 0);
  }

  private generateAssigneeKeys(name: string): { keys: string[]; tokens: string[] } {
    const keys = new Set<string>();
    const lower = name.trim();
    if (!lower) {
      return { keys: [], tokens: [] };
    }
    keys.add(this.normalizeKey(lower));
    const noParen = lower.replace(/\(.*?\)/g, "").trim();
    if (noParen !== lower) {
      keys.add(this.normalizeKey(noParen));
    }
    if (lower.includes("@")) {
      keys.add(lower.toLowerCase());
    }
    if (lower.includes(",")) {
      const parts = lower.split(",").map((part) => part.replace(/\(.*?\)/g, "").trim());
      if (parts.length === 2) {
        const first = parts[1];
        const last = parts[0];
        keys.add(this.normalizeKey(`${first} ${last}`));
        keys.add(this.normalizeKey(`${last} ${first}`));
      }
    }
    const sanitized = noParen.replace(/,/g, " ");
    const tokens = this.normalizeKey(sanitized).split(" ").filter((token) => token.length > 0);
    if (tokens.length >= 2) {
      const first = tokens[0];
      const last = tokens[tokens.length - 1];
      keys.add(this.normalizeKey(`${first} ${last}`));
      keys.add(this.normalizeKey(`${last} ${first}`));
    }
    return { keys: Array.from(keys).filter((key) => key.length > 0), tokens };
  }

  private buildMemberIndex(members: PlannerMember[]): {
    keyMap: Map<string, PlannerMember[]>;
    entries: { member: PlannerMember; searchText: string }[];
  } {
    const keyMap = new Map<string, PlannerMember[]>();
    const entries: { member: PlannerMember; searchText: string }[] = [];

    for (const member of members) {
      const memberKeys = this.generateMemberKeys(member);
      for (const key of memberKeys) {
        const existing = keyMap.get(key) || [];
        existing.push(member);
        keyMap.set(key, existing);
      }
      const searchText = this.normalizeKey(
        `${member.displayName || ""} ${member.mail || ""} ${member.userPrincipalName || ""} ${member.givenName || ""} ${member.surname || ""} ${member.companyName || ""}`
      );
      entries.push({ member, searchText });
    }

    return { keyMap, entries };
  }

  private findMemberMatches(
    name: string,
    index: { keyMap: Map<string, PlannerMember[]>; entries: { member: PlannerMember; searchText: string }[] }
  ): PlannerMember[] {
    const { keys, tokens } = this.generateAssigneeKeys(name);
    const matches: PlannerMember[] = [];

    for (const key of keys) {
      const membersForKey = index.keyMap.get(key);
      if (membersForKey && membersForKey.length > 0) {
        for (const member of membersForKey) {
          if (!matches.find((existing) => existing.id === member.id)) {
            matches.push(member);
          }
        }
      }
    }

    if (matches.length > 0) {
      return matches;
    }

    if (tokens.length === 0) {
      return matches;
    }

    for (const entry of index.entries) {
      if (tokens.every((token) => entry.searchText.includes(token))) {
        if (!matches.find((existing) => existing.id === entry.member.id)) {
          matches.push(entry.member);
        }
      }
    }

    return matches;
  }

  async lookupAssignees(
    tasks: ParsedTask[],
    auth: any,
    accessToken: string,
    plannerMembers: PlannerMember[] = []
  ): Promise<ProcessedTask[]> {
    const processedTasks: ProcessedTask[] = [];
    const memberIndex = this.buildMemberIndex(plannerMembers);

    for (const task of tasks) {
      const processedTask: ProcessedTask = { ...task };

      const evaluateAssignee = (assigneeName: string) => {
        const matches = this.findMemberMatches(assigneeName, memberIndex);
        if (matches.length === 1) {
          return {
            users: [{ ...matches[0], originalName: assigneeName, source: "planner-member" }],
            candidates: [] as (PlannerMember & { originalName: string; source: string })[],
            needsReview: false,
          };
        }
        if (matches.length > 1) {
          return {
            users: [] as (PlannerMember & { originalName: string; source: string })[],
            candidates: matches.map((member) => ({ ...member, originalName: assigneeName, source: "planner-member" })),
            needsReview: true,
          };
        }
        return {
          users: [] as (PlannerMember & { originalName: string; source: string })[],
          candidates: [] as (PlannerMember & { originalName: string; source: string })[],
          needsReview: true,
        };
      };

      if (task.assignees && task.assignees.length > 0) {
        const matchedUsers: (PlannerMember & { originalName: string; source: string })[] = [];
        const failed: string[] = [];
        const candidates: (PlannerMember & { originalName: string; source: string })[] = [];
        let needsReview = false;

        for (const assigneeName of task.assignees) {
          const result = evaluateAssignee(assigneeName);
          if (result.users.length > 0) {
            matchedUsers.push(...result.users);
          } else {
            failed.push(assigneeName);
          }
          if (result.candidates.length > 0) {
            candidates.push(...result.candidates);
          }
          if (result.needsReview) {
            needsReview = true;
          }
        }

        if (matchedUsers.length > 0) {
          processedTask.assigneeUsers = matchedUsers;
        }
        if (failed.length > 0) {
          processedTask.assigneeLookupFailedList = failed;
        }
        if (candidates.length > 0) {
          processedTask.assigneeCandidates = candidates;
        }
        if (needsReview) {
          processedTask.assigneeNeedsReview = true;
        }
      } else if (task.assignee) {
        const result = evaluateAssignee(task.assignee);
        if (result.users.length > 0) {
          processedTask.assigneeUsers = result.users;
        } else {
          processedTask.assigneeLookupFailed = true;
          if (result.candidates.length > 0) {
            processedTask.assigneeCandidates = result.candidates;
          }
          processedTask.assigneeNeedsReview = true;
        }
      }

      processedTasks.push(processedTask);
    }

    return processedTasks;
  }

  /**
   * Lookup buckets using the auth module
   */
  async lookupBuckets(
    tasks: ProcessedTask[],
    auth: any,
    accessToken: string,
    planId: string
  ): Promise<ProcessedTask[]> {
    const buckets = await auth.getPlannerBuckets(accessToken, planId);
    const bucketLookup: Record<string, any> = {};

    // Create case-insensitive lookup
    for (const bucket of buckets) {
      bucketLookup[bucket.name.toLowerCase()] = bucket;
    }

    // Get unique bucket names from tasks
    const uniqueBucketNames = new Set<string>();
    for (const task of tasks) {
      if (task.bucketName) {
        uniqueBucketNames.add(task.bucketName);
      }
    }

    // Create bucket mapping
    const bucketMapping: Record<string, any> = {};
    for (const bucketName of uniqueBucketNames) {
      const lowerName = bucketName.toLowerCase();
      
      if (bucketLookup[lowerName]) {
        // Exact match
        bucketMapping[bucketName] = {
          ...bucketLookup[lowerName],
          exactMatch: true,
        };
      } else {
        // Try fuzzy matching
        let bestMatch = null;
        let bestScore = 0;

        for (const [availableName, bucket] of Object.entries(bucketLookup)) {
          // Simple substring matching
          const score1 = lowerName.includes(availableName) || availableName.includes(lowerName);
          if (score1) {
            const score = Math.min(bucketName.length, availableName.length) / Math.max(bucketName.length, availableName.length);
            if (score > bestScore && score > 0.5) {
              bestScore = score;
              bestMatch = {
                ...bucket,
                exactMatch: false,
              };
            }
          }
        }

        bucketMapping[bucketName] = bestMatch;
      }
    }

    // Apply bucket information to tasks
    const enrichedTasks: ProcessedTask[] = [];
    for (const task of tasks) {
      const enrichedTask: ProcessedTask = { ...task };

      if (task.bucketName && bucketMapping[task.bucketName]) {
        const matchedBucket = bucketMapping[task.bucketName];
        enrichedTask.bucketInfo = {
          id: matchedBucket.id,
          name: matchedBucket.name,
          originalName: task.bucketName,
          exactMatch: matchedBucket.exactMatch,
        };
      } else if (task.bucketName) {
        enrichedTask.bucketLookupFailed = true;
      }

      enrichedTasks.push(enrichedTask);
    }

    return enrichedTasks;
  }
}