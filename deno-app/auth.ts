/**
 * Microsoft Graph Authentication Module for Deno
 * Handles OAuth2 authentication with Microsoft Graph API using device flow and browser authentication
 */

export interface User {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
  originalName?: string;
}

export interface Planner {
  id: string;
  title: string;
  groupId: string;
  groupName: string;
}

export interface Bucket {
  id: string;
  name: string;
  planId: string;
}

export interface PlannerMember {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
  givenName?: string;
  surname?: string;
  companyName?: string;
}

export interface TaskData {
  title: string;
  description?: string;
  dueDateTime?: string;
  startDateTime?: string;
  percentComplete?: number;
}

interface AuthCallbackResult {
  code?: string | null;
  state?: string | null;
  error?: string | null;
}

export class GraphAuth {
  private clientIds = [
    "1950a258-227b-4e31-a9cf-717495945fc2", // Microsoft Azure CLI
    "04b07795-8ddb-461a-bbee-02f9e1bf7b46", // Microsoft Graph Explorer
    "1fec8e78-bce4-4aaf-ab1b-5451cc387264", // Microsoft Graph PowerShell
  ];
  
  private tenantId = "common";
  private scopes = ["https://graph.microsoft.com/.default"];
  
  
  /**
   * Interactive authentication like MSAL - starts temporary server and opens browser
   */
  async authenticateInteractive(): Promise<string | null> {
    // Find an available port for the temporary server
    const redirectPort = await this.findAvailablePort(8081);
    const redirectUri = `http://localhost:${redirectPort}`;
    
    for (const clientId of this.clientIds) {
      try {
        console.log(`Trying authentication with client ${clientId}...`);
        
        // Start temporary server to catch the redirect
        const server = await this.startTemporaryServer(redirectPort);
        
        try {
          // Generate auth URL
          const state = crypto.randomUUID();
          const authUrl = this.generateAuthUrl(clientId, redirectUri, state);
          
          console.log(`Opening browser for authentication...`);
          try {
            if (Deno.build.os === "windows") {
              // Prefer rundll32 to avoid cmd parsing of '&' in query string
              try {
                const proc1 = new Deno.Command("rundll32", {
                  args: ["url.dll,FileProtocolHandler", authUrl],
                  stdout: "null",
                  stderr: "null",
                });
                proc1.spawn();
              } catch (_e) {
                // Fallback: cmd /c start with quoted URL and empty title argument
                const quotedUrl = `"${authUrl}"`;
                const proc2 = new Deno.Command("cmd", {
                  args: ["/c", "start", "", quotedUrl],
                  stdout: "null",
                  stderr: "null",
                });
                proc2.spawn();
              }
            } else {
              const openCommand = Deno.build.os === "darwin" ? "open" : "xdg-open";
              const proc = new Deno.Command(openCommand, {
                args: [authUrl],
                stdout: "null",
                stderr: "null",
              });
              proc.spawn();
            }
          } catch (e) {
            console.log("Failed to auto-open browser, please open this URL manually:", authUrl, e);
          }
          
          // Wait for the redirect with timeout
          const result = await Promise.race<AuthCallbackResult>([
            server.waitForCallback(),
            this.timeout(300000) // 5 minute timeout
          ]);
          
          if (result && result.code && result.state === state) {
            // Exchange code for token
            const accessToken = await this.exchangeCodeForToken(clientId, result.code, redirectUri);
            if (accessToken) {
              console.log(`✅ Authentication successful with client ${clientId}!`);
              return accessToken;
            }
          }
        } finally {
          server.close();
        }
      } catch (error) {
        console.log(`❌ Authentication failed with client ${clientId}:`, error);
        continue;
      }
    }
    
    return null;
  }

  private async findAvailablePort(startPort: number): Promise<number> {
    for (let port = startPort; port < startPort + 100; port++) {
      try {
        const listener = Deno.listen({ port });
        listener.close();
        return port;
      } catch {
        continue;
      }
    }
    throw new Error("No available ports found");
  }

  private async startTemporaryServer(port: number) {
    const listener = Deno.listen({ port });
    let resolveCallback!: (result: AuthCallbackResult) => void;
    let rejectCallback!: (error: unknown) => void;
    
    const callbackPromise = new Promise<AuthCallbackResult>((resolve, reject) => {
      resolveCallback = resolve;
      rejectCallback = reject;
    });
    
    const handleConnections = async () => {
      try {
        for await (const conn of listener) {
          this.handleConnection(conn, resolveCallback);
        }
      } catch (error) {
        if (!error.message.includes("Bad resource ID")) {
          rejectCallback(error);
        }
      }
    };
    
    handleConnections();
    
    return {
      waitForCallback: () => callbackPromise,
      close: () => {
        try {
          listener.close();
        } catch {
          // Ignore close errors
        }
      }
    };
  }

  private async handleConnection(conn: Deno.Conn, resolve: (result: AuthCallbackResult) => void) {
    try {
      const httpConn = Deno.serveHttp(conn);
      
      for await (const requestEvent of httpConn) {
        const url = new URL(requestEvent.request.url);
        const code = url.searchParams.get("code");
        const state = url.searchParams.get("state");
        const error = url.searchParams.get("error");
        
        const responseBody = error 
          ? `<html><body><h1>Authentication Failed</h1><p>Error: ${error}</p><p>You can close this window.</p></body></html>`
          : `<html><body><h1>Authentication Successful!</h1><p>You can close this window.</p><script>setTimeout(() => window.close(), 2000);</script></body></html>`;
        
        await requestEvent.respondWith(new Response(responseBody, {
          headers: { "Content-Type": "text/html" }
        }));
        
        resolve({ code, state, error });
        return;
      }
    } catch (error) {
      console.log("Connection handling error:", error);
    }
  }

  private generateAuthUrl(clientId: string, redirectUri: string, state: string): string {
    const authUrl = new URL(`https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/authorize`);
    
    authUrl.searchParams.set("client_id", clientId);
    authUrl.searchParams.set("response_type", "code");
    authUrl.searchParams.set("redirect_uri", redirectUri);
    authUrl.searchParams.set("scope", this.scopes.join(" "));
    authUrl.searchParams.set("state", state);
    authUrl.searchParams.set("prompt", "select_account");

    return authUrl.toString();
  }

  private async exchangeCodeForToken(clientId: string, code: string, redirectUri: string): Promise<string | null> {
    try {
      const response = await fetch(
        `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
          body: new URLSearchParams({
            grant_type: "authorization_code",
            client_id: clientId,
            code: code,
            redirect_uri: redirectUri,
          }),
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.access_token;
      } else {
        const errorData = await response.json();
        console.log("Token exchange failed:", errorData);
        return null;
      }
    } catch (error) {
      console.log(`Token exchange failed:`, error);
      return null;
    }
  }

  private timeout(ms: number): Promise<AuthCallbackResult> {
    return new Promise((_, reject) => {
      setTimeout(() => reject({ error: "Authentication timeout" }), ms);
    });
  }

  /**
   * Get user planners from Microsoft Graph
   */
  async getPlanners(accessToken: string): Promise<Planner[]> {
    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    try {
      // Try to get teams first
      const teamsResponse = await fetch(
        "https://graph.microsoft.com/v1.0/me/joinedTeams",
        { headers }
      );

      let planners: Planner[] = [];

      if (teamsResponse.ok) {
        const teamsData = await teamsResponse.json();
        const teams = teamsData.value || [];

        for (const team of teams) {
          try {
            const plansResponse = await fetch(
              `https://graph.microsoft.com/v1.0/groups/${team.id}/planner/plans`,
              { headers }
            );

            if (plansResponse.ok) {
              const plansData = await plansResponse.json();
              const plans = plansData.value || [];

              for (const plan of plans) {
                planners.push({
                  id: plan.id,
                  title: plan.title,
                  groupId: team.id,
                  groupName: team.displayName || "Unknown Group",
                });
              }
            }
          } catch (error) {
            console.log(`Error getting plans for team ${team.id}:`, error);
          }
        }
      } else {
        // Fallback to groups
        const groupsResponse = await fetch(
          "https://graph.microsoft.com/v1.0/me/memberOf?$filter=groupTypes/any(c:c eq 'Unified')",
          { headers }
        );

        if (groupsResponse.ok) {
          const groupsData = await groupsResponse.json();
          const groups = groupsData.value || [];

          for (const group of groups) {
            try {
              const plansResponse = await fetch(
                `https://graph.microsoft.com/v1.0/groups/${group.id}/planner/plans`,
                { headers }
              );

              if (plansResponse.ok) {
                const plansData = await plansResponse.json();
                const plans = plansData.value || [];

                        for (const plan of plans) {
                          planners.push({
                            id: plan.id,
                            title: plan.title,
                            groupId: group.id,
                            groupName: group.displayName || "Unknown Group",
                          });
                }
              }
            } catch (error) {
              console.log(`Error getting plans for group ${group.id}:`, error);
            }
          }
        }
      }

      return planners;
    } catch (error) {
      console.error("Error getting planners:", error);
      return [];
    }
  }

  /**
   * Get buckets for a specific planner
   */
  async getPlannerBuckets(accessToken: string, planId: string): Promise<Bucket[]> {
    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets`,
        { headers }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value || [];
      }
    } catch (error) {
      console.error("Error getting buckets:", error);
    }

    return [];
  }

  /**
   * Create a new bucket in a planner
   */
  async createBucket(accessToken: string, planId: string, bucketName: string): Promise<Bucket | null> {
    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    const bucketData = {
      name: bucketName,
      planId: planId,
      orderHint: " !",
    };

    try {
      const response = await fetch(
        "https://graph.microsoft.com/v1.0/planner/buckets",
        {
          method: "POST",
          headers,
          body: JSON.stringify(bucketData),
        }
      );

      if (response.ok) {
        return await response.json();
      }
    } catch (error) {
      console.error(`Error creating bucket ${bucketName}:`, error);
    }

    return null;
  }

  /**
   * Parse display name to extract components
   */
  parseDisplayName(displayName: string): { firstName: string; lastName: string; company: string; fullName: string } {
    if (!displayName) {
      return { firstName: "", lastName: "", company: "", fullName: "" };
    }

    // Pattern to match "FirstName LastName (COMPANY)" format
    const pattern = /^(.+?)\s+(.+?)\s*\((.+?)\)\s*$/;
    const match = displayName.trim().match(pattern);

    if (match) {
      const firstName = match[1].trim();
      const lastName = match[2].trim();
      const company = match[3].trim();

      return {
        firstName,
        lastName,
        company,
        fullName: `${firstName} ${lastName}`,
      };
    }

    // Try simpler pattern "FirstName LastName"
    const simplePattern = /^(.+?)\s+(.+?)$/;
    const simpleMatch = displayName.trim().match(simplePattern);

    if (simpleMatch) {
      const firstName = simpleMatch[1].trim();
      const lastName = simpleMatch[2].trim();

      return {
        firstName,
        lastName,
        company: "",
        fullName: `${firstName} ${lastName}`,
      };
    }

    // If no pattern matches, return the original as full name
    return {
      firstName: "",
      lastName: "",
      company: "",
      fullName: displayName,
    };
  }

  async getGroupMembers(accessToken: string, groupId: string): Promise<PlannerMember[]> {
    if (!groupId) {
      return [];
    }

    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    const members: PlannerMember[] = [];
    let nextUrl: string | null = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName,givenName,surname,companyName`;

    while (nextUrl) {
      try {
        const response = await fetch(nextUrl, { headers });
        if (!response.ok) {
          break;
        }

        const data = await response.json();
        const values = data.value || [];
        for (const entry of values) {
          if (entry.id) {
            members.push({
              id: entry.id,
              displayName: entry.displayName || "",
              mail: entry.mail || entry.userPrincipalName,
              userPrincipalName: entry.userPrincipalName,
              givenName: entry.givenName,
              surname: entry.surname,
              companyName: entry.companyName,
            });
          }
        }

        nextUrl = data["@odata.nextLink"] || null;
      } catch (_error) {
        break;
      }
    }

    return members;
  }

  async getPlannerDetails(accessToken: string, planId: string): Promise<any | null> {
    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/planner/plans/${planId}/details`,
        { headers }
      );

      if (!response.ok) {
        return null;
      }

      return await response.json();
    } catch (_error) {
      return null;
    }
  }

  async getPlannerMembers(accessToken: string, planId: string, groupId?: string): Promise<PlannerMember[]> {
    const members: PlannerMember[] = [];
    const seen = new Map<string, PlannerMember>();

    const groupMembers = await this.getGroupMembers(accessToken, groupId || "");
    for (const member of groupMembers) {
      if (!seen.has(member.id)) {
        seen.set(member.id, member);
        members.push(member);
      }
    }

    const details = await this.getPlannerDetails(accessToken, planId);
    const sharedWith = details?.sharedWith ? Object.keys(details.sharedWith) : [];

    for (const userId of sharedWith) {
      if (seen.has(userId)) {
        continue;
      }

      const headers = {
        "Authorization": `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      };

      try {
        const response = await fetch(
          `https://graph.microsoft.com/v1.0/users/${userId}?$select=id,displayName,mail,userPrincipalName,givenName,surname,companyName`,
          { headers }
        );

        if (!response.ok) {
          continue;
        }

        const user = await response.json();
        const member: PlannerMember = {
          id: user.id,
          displayName: user.displayName || "",
          mail: user.mail || user.userPrincipalName,
          userPrincipalName: user.userPrincipalName,
          givenName: user.givenName,
          surname: user.surname,
          companyName: user.companyName,
        };

        seen.set(member.id, member);
        members.push(member);
      } catch (_error) {
        continue;
      }
    }

    return members;
  }

  /**
   * Search for a user by display name
   */
  async searchUser(accessToken: string, assigneeName: string): Promise<User | null> {
    if (!assigneeName || assigneeName.trim() === "") {
      return null;
    }

    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    const parsedName = this.parseDisplayName(assigneeName);
    
    const searchTerms = [
      parsedName.fullName,
      assigneeName,
      parsedName.firstName,
      parsedName.lastName,
    ].filter(term => term);

    for (const searchTerm of searchTerms) {
      const searchQueries = [
        `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${searchTerm}'&$select=id,displayName,mail,userPrincipalName`,
        `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${searchTerm}')&$select=id,displayName,mail,userPrincipalName`,
        `https://graph.microsoft.com/v1.0/users?$search="displayName:${searchTerm}"&$select=id,displayName,mail,userPrincipalName`,
      ];

      for (const queryUrl of searchQueries) {
        try {
          const response = await fetch(queryUrl, { headers });

          if (response.ok) {
            const data = await response.json();
            const users = data.value || [];

            // Look for exact or close matches
            for (const user of users) {
              const userDisplayName = user.displayName || "";

              // Check for exact match first
              if (userDisplayName.toLowerCase() === parsedName.fullName.toLowerCase()) {
                return {
                  ...user,
                  originalName: assigneeName,
                };
              }

              // Check if user display name contains the search components
              if (
                parsedName.firstName &&
                parsedName.lastName &&
                userDisplayName.toLowerCase().includes(parsedName.firstName.toLowerCase()) &&
                userDisplayName.toLowerCase().includes(parsedName.lastName.toLowerCase())
              ) {
                return {
                  ...user,
                  originalName: assigneeName,
                };
              }
            }

            // If exact matches not found, return first result for exact search
            if (users.length > 0 && queryUrl.includes("eq")) {
              return {
                ...users[0],
                originalName: assigneeName,
              };
            }
          }
        } catch (error) {
          continue; // Try next search approach
        }
      }
    }

    return null;
  }

  /**
   * Create a task in Microsoft Planner
   */
  async createTask(
    accessToken: string,
    planId: string,
    bucketId: string,
    taskData: TaskData,
    assignees?: PlannerMember[]
  ): Promise<any> {
    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    // Step 1: Create the basic task
    const basicTaskData = {
      planId,
      bucketId,
      title: taskData.title,
      ...(taskData.dueDateTime && { dueDateTime: taskData.dueDateTime }),
      ...(taskData.startDateTime && { startDateTime: taskData.startDateTime }),
      ...(taskData.percentComplete !== undefined && { percentComplete: taskData.percentComplete }),
    };

    try {
      const response = await fetch(
        "https://graph.microsoft.com/v1.0/planner/tasks",
        {
          method: "POST",
          headers,
          body: JSON.stringify(basicTaskData),
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to create task: ${response.status} ${response.statusText}`);
      }

      const task = await response.json();
      const taskId = task.id;

      // Step 2: Update description if provided
      if (taskData.description) {
        await this.updateTaskDescription(accessToken, taskId, taskData.description);
      }

      // Step 3: Assign users if provided
      const assignedUsers: User[] = [];
      if (assignees && assignees.length > 0) {
        for (const assignee of assignees) {
          if (!assignee || !assignee.id) {
            continue;
          }

          const assignSuccess = await this.assignTask(accessToken, taskId, assignee.id);
          if (assignSuccess) {
            assignedUsers.push({
              id: assignee.id,
              displayName: assignee.displayName,
              mail: assignee.mail,
              userPrincipalName: assignee.userPrincipalName,
              originalName: assignee.displayName,
            });
          }
        }
      }

      return {
        ...task,
        assignedUsers,
      };
    } catch (error) {
      console.error("Error creating task:", error);
      return null;
    }
  }

  /**
   * Update task description
   */
  private async updateTaskDescription(accessToken: string, taskId: string, description: string): Promise<boolean> {
    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    try {
      // Get task details to retrieve ETag
      const getResponse = await fetch(
        `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`,
        { headers }
      );

      if (!getResponse.ok) {
        return false;
      }

      const etag = getResponse.headers.get("ETag");
      if (!etag) {
        return false;
      }

      // Update the description
      const updateResponse = await fetch(
        `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`,
        {
          method: "PATCH",
          headers: {
            ...headers,
            "If-Match": etag,
          },
          body: JSON.stringify({ description }),
        }
      );

      return updateResponse.ok;
    } catch (error) {
      console.error("Error updating task description:", error);
      return false;
    }
  }

  /**
   * Assign a task to a user
   */
  private async assignTask(accessToken: string, taskId: string, userId: string): Promise<boolean> {
    const headers = {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    try {
      // Get task to retrieve ETag
      const taskResponse = await fetch(
        `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`,
        { headers }
      );

      if (!taskResponse.ok) {
        return false;
      }

      const etag = taskResponse.headers.get("ETag");
      if (!etag) {
        return false;
      }

      // Create assignment
      const assignments = {
        [userId]: {
          "@odata.type": "microsoft.graph.plannerAssignment",
          orderHint: " !",
        },
      };

      const assignmentResponse = await fetch(
        `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`,
        {
          method: "PATCH",
          headers: {
            ...headers,
            "If-Match": etag,
          },
          body: JSON.stringify({ assignments }),
        }
      );

      return assignmentResponse.ok;
    } catch (error) {
      console.error("Error assigning task:", error);
      return false;
    }
  }
}