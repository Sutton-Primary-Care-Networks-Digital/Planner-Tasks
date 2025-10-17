/**
 * Microsoft Planner Task Creator - Native Deno HTTP Server
 * Using native Deno APIs instead of Oak for reliable form handling
 */

import { GraphAuth } from "./auth.ts";
import { FileParser, type ParsedTask, type ProcessedTask } from "./file-parser.ts";

const auth = new GraphAuth();
const fileParser = new FileParser();

// Session storage with file persistence
const SESSIONS_FILE = "./sessions.json";
const sessions: Map<string, any> = new Map();

// Load existing sessions on startup
function loadSessions() {
  try {
    const data = Deno.readTextFileSync(SESSIONS_FILE);
    const sessionData = JSON.parse(data);
    for (const [key, value] of Object.entries(sessionData)) {
      sessions.set(key, value);
    }
    console.log(`Loaded ${sessions.size} existing sessions`);
  } catch {
    console.log("No existing sessions found, starting fresh");
  }
}

// Save sessions to file
function saveSessions() {
  try {
    const sessionData = Object.fromEntries(sessions);
    Deno.writeTextFileSync(SESSIONS_FILE, JSON.stringify(sessionData, null, 2));
  } catch (error) {
    console.error("Failed to save sessions:", error);
  }
}

// Load sessions on startup
loadSessions();

function generateSessionId(): string {
  return crypto.randomUUID();
}

// CORS headers
const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Session-ID",
};

// Parse multipart form data
async function parseMultipartForm(request: Request): Promise<FormData | null> {
  const contentType = request.headers.get("content-type");
  if (!contentType?.startsWith("multipart/form-data")) {
    return null;
  }

  try {
    return await request.formData();
  } catch (error) {
    console.error("Failed to parse form data:", error);
    return null;
  }
}

// Parse JSON body
async function parseJSON(request: Request): Promise<any> {
  try {
    return await request.json();
  } catch (error) {
    console.error("Failed to parse JSON:", error);
    return null;
  }
}

// Serve static files
async function serveStaticFile(pathname: string): Promise<Response> {
  try {
    const filePath = pathname.startsWith("/static/")
      ? `.${pathname}`
      : `./static${pathname}`;
    console.log(`Serving static file: ${filePath}`);
    const file = await Deno.readTextFile(filePath);
    
    let contentType = "text/plain";
    if (pathname.endsWith(".html")) contentType = "text/html";
    else if (pathname.endsWith(".css")) contentType = "text/css";
    else if (pathname.endsWith(".js")) contentType = "application/javascript";
    else if (pathname.endsWith(".json")) contentType = "application/json";
    
    console.log(`✅ Static file served: ${filePath} (${contentType})`);
    return new Response(file, {
      headers: { ...corsHeaders, "Content-Type": contentType }
    });
  } catch (error) {
    console.error(`❌ Failed to serve static file ${pathname}:`, error);
    return new Response("File not found", { status: 404, headers: corsHeaders });
  }
}

// Main request handler
async function handleRequest(request: Request): Promise<Response> {
  const url = new URL(request.url);
  const { pathname, searchParams } = url;
  const method = request.method;
  
  console.log(`${method} ${pathname}`);

  // Handle CORS preflight
  if (method === "OPTIONS") {
    return new Response(null, { status: 200, headers: corsHeaders });
  }

  // Serve static files
  if (pathname === "/" || pathname === "/index.html") {
    return serveStaticFile("/index.html");
  }
  if (pathname.startsWith("/static/")) {
    return serveStaticFile(pathname);
  }

  // API routes
  try {
    // Auth status check
    if (pathname === "/api/auth/status" && method === "GET") {
      const sessionId = request.headers.get("X-Session-ID");
      if (sessionId && sessions.has(sessionId)) {
        const sessionData = sessions.get(sessionId);
        if (sessionData.authenticated && sessionData.accessToken) {
          return new Response(JSON.stringify({ authenticated: true, sessionId }), {
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }
      }
      return new Response(JSON.stringify({ authenticated: false }), {
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    // Interactive authentication
    if (pathname === "/api/auth/interactive" && method === "POST") {
      try {
        const accessToken = await auth.authenticateInteractive();
        if (accessToken) {
          const sessionId = generateSessionId();
          sessions.set(sessionId, {
            accessToken,
            authenticated: true,
            created: Date.now(),
          });
          saveSessions();

          return new Response(JSON.stringify({ success: true, sessionId }), {
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        } else {
          return new Response(JSON.stringify({ error: "Authentication failed" }), {
            status: 500,
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }
      } catch (error) {
        console.error("Interactive auth error:", error);
        return new Response(JSON.stringify({ error: "Authentication failed" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Create buckets for a plan
    if (pathname === "/api/buckets/create" && method === "POST") {
      const sessionId = request.headers.get("X-Session-ID");
      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const sessionData = sessions.get(sessionId);
      if (!sessionData.authenticated || !sessionData.accessToken) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const body = await parseJSON(request);
        const planId = body?.planId as string | undefined;
        const bucketNames: string[] = Array.isArray(body?.bucketNames) ? body.bucketNames : [];
        if (!planId || bucketNames.length === 0) {
          return new Response(JSON.stringify({ error: "Missing planId or bucketNames" }), {
            status: 400,
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }

        // Avoid duplicates by comparing lower-cased names
        const existing = await auth.getPlannerBuckets(sessionData.accessToken, planId);
        const existingLower = new Set((existing || []).map((b: any) => (b.name || "").toLowerCase()));
        const toCreate = bucketNames
          .map((n) => (typeof n === 'string' ? n.trim() : ''))
          .filter((n) => n.length > 0 && !existingLower.has(n.toLowerCase()));

        const created: any[] = [];
        for (const name of toCreate) {
          try {
            const res = await auth.createBucket(sessionData.accessToken, planId, name);
            if (res) created.push(res);
          } catch (e) {
            console.error("Create bucket failed for", name, e);
          }
        }

        const buckets = await auth.getPlannerBuckets(sessionData.accessToken, planId);
        return new Response(JSON.stringify({ success: true, created: created.length, buckets }), {
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      } catch (error) {
        console.error("Create buckets error:", error);
        return new Response(JSON.stringify({ error: "Failed to create buckets" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Lookup buckets for current tasks in a plan
    if (pathname === "/api/lookup-buckets" && method === "POST") {
      const sessionId = request.headers.get("X-Session-ID");
      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
      const sessionData = sessions.get(sessionId);
      if (!sessionData.authenticated || !sessionData.accessToken) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
      try {
        const body = await parseJSON(request);
        const { planId } = body || {};
        if (!planId) {
          return new Response(JSON.stringify({ error: "Missing planId" }), { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } });
        }
        const tasks = sessionData.enrichedTasks || sessionData.processedTasks || [];
        const enriched = await fileParser.lookupBuckets(tasks, auth, sessionData.accessToken, planId);
        sessionData.enrichedTasks = enriched;
        return new Response(JSON.stringify({ tasks: enriched }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
      } catch (error) {
        console.error("lookup-buckets failed:", error);
        return new Response(JSON.stringify({ error: "Failed to lookup buckets" }), { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } });
      }
    }

    // Get planners
    if (pathname === "/api/planners" && method === "GET") {
      const sessionId = request.headers.get("X-Session-ID");
      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const sessionData = sessions.get(sessionId);
      if (!sessionData.authenticated || !sessionData.accessToken) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const planners = await auth.getPlanners(sessionData.accessToken);
        return new Response(JSON.stringify({ planners }), {
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      } catch (error) {
        console.error("Get planners error:", error);
        return new Response(JSON.stringify({ error: "Failed to get planners" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Get planner buckets
    if (pathname.startsWith("/api/planners/") && pathname.endsWith("/buckets") && method === "GET") {
      const sessionId = request.headers.get("X-Session-ID");
      const planId = pathname.split("/")[3];
      
      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const sessionData = sessions.get(sessionId);
      if (!sessionData.authenticated || !sessionData.accessToken) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const buckets = await auth.getPlannerBuckets(sessionData.accessToken, planId);
        return new Response(JSON.stringify({ buckets }), {
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      } catch (error) {
        console.error("Get buckets error:", error);
        return new Response(JSON.stringify({ error: "Failed to get buckets" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Parse uploaded file
    if (pathname === "/api/parse-file" && method === "POST") {
      const sessionId = request.headers.get("X-Session-ID");
      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        console.log("Request content type:", request.headers.get("content-type"));
        
        const formData = await parseMultipartForm(request);
        if (!formData) {
          return new Response(JSON.stringify({ error: "Expected form data" }), {
            status: 400,
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }

        console.log("Processing form data...");
        
        const file = formData.get("file") as File;
        if (!file) {
          return new Response(JSON.stringify({ error: "No file uploaded" }), {
            status: 400,
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }

        const fileName = file.name;
        const fileContent = await file.text();
        console.log(`File extracted: ${fileName}, size: ${fileContent.length} chars`);

        console.log(`Parsing file: ${fileName}`);
        let data: Record<string, string>[];
        if (fileName.endsWith('.csv')) {
          data = await fileParser.parseCSV(fileContent);
        } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
          data = await fileParser.parseExcel(fileContent);
        } else {
          return new Response(JSON.stringify({ error: "Unsupported file format" }), {
            status: 400,
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }

        const availableColumns = fileParser.getAvailableColumns(data);
        
        // Store parsed data in session
        const sessionData = sessions.get(sessionId);
        sessionData.parsedData = data;
        sessionData.availableColumns = availableColumns;
        sessions.set(sessionId, sessionData);
        saveSessions();

        console.log(`File processed successfully: ${data.length} rows, ${availableColumns.length} columns`);
        return new Response(JSON.stringify({
          success: true,
          rowCount: data.length,
          availableColumns,
          preview: data.slice(0, 5),
        }), {
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });

      } catch (error) {
        console.error("File parse error:", error);
        return new Response(JSON.stringify({ error: "Failed to parse file" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Process file data with column mapping
    if (pathname === "/api/process-data" && method === "POST") {
      const sessionId = request.headers.get("X-Session-ID");
      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const sessionData = sessions.get(sessionId);
      if (!sessionData.parsedData) {
        return new Response(JSON.stringify({ error: "No parsed data available" }), {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const body = await parseJSON(request);
        if (!body) {
          return new Response(JSON.stringify({ error: "Invalid JSON" }), {
            status: 400,
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }

        const { columnMapping } = body;
        const tasks = fileParser.processData(sessionData.parsedData, columnMapping);
        const validation = fileParser.validateTasks(tasks);

        if (!validation.valid) {
          return new Response(JSON.stringify({ error: "Invalid task data", errors: validation.errors }), {
            status: 400,
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }

        sessionData.processedTasks = tasks;
        sessions.set(sessionId, sessionData);
        saveSessions();

        const assigneeStats = fileParser.getAssigneeStatistics(tasks);
        const bucketStats = fileParser.getBucketStatistics(tasks);

        return new Response(JSON.stringify({
          success: true,
          taskCount: tasks.length,
          tasks: tasks.slice(0, 10),
          assigneeStats,
          bucketStats,
        }), {
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });

      } catch (error) {
        console.error("Process data error:", error);
        return new Response(JSON.stringify({ error: "Failed to process data" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Lookup assignees
    if (pathname === "/api/lookup-assignees" && method === "POST") {
      const sessionId = request.headers.get("X-Session-ID");
      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const sessionData = sessions.get(sessionId);
      if (!sessionData.processedTasks || !sessionData.authenticated) {
        return new Response(JSON.stringify({ error: "No processed tasks or not authenticated" }), {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const body = await parseJSON(request);
        const planId = body?.planId as string | undefined;
        const groupId = body?.groupId as string | undefined;

        let members: any[] = [];
        try {
          if (planId) {
            members = await auth.getPlannerMembers(sessionData.accessToken, planId, groupId);
          }
        } catch (_e) {
          members = [];
        }

        const enrichedTasks = await fileParser.lookupAssignees(
          sessionData.processedTasks,
          auth,
          sessionData.accessToken,
          members
        );

        sessionData.enrichedTasks = enrichedTasks;
        sessions.set(sessionId, sessionData);
        saveSessions();

        return new Response(JSON.stringify({
          success: true,
          tasks: enrichedTasks.slice(0, 10),
        }), {
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });

      } catch (error) {
        console.error("Lookup assignees error:", error);
        return new Response(JSON.stringify({ error: "Failed to lookup assignees" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Get planner members for a plan (and optional groupId)
    if (pathname.startsWith("/api/planners/") && pathname.endsWith("/members") && method === "GET") {
      const sessionId = request.headers.get("X-Session-ID");
      const planId = pathname.split("/")[3];
      const groupId = new URL(request.url).searchParams.get("groupId") || undefined;

      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const sessionData = sessions.get(sessionId);
      if (!sessionData.authenticated || !sessionData.accessToken) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const members = await auth.getPlannerMembers(sessionData.accessToken, planId, groupId);
        return new Response(JSON.stringify({ members }), {
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      } catch (error) {
        console.error("Get members error:", error);
        return new Response(JSON.stringify({ error: "Failed to get members" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Create tasks
    if (pathname === "/api/create-tasks" && method === "POST") {
      const sessionId = request.headers.get("X-Session-ID");
      if (!sessionId || !sessions.has(sessionId)) {
        return new Response(JSON.stringify({ error: "Not authenticated" }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const sessionData = sessions.get(sessionId);
      if (!sessionData.enrichedTasks || !sessionData.authenticated) {
        return new Response(JSON.stringify({ error: "No processed tasks or not authenticated" }), {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const body = await parseJSON(request);
        if (!body) {
          return new Response(JSON.stringify({ error: "Invalid JSON" }), {
            status: 400,
            headers: { ...corsHeaders, "Content-Type": "application/json" }
          });
        }

        const { planId, bucketId, overrides } = body;
        const tasks = sessionData.enrichedTasks;
        const results: any[] = [];
        let created = 0;
        let failed = 0;
        let assigned = 0;

        // Prepare member map for overrides if provided
        let memberById: Record<string, any> = {};
        if (overrides && planId) {
          try {
            const members = await auth.getPlannerMembers(sessionData.accessToken, planId);
            memberById = (members || []).reduce((acc: Record<string, any>, m: any) => {
              acc[m.id] = m;
              return acc;
            }, {});
          } catch (_) {
            memberById = {};
          }
        }

        for (let idx = 0; idx < tasks.length; idx++) {
          const task = tasks[idx];
          try {
            const taskData = {
              title: task.title,
              description: task.description,
              dueDateTime: task.dueDate,
              startDateTime: task.startDate,
              percentComplete: task.status ? fileParser.mapStatusToPercentage(task.status) : 0,
            };

            // Apply status override if provided
            if (body?.statusOverrides && Object.prototype.hasOwnProperty.call(body.statusOverrides, idx)) {
              const statusKey = body.statusOverrides[idx] as string; // e.g. 'not_started' | 'in_progress' | 'complete'
              if (typeof statusKey === 'string') {
                taskData.percentComplete = fileParser.mapStatusToPercentage(statusKey);
              }
            }

            // Apply overrides if provided for this task index
            let assignees = Array.isArray(task.assigneeUsers) ? task.assigneeUsers : [];
            if (overrides && overrides.hasOwnProperty(idx)) {
              const overrideIds: string[] = overrides[idx] || [];
              const resolved = overrideIds
                .map((id) => memberById[id])
                .filter(Boolean)
                .map((m) => ({ ...m, originalName: m.displayName, source: 'override' }));
              if (resolved.length > 0) {
                assignees = resolved;
              }
            }

            const taskBucketId = task.bucketInfo ? task.bucketInfo.id : bucketId;

            const result = await auth.createTask(
              sessionData.accessToken,
              planId,
              taskBucketId,
              taskData,
              assignees.length > 0 ? assignees : undefined
            );

            if (result) {
              created++;
              if (result.assignedUsers && result.assignedUsers.length > 0) {
                assigned += result.assignedUsers.length;
              }
              results.push({
                success: true,
                task: task.title,
                assignedUsers: result.assignedUsers || [],
              });
            } else {
              failed++;
              results.push({
                success: false,
                task: task.title,
                error: "Failed to create task",
              });
            }
          } catch (error) {
            failed++;
            results.push({
              success: false,
              task: task.title,
              error: error.message || "Unknown error",
            });
          }
        }

        return new Response(JSON.stringify({
          success: true,
          summary: {
            total: tasks.length,
            created,
            failed,
            assigned,
          },
          results,
        }), {
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });

      } catch (error) {
        console.error("Create tasks error:", error);
        return new Response(JSON.stringify({ error: "Failed to create tasks" }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    // Sign out
    if (pathname === "/api/auth/signout" && method === "POST") {
      const sessionId = request.headers.get("X-Session-ID");
      if (sessionId && sessions.has(sessionId)) {
        sessions.delete(sessionId);
        saveSessions();
      }
      
      return new Response(JSON.stringify({ success: true }), {
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

  } catch (error) {
    console.error("Server error:", error);
    return new Response(JSON.stringify({ error: "Internal server error" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" }
    });
  }

  // 404 for unmatched routes
  return new Response("Not Found", { status: 404, headers: corsHeaders });
}

const port = 8080;
console.log(`🚀 Microsoft Planner Task Creator running on http://localhost:${port}`);
console.log(`📝 Ready to process CSV/Excel files and create Microsoft Planner tasks!`);

Deno.serve({ port }, handleRequest);