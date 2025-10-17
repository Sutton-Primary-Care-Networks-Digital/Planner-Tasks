/**
 * Microsoft Planner Task Creator - Deno Web Application
 * Main server entry point
 */

// Global error handler for stream controller issues
globalThis.addEventListener("unhandledrejection", (event) => {
  if (event.reason?.message?.includes("stream controller")) {
    console.log("Ignoring stream controller error (Oak Node.js compatibility issue)");
    event.preventDefault();
  }
});

globalThis.addEventListener("error", (event) => {
  if (event.error?.message?.includes("stream controller")) {
    console.log("Ignoring stream controller error (Oak Node.js compatibility issue)");
    event.preventDefault();
  }
});

import { Application, Router } from "https://deno.land/x/oak@v12.6.2/mod.ts";
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

function generateState(): string {
  return crypto.randomUUID();
}


const router = new Router();

// Serve the main HTML page
router.get("/", async (ctx) => {
  const html = await Deno.readTextFile("./static/index.html");
  ctx.response.body = html;
  ctx.response.headers.set("Content-Type", "text/html");
});

// Serve static assets
router.get("/static/:filename", async (ctx) => {
  const filename = ctx.params.filename;
  if (!filename) {
    ctx.response.status = 404;
    return;
  }

  const ext = filename.split('.').pop()?.toLowerCase();
  let contentType = "text/plain";

  switch (ext) {
    case "css":
      contentType = "text/css";
      break;
    case "js":
      contentType = "application/javascript";
      break;
    case "html":
      contentType = "text/html";
      break;
    case "json":
      contentType = "application/json";
      break;
  }

  try {
    const file = await Deno.readTextFile(`./static/${filename}`);
    ctx.response.body = file;
    ctx.response.headers.set("Content-Type", contentType);
  } catch {
    ctx.response.status = 404;
    ctx.response.body = "File not found";
  }
});


// Check authentication status
router.get("/api/auth/status", (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  if (sessionId && sessions.has(sessionId)) {
    const sessionData = sessions.get(sessionId);
    if (sessionData.authenticated && sessionData.accessToken) {
      ctx.response.body = { authenticated: true, sessionId };
      return;
    }
  }
  ctx.response.body = { authenticated: false };
});

// Interactive authentication - like MSAL
router.post("/api/auth/interactive", async (ctx) => {
  try {
    const accessToken = await auth.authenticateInteractive();
    if (accessToken) {
      // Store token in session
      const sessionId = generateSessionId();
      sessions.set(sessionId, {
        accessToken,
        authenticated: true,
        created: Date.now(),
      });
      saveSessions(); // Persist authentication

      ctx.response.body = { success: true, sessionId };
    } else {
      ctx.response.status = 500;
      ctx.response.body = { error: "Authentication failed" };
    }
  } catch (error) {
    console.error("Interactive auth error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Authentication failed" };
  }
});

// Get planners
router.get("/api/planners", async (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  if (!sessionId || !sessions.has(sessionId)) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  const sessionData = sessions.get(sessionId);
  if (!sessionData.authenticated || !sessionData.accessToken) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  try {
    const planners = await auth.getPlanners(sessionData.accessToken);
    ctx.response.body = { planners };
  } catch (error) {
    console.error("Get planners error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Failed to get planners" };
  }
});

// Get planner buckets
router.get("/api/planners/:planId/buckets", async (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  const planId = ctx.params.planId;
  
  if (!sessionId || !sessions.has(sessionId)) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  const sessionData = sessions.get(sessionId);
  if (!sessionData.authenticated || !sessionData.accessToken) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  try {
    const buckets = await auth.getPlannerBuckets(sessionData.accessToken, planId);
    ctx.response.body = { buckets };
  } catch (error) {
    console.error("Get buckets error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Failed to get buckets" };
  }
});

// Parse uploaded file
router.post("/api/parse-file", async (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  if (!sessionId || !sessions.has(sessionId)) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  try {
    console.log("Request content type:", ctx.request.headers.get("content-type"));
    
    // Oak v12 body parsing - body is a function with type options
    const body = ctx.request.body({ type: "form-data" });
    const formData = await body.value;
    console.log("Processing form data...");
    
    let fileContent = "";
    let fileName = "";
    
    console.log("Extracting file from form data...");
    // Oak v12 FormDataReader - use .read() method
    const formDataBody = await formData.read();
    console.log("FormData files:", formDataBody.files?.length || 0);
    
    if (formDataBody.files && formDataBody.files.length > 0) {
      const file = formDataBody.files[0];
      fileName = file.originalName || file.filename || "unknown";
      if (file.content) {
        fileContent = new TextDecoder().decode(file.content);
      } else if (file.filename) {
        // File was written to disk, read it
        fileContent = await Deno.readTextFile(file.filename);
      }
      console.log(`File extracted: ${fileName}, size: ${fileContent.length} chars`);
    }

    if (!fileContent || !fileName) {
      console.log("No file content or name:", { fileContent: !!fileContent, fileName });
      ctx.response.status = 400;
      ctx.response.body = { error: "No file uploaded" };
      return;
    }
    
    console.log(`Parsing file: ${fileName}`);
    let data: Record<string, string>[];
    if (fileName.endsWith('.csv')) {
      data = await fileParser.parseCSV(fileContent);
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      data = await fileParser.parseExcel(fileContent);
    } else {
      ctx.response.status = 400;
      ctx.response.body = { error: "Unsupported file format" };
      return;
    }

    const availableColumns = fileParser.getAvailableColumns(data);
    
    // Store parsed data in session
    const sessionData = sessions.get(sessionId);
    sessionData.parsedData = data;
    sessionData.availableColumns = availableColumns;
    sessions.set(sessionId, sessionData);
    saveSessions(); // Persist session data

    console.log(`File processed successfully: ${data.length} rows, ${availableColumns.length} columns`);
    ctx.response.headers.set("Content-Type", "application/json");
    ctx.response.status = 200;
    ctx.response.body = {
      success: true,
      rowCount: data.length,
      availableColumns,
      preview: data.slice(0, 5), // First 5 rows for preview
    };

  } catch (error) {
    console.error("File parse error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Failed to parse file" };
  }
});

// Process file data with column mapping
router.post("/api/process-data", async (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  if (!sessionId || !sessions.has(sessionId)) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  const sessionData = sessions.get(sessionId);
  if (!sessionData.parsedData) {
    ctx.response.status = 400;
    ctx.response.body = { error: "No parsed data available" };
    return;
  }

  try {
    const body = ctx.request.body({ type: "json" });
    const data = await body.value;
    const { columnMapping } = data;

    const tasks = fileParser.processData(sessionData.parsedData, columnMapping);
    const validation = fileParser.validateTasks(tasks);

    if (!validation.valid) {
      ctx.response.status = 400;
      ctx.response.body = { error: "Invalid task data", errors: validation.errors };
      return;
    }

    // Store processed tasks
    sessionData.processedTasks = tasks;
    sessions.set(sessionId, sessionData);
    saveSessions(); // Persist session data

    const assigneeStats = fileParser.getAssigneeStatistics(tasks);
    const bucketStats = fileParser.getBucketStatistics(tasks);

    ctx.response.body = {
      success: true,
      taskCount: tasks.length,
      tasks: tasks.slice(0, 10), // First 10 for preview
      assigneeStats,
      bucketStats,
    };

  } catch (error) {
    console.error("Process data error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Failed to process data" };
  }
});

// Lookup assignees
router.post("/api/lookup-assignees", async (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  if (!sessionId || !sessions.has(sessionId)) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  const sessionData = sessions.get(sessionId);
  if (!sessionData.processedTasks || !sessionData.authenticated) {
    ctx.response.status = 400;
    ctx.response.body = { error: "No processed tasks or not authenticated" };
    return;
  }

  try {
    const body = ctx.request.body({ type: "json" });
    const { planId, groupId } = await body.value;

    // Fetch planner members (scoped to selected planner)
    let members: any[] = [];
    try {
      members = await auth.getPlannerMembers(sessionData.accessToken, planId, groupId);
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
    saveSessions(); // Persist session data

    ctx.response.body = {
      success: true,
      tasks: enrichedTasks.slice(0, 10), // First 10 for preview
    };

  } catch (error) {
    console.error("Lookup assignees error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Failed to lookup assignees" };
  }
});

// Get planner members for a plan (and its group)
router.get("/api/planners/:planId/members", async (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  const planId = ctx.params.planId;
  const url = new URL(ctx.request.url);
  const groupId = url.searchParams.get("groupId") || undefined;

  if (!sessionId || !sessions.has(sessionId)) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  const sessionData = sessions.get(sessionId);
  if (!sessionData.authenticated || !sessionData.accessToken) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  try {
    const members = await auth.getPlannerMembers(sessionData.accessToken, planId!, groupId);
    ctx.response.body = { members };
  } catch (error) {
    console.error("Get members error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Failed to get members" };
  }
});

// Create tasks
router.post("/api/create-tasks", async (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  if (!sessionId || !sessions.has(sessionId)) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  const sessionData = sessions.get(sessionId);
  if (!sessionData.enrichedTasks || !sessionData.authenticated) {
    ctx.response.status = 400;
    ctx.response.body = { error: "No processed tasks or not authenticated" };
    return;
  }

  try {
    const body = ctx.request.body({ type: "json" });
    const data = await body.value;
    const { planId, bucketId } = data;

    const tasks = sessionData.enrichedTasks;
    const results = [];
    let created = 0;
    let failed = 0;
    let assigned = 0;

    for (const task of tasks) {
      try {
        const taskData = {
          title: task.title,
          description: task.description,
          dueDateTime: task.dueDate,
          startDateTime: task.startDate,
          percentComplete: task.status ? fileParser.mapStatusToPercentage(task.status) : 0,
        };

        // Determine assignees (resolved PlannerMember objects)
        const assignees = Array.isArray(task.assigneeUsers) ? task.assigneeUsers : [];

        // Determine bucket ID
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

    ctx.response.body = {
      success: true,
      summary: {
        total: tasks.length,
        created,
        failed,
        assigned,
      },
      results,
    };

  } catch (error) {
    console.error("Create tasks error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Failed to create tasks" };
  }
});

// Create bucket
router.post("/api/create-bucket", async (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  if (!sessionId || !sessions.has(sessionId)) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  const sessionData = sessions.get(sessionId);
  if (!sessionData.authenticated) {
    ctx.response.status = 401;
    ctx.response.body = { error: "Not authenticated" };
    return;
  }

  try {
    const body = ctx.request.body({ type: "json" });
    const data = await body.value;
    const { planId, bucketName } = data;

    const bucket = await auth.createBucket(sessionData.accessToken, planId, bucketName);
    
    if (bucket) {
      ctx.response.body = { success: true, bucket };
    } else {
      ctx.response.status = 500;
      ctx.response.body = { error: "Failed to create bucket" };
    }

  } catch (error) {
    console.error("Create bucket error:", error);
    ctx.response.status = 500;
    ctx.response.body = { error: "Failed to create bucket" };
  }
});

// Sign out
router.post("/api/auth/signout", (ctx) => {
  const sessionId = ctx.request.headers.get("X-Session-ID");
  if (sessionId && sessions.has(sessionId)) {
    sessions.delete(sessionId);
  }
  
  ctx.response.body = { success: true };
});

// Create the Oak application
const app = new Application();

// CORS middleware
app.use(async (ctx, next) => {
  ctx.response.headers.set("Access-Control-Allow-Origin", "*");
  ctx.response.headers.set("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
  ctx.response.headers.set("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Session-ID");
  
  if (ctx.request.method === "OPTIONS") {
    ctx.response.status = 200;
    return;
  }
  
  await next();
});

// Error handling middleware
app.use(async (ctx, next) => {
  try {
    await next();
  } catch (error) {
    console.error("Server error:", error);
    if (!ctx.response.headersSent) {
      ctx.response.status = 500;
      ctx.response.body = { error: "Internal server error" };
    }
  }
});

// Add global response headers and handle stream errors
app.use(async (ctx, next) => {
  try {
    // Set default headers before processing
    if (ctx.request.url.pathname.startsWith("/api/")) {
      ctx.response.headers.set("Content-Type", "application/json");
    }
    await next();
  } catch (error) {
    // Silently handle stream controller errors
    if (error.message?.includes("stream controller")) {
      console.log("Stream completed (ignoring controller error)");
      return;
    }
    throw error;
  }
});

// Use the router
app.use(router.routes());
app.use(router.allowedMethods());

const port = 8080;
console.log(`🚀 Microsoft Planner Task Creator running on http://localhost:${port}`);
console.log(`📝 Ready to process CSV/Excel files and create Microsoft Planner tasks!`);

await app.listen({ port });