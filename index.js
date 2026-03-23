import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  startDeviceCode,
  pollDeviceCode,
  deleteStoredAccount,
  getAccessToken,
  listStoredAccounts,
  setDefaultStoredAccount,
} from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

// Coercion helpers — MCP callers sometimes pass booleans/arrays/objects as JSON strings
const zBool = z.union([z.boolean(), z.string().transform((s) => s === "true")]);
const zStringArray = z.union([
  z.array(z.string()),
  z.string().transform((s) => JSON.parse(s)),
]);

const recurrenceObjectSchema = z.object({
  pattern: z.object({
    type:         z.enum(["daily", "weekly", "absoluteMonthly", "absoluteYearly", "relativeMonthly", "relativeYearly"]),
    interval:     z.number().int().positive().default(1),
    days_of_week: z.array(z.enum(["sunday","monday","tuesday","wednesday","thursday","friday","saturday"])).optional()
      .describe("Required for weekly recurrence, e.g. [\"monday\", \"wednesday\"]"),
    day_of_month: z.number().int().optional().describe("Day of month for absoluteMonthly/absoluteYearly (1-31)"),
    month:        z.number().int().optional().describe("Month for absoluteYearly (1-12)"),
  }),
  range: z.object({
    type:                  z.enum(["noEnd", "endDate", "numbered"]),
    start_date:            z.string().describe("YYYY-MM-DD — date of first occurrence"),
    end_date:              z.string().optional().describe("YYYY-MM-DD — required if type is endDate"),
    number_of_occurrences: z.number().int().optional().describe("Required if type is numbered"),
  }),
});

// Accept recurrence as an object OR a JSON string (since callers may stringify it)
const recurrenceSchema = z.union([
  recurrenceObjectSchema,
  z.string().transform((s) => recurrenceObjectSchema.parse(JSON.parse(s))),
]).describe("Recurrence rule (only settable at creation, not via update). pattern.type: daily|weekly|absoluteMonthly|absoluteYearly. range.type: noEnd|endDate|numbered");

const optionalAccountArg = {
  account: z.string().optional().describe("Stored account alias to use. Defaults to the configured default account."),
};

function isGraphError(payload) {
  return payload !== null
    && typeof payload === "object"
    && !Array.isArray(payload)
    && payload.error !== undefined;
}

async function readResponseBody(resp) {
  if (resp.status === 204 || resp.status === 205) return null;
  const text = await resp.text();
  if (!text) return null;
  const contentType = resp.headers.get("content-type") || "";
  if (contentType.includes("application/json")) return JSON.parse(text);
  return text;
}

function formatGraphError(method, pathOrUrl, resp, payload) {
  const status = `${resp.status} ${resp.statusText}`.trim();
  if (isGraphError(payload)) {
    const code = payload.error?.code;
    const message = payload.error?.message;
    const detail = [code, message].filter(Boolean).join(": ");
    return `${method} ${pathOrUrl} failed (${status})${detail ? `: ${detail}` : ""}`;
  }
  if (typeof payload === "string" && payload.trim()) {
    return `${method} ${pathOrUrl} failed (${status}): ${payload.trim()}`;
  }
  return `${method} ${pathOrUrl} failed (${status})`;
}

function toGraphUrl(pathOrUrl) {
  if (pathOrUrl.startsWith("http://") || pathOrUrl.startsWith("https://")) return pathOrUrl;
  return `${GRAPH}${pathOrUrl}`;
}

function buildTasksPath(listId) {
  return `/me/todo/lists/${encodeURIComponent(listId)}/tasks`;
}

function buildTaskPath(listId, taskId) {
  return `${buildTasksPath(listId)}/${encodeURIComponent(taskId)}`;
}

function buildChecklistPath(listId, taskId) {
  return `${buildTaskPath(listId, taskId)}/checklistItems`;
}

function buildChecklistItemPath(listId, taskId, itemId) {
  return `${buildChecklistPath(listId, taskId)}/${encodeURIComponent(itemId)}`;
}

function toDateOnly(dateTime) {
  return dateTime?.substring(0, 10) || null;
}

async function graphRequest(method, pathOrUrl, body, account) {
  const token = await getAccessToken(account);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  try {
    const resp = await fetch(toGraphUrl(pathOrUrl), {
      method,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined,
      signal: controller.signal,
    });
    const payload = await readResponseBody(resp);
    if (!resp.ok || isGraphError(payload)) {
      throw new Error(formatGraphError(method, pathOrUrl, resp, payload));
    }
    return payload ?? {};
  } finally {
    clearTimeout(timer);
  }
}

const graphGet  = (path, account)       => graphRequest("GET",    path, undefined, account);
const graphPost = (path, body, account) => graphRequest("POST",   path, body,      account);
const graphPatch= (path, body, account) => graphRequest("PATCH",  path, body,      account);
const graphDel  = (path, account)       => graphRequest("DELETE", path, undefined, account);

async function graphGetAllPages(pathOrUrl, account) {
  const items = [];
  let nextPage = pathOrUrl;
  while (nextPage) {
    const resp = await graphGet(nextPage, account);
    items.push(...(resp.value || []));
    nextPage = resp["@odata.nextLink"] || null;
  }
  return items;
}

function buildTaskPayload({ title, status, due_date, start_date, importance, body, reminder_date_time, categories, recurrence }) {
  const payload = {};
  if (title !== undefined)      payload.title = title;
  if (status !== undefined)     payload.status = status;
  if (importance !== undefined) payload.importance = importance;
  if (due_date !== undefined)   payload.dueDateTime   = due_date   ? { dateTime: `${due_date}T00:00:00.000Z`,   timeZone: "UTC" } : null;
  if (start_date !== undefined) payload.startDateTime = start_date ? { dateTime: `${start_date}T00:00:00.000Z`, timeZone: "UTC" } : null;
  if (body !== undefined)       payload.body = { contentType: "text", content: body };

  if (reminder_date_time !== undefined) {
    if (reminder_date_time) {
      payload.isReminderOn = true;
      payload.reminderDateTime = { dateTime: reminder_date_time, timeZone: "UTC" };
    } else {
      payload.isReminderOn = false;
    }
  }

  if (categories !== undefined) payload.categories = categories;
  if (recurrence !== undefined) payload.recurrence = recurrence;

  return payload;
}

function formatTask(task) {
  return {
    id:               task.id,
    title:            task.title,
    status:           task.status,
    importance:       task.importance,
    dueDate:          toDateOnly(task.dueDateTime?.dateTime),
    startDate:        toDateOnly(task.startDateTime?.dateTime),
    reminderOn:       task.isReminderOn ?? false,
    reminderDateTime: task.reminderDateTime?.dateTime ?? null,
    categories:       task.categories ?? [],
    recurrence:       task.recurrence ?? null,
    body:             task.body?.content || null,
  };
}

function buildRecurrence({ pattern, range }) {
  // Build pattern — include all fields the Graph API expects
  const p = {
    type:           pattern.type,
    interval:       pattern.interval ?? 1,
    month:          pattern.month ?? 0,
    dayOfMonth:     pattern.day_of_month ?? 0,
    daysOfWeek:     pattern.days_of_week ?? [],
    firstDayOfWeek: "sunday",
    index:          "first",
  };

  // Build range — recurrenceTimeZone is required by the API
  const r = {
    type:                 range.type,
    startDate:            range.start_date,
    endDate:              range.end_date ?? "0001-01-01",
    recurrenceTimeZone:   "UTC",
    numberOfOccurrences:  range.number_of_occurrences ?? 0,
  };

  return { pattern: p, range: r };
}

// ── Server ────────────────────────────────────────────────────────────────────

const server = new McpServer({ name: "todo-mcp", version: "1.0.0" });

// ── Account management ───────────────────────────────────────────────────────

server.tool("list_accounts", "List stored Microsoft To Do accounts", {}, async () => {
  return { content: [{ type: "text", text: JSON.stringify(listStoredAccounts(), null, 2) }] };
});

server.tool("authenticate_account", "Authenticate and store a Microsoft To Do account", {
  account:     z.string().min(1).describe("Alias for this account, e.g. 'work' or 'personal'"),
  client_id:   z.string().optional().describe("Optional Azure app client ID override"),
  tenant_id:   z.string().optional().describe("Optional tenant ID override. Use 'consumers' for personal Microsoft accounts, your tenant UUID for work accounts."),
  scope:       z.string().optional().describe("Optional OAuth scope override"),
  set_default: zBool.optional().default(false),
}, async ({ account, client_id, tenant_id, scope, set_default }) => {
  const state = await startDeviceCode({ account, client_id, tenant_id, scope, set_default });
  pollDeviceCode(state).catch((err) => {
    process.stderr.write(`Background auth polling failed for "${state.accountName}": ${err.message}\n`);
  });
  const { dcResp, accountName } = state;
  return {
    content: [{
      type: "text",
      text: JSON.stringify({
        status: "pending",
        account: accountName,
        message: dcResp.message,
        verification_url: dcResp.verification_uri,
        user_code: dcResp.user_code,
        expires_in_seconds: dcResp.expires_in,
        instructions: `Go to ${dcResp.verification_uri} and enter code: ${dcResp.user_code}. Authentication completes automatically in the background.`,
      }, null, 2),
    }],
  };
});

server.tool("set_default_account", "Set the default stored Microsoft To Do account", {
  account: z.string().min(1),
}, async ({ account }) => {
  return { content: [{ type: "text", text: JSON.stringify(setDefaultStoredAccount(account), null, 2) }] };
});

server.tool("delete_account", "Delete a stored Microsoft To Do account", {
  account: z.string().min(1),
}, async ({ account }) => {
  return { content: [{ type: "text", text: JSON.stringify(deleteStoredAccount(account), null, 2) }] };
});

// ── Task lists ───────────────────────────────────────────────────────────────

server.tool("list_todo_lists", "List all Microsoft To Do task lists", {
  ...optionalAccountArg,
}, async ({ account }) => {
  const resp = await graphGet("/me/todo/lists", account);
  const lists = (resp.value || []).map((l) => ({
    id: l.id, name: l.displayName, isOwner: l.isOwner, isShared: l.isShared,
  }));
  return { content: [{ type: "text", text: JSON.stringify(lists, null, 2) }] };
});

// ── Tasks ────────────────────────────────────────────────────────────────────

server.tool("get_tasks", "Get tasks from a specific To Do list", {
  ...optionalAccountArg,
  list_id: z.string(),
  status:  z.enum(["all", "notStarted", "inProgress", "waitingOnOthers", "deferred", "completed"]).optional().default("all"),
  top:     z.number().optional().default(50),
}, async ({ account, list_id, status, top }) => {
  const query = new URLSearchParams({ "$top": String(top) });
  if (status !== "all") query.set("$filter", `status eq '${status}'`);
  const resp = await graphGet(`${buildTasksPath(list_id)}?${query}`, account);
  return { content: [{ type: "text", text: JSON.stringify((resp.value || []).map(formatTask), null, 2) }] };
});

server.tool("get_all_pending_tasks", "Get all pending (not completed) tasks across all To Do lists", {
  ...optionalAccountArg,
  include_in_progress: zBool.optional().default(true),
}, async ({ account, include_in_progress }) => {
  const listsResp = await graphGet("/me/todo/lists", account);
  const lists = listsResp.value || [];
  const result = [];
  const filter = include_in_progress ? `status ne 'completed'` : `status eq 'notStarted'`;

  await Promise.all(lists.map(async (list) => {
    const query = new URLSearchParams({ "$filter": filter, "$top": "100" });
    const tasks = await graphGetAllPages(`${buildTasksPath(list.id)}?${query}`, account);
    for (const task of tasks) {
      result.push({ list: list.displayName, list_id: list.id, ...formatTask(task) });
    }
  }));

  result.sort((a, b) => {
    if (a.importance === "high" && b.importance !== "high") return -1;
    if (b.importance === "high" && a.importance !== "high") return 1;
    if (a.dueDate && b.dueDate) return a.dueDate.localeCompare(b.dueDate);
    if (a.dueDate) return -1;
    if (b.dueDate) return 1;
    return 0;
  });

  return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
});

server.tool("create_task", "Create a new task in a To Do list. Note: recurrence can only be set at creation time, not updated later.", {
  ...optionalAccountArg,
  list_id:            z.string(),
  title:              z.string(),
  status:             z.enum(["notStarted", "inProgress", "waitingOnOthers", "deferred"]).optional(),
  due_date:           z.string().optional().describe("YYYY-MM-DD"),
  start_date:         z.string().optional().describe("YYYY-MM-DD"),
  importance:         z.enum(["low", "normal", "high"]).optional().default("normal"),
  body:               z.string().optional().describe("Plain text note/description"),
  reminder_date_time: z.string().optional().describe("ISO 8601 datetime for reminder, e.g. 2026-04-01T09:00:00"),
  categories:         zStringArray.optional().describe("Category/tag strings, e.g. [\"work\", \"urgent\"]"),
  recurrence:         recurrenceSchema.optional(),
}, async ({ account, list_id, title, status, due_date, start_date, importance, body, reminder_date_time, categories, recurrence }) => {
  const payload = buildTaskPayload({
    title, status, due_date, start_date, importance, body, reminder_date_time, categories,
    recurrence: recurrence ? buildRecurrence(recurrence) : undefined,
  });
  const resp = await graphPost(buildTasksPath(list_id), payload, account);
  return { content: [{ type: "text", text: JSON.stringify(formatTask(resp), null, 2) }] };
});

server.tool("update_task", "Update a task's fields. Note: recurrence cannot be changed after task creation (Microsoft Graph API limitation).", {
  ...optionalAccountArg,
  list_id:            z.string(),
  task_id:            z.string(),
  title:              z.string().optional(),
  status:             z.enum(["notStarted", "inProgress", "waitingOnOthers", "deferred"]).optional(),
  due_date:           z.string().optional().describe("YYYY-MM-DD, or empty string to clear"),
  start_date:         z.string().optional().describe("YYYY-MM-DD, or empty string to clear"),
  importance:         z.enum(["low", "normal", "high"]).optional(),
  body:               z.string().optional(),
  reminder_date_time: z.string().optional().describe("ISO 8601 datetime, or empty string to turn off reminder"),
  categories:         zStringArray.optional().describe("Category/tag strings, e.g. [\"work\", \"urgent\"]"),
}, async ({ account, list_id, task_id, title, status, due_date, start_date, importance, body, reminder_date_time, categories }) => {
  const payload = buildTaskPayload({
    title, status, due_date, start_date, importance, body, reminder_date_time, categories,
  });
  const resp = await graphPatch(buildTaskPath(list_id, task_id), payload, account);
  return { content: [{ type: "text", text: JSON.stringify(formatTask(resp), null, 2) }] };
});

server.tool("complete_task", "Mark a task as completed", {
  ...optionalAccountArg,
  list_id: z.string(),
  task_id: z.string(),
}, async ({ account, list_id, task_id }) => {
  const resp = await graphPatch(buildTaskPath(list_id, task_id), {
    status: "completed",
    completedDateTime: { dateTime: new Date().toISOString(), timeZone: "UTC" },
  }, account);
  return { content: [{ type: "text", text: `Completed: "${resp.title}"` }] };
});

server.tool("delete_task", "Delete a task from a To Do list", {
  ...optionalAccountArg,
  list_id: z.string(),
  task_id: z.string(),
}, async ({ account, list_id, task_id }) => {
  await graphDel(buildTaskPath(list_id, task_id), account);
  return { content: [{ type: "text", text: "Task deleted." }] };
});

// ── Checklist items (subtasks) ───────────────────────────────────────────────

server.tool("get_checklist_items", "Get subtask checklist items for a task", {
  ...optionalAccountArg,
  list_id: z.string(),
  task_id: z.string(),
}, async ({ account, list_id, task_id }) => {
  const resp = await graphGet(buildChecklistPath(list_id, task_id), account);
  const items = (resp.value || []).map((i) => ({
    id: i.id, title: i.displayName, isChecked: i.isChecked,
  }));
  return { content: [{ type: "text", text: JSON.stringify(items, null, 2) }] };
});

server.tool("add_checklist_item", "Add a subtask checklist item to a task", {
  ...optionalAccountArg,
  list_id:    z.string(),
  task_id:    z.string(),
  title:      z.string().describe("Checklist item text"),
  is_checked: zBool.optional().default(false),
}, async ({ account, list_id, task_id, title, is_checked }) => {
  const resp = await graphPost(buildChecklistPath(list_id, task_id), {
    displayName: title, isChecked: is_checked,
  }, account);
  return { content: [{ type: "text", text: JSON.stringify({ id: resp.id, title: resp.displayName, isChecked: resp.isChecked }, null, 2) }] };
});

server.tool("update_checklist_item", "Update a checklist item's text or checked state", {
  ...optionalAccountArg,
  list_id:    z.string(),
  task_id:    z.string(),
  item_id:    z.string(),
  title:      z.string().optional(),
  is_checked: zBool.optional(),
}, async ({ account, list_id, task_id, item_id, title, is_checked }) => {
  const body = {};
  if (title      !== undefined) body.displayName = title;
  if (is_checked !== undefined) body.isChecked   = is_checked;
  const resp = await graphPatch(buildChecklistItemPath(list_id, task_id, item_id), body, account);
  return { content: [{ type: "text", text: JSON.stringify({ id: resp.id, title: resp.displayName, isChecked: resp.isChecked }, null, 2) }] };
});

server.tool("delete_checklist_item", "Delete a checklist item from a task", {
  ...optionalAccountArg,
  list_id: z.string(),
  task_id: z.string(),
  item_id: z.string(),
}, async ({ account, list_id, task_id, item_id }) => {
  await graphDel(buildChecklistItemPath(list_id, task_id, item_id), account);
  return { content: [{ type: "text", text: "Checklist item deleted." }] };
});

// ── Start ─────────────────────────────────────────────────────────────────────

const transport = new StdioServerTransport();
await server.connect(transport);
process.stderr.write("todo-mcp server running\n");
