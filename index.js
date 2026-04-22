import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { randomUUID } from "crypto";
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

function buildPlannerPlansPath() {
  return "/me/planner/plans";
}

function buildPlannerPlansByGroupPath(groupId) {
  return `/groups/${encodeURIComponent(groupId)}/planner/plans`;
}

function buildPlannerPlanPath(planId) {
  return `/planner/plans/${encodeURIComponent(planId)}`;
}

function buildPlannerPlanBucketsPath(planId) {
  return `${buildPlannerPlanPath(planId)}/buckets`;
}

function buildPlannerPlanTasksPath(planId) {
  return `${buildPlannerPlanPath(planId)}/tasks`;
}

function buildPlannerTaskPath(taskId) {
  return `/planner/tasks/${encodeURIComponent(taskId)}`;
}

function buildPlannerTaskDetailsPath(taskId) {
  return `${buildPlannerTaskPath(taskId)}/details`;
}

function toDateOnly(dateTime) {
  return dateTime?.substring(0, 10) || null;
}

async function graphRequest(method, pathOrUrl, body, account, extraHeaders = {}) {
  const token = await getAccessToken(account);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      ...extraHeaders,
    };

    if (body !== undefined && headers["Content-Type"] === undefined) {
      headers["Content-Type"] = "application/json";
    }

    const resp = await fetch(toGraphUrl(pathOrUrl), {
      method,
      headers,
      body: body !== undefined ? JSON.stringify(body) : undefined,
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
const graphPatchWithHeaders = (path, body, account, headers) => graphRequest("PATCH", path, body, account, headers);
const graphDelWithHeaders = (path, account, headers) => graphRequest("DELETE", path, undefined, account, headers);

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

function normalizeStringArray(values, fieldName) {
  if (!Array.isArray(values)) {
    throw new Error(`${fieldName} must be an array of strings.`);
  }

  const normalized = [];
  for (const value of values) {
    if (typeof value !== "string" || !value.trim()) {
      throw new Error(`${fieldName} must contain non-empty strings.`);
    }
    normalized.push(value.trim());
  }

  return [...new Set(normalized)];
}

const plannerCategoryPattern = /^category([1-9]|1\d|2[0-5])$/;

function normalizePlannerCategories(categories) {
  const normalized = normalizeStringArray(categories, "categories");
  for (const category of normalized) {
    if (!plannerCategoryPattern.test(category)) {
      throw new Error(`Invalid Planner category "${category}". Use category1 through category25.`);
    }
  }
  return normalized;
}

function buildPlannerAppliedCategories(categories, clearUnset = false) {
  const appliedCategories = {};

  if (clearUnset) {
    for (let index = 1; index <= 25; index += 1) {
      appliedCategories[`category${index}`] = false;
    }
  }

  for (const category of categories) {
    appliedCategories[category] = true;
  }

  return appliedCategories;
}

function buildPlannerAssignments(userIds) {
  const assignments = {};
  for (const userId of userIds) {
    assignments[userId] = {
      "@odata.type": "microsoft.graph.plannerAssignment",
      orderHint: " !",
    };
  }
  return assignments;
}

function buildPlannerAssignmentsPatch(existingAssignments, nextUserIds) {
  const patch = {};
  const existingUserIds = Object.keys(existingAssignments || {});
  const nextUserIdSet = new Set(nextUserIds);

  for (const userId of existingUserIds) {
    if (!nextUserIdSet.has(userId)) patch[userId] = null;
  }

  for (const userId of nextUserIdSet) {
    patch[userId] = {
      "@odata.type": "microsoft.graph.plannerAssignment",
      orderHint: " !",
    };
  }

  return patch;
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

function formatPlannerPlan(plan) {
  return {
    id: plan.id,
    title: plan.title,
    ownerGroupId: plan.owner ?? null,
    createdDateTime: plan.createdDateTime ?? null,
  };
}

function formatPlannerBucket(bucket) {
  return {
    id: bucket.id,
    name: bucket.name,
    planId: bucket.planId,
    orderHint: bucket.orderHint ?? null,
  };
}

function comparePlannerOrderHint(leftOrderHint, rightOrderHint) {
  const leftHasOrderHint = typeof leftOrderHint === "string" && leftOrderHint.length > 0;
  const rightHasOrderHint = typeof rightOrderHint === "string" && rightOrderHint.length > 0;

  if (leftHasOrderHint && rightHasOrderHint) {
    if (leftOrderHint > rightOrderHint) return -1;
    if (leftOrderHint < rightOrderHint) return 1;
    return 0;
  }

  if (leftHasOrderHint) return -1;
  if (rightHasOrderHint) return 1;
  return 0;
}

function sortPlannerBucketsLeftToRight(buckets) {
  return [...buckets].sort((left, right) => {
    const orderHintComparison = comparePlannerOrderHint(left.orderHint, right.orderHint);
    if (orderHintComparison !== 0) return orderHintComparison;

    const nameComparison = left.name.localeCompare(right.name);
    if (nameComparison !== 0) return nameComparison;

    return left.id.localeCompare(right.id);
  });
}

async function resolvePlannerBucketId(planId, requestedBucketId, account) {
  if (requestedBucketId !== undefined) return requestedBucketId;

  const buckets = await graphGetAllPages(buildPlannerPlanBucketsPath(planId), account);
  const orderedBuckets = sortPlannerBucketsLeftToRight(buckets.map(formatPlannerBucket));
  const firstBucket = orderedBuckets[0];

  return firstBucket?.id;
}

function formatPlannerTask(task) {
  return {
    id: task.id,
    title: task.title,
    planId: task.planId,
    bucketId: task.bucketId,
    percentComplete: task.percentComplete ?? 0,
    priority: task.priority ?? null,
    startDateTime: task.startDateTime ?? null,
    dueDateTime: task.dueDateTime ?? null,
    completedDateTime: task.completedDateTime ?? null,
    hasDescription: task.hasDescription ?? false,
    assigneeUserIds: Object.keys(task.assignments || {}),
    categories: Object.entries(task.appliedCategories || {})
      .filter(([, enabled]) => Boolean(enabled))
      .map(([category]) => category)
      .sort((left, right) => left.localeCompare(right)),
    etag: task["@odata.etag"] ?? null,
  };
}

function formatDirectoryUser(user) {
  const displayName = typeof user.displayName === "string" && user.displayName.trim()
    ? user.displayName.trim()
    : [user.givenName, user.surname]
      .filter((part) => typeof part === "string" && part.trim())
      .join(" ")
      .trim();

  return {
    id: user.id,
    name: displayName || user.mail || user.userPrincipalName || user.id,
  };
}

function decodeReferenceKey(encodedKey) {
  try {
    return decodeURIComponent(encodedKey);
  } catch {
    return encodedKey;
  }
}

function formatPlannerTaskDetails(details) {
  const checklist = Object.entries(details.checklist || {}).map(([id, item]) => ({
    id,
    title: item?.title ?? null,
    isChecked: item?.isChecked ?? false,
    orderHint: item?.orderHint ?? null,
    lastModifiedDateTime: item?.lastModifiedDateTime ?? null,
  }));

  const references = Object.entries(details.references || {}).map(([encodedKey, reference]) => ({
    url: decodeReferenceKey(encodedKey),
    alias: reference?.alias ?? null,
    previewPriority: reference?.previewPriority ?? null,
    type: reference?.type ?? null,
  }));

  return {
    id: details.id,
    description: details.description ?? "",
    previewType: details.previewType ?? null,
    checklist,
    references,
    etag: details["@odata.etag"] ?? null,
  };
}

function getGraphEtag(entity, context) {
  const etag = entity?.["@odata.etag"];
  if (typeof etag !== "string" || !etag) {
    throw new Error(`Microsoft Graph did not return an ETag for ${context}.`);
  }
  return etag;
}

async function patchPlannerTask(taskId, patchBody, account) {
  const path = buildPlannerTaskPath(taskId);
  const currentTask = await graphGet(path, account);
  const etag = getGraphEtag(currentTask, `planner task "${taskId}"`);
  await graphPatchWithHeaders(path, patchBody, account, { "If-Match": etag });
  return graphGet(path, account);
}

async function patchPlannerTaskDetails(taskId, patchBody, account) {
  const path = buildPlannerTaskDetailsPath(taskId);
  const currentDetails = await graphGet(path, account);
  const etag = getGraphEtag(currentDetails, `planner task details "${taskId}"`);
  await graphPatchWithHeaders(path, patchBody, account, { "If-Match": etag });
  return graphGet(path, account);
}

async function deletePlannerTask(taskId, account) {
  const path = buildPlannerTaskPath(taskId);
  const currentTask = await graphGet(path, account);
  const etag = getGraphEtag(currentTask, `planner task "${taskId}"`);
  await graphDelWithHeaders(path, account, { "If-Match": etag });
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

// ── Microsoft Planner ────────────────────────────────────────────────────────

server.tool("list_planner_plans", "List Microsoft Planner plans available to the signed-in user", {
  ...optionalAccountArg,
  owner_group_id: z.string().optional().describe("Optional Microsoft 365 group ID to list plans for a specific group."),
}, async ({ account, owner_group_id }) => {
  const path = owner_group_id ? buildPlannerPlansByGroupPath(owner_group_id) : buildPlannerPlansPath();
  const plans = await graphGetAllPages(path, account);
  return { content: [{ type: "text", text: JSON.stringify(plans.map(formatPlannerPlan), null, 2) }] };
});

server.tool("list_planner_buckets", "List buckets in a Microsoft Planner plan", {
  ...optionalAccountArg,
  plan_id: z.string(),
}, async ({ account, plan_id }) => {
  const buckets = await graphGetAllPages(buildPlannerPlanBucketsPath(plan_id), account);
  const sortedBuckets = sortPlannerBucketsLeftToRight(buckets.map(formatPlannerBucket));
  return { content: [{ type: "text", text: JSON.stringify(sortedBuckets, null, 2) }] };
});

server.tool("list_employee_ids", "List employee IDs with names for Planner assignment and lookup", {
  ...optionalAccountArg,
  search: z.string().optional().describe("Optional text filter applied to display name, given name, surname, mail, and UPN."),
  max_results: z.number().int().positive().max(5000).optional().default(500),
}, async ({ account, search, max_results }) => {
  const users = await graphGetAllPages("/users?$select=id,displayName,givenName,surname,mail,userPrincipalName&$top=999", account);
  const normalizedSearch = typeof search === "string" ? search.trim().toLocaleLowerCase() : "";

  const filteredUsers = users
    .filter((user) => {
      if (!normalizedSearch) return true;
      const lookupValues = [user.displayName, user.givenName, user.surname, user.mail, user.userPrincipalName]
        .filter((value) => typeof value === "string" && value.trim());
      return lookupValues.some((value) => value.toLocaleLowerCase().includes(normalizedSearch));
    })
    .map(formatDirectoryUser)
    .sort((left, right) => {
      const nameComparison = left.name.localeCompare(right.name);
      if (nameComparison !== 0) return nameComparison;
      return left.id.localeCompare(right.id);
    })
    .slice(0, max_results);

  return { content: [{ type: "text", text: JSON.stringify(filteredUsers, null, 2) }] };
});

server.tool("get_user_planner_tasks", "Get pending Planner tasks assigned to a specific user from plans visible to the signed-in user", {
  ...optionalAccountArg,
  user_id: z.string().describe("Microsoft Entra user ID."),
  plan_id: z.string().optional().describe("Optional Planner plan ID. If the signed-in user cannot access this plan, no tasks are returned."),
  max_results: z.number().int().positive().max(5000).optional().default(500),
}, async ({ account, user_id, plan_id, max_results }) => {
  const accessiblePlans = await graphGetAllPages(buildPlannerPlansPath(), account);
  const selectedPlans = plan_id
    ? accessiblePlans.filter((plan) => plan.id === plan_id)
    : accessiblePlans;

  if (selectedPlans.length === 0) {
    return { content: [{ type: "text", text: JSON.stringify([], null, 2) }] };
  }

  const tasksByPlan = await Promise.all(selectedPlans.map(async (plan) => {
    try {
      return await graphGetAllPages(`${buildPlannerPlanTasksPath(plan.id)}?$top=100`, account);
    } catch {
      return [];
    }
  }));

  const normalizedUserId = user_id.trim().toLocaleLowerCase();
  const filtered = tasksByPlan
    .flat()
    .filter((task) => Number(task.percentComplete || 0) < 100)
    .filter((task) => Object.keys(task.assignments || {}).some((assigneeId) => assigneeId.toLocaleLowerCase() === normalizedUserId));

  const planTitles = new Map(selectedPlans.map((plan) => [plan.id, plan.title || null]));

  const result = filtered
    .map((task) => ({
      ...formatPlannerTask(task),
      planTitle: task.planId ? (planTitles.get(task.planId) ?? null) : null,
    }))
    .sort((left, right) => {
      const leftPriority = Number.isFinite(left.priority) ? left.priority : 99;
      const rightPriority = Number.isFinite(right.priority) ? right.priority : 99;
      if (leftPriority !== rightPriority) return leftPriority - rightPriority;

      if (left.dueDateTime && right.dueDateTime) return left.dueDateTime.localeCompare(right.dueDateTime);
      if (left.dueDateTime) return -1;
      if (right.dueDateTime) return 1;
      return left.title.localeCompare(right.title);
    })
    .slice(0, max_results);

  return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
});

server.tool("get_planner_tasks", "Get tasks from a Microsoft Planner plan", {
  ...optionalAccountArg,
  plan_id: z.string(),
  bucket_id: z.string().optional().describe("Optional bucket ID to limit tasks to one bucket."),
  status: z.enum(["all", "incomplete", "completed"]).optional().default("all"),
  top: z.number().int().positive().optional().default(100),
}, async ({ account, plan_id, bucket_id, status, top }) => {
  const query = new URLSearchParams({ "$top": String(top) });
  if (bucket_id) {
    query.set("$filter", `bucketId eq '${bucket_id}'`);
  }

  const resp = await graphGet(`${buildPlannerPlanTasksPath(plan_id)}?${query}`, account);
  let tasks = (resp.value || []).map(formatPlannerTask);

  if (status === "completed") {
    tasks = tasks.filter((task) => (task.percentComplete ?? 0) >= 100);
  }
  if (status === "incomplete") {
    tasks = tasks.filter((task) => (task.percentComplete ?? 0) < 100);
  }

  return { content: [{ type: "text", text: JSON.stringify(tasks, null, 2) }] };
});

server.tool("get_all_pending_planner_tasks", "Get all pending (not completed) Microsoft Planner tasks assigned to the signed-in user", {
  ...optionalAccountArg,
  include_in_progress: zBool.optional().default(true),
}, async ({ account, include_in_progress }) => {
  const tasks = await graphGetAllPages("/me/planner/tasks?$top=100", account);
  const filtered = tasks.filter((task) => {
    const percentComplete = Number(task.percentComplete || 0);
    if (include_in_progress) return percentComplete < 100;
    return percentComplete === 0;
  });

  const uniquePlanIds = [...new Set(filtered.map((task) => task.planId).filter(Boolean))];
  const planTitles = new Map();
  await Promise.all(uniquePlanIds.map(async (planId) => {
    try {
      const plan = await graphGet(buildPlannerPlanPath(planId), account);
      planTitles.set(planId, plan.title || null);
    } catch {
      planTitles.set(planId, null);
    }
  }));

  const result = filtered
    .map((task) => ({
      ...formatPlannerTask(task),
      planTitle: task.planId ? (planTitles.get(task.planId) ?? null) : null,
    }))
    .sort((left, right) => {
      const leftPriority = Number.isFinite(left.priority) ? left.priority : 99;
      const rightPriority = Number.isFinite(right.priority) ? right.priority : 99;
      if (leftPriority !== rightPriority) return leftPriority - rightPriority;

      if (left.dueDateTime && right.dueDateTime) return left.dueDateTime.localeCompare(right.dueDateTime);
      if (left.dueDateTime) return -1;
      if (right.dueDateTime) return 1;
      return left.title.localeCompare(right.title);
    });

  return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
});

server.tool("create_planner_task", "Create a Microsoft Planner task", {
  ...optionalAccountArg,
  plan_id: z.string(),
  title: z.string(),
  bucket_id: z.string().optional().describe("Optional bucket ID. If omitted, task is placed in the leftmost bucket in the plan when one exists; otherwise created without a bucket."),
  start_date_time: z.string().optional().describe("ISO 8601 datetime, e.g. 2026-04-16T09:00:00Z"),
  due_date_time: z.string().optional().describe("ISO 8601 datetime, e.g. 2026-04-20T17:00:00Z"),
  priority: z.number().int().min(0).max(10).optional().describe("Planner priority 0-10, where lower is higher priority"),
  assign_to_user_ids: zStringArray.optional().describe("AAD user IDs to assign, e.g. [\"<user-id>\"]"),
  categories: zStringArray.optional().describe("Planner category keys, e.g. [\"category1\", \"category3\"]"),
}, async ({ account, plan_id, title, bucket_id, start_date_time, due_date_time, priority, assign_to_user_ids, categories }) => {
  const payload = { planId: plan_id, title };
  const resolvedBucketId = await resolvePlannerBucketId(plan_id, bucket_id, account);
  if (resolvedBucketId !== undefined) payload.bucketId = resolvedBucketId;

  if (start_date_time !== undefined) payload.startDateTime = start_date_time;
  if (due_date_time !== undefined) payload.dueDateTime = due_date_time;
  if (priority !== undefined) payload.priority = priority;

  if (assign_to_user_ids !== undefined) {
    const assignees = normalizeStringArray(assign_to_user_ids, "assign_to_user_ids");
    if (assignees.length > 0) payload.assignments = buildPlannerAssignments(assignees);
  }

  if (categories !== undefined) {
    const plannerCategories = normalizePlannerCategories(categories);
    payload.appliedCategories = buildPlannerAppliedCategories(plannerCategories);
  }

  const createdTask = await graphPost("/planner/tasks", payload, account);
  return { content: [{ type: "text", text: JSON.stringify(formatPlannerTask(createdTask), null, 2) }] };
});

server.tool("update_planner_task", "Update a Microsoft Planner task", {
  ...optionalAccountArg,
  task_id: z.string(),
  title: z.string().optional(),
  bucket_id: z.string().optional(),
  start_date_time: z.string().optional().describe("ISO 8601 datetime, or empty string to clear"),
  due_date_time: z.string().optional().describe("ISO 8601 datetime, or empty string to clear"),
  priority: z.number().int().min(0).max(10).optional(),
  percent_complete: z.number().int().min(0).max(100).optional(),
  set_assigned_user_ids: zStringArray.optional().describe("Replace assignees with this user ID list"),
  categories: zStringArray.optional().describe("Set Planner categories, e.g. [\"category1\"]. Use [] to clear all."),
}, async ({ account, task_id, title, bucket_id, start_date_time, due_date_time, priority, percent_complete, set_assigned_user_ids, categories }) => {
  const currentTask = await graphGet(buildPlannerTaskPath(task_id), account);
  const etag = getGraphEtag(currentTask, `planner task "${task_id}"`);

  const payload = {};
  if (title !== undefined) payload.title = title;
  if (bucket_id !== undefined) payload.bucketId = bucket_id;
  if (start_date_time !== undefined) payload.startDateTime = start_date_time ? start_date_time : null;
  if (due_date_time !== undefined) payload.dueDateTime = due_date_time ? due_date_time : null;
  if (priority !== undefined) payload.priority = priority;
  if (percent_complete !== undefined) payload.percentComplete = percent_complete;

  if (set_assigned_user_ids !== undefined) {
    const nextAssignees = normalizeStringArray(set_assigned_user_ids, "set_assigned_user_ids");
    const assignmentPatch = buildPlannerAssignmentsPatch(currentTask.assignments || {}, nextAssignees);
    if (Object.keys(assignmentPatch).length > 0) payload.assignments = assignmentPatch;
  }

  if (categories !== undefined) {
    const plannerCategories = normalizePlannerCategories(categories);
    payload.appliedCategories = buildPlannerAppliedCategories(plannerCategories, true);
  }

  if (Object.keys(payload).length === 0) {
    throw new Error("No fields provided to update_planner_task.");
  }

  await graphPatchWithHeaders(buildPlannerTaskPath(task_id), payload, account, { "If-Match": etag });
  const updatedTask = await graphGet(buildPlannerTaskPath(task_id), account);
  return { content: [{ type: "text", text: JSON.stringify(formatPlannerTask(updatedTask), null, 2) }] };
});

server.tool("complete_planner_task", "Mark a Microsoft Planner task as completed", {
  ...optionalAccountArg,
  task_id: z.string(),
}, async ({ account, task_id }) => {
  const updatedTask = await patchPlannerTask(task_id, { percentComplete: 100 }, account);
  return { content: [{ type: "text", text: `Completed: "${updatedTask.title}"` }] };
});

server.tool("delete_planner_task", "Delete a Microsoft Planner task", {
  ...optionalAccountArg,
  task_id: z.string(),
}, async ({ account, task_id }) => {
  await deletePlannerTask(task_id, account);
  return { content: [{ type: "text", text: "Planner task deleted." }] };
});

server.tool("get_planner_task_details", "Get Microsoft Planner task details (description, checklist, references)", {
  ...optionalAccountArg,
  task_id: z.string(),
}, async ({ account, task_id }) => {
  const details = await graphGet(buildPlannerTaskDetailsPath(task_id), account);
  return { content: [{ type: "text", text: JSON.stringify(formatPlannerTaskDetails(details), null, 2) }] };
});

server.tool("update_planner_task_details", "Update Microsoft Planner task details", {
  ...optionalAccountArg,
  task_id: z.string(),
  description: z.string().optional().describe("Task description / notes"),
  preview_type: z.enum(["automatic", "noPreview", "checklist", "description", "reference"]).optional(),
}, async ({ account, task_id, description, preview_type }) => {
  const payload = {};
  if (description !== undefined) payload.description = description;
  if (preview_type !== undefined) payload.previewType = preview_type;

  if (Object.keys(payload).length === 0) {
    throw new Error("No fields provided to update_planner_task_details.");
  }

  const updatedDetails = await patchPlannerTaskDetails(task_id, payload, account);
  return { content: [{ type: "text", text: JSON.stringify(formatPlannerTaskDetails(updatedDetails), null, 2) }] };
});

server.tool("add_planner_checklist_item", "Add a checklist item to a Microsoft Planner task", {
  ...optionalAccountArg,
  task_id: z.string(),
  title: z.string(),
  is_checked: zBool.optional().default(false),
}, async ({ account, task_id, title, is_checked }) => {
  const checklistItemId = randomUUID();
  const updatedDetails = await patchPlannerTaskDetails(task_id, {
    checklist: {
      [checklistItemId]: {
        "@odata.type": "microsoft.graph.plannerChecklistItem",
        title,
        isChecked: is_checked,
        orderHint: " !",
      },
    },
  }, account);

  const newItem = updatedDetails.checklist?.[checklistItemId];
  return {
    content: [{
      type: "text",
      text: JSON.stringify({
        id: checklistItemId,
        title: newItem?.title ?? title,
        isChecked: newItem?.isChecked ?? is_checked,
      }, null, 2),
    }],
  };
});

server.tool("update_planner_checklist_item", "Update a checklist item in a Microsoft Planner task", {
  ...optionalAccountArg,
  task_id: z.string(),
  checklist_item_id: z.string(),
  title: z.string().optional(),
  is_checked: zBool.optional(),
}, async ({ account, task_id, checklist_item_id, title, is_checked }) => {
  const currentDetails = await graphGet(buildPlannerTaskDetailsPath(task_id), account);
  const existingChecklistItem = currentDetails.checklist?.[checklist_item_id];

  if (!existingChecklistItem) {
    throw new Error(`Checklist item "${checklist_item_id}" was not found on planner task "${task_id}".`);
  }

  const patchItem = {
    "@odata.type": "microsoft.graph.plannerChecklistItem",
    title: title !== undefined ? title : (existingChecklistItem.title || ""),
    isChecked: is_checked !== undefined ? is_checked : Boolean(existingChecklistItem.isChecked),
    orderHint: existingChecklistItem.orderHint || " !",
  };

  const updatedDetails = await patchPlannerTaskDetails(task_id, {
    checklist: {
      [checklist_item_id]: patchItem,
    },
  }, account);

  const updatedItem = updatedDetails.checklist?.[checklist_item_id];
  return {
    content: [{
      type: "text",
      text: JSON.stringify({
        id: checklist_item_id,
        title: updatedItem?.title ?? patchItem.title,
        isChecked: updatedItem?.isChecked ?? patchItem.isChecked,
      }, null, 2),
    }],
  };
});

server.tool("delete_planner_checklist_item", "Delete a checklist item from a Microsoft Planner task", {
  ...optionalAccountArg,
  task_id: z.string(),
  checklist_item_id: z.string(),
}, async ({ account, task_id, checklist_item_id }) => {
  const currentDetails = await graphGet(buildPlannerTaskDetailsPath(task_id), account);
  if (!currentDetails.checklist?.[checklist_item_id]) {
    throw new Error(`Checklist item "${checklist_item_id}" was not found on planner task "${task_id}".`);
  }

  await patchPlannerTaskDetails(task_id, {
    checklist: {
      [checklist_item_id]: null,
    },
  }, account);

  return { content: [{ type: "text", text: "Planner checklist item deleted." }] };
});

// ── Start ─────────────────────────────────────────────────────────────────────

const transport = new StdioServerTransport();
await server.connect(transport);
process.stderr.write("todo-mcp server running\n");
