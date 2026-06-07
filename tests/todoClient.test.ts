import test from "node:test";
import assert from "node:assert/strict";

import {
  GraphApiError,
  MicrosoftTodoClient,
  type CreateTodoTaskInput,
  type ShoppingTaskInput,
  type TodoTaskList
} from "../src/todoClient.ts";

type FetchCall = {
  input: string | URL | Request;
  init?: RequestInit;
};

type FetchHandler = (
  input: string | URL | Request,
  init?: RequestInit
) => Response | Promise<Response>;

const originalFetch = globalThis.fetch;

test.afterEach(() => {
  globalThis.fetch = originalFetch;
});

function installFetch(handler: FetchHandler): FetchCall[] {
  const calls: FetchCall[] = [];

  globalThis.fetch = (async (input, init) => {
    calls.push({ input, init });
    return handler(input, init);
  }) as typeof fetch;

  return calls;
}

function jsonResponse(body: unknown, init: ResponseInit = {}): Response {
  const headers = new Headers(init.headers);
  headers.set("content-type", "application/json");

  return new Response(JSON.stringify(body), {
    ...init,
    headers
  });
}

function inputToString(input: string | URL | Request): string {
  if (typeof input === "string") {
    return input;
  }

  if (input instanceof URL) {
    return input.toString();
  }

  return input.url;
}

function readJsonBody(init?: RequestInit): Record<string, unknown> {
  assert.equal(typeof init?.body, "string");
  return JSON.parse(init.body as string) as Record<string, unknown>;
}

test("request prefixes relative paths, adds auth and JSON headers, and parses JSON", async () => {
  const calls = installFetch(() =>
    jsonResponse({ id: "list-1", displayName: "Groceries" })
  );
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0/"
  });

  const result = await client.request<TodoTaskList>("/me/todo/lists", {
    method: "POST",
    body: JSON.stringify({ displayName: "Groceries" })
  });

  assert.deepEqual(result, { id: "list-1", displayName: "Groceries" });
  assert.equal(inputToString(calls[0].input), "https://graph.test/v1.0/me/todo/lists");
  assert.equal(calls[0].init?.method, "POST");

  const headers = new Headers(calls[0].init?.headers);
  assert.equal(headers.get("authorization"), "Bearer token-123");
  assert.equal(headers.get("content-type"), "application/json");
});

test("request throws GraphApiError with status, message, body, and request id", async () => {
  const errorBody = { error: { message: "Too many requests" } };

  installFetch(() =>
    jsonResponse(errorBody, {
      status: 429,
      statusText: "Too Many Requests",
      headers: { "request-id": "request-123" }
    })
  );
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  await assert.rejects(
    () => client.request("/me/todo/lists"),
    (error: unknown) => {
      assert.ok(error instanceof GraphApiError);
      assert.equal(error.status, 429);
      assert.equal(error.message, "Too many requests");
      assert.deepEqual(error.body, errorBody);
      assert.equal(error.requestId, "request-123");
      return true;
    }
  );
});

test("listTaskLists follows @odata.nextLink collection pages", async () => {
  const nextLink = "https://graph.test/v1.0/me/todo/lists?page=2";
  const calls = installFetch((input) => {
    const url = inputToString(input);

    if (url === "https://graph.test/v1.0/me/todo/lists") {
      return jsonResponse({
        value: [{ id: "list-1", displayName: "Groceries" }],
        "@odata.nextLink": nextLink
      });
    }

    assert.equal(url, nextLink);
    return jsonResponse({
      value: [{ id: "list-2", displayName: "Hardware" }]
    });
  });
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  const lists = await client.listTaskLists();

  assert.deepEqual(
    lists.map((list) => list.displayName),
    ["Groceries", "Hardware"]
  );
  assert.equal(calls.length, 2);
  assert.equal(inputToString(calls[1].input), nextLink);
});

test("findTaskListByName matches displayName case-insensitively", async () => {
  installFetch(() =>
    jsonResponse({
      value: [
        { id: "list-1", displayName: "Groceries" },
        { id: "list-2", displayName: "Family Meal Shopping - Aldi + Costco" }
      ]
    })
  );
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  const list = await client.findTaskListByName("family meal shopping - aldi + costco");

  assert.equal(list?.id, "list-2");
});

test("createTaskList posts the displayName", async () => {
  const calls = installFetch((input, init) => {
    assert.equal(inputToString(input), "https://graph.test/v1.0/me/todo/lists");
    assert.equal(init?.method, "POST");
    assert.deepEqual(readJsonBody(init), { displayName: "Errands" });
    return jsonResponse({ id: "list-1", displayName: "Errands" });
  });
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  const list = await client.createTaskList({ displayName: "Errands" });

  assert.equal(list.id, "list-1");
  assert.equal(calls.length, 1);
});

test("getOrCreateTaskList creates a list only when no matching list exists", async () => {
  const calls = installFetch((input, init) => {
    if (init?.method === "POST") {
      assert.deepEqual(readJsonBody(init), { displayName: "Meal Plan" });
      return jsonResponse({ id: "created-list", displayName: "Meal Plan" });
    }

    return jsonResponse({
      value: [{ id: "existing-list", displayName: "House" }]
    });
  });
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  const list = await client.getOrCreateTaskList("Meal Plan");

  assert.equal(list.id, "created-list");
  assert.equal(calls.length, 2);
});

test("listTasks returns all paged tasks in a task list", async () => {
  const nextLink = "https://graph.test/v1.0/me/todo/lists/list-1/tasks?page=2";
  installFetch((input) => {
    const url = inputToString(input);

    if (url === "https://graph.test/v1.0/me/todo/lists/list-1/tasks") {
      return jsonResponse({
        value: [{ id: "task-1", title: "Buy chicken", status: "notStarted" }],
        "@odata.nextLink": nextLink
      });
    }

    assert.equal(url, nextLink);
    return jsonResponse({
      value: [{ id: "task-2", title: "Buy foil", status: "completed" }]
    });
  });
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  const tasks = await client.listTasks("list-1");

  assert.deepEqual(
    tasks.map((task) => task.id),
    ["task-1", "task-2"]
  );
});

test("createTask posts a single task input", async () => {
  const taskInput: CreateTodoTaskInput = {
    title: "Buy peaches",
    importance: "normal",
    categories: ["shopping"]
  };
  installFetch((input, init) => {
    assert.equal(inputToString(input), "https://graph.test/v1.0/me/todo/lists/list-1/tasks");
    assert.equal(init?.method, "POST");
    assert.deepEqual(readJsonBody(init), taskInput);
    return jsonResponse({ id: "task-1", title: "Buy peaches", importance: "normal" });
  });
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  const task = await client.createTask("list-1", taskInput);

  assert.equal(task.id, "task-1");
});

test("createTasks creates tasks sequentially and returns created tasks", async () => {
  let activeRequests = 0;
  const postedTitles: string[] = [];

  installFetch(async (_input, init) => {
    activeRequests += 1;
    assert.equal(activeRequests, 1);

    const body = readJsonBody(init);
    postedTitles.push(body.title as string);
    await new Promise((resolve) => setTimeout(resolve, 5));
    activeRequests -= 1;

    return jsonResponse({
      id: `task-${postedTitles.length}`,
      title: body.title
    });
  });
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  const tasks = await client.createTasks("list-1", [
    { title: "First" },
    { title: "Second" }
  ]);

  assert.deepEqual(postedTitles, ["First", "Second"]);
  assert.deepEqual(
    tasks.map((task) => task.id),
    ["task-1", "task-2"]
  );
});

test("createShoppingTasks gets or creates the list and formats one task per item", async () => {
  const postedTasks: Record<string, unknown>[] = [];
  const shoppingItems: ShoppingTaskInput[] = [
    {
      store: "Costco",
      section: "Protein",
      item: "Chicken thighs",
      quantity: "1 large pack"
    },
    {
      store: "Aldi",
      section: "Produce",
      item: "Peaches",
      quantity: "4-5",
      notes: "Choose ripe fruit"
    }
  ];

  installFetch((input, init) => {
    const url = inputToString(input);

    if (url === "https://graph.test/v1.0/me/todo/lists" && init?.method !== "POST") {
      return jsonResponse({ value: [] });
    }

    if (url === "https://graph.test/v1.0/me/todo/lists" && init?.method === "POST") {
      return jsonResponse({
        id: "shopping-list",
        displayName: "Family Meal Shopping - Aldi + Costco"
      });
    }

    assert.equal(url, "https://graph.test/v1.0/me/todo/lists/shopping-list/tasks");
    const body = readJsonBody(init);
    postedTasks.push(body);
    return jsonResponse({
      id: `task-${postedTasks.length}`,
      title: body.title
    });
  });
  const client = new MicrosoftTodoClient("token-123", {
    baseUrl: "https://graph.test/v1.0"
  });

  const tasks = await client.createShoppingTasks(
    "Family Meal Shopping - Aldi + Costco",
    shoppingItems
  );

  assert.deepEqual(
    postedTasks.map((task) => task.title),
    ["[Costco] Chicken thighs - 1 large pack", "[Aldi] Peaches - 4-5"]
  );
  const firstTaskBody = postedTasks[0].body as { content: string };
  const secondTaskBody = postedTasks[1].body as { content: string };

  assert.match(firstTaskBody.content, /Store: Costco/);
  assert.match(firstTaskBody.content, /Section: Protein/);
  assert.match(firstTaskBody.content, /Quantity: 1 large pack/);
  assert.match(secondTaskBody.content, /Notes: Choose ripe fruit/);
  assert.deepEqual(
    tasks.map((task) => task.id),
    ["task-1", "task-2"]
  );
});
