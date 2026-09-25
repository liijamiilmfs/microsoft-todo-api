export interface DateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

export interface TodoTaskBody {
  content?: string;
  contentType?: "text" | "html" | string;
}

export interface TodoTaskList {
  id: string;
  displayName: string;
  isOwner?: boolean;
  isShared?: boolean;
  wellknownListName?: string;
}

export interface TodoTask {
  id: string;
  title: string;
  status?: string;
  importance?: "low" | "normal" | "high" | string;
  body?: TodoTaskBody;
  dueDateTime?: DateTimeTimeZone;
  reminderDateTime?: DateTimeTimeZone;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

export interface CreateTodoTaskInput {
  title: string;
  body?: {
    content: string;
    contentType: "text" | "html";
  };
  importance?: "low" | "normal" | "high";
  dueDateTime?: DateTimeTimeZone;
  reminderDateTime?: DateTimeTimeZone;
  isReminderOn?: boolean;
  categories?: string[];
}

export interface CreateTodoListInput {
  displayName: string;
}

export interface ShoppingTaskInput {
  store: string;
  section: string;
  item: string;
  quantity?: string;
  notes?: string;
}

interface GraphCollectionResponse<T> {
  value?: T[];
  "@odata.nextLink"?: string;
}

export class GraphApiError extends Error {
  status: number;
  body: unknown;
  requestId?: string;

  constructor(status: number, message: string, body: unknown, requestId?: string) {
    super(message);
    this.name = "GraphApiError";
    this.status = status;
    this.body = body;
    this.requestId = requestId;
  }
}

export class MicrosoftTodoClient {
  private accessToken: string;
  private baseUrl: string;

  constructor(accessToken: string, options: { baseUrl?: string } = {}) {
    this.accessToken = accessToken;
    this.baseUrl = trimTrailingSlashes(
      options.baseUrl ?? "https://graph.microsoft.com/v1.0"
    );
  }

  async request<T>(pathOrUrl: string, init: RequestInit = {}): Promise<T> {
    const url = this.resolveUrl(pathOrUrl);
    const headers = new Headers(init.headers);

    headers.set("Authorization", `Bearer ${this.accessToken}`);

    if (init.body !== undefined && init.body !== null && !headers.has("Content-Type")) {
      headers.set("Content-Type", "application/json");
    }

    const response = await fetch(url, {
      ...init,
      headers
    });
    const body = await parseResponseBody(response);

    if (!response.ok) {
      const requestId = getRequestId(response, body);
      const message = getErrorMessage(response, body);

      throw new GraphApiError(response.status, message, body, requestId);
    }

    return body as T;
  }

  async getAllPages<T>(path: string): Promise<T[]> {
    const items: T[] = [];
    let nextPage: string | undefined = path;

    while (nextPage) {
      const page: GraphCollectionResponse<T> =
        await this.request<GraphCollectionResponse<T>>(nextPage);

      if (!Array.isArray(page.value)) {
        throw new Error("Expected Microsoft Graph collection response with a value array.");
      }

      items.push(...page.value);
      nextPage = page["@odata.nextLink"];
    }

    return items;
  }

  listTaskLists(): Promise<TodoTaskList[]> {
    return this.getAllPages<TodoTaskList>("/me/todo/lists");
  }

  async findTaskListByName(displayName: string): Promise<TodoTaskList | undefined> {
    const targetName = displayName.toLocaleLowerCase();
    const lists = await this.listTaskLists();

    return lists.find((list) => list.displayName.toLocaleLowerCase() === targetName);
  }

  createTaskList(input: CreateTodoListInput): Promise<TodoTaskList> {
    return this.request<TodoTaskList>("/me/todo/lists", {
      method: "POST",
      body: JSON.stringify(input)
    });
  }

  async getOrCreateTaskList(displayName: string): Promise<TodoTaskList> {
    const existingList = await this.findTaskListByName(displayName);

    if (existingList) {
      return existingList;
    }

    return this.createTaskList({ displayName });
  }

  listTasks(listId: string): Promise<TodoTask[]> {
    return this.getAllPages<TodoTask>(
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks`
    );
  }

  createTask(listId: string, input: CreateTodoTaskInput): Promise<TodoTask> {
    return this.request<TodoTask>(
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks`,
      {
        method: "POST",
        body: JSON.stringify(input)
      }
    );
  }

  async createTasks(
    listId: string,
    tasks: CreateTodoTaskInput[]
  ): Promise<TodoTask[]> {
    const createdTasks: TodoTask[] = [];

    for (const task of tasks) {
      createdTasks.push(await this.createTask(listId, task));
    }

    return createdTasks;
  }

  async createShoppingTasks(
    listName: string,
    shoppingItems: ShoppingTaskInput[]
  ): Promise<TodoTask[]> {
    const list = await this.getOrCreateTaskList(listName);
    const tasks = shoppingItems.map(toShoppingTaskInput);

    return this.createTasks(list.id, tasks);
  }

  private resolveUrl(pathOrUrl: string): string {
    if (pathOrUrl.startsWith("http")) {
      return pathOrUrl;
    }

    const path = pathOrUrl.startsWith("/") ? pathOrUrl : `/${pathOrUrl}`;
    return `${this.baseUrl}${path}`;
  }
}

function toShoppingTaskInput(input: ShoppingTaskInput): CreateTodoTaskInput {
  const quantitySuffix = input.quantity ? ` - ${input.quantity}` : "";
  const lines = [
    `Store: ${input.store}`,
    `Section: ${input.section}`,
    `Item: ${input.item}`,
    `Quantity: ${input.quantity ?? "Not specified"}`
  ];

  if (input.notes) {
    lines.push(`Notes: ${input.notes}`);
  }

  return {
    title: `[${input.store}] ${input.item}${quantitySuffix}`,
    body: {
      content: lines.join("\n"),
      contentType: "text"
    }
  };
}

async function parseResponseBody(response: Response): Promise<unknown> {
  const text = await response.text();

  if (!text) {
    return undefined;
  }

  const contentType = response.headers.get("content-type") ?? "";

  if (contentType.includes("application/json")) {
    return JSON.parse(text);
  }

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function getErrorMessage(response: Response, body: unknown): string {
  const graphMessage = getGraphErrorMessage(body);

  if (graphMessage) {
    return graphMessage;
  }

  if (typeof body === "string" && body.trim()) {
    return body;
  }

  return response.statusText || `Microsoft Graph request failed with status ${response.status}`;
}

function getGraphErrorMessage(body: unknown): string | undefined {
  if (!isRecord(body)) {
    return undefined;
  }

  const error = body.error;

  if (isRecord(error) && typeof error.message === "string") {
    return error.message;
  }

  if (typeof body.message === "string") {
    return body.message;
  }

  return undefined;
}

function getRequestId(response: Response, body: unknown): string | undefined {
  return (
    response.headers.get("request-id") ??
    response.headers.get("client-request-id") ??
    getRequestIdFromBody(body)
  );
}

function getRequestIdFromBody(body: unknown): string | undefined {
  if (!isRecord(body) || !isRecord(body.error)) {
    return undefined;
  }

  const innerError = body.error.innerError;

  if (!isRecord(innerError)) {
    return undefined;
  }

  const requestId = innerError["request-id"] ?? innerError.requestId;
  return typeof requestId === "string" ? requestId : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function trimTrailingSlashes(value: string): string {
  return value.replace(/\/+$/, "");
}
