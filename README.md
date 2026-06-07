# Microsoft To Do API Client

Lightweight TypeScript client for Microsoft To Do through Microsoft Graph v1.0.
It accepts an already-acquired delegated Microsoft Graph access token and uses
native `fetch`.

## Requirements

- Environment variable: `GRAPH_ACCESS_TOKEN`
- Microsoft Graph delegated permission: `Tasks.ReadWrite`
- Node.js 24 or newer for the included TypeScript test/example workflow

This module does not implement OAuth. Acquire the access token in your app, then
pass it to `MicrosoftTodoClient`.

## Basic Usage

```ts
import { MicrosoftTodoClient } from "./src/todoClient.ts";

const accessToken = process.env.GRAPH_ACCESS_TOKEN;

if (!accessToken) {
  throw new Error("Set GRAPH_ACCESS_TOKEN before running.");
}

const client = new MicrosoftTodoClient(accessToken);

const list = await client.getOrCreateTaskList("Family Meal Shopping - Aldi + Costco");
const task = await client.createTask(list.id, {
  title: "Buy chicken thighs",
  importance: "normal"
});

console.log(task.id);
```

## Shopping Example

```sh
GRAPH_ACCESS_TOKEN="<token>" node examples/example-shopping.ts
```

## Test

```sh
npm test
```
