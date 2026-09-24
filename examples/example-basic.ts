import { resolveGraphAccessToken } from "../src/graphAuth.ts";
import { MicrosoftTodoClient } from "../src/todoClient.ts";

const accessToken = await resolveGraphAccessToken(process.env);
const client = new MicrosoftTodoClient(accessToken);
const list = await client.getOrCreateTaskList("Microsoft To Do API Example");
const task = await client.createTask(list.id, {
  title: "Created by the Microsoft To Do API example",
  importance: "normal"
});

console.log(`Created task ${task.id}.`);
