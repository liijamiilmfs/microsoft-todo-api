import {
  MicrosoftTodoClient,
  type ShoppingTaskInput
} from "../src/todoClient.ts";
import { resolveGraphAccessToken } from "../src/graphAuth.ts";

const accessToken = await resolveGraphAccessToken(process.env);
const client = new MicrosoftTodoClient(accessToken);

const shoppingItems: ShoppingTaskInput[] = [
  {
    store: "Costco",
    section: "Protein",
    item: "Chicken thighs",
    quantity: "1 large pack"
  },
  {
    store: "Costco",
    section: "Protein",
    item: "Chuck roast or brisket",
    quantity: "6-8 lbs"
  },
  {
    store: "Aldi",
    section: "Produce",
    item: "Peaches",
    quantity: "4-5"
  },
  {
    store: "Aldi",
    section: "Produce",
    item: "Corn",
    quantity: "8-12 ears"
  },
  {
    store: "Supplies",
    section: "Kitchen",
    item: "Aluminum foil",
    quantity: "1 roll"
  }
];

const tasks = await client.createShoppingTasks(
  "Family Meal Shopping - Aldi + Costco",
  shoppingItems
);

console.log(`Created ${tasks.length} shopping tasks.`);
