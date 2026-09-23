import { resolveGraphAccessToken } from "../src/graphAuth.ts";

await resolveGraphAccessToken(process.env);
console.log("Microsoft Graph authentication succeeded.");
