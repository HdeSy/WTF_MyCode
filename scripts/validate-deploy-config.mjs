import { readFileSync } from "node:fs";

const configPath = new URL("../wrangler.jsonc", import.meta.url);
const config = JSON.parse(readFileSync(configPath, "utf8"));
const database = config.d1_databases?.find((item) => item.binding === "DB");
const validD1Id = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

if (!config.name || !config.main || !database || !database.database_name) {
  console.error("Deployment blocked: wrangler.jsonc is missing the Worker entry point or D1 binding.");
  process.exit(1);
}
if (!validD1Id.test(database.database_id ?? "")) {
  console.error("Deployment blocked: replace the D1 database_id placeholder in wrangler.jsonc with the actual Cloudflare D1 database ID.");
  process.exit(1);
}
console.log(`Deployment config validated for Worker "${config.name}" and D1 database "${database.database_name}".`);
console.log("Reminder: configure BETTER_AUTH_SECRET as an encrypted Cloudflare Worker secret before accepting real users.");
