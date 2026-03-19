import { existsSync } from "node:fs";
import { join } from "node:path";

if (process.env.npm_config_package_lock_only === "true") {
  console.log("Skipping prepare during package-lock-only install");
  process.exit(0);
}

if (existsSync(join(import.meta.dirname, "..", "dist", "index.js"))) {
  console.log("Skipping prepare because dist/index.js already exists");
  process.exit(0);
}

process.exitCode = 1;
const { spawn } = await import("node:child_process");

const child = spawn("npm", ["run", "build"], {
  stdio: "inherit",
  shell: process.platform === "win32",
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 1);
});
