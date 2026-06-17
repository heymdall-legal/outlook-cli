#!/usr/bin/env node

import { createCli } from "./cli.js";

async function main() {
  const cli = createCli();
  const exitCode = await cli.run(process.argv.slice(2));
  if (exitCode !== 0) {
    process.exitCode = exitCode;
  }
}

main();
