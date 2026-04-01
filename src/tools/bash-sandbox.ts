import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";

import { getErrorMessage } from "../utils/errors.js";
import { getBash } from "../vfs/index.js";

const schema = Type.Object({
  command: Type.String({
    minLength: 1,
    description:
      "Bash command(s) to execute in the sandboxed VFS. Supports pipes (|), redirections (>, >>), chaining (&&, ||, ;), variables, loops, and conditionals. "
      + "Common commands include ls, cat, grep, find, awk, sed, jq, sort, uniq, wc, cut, head, tail, rg, yq, and xan. "
      + "Custom commands: csv-to-sheet and sheet-to-csv. No network access. No external runtimes like node or python.",
  }),
});

type Params = Static<typeof schema>;

function formatOutput(result: {
  stdout: string;
  stderr: string;
  exitCode: number;
}): string {
  let output = "";

  if (result.stdout) {
    output += result.stdout;
  }

  if (result.stderr) {
    if (output && !output.endsWith("\n")) output += "\n";
    output += `stderr: ${result.stderr}`;
  }

  if (result.exitCode !== 0) {
    if (output && !output.endsWith("\n")) output += "\n";
    output += `[exit code: ${result.exitCode}]`;
  }

  return output.trim() || "[no output]";
}

export function createBashSandboxTool(): AgentTool<typeof schema, undefined> {
  return {
    name: "bash",
    label: "Bash",
    description:
      "Execute bash commands in a sandboxed in-memory virtual filesystem. "
      + "Useful for file operations, text processing, CSV/JSON transformations, and Office.js API lookups. "
      + "User uploads live in /home/user/uploads/. "
      + "Custom commands: csv-to-sheet <file> <sheetName> [startCell] [--force] and sheet-to-csv <sheetName> [range] [file]. "
      + "No network access.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await getBash().exec(params.command);
        return {
          content: [{ type: "text", text: formatOutput(result) }],
          details: undefined,
        };
      } catch (error: unknown) {
        return {
          content: [{ type: "text", text: `Error: ${getErrorMessage(error)}` }],
          details: undefined,
        };
      }
    },
  };
}
