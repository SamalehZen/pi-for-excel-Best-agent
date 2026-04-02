/**
 * delegate_task — Orchestrator tool for invoking specialized sub-agents.
 *
 * The main Pi agent calls this tool to delegate tasks to sub-agents
 * (Analyst, Builder, Stylist, Template Builder, Researcher, Modeler, Debugger).
 * Each sub-agent runs as an inner LLM loop with a focused system prompt and
 * restricted tool set, then returns a structured result.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult, StreamFn } from "@mariozechner/pi-agent-core";
import type { Model, Api } from "@mariozechner/pi-ai";

import { SUB_AGENT_ROLE_IDS, isValidRoleId } from "../agents/roles.js";
import type { SubAgentRoleId } from "../agents/types.js";
import { runSubAgent, type SubAgentRunnerDependencies } from "../agents/sub-agent-runner.js";
import type { DelegateTaskDetails } from "./tool-details.js";

function StringEnum<T extends string[]>(values: [...T], opts?: { description?: string }) {
  return Type.Union(
    values.map((v) => Type.Literal(v)),
    opts,
  );
}

const schema = Type.Object({
  role: StringEnum(
    [...SUB_AGENT_ROLE_IDS],
    {
      description:
        "Sub-agent role to delegate to. "
        + "analyst = read-only data comprehension. "
        + "builder = create structures and formulas. "
        + "stylist = formatting and visual design. "
        + "template-builder = analyze data layout and apply template designs. "
        + "researcher = web search and external data. "
        + "modeler = financial modeling and quantitative analysis. "
        + "debugger = error diagnosis and formula repair.",
    },
  ),
  task: Type.String({
    minLength: 1,
    maxLength: 4000,
    description: "Natural language description of the task for the sub-agent to accomplish.",
  }),
  context: Type.Optional(
    Type.String({
      maxLength: 8000,
      description:
        "Additional context for the sub-agent (e.g. template palette JSON, data from a previous sub-agent, specific instructions).",
    }),
  ),
  tools: Type.Optional(
    Type.Array(
      Type.String({ minLength: 1 }),
      {
        description:
          "Subset of tools to give the sub-agent. Only include tools strictly needed for this specific task. "
          + "If omitted, the sub-agent gets all its role's default tools. "
          + "Example: for a simple formatting task, give stylist only ['read_range', 'format_cells'] instead of all 8 tools.",
        minItems: 1,
        maxItems: 15,
      },
    ),
  ),
});

type Params = Static<typeof schema>;

export interface DelegateTaskToolDependencies {
  getStreamFn: () => StreamFn;
  getModel: () => Model<Api>;
  getAllTools: () => readonly AgentTool[];
  getApiKey?: (provider: string) => Promise<string | undefined> | string | undefined;
  getWorkbookContext?: () => string | undefined;
}

export function createDelegateTaskTool(
  deps: DelegateTaskToolDependencies,
): AgentTool<typeof schema, DelegateTaskDetails> {
  return {
    name: "delegate_task",
    label: "Delegate Task",
    description:
      "Delegate a task to a specialized sub-agent. Use for complex multi-step work that benefits from a focused specialist: "
      + "analysis (analyst), structure building (builder), formatting (stylist), template application (template-builder), "
      + "web research (researcher), financial modeling (modeler), or error diagnosis (debugger).",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
      signal?: AbortSignal,
    ): Promise<AgentToolResult<DelegateTaskDetails>> => {
      const roleId = params.role;

      if (!isValidRoleId(roleId)) {
        const invalidRole = String(roleId);
        return {
          content: [{ type: "text", text: `Unknown sub-agent role: "${invalidRole}". Available roles: ${SUB_AGENT_ROLE_IDS.join(", ")}` }],
          details: {
            kind: "delegate_task",
            roleId: invalidRole as SubAgentRoleId,
            roleName: invalidRole,
            status: "failed",
            summary: `Unknown role: ${invalidRole}`,
            toolCallCount: 0,
            turnsUsed: 0,
            errors: [`Unknown role: ${invalidRole}`],
          },
        };
      }

      const runnerDeps: SubAgentRunnerDependencies = {
        streamFn: deps.getStreamFn(),
        model: deps.getModel(),
        allTools: deps.getAllTools(),
        getApiKey: deps.getApiKey,
        workbookContext: deps.getWorkbookContext?.(),
      };

      const result = await runSubAgent(
        {
          roleId,
          task: params.task,
          context: params.context,
          tools: params.tools,
        },
        runnerDeps,
        signal,
      );

      const statusLabel = result.status === "completed"
        ? "completed"
        : "failed";

      const errorSection = result.errors.length > 0
        ? `\n\nErrors:\n${result.errors.map((e) => `- ${e}`).join("\n")}`
        : "";

      const content = [
        {
          type: "text" as const,
          text:
            `**${result.roleName}** — ${statusLabel}\n\n`
            + `${result.summary}\n\n`
            + `_${result.toolCallCount} tool calls in ${result.turnsUsed} turns_`
            + errorSection,
        },
      ];

      return {
        content,
        details: {
          kind: "delegate_task",
          roleId: result.roleId,
          roleName: result.roleName,
          status: result.status,
          summary: result.summary,
          toolCallCount: result.toolCallCount,
          turnsUsed: result.turnsUsed,
          errors: result.errors,
        },
      };
    },
  };
}
