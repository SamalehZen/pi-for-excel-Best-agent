/**
 * Sub-agent runner.
 *
 * Executes a sub-agent as an inner LLM loop with a specialized system prompt
 * and restricted tool set. Uses the same streamFn and provider as the parent
 * agent, and shares the WorkbookCoordinator for serialized mutations.
 *
 * The sub-agent runs synchronously within the parent's tool execution —
 * from the parent's perspective, `delegate_task` is just another tool call
 * that returns a structured result.
 */

import { runAgentLoop } from "@mariozechner/pi-agent-core/dist/agent-loop.js";
import type {
  AgentEvent,
  AgentMessage,
  AgentTool,
  StreamFn,
  BeforeToolCallContext,
  BeforeToolCallResult,
  AfterToolCallContext,
  AfterToolCallResult,
} from "@mariozechner/pi-agent-core";
import type { Message, Model, Api } from "@mariozechner/pi-ai";

import type {
  SubAgentRequest,
  SubAgentResult,
  SubAgentRole,
} from "./types.js";
import { getRole } from "./roles.js";
import { getErrorMessage } from "../utils/errors.js";

const DEFAULT_MAX_TURNS = 10;
const ABSOLUTE_MAX_TURNS = 15;

export interface SubAgentRunnerDependencies {
  streamFn: StreamFn;
  model: Model<Api>;
  allTools: readonly AgentTool[];
  getApiKey?: (provider: string) => Promise<string | undefined> | string | undefined;
  convertToLlm?: (messages: AgentMessage[]) => Message[] | Promise<Message[]>;
  beforeToolCall?: (context: BeforeToolCallContext, signal?: AbortSignal) => Promise<BeforeToolCallResult | undefined>;
  afterToolCall?: (context: AfterToolCallContext, signal?: AbortSignal) => Promise<AfterToolCallResult | undefined>;
  workbookContext?: string;
}

function filterToolsForRole(
  allTools: readonly AgentTool[],
  allowedNames: readonly string[],
): AgentTool[] {
  const allowed = new Set(allowedNames);
  return allTools.filter((tool) => allowed.has(tool.name));
}

function buildSubAgentSystemPrompt(
  role: SubAgentRole,
  taskDescription: string,
  taskContext: string | undefined,
  workbookContext: string | undefined,
): string {
  const sections: string[] = [];

  sections.push(role.systemPrompt);

  sections.push(`## Current Task\n\n${taskDescription}`);

  if (taskContext) {
    sections.push(`## Task Context\n\n${taskContext}`);
  }

  if (workbookContext && role.requiredContext.workbookBlueprint) {
    sections.push(`## Workbook Context (already loaded — do NOT call get_workbook_overview)\n\n${workbookContext}`);
  }

  sections.push(
    `## Constraints\n\n`
    + `- **CRITICAL: Minimum tool calls.** Having a tool available does NOT mean you should use it. Only call tools that are strictly necessary to complete the task. If you can finish in 3 calls, do NOT make 8.\n`
    + `- **Plan before acting.** Before your first tool call, mentally plan the minimum sequence of calls needed. State your plan in 1-2 lines, then execute.\n`
    + `- **Batch operations**: combine multiple writes/formats into single tool calls. One write_cells with 20 cells beats 20 separate calls.\n`
    + `- **Read once**: if the workbook blueprint is in context, do NOT call get_workbook_overview again. Read target ranges once, then work from memory.\n`
    + `- **No unnecessary verification**: do NOT re-read cells just to confirm a write succeeded — write_cells auto-verifies. Only re-read if you need the value for a subsequent formula.\n`
    + `- **Stop immediately** when the task is complete. Do not use remaining turns.\n`
    + `- You have a maximum of ${role.maxTurns} turns — but fewer is better. Target: complete in ${Math.ceil(role.maxTurns * 0.5)} turns or less.\n`
    + `- When done, provide a brief summary of what you accomplished with cell references.\n`
    + `- If you encounter an error you cannot resolve, stop and report it.\n`
    + `- Do not ask the user questions — you are a background worker. Make reasonable decisions.`,
  );

  return sections.join("\n\n");
}

function defaultConvertToLlm(messages: AgentMessage[]): Message[] {
  return messages.filter((m): m is Message => {
    if (typeof m !== "object" || m === null) return false;
    const msg = m as Record<string, unknown>;
    return msg.role === "user" || msg.role === "assistant" || msg.role === "toolResult";
  });
}

export async function runSubAgent(
  request: SubAgentRequest,
  deps: SubAgentRunnerDependencies,
  signal?: AbortSignal,
): Promise<SubAgentResult> {
  const role = getRole(request.roleId);
  if (!role) {
    return {
      roleId: request.roleId,
      roleName: request.roleId,
      status: "failed",
      summary: `Unknown sub-agent role: "${request.roleId}"`,
      toolCallCount: 0,
      turnsUsed: 0,
      errors: [`Unknown role: ${request.roleId}`],
    };
  }

  const maxTurns = Math.min(
    request.maxTurns ?? role.maxTurns ?? DEFAULT_MAX_TURNS,
    ABSOLUTE_MAX_TURNS,
  );

  const toolAllowList = request.tools && request.tools.length > 0
    ? request.tools.filter((t) => role.allowedTools.includes(t))
    : [...role.allowedTools];

  const tools = filterToolsForRole(deps.allTools, toolAllowList);
  if (tools.length === 0) {
    return {
      roleId: role.id,
      roleName: role.name,
      status: "failed",
      summary: `No tools available for role "${role.name}". Required: ${(request.tools ?? role.allowedTools).join(", ")}`,
      toolCallCount: 0,
      turnsUsed: 0,
      errors: ["No matching tools found in the current tool set"],
    };
  }

  const systemPrompt = buildSubAgentSystemPrompt(
    role,
    request.task,
    request.context,
    deps.workbookContext,
  );

  const convertToLlm = deps.convertToLlm ?? defaultConvertToLlm;

  let toolCallCount = 0;
  let turnsUsed = 0;
  const errors: string[] = [];
  let lastAssistantText = "";

  const emit = (event: AgentEvent): void => {
    if (event.type === "tool_execution_end") {
      toolCallCount += 1;
      if (event.isError) {
        const errText = typeof event.result === "string"
          ? event.result
          : event.toolName + " failed";
        errors.push(errText);
      }
    }

    if (event.type === "turn_end") {
      turnsUsed += 1;
    }

    if (event.type === "message_end") {
      const msg = event.message;
      if (
        typeof msg === "object"
        && msg !== null
        && (msg as Record<string, unknown>).role === "assistant"
      ) {
        const content = (msg as Record<string, unknown>).content;
        if (Array.isArray(content)) {
          const textParts = content
            .filter((c: unknown) => {
              const block = c as Record<string, unknown>;
              return block.type === "text" && typeof block.text === "string";
            })
            .map((c: unknown) => (c as { text: string }).text);

          if (textParts.length > 0) {
            lastAssistantText = textParts.join("\n");
          }
        }
      }
    }
  };

  const context = {
    systemPrompt,
    messages: [] as AgentMessage[],
    tools,
  };

  const config: Parameters<typeof runAgentLoop>[2] = {
    model: deps.model,
    convertToLlm,
    getApiKey: deps.getApiKey,
    beforeToolCall: async (ctx, sig) => {
      if (turnsUsed >= maxTurns) {
        return { block: true, reason: `Sub-agent turn limit reached (${maxTurns} turns). Stopping execution.` };
      }
      if (deps.beforeToolCall) {
        return deps.beforeToolCall(ctx, sig);
      }
      return undefined;
    },
    afterToolCall: deps.afterToolCall,
    toolExecution: "sequential",
  };

  const userPrompt: AgentMessage = {
    role: "user",
    content: [{ type: "text", text: request.task }],
    timestamp: Date.now(),
  } satisfies Message as AgentMessage;

  try {
    await runAgentLoop(
      [userPrompt],
      context,
      config,
      emit,
      signal,
      deps.streamFn,
    );
  } catch (error: unknown) {
    const errMsg = getErrorMessage(error);
    errors.push(errMsg);

    return {
      roleId: role.id,
      roleName: role.name,
      status: "failed",
      summary: lastAssistantText || `Sub-agent failed: ${errMsg}`,
      toolCallCount,
      turnsUsed,
      errors,
    };
  }

  const status = turnsUsed >= maxTurns ? "max_turns_reached" : "completed";
  const summary = lastAssistantText || `${role.name} completed with ${toolCallCount} tool calls.`;

  return {
    roleId: role.id,
    roleName: role.name,
    status,
    summary,
    toolCallCount,
    turnsUsed,
    errors,
  };
}
