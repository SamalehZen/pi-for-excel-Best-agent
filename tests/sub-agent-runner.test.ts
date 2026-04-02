import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentEvent, AgentTool, StreamFn } from "@mariozechner/pi-agent-core";
import { Type } from "@sinclair/typebox";
import type { Api, Model } from "@mariozechner/pi-ai";

import { runSubAgent, type SubAgentRunnerDependencies } from "../src/agents/sub-agent-runner.ts";

const EMPTY_PARAMS = Type.Object({});

function makeTool(name: string): AgentTool<typeof EMPTY_PARAMS, { ok: true }> {
  return {
    name,
    label: name,
    description: `test ${name}`,
    parameters: EMPTY_PARAMS,
    execute: () => Promise.resolve({
      content: [{ type: "text", text: `${name}:ok` }],
      details: { ok: true },
    }),
  };
}

function makeDeps(overrides: Partial<SubAgentRunnerDependencies> = {}): SubAgentRunnerDependencies {
  const deps: SubAgentRunnerDependencies = {
    streamFn: (() => {
      throw new Error("streamFn should not be called in tests");
    }) as StreamFn,
    model: {} as Model<Api>,
    allTools: [makeTool("read_range")],
    ...overrides,
  };

  return deps;
}

void test("runSubAgent returns completed summary when loop emits assistant text", async () => {
  const result = await runSubAgent(
    {
      roleId: "analyst",
      task: "Summarize the sheet",
    },
    makeDeps({
      runAgentLoopImpl: (_messages, _context, _config, emit) => {
        const messageEnd: AgentEvent = {
          type: "message_end",
          message: {
            role: "assistant",
            content: [{ type: "text", text: "Summary complete." }],
          },
        };
        const turnEnd: AgentEvent = { type: "turn_end" };

        void emit(messageEnd);
        void emit(turnEnd);
        return Promise.resolve();
      },
    }),
  );

  assert.equal(result.status, "completed");
  assert.equal(result.summary, "Summary complete.");
  assert.equal(result.turnsUsed, 1);
});

void test("runSubAgent fails with inactivity timeout when loop stops making progress", async () => {
  const result = await runSubAgent(
    {
      roleId: "analyst",
      task: "Inspect data quality",
    },
    makeDeps({
      inactivityTimeoutMs: 20,
      runAgentLoopImpl: (_messages, _context, _config, _emit, signal) => {
        return new Promise<void>((_resolve, reject) => {
          const onAbort = () => {
            signal?.removeEventListener("abort", onAbort);
            reject(new Error("Aborted"));
          };

          if (signal?.aborted) {
            reject(new Error("Aborted"));
            return;
          }

          signal?.addEventListener("abort", onAbort, { once: true });
        });
      },
    }),
  );

  assert.equal(result.status, "failed");
  assert.match(result.summary, /timed out after 20ms of inactivity/i);
  assert.deepEqual(result.errors, ["Sub-agent timed out after 20ms of inactivity."]);
});
