/**
 * Sub-agent role registry.
 *
 * Defines the available sub-agent roles, each with a specialized system prompt,
 * restricted tool set, and execution constraints. The orchestrator (main Pi agent)
 * delegates tasks to sub-agents via the `delegate_task` tool.
 */

import type { SubAgentRoleId, SubAgentRole } from "./types.js";
import { ANALYST_ROLE } from "./roles/analyst.js";
import { BUILDER_ROLE } from "./roles/builder.js";
import { STYLIST_ROLE } from "./roles/stylist.js";
import { TEMPLATE_BUILDER_ROLE } from "./roles/template-builder.js";
import { RESEARCHER_ROLE } from "./roles/researcher.js";
import { MODELER_ROLE } from "./roles/modeler.js";
import { DEBUGGER_ROLE } from "./roles/debugger.js";

export type { SubAgentRoleId, SubAgentRole } from "./types.js";

const ROLE_REGISTRY = new Map<SubAgentRoleId, SubAgentRole>([
  ["analyst", ANALYST_ROLE],
  ["builder", BUILDER_ROLE],
  ["stylist", STYLIST_ROLE],
  ["template-builder", TEMPLATE_BUILDER_ROLE],
  ["researcher", RESEARCHER_ROLE],
  ["modeler", MODELER_ROLE],
  ["debugger", DEBUGGER_ROLE],
]);

export const SUB_AGENT_ROLE_IDS: readonly SubAgentRoleId[] = [
  "analyst",
  "builder",
  "stylist",
  "template-builder",
  "researcher",
  "modeler",
  "debugger",
] as const;

export function getRole(id: SubAgentRoleId): SubAgentRole | null {
  return ROLE_REGISTRY.get(id) ?? null;
}

export function listRoles(): readonly SubAgentRole[] {
  return SUB_AGENT_ROLE_IDS.map((id) => {
    const role = ROLE_REGISTRY.get(id);
    if (!role) throw new Error(`Missing role definition for ${id}`);
    return role;
  });
}

export function isValidRoleId(value: string): value is SubAgentRoleId {
  return ROLE_REGISTRY.has(value as SubAgentRoleId);
}
