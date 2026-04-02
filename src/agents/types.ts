/**
 * Sub-agent type definitions.
 */

export type SubAgentRoleId =
  | "analyst"
  | "builder"
  | "stylist"
  | "template-builder"
  | "researcher"
  | "modeler"
  | "debugger";

export interface SubAgentRole {
  id: SubAgentRoleId;
  name: string;
  description: string;
  systemPrompt: string;
  allowedTools: readonly string[];
  requiredContext: SubAgentContextRequirements;
  skillsToPreload: readonly string[];
}

export interface SubAgentContextRequirements {
  workbookBlueprint: boolean;
  selectionState: boolean;
  recentChanges: boolean;
}

export interface SubAgentRequest {
  roleId: SubAgentRoleId;
  task: string;
  context?: string;
  /** Optional subset of tools to give the sub-agent. If omitted, uses the role's full allowedTools. */
  tools?: string[];
}

export type SubAgentStatus = "completed" | "failed";

export interface SubAgentResult {
  roleId: SubAgentRoleId;
  roleName: string;
  status: SubAgentStatus;
  summary: string;
  toolCallCount: number;
  turnsUsed: number;
  errors: string[];
}
