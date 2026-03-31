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
  maxTurns: number;
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
  maxTurns?: number;
}

export type SubAgentStatus = "completed" | "failed" | "max_turns_reached";

export interface SubAgentResult {
  roleId: SubAgentRoleId;
  roleName: string;
  status: SubAgentStatus;
  summary: string;
  toolCallCount: number;
  turnsUsed: number;
  errors: string[];
}
