// Type stubs for openclaw/plugin-sdk — resolved at runtime by jiti
declare module "openclaw/plugin-sdk/core" {
  export type AnyAgentTool = {
    name: string;
    label?: string;
    description: string;
    parameters: unknown;
    execute: (toolCallId: string, args: Record<string, unknown>) => Promise<unknown>;
  };

  export function stringEnum<T extends string>(values: readonly T[], opts?: Record<string, unknown>): unknown;

  export function definePluginEntry(entry: {
    id: string;
    name: string;
    description: string;
    register: (api: {
      registerTool: (tool: AnyAgentTool) => void;
      registerHttpRoute: (params: {
        path: string;
        auth: "gateway" | "plugin";
        match?: "exact" | "prefix";
        handler: (req: import("node:http").IncomingMessage, res: import("node:http").ServerResponse) => unknown;
      }) => void;
      config: unknown;
      pluginConfig?: Record<string, unknown>;
    }) => void;
  }): unknown;
}

declare module "openclaw/plugin-sdk/agent-runtime" {
  export function jsonResult(data: unknown): unknown;
  export function readStringParam(params: Record<string, unknown>, key: string, opts?: { required?: boolean }): string;
  export function readNumberParam(params: Record<string, unknown>, key: string, opts?: { integer?: boolean }): number | null;
  export function readStringArrayParam(params: Record<string, unknown>, key: string): string[] | undefined;
}
