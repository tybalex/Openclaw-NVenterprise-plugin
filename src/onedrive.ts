/**
 * OneDrive tool.
 *
 * Upload files to OneDrive, list files, and create sharing links.
 * Uses Azure AD refresh token -> OBO for Graph Files.ReadWrite scope.
 *
 * Supports:
 * - Simple upload (< 4 MB)
 * - Chunked upload (>= 4 MB) via upload session
 * - File listing
 * - Sharing: invite people (read-only) or create sharing links
 */

import { Type } from "@sinclair/typebox";
import { jsonResult, readStringParam, readStringArrayParam } from "openclaw/plugin-sdk/agent-runtime";
import { stringEnum, type AnyAgentTool } from "openclaw/plugin-sdk/core";
import { acquireDownstreamToken, isAzureOBOConfigured } from "./azure-obo.js";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";

// =============================================================================
// Configuration
// =============================================================================

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const ONEDRIVE_SCOPES = "https://graph.microsoft.com/Files.ReadWrite";
const DEFAULT_TIMEOUT_MS = 120_000;
const SMALL_FILE_LIMIT = 4 * 1024 * 1024; // 4 MB
const UPLOAD_CHUNK_SIZE = 3_276_800; // ~3.1 MB (320 KiB aligned)

// =============================================================================
// Actions & Schema
// =============================================================================

const ONEDRIVE_ACTIONS = ["upload_file", "list_files", "create_sharing_link", "share_with_people"] as const;

const OneDriveSchema = Type.Object({
  action: stringEnum(ONEDRIVE_ACTIONS, {
    description:
      "Action: upload_file, list_files, create_sharing_link, share_with_people.",
  }),
  file_path: Type.Optional(
    Type.String({ description: "Local file path to upload (for upload_file). Supports ~ for home dir." }),
  ),
  destination_path: Type.Optional(
    Type.String({ description: 'Destination path in OneDrive (e.g. "Documents/report.pdf"). Default: root with original name.' }),
  ),
  folder_path: Type.Optional(
    Type.String({ description: 'OneDrive folder to list (for list_files). Default: root.' }),
  ),
  item_id: Type.Optional(
    Type.String({ description: "OneDrive item ID (for sharing actions)." }),
  ),
  share_with: Type.Optional(
    Type.Array(Type.String(), { description: 'Email addresses to share with (for share_with_people).' }),
  ),
  share_link_type: Type.Optional(
    Type.String({ description: '"anonymous" or "organization" (for create_sharing_link).' }),
  ),
  share_message: Type.Optional(
    Type.String({ description: "Optional message for sharing invitation." }),
  ),
});

// =============================================================================
// Graph API Helpers
// =============================================================================

function encodeURIPath(filePath: string): string {
  return filePath.split("/").map((s) => encodeURIComponent(s)).join("/");
}

function formatBytes(bytes: number): string {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

// =============================================================================
// Upload
// =============================================================================

async function simpleUpload(
  localPath: string,
  oneDrivePath: string,
  token: string,
): Promise<{ itemId: string; name: string; webUrl: string; size: number }> {
  const content = await fs.readFile(localPath);
  const url = `${GRAPH_BASE_URL}/me/drive/root:/${encodeURIPath(oneDrivePath)}:/content`;

  const res = await fetch(url, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" },
    body: new Uint8Array(content.buffer, content.byteOffset, content.byteLength),
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });

  if (!res.ok) throw new Error(`Upload failed (${res.status}): ${await res.text()}`);
  const data = (await res.json()) as any;
  return { itemId: data.id || "", name: data.name || path.basename(oneDrivePath), webUrl: data.webUrl || "", size: content.length };
}

async function chunkedUpload(
  localPath: string,
  oneDrivePath: string,
  token: string,
  fileSize: number,
): Promise<{ itemId: string; name: string; webUrl: string; size: number }> {
  // Create upload session
  const sessionUrl = `${GRAPH_BASE_URL}/me/drive/root:/${encodeURIPath(oneDrivePath)}:/createUploadSession`;
  const sessionRes = await fetch(sessionUrl, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename", name: path.basename(oneDrivePath) } }),
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });
  if (!sessionRes.ok) throw new Error(`Session creation failed (${sessionRes.status}): ${await sessionRes.text()}`);
  const { uploadUrl } = (await sessionRes.json()) as any;
  if (!uploadUrl) throw new Error("No upload URL in session response");

  // Upload chunks
  const handle = await fs.open(localPath, "r");
  let offset = 0;
  try {
    while (offset < fileSize) {
      const chunkSize = Math.min(UPLOAD_CHUNK_SIZE, fileSize - offset);
      const buffer = new Uint8Array(chunkSize);
      const { bytesRead } = await handle.read(buffer, 0, chunkSize, offset);
      const chunk = buffer.subarray(0, bytesRead);

      const res = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(bytesRead),
          "Content-Range": `bytes ${offset}-${offset + bytesRead - 1}/${fileSize}`,
        },
        body: chunk,
        signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
      });

      if (res.status === 200 || res.status === 201) {
        const data = (await res.json()) as any;
        return { itemId: data.id || "", name: data.name || path.basename(oneDrivePath), webUrl: data.webUrl || "", size: fileSize };
      }
      if (!res.ok && res.status !== 202) {
        throw new Error(`Chunk upload failed (${res.status}): ${await res.text()}`);
      }
      offset += bytesRead;
    }
  } finally {
    await handle.close();
  }

  return { itemId: "", name: path.basename(oneDrivePath), webUrl: "", size: fileSize };
}

// =============================================================================
// Action Handlers
// =============================================================================

async function handleUploadFile(token: string, params: Record<string, unknown>): Promise<unknown> {
  const filePath = readStringParam(params, "file_path", { required: true });
  const destPath = readStringParam(params, "destination_path");

  const expanded = filePath.startsWith("~") ? filePath.replace(/^~/, os.homedir()) : filePath;
  const resolved = path.resolve(expanded);

  const stats = await fs.stat(resolved);
  if (!stats.isFile()) return { error: `Not a file: ${resolved}` };

  const oneDrivePath = destPath?.trim().replace(/^\/+|\/+$/g, "") || path.basename(resolved);

  const result =
    stats.size < SMALL_FILE_LIMIT
      ? await simpleUpload(resolved, oneDrivePath, token)
      : await chunkedUpload(resolved, oneDrivePath, token, stats.size);

  return {
    name: result.name,
    size: formatBytes(result.size),
    oneDrivePath,
    webUrl: result.webUrl,
    itemId: result.itemId,
  };
}

async function handleListFiles(token: string, params: Record<string, unknown>): Promise<unknown> {
  const folderPath = readStringParam(params, "folder_path");
  let url: string;
  if (folderPath?.trim()) {
    const clean = folderPath.trim().replace(/^\/+|\/+$/g, "");
    url = `${GRAPH_BASE_URL}/me/drive/root:/${encodeURIPath(clean)}:/children`;
  } else {
    url = `${GRAPH_BASE_URL}/me/drive/root/children`;
  }
  url += "?$select=name,size,lastModifiedDateTime,folder,file,webUrl&$top=100";

  const res = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });
  if (!res.ok) {
    if (res.status === 404) return { error: `Folder not found: ${folderPath || "/"}` };
    throw new Error(`List failed (${res.status}): ${await res.text()}`);
  }

  const data = (await res.json()) as any;
  const items = (data.value || []).map((item: any) => ({
    name: item.name,
    type: item.folder ? "folder" : "file",
    size: item.folder ? `${item.folder.childCount || 0} items` : formatBytes(item.size || 0),
    lastModified: item.lastModifiedDateTime,
    webUrl: item.webUrl,
  }));
  return { folder: folderPath || "/", count: items.length, items };
}

async function handleShareWithPeople(token: string, params: Record<string, unknown>): Promise<unknown> {
  const itemId = readStringParam(params, "item_id", { required: true });
  const shareWith = params.share_with as string[] | undefined;
  const message = readStringParam(params, "share_message");

  if (!shareWith?.length) return { error: "share_with is required" };

  const body: any = {
    recipients: shareWith.map((email) => ({ "@odata.type": "microsoft.graph.driveRecipient", email: email.trim() })),
    roles: ["read"],
    requireSignIn: true,
    sendInvitation: true,
  };
  if (message) body.message = message;

  const res = await fetch(`${GRAPH_BASE_URL}/me/drive/items/${itemId}/invite`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });

  if (!res.ok) return { error: `Share failed (${res.status}): ${await res.text()}` };
  const data = (await res.json()) as any;
  const invited = data.value || [];
  return { shared: invited.filter((p: any) => !p.error).length, total: shareWith.length };
}

async function handleCreateSharingLink(token: string, params: Record<string, unknown>): Promise<unknown> {
  const itemId = readStringParam(params, "item_id", { required: true });
  const scope = readStringParam(params, "share_link_type") || "organization";

  const res = await fetch(`${GRAPH_BASE_URL}/me/drive/items/${itemId}/createLink`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ type: "view", scope }),
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });

  if (!res.ok) return { error: `Link creation failed (${res.status}): ${await res.text()}` };
  const data = (await res.json()) as any;
  return { link: data.link?.webUrl || "", scope };
}

// =============================================================================
// Tool Factory
// =============================================================================

export function createOneDriveTool(options: {
  getRefreshToken: () => string | null | Promise<string | null>;
  enabled?: boolean;
}): AnyAgentTool | null {
  if (options.enabled === false) return null;
  if (!isAzureOBOConfigured()) return null;

  return {
    label: "OneDrive",
    name: "onedrive",
    description:
      "Upload files to OneDrive, list files, and create sharing links via Microsoft Graph. Supports simple (<4MB) and chunked (>=4MB) uploads, invite people for read access, and create anonymous/organization sharing links. Requires NVIDIA SSO authentication.",
    parameters: OneDriveSchema,
    execute: async (_toolCallId, args) => {
      const params = args as Record<string, unknown>;
      const action = readStringParam(params, "action", { required: true });

      try {
        const refreshToken = await options.getRefreshToken();
        if (!refreshToken) return jsonResult({ error: "not_authenticated", message: "Please log in first." });

        const tokenResult = await acquireDownstreamToken(refreshToken, ONEDRIVE_SCOPES);
        if (!tokenResult.ok) {
          return jsonResult({ error: "token_exchange_failed", message: `Failed to acquire token: ${"error" in tokenResult ? tokenResult.error : "unknown"}` });
        }
        const graphToken = tokenResult.accessToken;

        let result: unknown;
        switch (action) {
          case "upload_file":
            result = await handleUploadFile(graphToken, params);
            break;
          case "list_files":
            result = await handleListFiles(graphToken, params);
            break;
          case "share_with_people":
            result = await handleShareWithPeople(graphToken, params);
            break;
          case "create_sharing_link":
            result = await handleCreateSharingLink(graphToken, params);
            break;
          default:
            return jsonResult({ error: "invalid_action", message: `Unknown action: ${action}. Valid: ${ONEDRIVE_ACTIONS.join(", ")}` });
        }

        return jsonResult({ action, result });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return jsonResult({ error: "onedrive_error", message });
      }
    },
  } as AnyAgentTool;
}
