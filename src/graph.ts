import { requestUrl } from "obsidian";
import type { RemoteItem } from "./types";
import { MicrosoftAuth } from "./auth";

const GRAPH = "https://graph.microsoft.com/v1.0";

interface DriveItemJson {
  id: string;
  name: string;
  eTag?: string;
  size?: number;
  lastModifiedDateTime?: string;
  folder?: { childCount: number };
}

interface ChildrenResponse {
  value: DriveItemJson[];
  "@odata.nextLink"?: string;
}

export class GraphClient {
  constructor(private readonly auth: MicrosoftAuth) {}

  async ensureVaultFolder(vaultId: string): Promise<DriveItemJson> {
    const appRoot = await this.json<DriveItemJson>("GET", "/me/drive/special/approot");
    const vaults = await this.ensureChildFolder(appRoot.id, "vaults");
    return this.ensureChildFolder(vaults.id, vaultId);
  }

  async listTree(root: DriveItemJson): Promise<Map<string, RemoteItem>> {
    const result = new Map<string, RemoteItem>();
    await this.walk(root.id, "", result);
    return result;
  }

  async download(itemId: string): Promise<ArrayBuffer> {
    const response = await this.request("GET", `/me/drive/items/${itemId}/content`);
    if (response.status < 200 || response.status >= 300) throw graphError(response.status, response.json);
    return response.arrayBuffer;
  }

  async upload(rootId: string, path: string, data: ArrayBuffer): Promise<RemoteItem> {
    const parts = splitPath(path);
    const name = parts.pop();
    if (!name) throw new Error(`无效路径: ${path}`);
    const parentId = await this.ensureFolderPath(rootId, parts);
    const item = await this.json<DriveItemJson>(
      "PUT",
      `/me/drive/items/${parentId}:/${encodeURIComponent(name)}:/content`,
      data,
      "application/octet-stream"
    );
    return toRemote(item, path);
  }

  async delete(itemId: string): Promise<void> {
    const response = await this.request("DELETE", `/me/drive/items/${itemId}`);
    if (response.status < 200 || response.status >= 300) throw graphError(response.status, response.json);
  }

  private async walk(parentId: string, parentPath: string, result: Map<string, RemoteItem>): Promise<void> {
    let url: string | undefined = `/me/drive/items/${parentId}/children?$top=200`;
    while (url) {
      const page: ChildrenResponse = await this.json<ChildrenResponse>("GET", url);
      for (const item of page.value) {
        const path = parentPath ? `${parentPath}/${item.name}` : item.name;
        if (item.folder) await this.walk(item.id, path, result);
        else result.set(path, toRemote(item, path));
      }
      url = page["@odata.nextLink"];
    }
  }

  private async ensureFolderPath(rootId: string, parts: string[]): Promise<string> {
    let currentId = rootId;
    for (const part of parts) currentId = (await this.ensureChildFolder(currentId, part)).id;
    return currentId;
  }

  private async ensureChildFolder(parentId: string, name: string): Promise<DriveItemJson> {
    const response = await this.request(
      "POST",
      `/me/drive/items/${parentId}/children`,
      JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
      "application/json"
    );
    if (response.status >= 200 && response.status < 300) return response.json as DriveItemJson;
    if (response.status !== 409) throw graphError(response.status, response.json);

    let url: string | undefined = `/me/drive/items/${parentId}/children?$select=id,name,folder&$top=200`;
    while (url) {
      const page: ChildrenResponse = await this.json<ChildrenResponse>("GET", url);
      const found = page.value.find((item) => item.name === name && item.folder);
      if (found) return found;
      url = page["@odata.nextLink"];
    }
    throw new Error(`无法找到远端目录: ${name}`);
  }

  private async json<T>(method: string, pathOrUrl: string, body?: string | ArrayBuffer, contentType?: string): Promise<T> {
    const response = await this.request(method, pathOrUrl, body, contentType);
    if (response.status < 200 || response.status >= 300) throw graphError(response.status, response.json);
    return response.json as T;
  }

  private async request(method: string, pathOrUrl: string, body?: string | ArrayBuffer, contentType?: string) {
    const token = await this.auth.accessToken();
    return requestUrl({
      url: pathOrUrl.startsWith("http") ? pathOrUrl : `${GRAPH}${pathOrUrl}`,
      method,
      headers: { Authorization: `Bearer ${token}` },
      contentType,
      body,
      throw: false
    });
  }
}

function toRemote(item: DriveItemJson, path: string): RemoteItem {
  return {
    id: item.id,
    name: item.name,
    path,
    eTag: item.eTag ?? "",
    mtime: Date.parse(item.lastModifiedDateTime ?? "") || 0,
    size: item.size ?? 0,
    folder: Boolean(item.folder)
  };
}

function splitPath(path: string): string[] {
  return path.split("/").filter(Boolean);
}

function graphError(status: number, value: unknown): Error {
  const graph = value as { error?: { message?: string; code?: string } } | undefined;
  return new Error(`OneDrive 请求失败 (${status}): ${graph?.error?.message ?? graph?.error?.code ?? "未知错误"}`);
}
