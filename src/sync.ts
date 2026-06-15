import { App, normalizePath, Platform } from "obsidian";
import { GraphClient } from "./graph";
import type { LocalItem, RemoteItem, SyncEntry, SyncSettings, SyncSummary } from "./types";

export class SyncEngine {
  private running = false;

  constructor(
    private readonly app: App,
    private readonly graph: GraphClient,
    private readonly settings: SyncSettings,
    private readonly save: () => Promise<void>
  ) {}

  async sync(): Promise<SyncSummary> {
    if (this.running) throw new Error("同步已在进行中");
    this.running = true;
    const summary: SyncSummary = {
      uploaded: 0,
      downloaded: 0,
      deletedLocal: 0,
      deletedRemote: 0,
      conflicts: 0,
      skipped: 0
    };

    try {
      const root = await this.graph.ensureVaultFolder(this.settings.vaultId);
      const [local, remote] = await Promise.all([this.scanLocal(), this.graph.listTree(root)]);
      const paths = new Set([...local.keys(), ...remote.keys(), ...Object.keys(this.settings.entries)]);

      for (const path of [...paths].sort()) {
        if (this.excluded(path)) {
          summary.skipped++;
          continue;
        }
        await this.reconcile(path, root.id, local.get(path), remote.get(path), summary);
      }

      this.settings.lastSyncAt = Date.now();
      await this.save();
      return summary;
    } finally {
      this.running = false;
    }
  }

  private async reconcile(
    path: string,
    rootId: string,
    local: LocalItem | undefined,
    remote: RemoteItem | undefined,
    summary: SyncSummary
  ): Promise<void> {
    const previous = this.settings.entries[path];

    if (local && remote) {
      if (!previous) {
        if (local.size === remote.size && Math.abs(local.mtime - remote.mtime) < 2000) {
          this.record(path, local, remote);
        } else if (local.mtime > remote.mtime) {
          const uploaded = await this.upload(rootId, local);
          this.record(path, local, uploaded);
          summary.uploaded++;
        } else {
          const downloaded = await this.download(remote);
          this.record(path, downloaded, remote);
          summary.downloaded++;
        }
        return;
      }

      const localChanged = changedLocal(local, previous);
      const remoteChanged = remote.eTag !== previous.remoteETag;
      if (localChanged && remoteChanged) {
        await this.createConflictCopy(path);
        const downloaded = await this.download(remote);
        this.record(path, downloaded, remote);
        summary.conflicts++;
      } else if (localChanged) {
        const uploaded = await this.upload(rootId, local);
        this.record(path, local, uploaded);
        summary.uploaded++;
      } else if (remoteChanged) {
        const downloaded = await this.download(remote);
        this.record(path, downloaded, remote);
        summary.downloaded++;
      }
      return;
    }

    if (local) {
      if (previous && !changedLocal(local, previous)) {
        await this.removeLocal(path);
        delete this.settings.entries[path];
        summary.deletedLocal++;
      } else {
        const uploaded = await this.upload(rootId, local);
        this.record(path, local, uploaded);
        summary.uploaded++;
      }
      return;
    }

    if (remote) {
      if (previous && remote.eTag === previous.remoteETag) {
        await this.graph.delete(remote.id);
        delete this.settings.entries[path];
        summary.deletedRemote++;
      } else {
        const downloaded = await this.download(remote);
        this.record(path, downloaded, remote);
        summary.downloaded++;
      }
      return;
    }

    delete this.settings.entries[path];
  }

  private async scanLocal(): Promise<Map<string, LocalItem>> {
    const result = new Map<string, LocalItem>();
    const walk = async (folder: string): Promise<void> => {
      const listing = await this.app.vault.adapter.list(folder);
      for (const file of listing.files) {
        const path = normalizePath(file);
        if (this.excluded(path)) continue;
        const stat = await this.app.vault.adapter.stat(path);
        if (stat) result.set(path, { path, mtime: stat.mtime, size: stat.size });
      }
      for (const child of listing.folders) {
        const path = normalizePath(child);
        if (!this.excluded(`${path}/`)) await walk(path);
      }
    };
    await walk("");
    return result;
  }

  private async upload(rootId: string, local: LocalItem): Promise<RemoteItem> {
    const data = await this.app.vault.adapter.readBinary(local.path);
    return this.graph.upload(rootId, local.path, data);
  }

  private async download(remote: RemoteItem): Promise<LocalItem> {
    const parent = remote.path.includes("/") ? remote.path.slice(0, remote.path.lastIndexOf("/")) : "";
    if (parent) await this.ensureLocalFolder(parent);
    await this.app.vault.adapter.writeBinary(remote.path, await this.graph.download(remote.id));
    const stat = await this.app.vault.adapter.stat(remote.path);
    if (!stat) throw new Error(`写入本地文件失败: ${remote.path}`);
    return { path: remote.path, mtime: stat.mtime, size: stat.size };
  }

  private async removeLocal(path: string): Promise<void> {
    if (await this.app.vault.adapter.exists(path)) await this.app.vault.adapter.remove(path);
  }

  private async createConflictCopy(path: string): Promise<void> {
    const dot = path.lastIndexOf(".");
    const suffix = ` (本地冲突 ${deviceLabel()} ${timestamp()})`;
    const conflict = dot > path.lastIndexOf("/") ? `${path.slice(0, dot)}${suffix}${path.slice(dot)}` : `${path}${suffix}`;
    const parent = conflict.includes("/") ? conflict.slice(0, conflict.lastIndexOf("/")) : "";
    if (parent) await this.ensureLocalFolder(parent);
    await this.app.vault.adapter.writeBinary(conflict, await this.app.vault.adapter.readBinary(path));
  }

  private async ensureLocalFolder(path: string): Promise<void> {
    let current = "";
    for (const part of path.split("/")) {
      current = current ? `${current}/${part}` : part;
      if (!(await this.app.vault.adapter.exists(current))) await this.app.vault.adapter.mkdir(current);
    }
  }

  private record(path: string, local: LocalItem, remote: RemoteItem): void {
    this.settings.entries[path] = {
      localMtime: local.mtime,
      localSize: local.size,
      remoteETag: remote.eTag,
      remoteMtime: remote.mtime,
      remoteSize: remote.size
    };
  }

  private excluded(path: string): boolean {
    const configDir = normalizePath(this.app.vault.configDir);
    if (path === `${configDir}/plugins/onedrive-bidirectional-sync/data.json`) return true;
    if (!this.settings.syncConfigDir && (path === configDir || path.startsWith(`${configDir}/`))) return true;
    const patterns = this.settings.excludedPatterns.split(/\r?\n/).map((p) => p.trim()).filter(Boolean);
    return patterns.some((pattern) => wildcard(pattern, path));
  }
}

function changedLocal(local: LocalItem, previous: SyncEntry): boolean {
  return local.size !== previous.localSize || Math.abs(local.mtime - previous.localMtime) >= 1000;
}

function wildcard(pattern: string, value: string): boolean {
  const regex = pattern
    .split("**")
    .map((part) => part.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, "[^/]*"))
    .join(".*");
  return new RegExp(`^${regex}$`).test(value);
}

function timestamp(): string {
  return new Date().toISOString().replace(/[:.]/g, "-");
}

function deviceLabel(): string {
  if (Platform.isIosApp) return "iOS";
  if (Platform.isAndroidApp) return "Android";
  if (Platform.isMacOS) return "macOS";
  if (Platform.isWin) return "Windows";
  if (Platform.isLinux) return "Linux";
  return Platform.isMobileApp ? "mobile" : "desktop";
}
