export interface TokenState {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

export interface SyncEntry {
  localMtime: number;
  localSize: number;
  remoteETag: string;
  remoteMtime: number;
  remoteSize: number;
}

export interface SyncSettings {
  clientId: string;
  tenant: string;
  vaultId: string;
  intervalMinutes: number;
  syncOnStartup: boolean;
  syncConfigDir: boolean;
  excludedPatterns: string;
  token: TokenState | null;
  entries: Record<string, SyncEntry>;
  lastSyncAt: number;
}

export interface RemoteItem {
  id: string;
  name: string;
  path: string;
  eTag: string;
  mtime: number;
  size: number;
  folder: boolean;
}

export interface LocalItem {
  path: string;
  mtime: number;
  size: number;
}

export interface SyncSummary {
  uploaded: number;
  downloaded: number;
  deletedLocal: number;
  deletedRemote: number;
  conflicts: number;
  skipped: number;
}
