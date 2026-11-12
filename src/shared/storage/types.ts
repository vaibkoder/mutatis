/** A simple async key-value store with export/import. */
export interface Store<T> {
  get(key: string): Promise<T | undefined>;
  set(key: string, value: T): Promise<void>;
  delete(key: string): Promise<void>;
  list(prefix?: string): Promise<string[]>;
  /** Export all data as a JSON Blob for user download. */
  export(): Promise<Blob>;
  /** Import from a JSON Blob (merges; does not clear existing entries). */
  import(blob: Blob): Promise<void>;
}

/** Pluggable remote adapter (GitHub / SharePoint). Carries only user-authored library data. */
export interface RemoteAdapter {
  readonly id: string;
  isConfigured(): boolean;
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
  list(prefix?: string): Promise<string[]>;
}

export interface StorageProvider {
  local: Store<unknown>;
  docSettings: Store<unknown>;
  remote?: RemoteAdapter;
}
