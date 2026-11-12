import type { UIMessage } from 'ai';

/**
 * The shape of the preload bridge. Types, plus the one sentinel string all three
 * processes must agree on. Nothing else belongs here: this file is imported by
 * main, preload and the renderer, so anything with behaviour would be running in
 * three places at once.
 *
 * It is deliberately a closed list of named operations. There is no
 * `invoke(channel, ...args)`, which would be `ipcRenderer` wearing a hat, and
 * there is no operation that returns an API key. The renderer can say "store
 * this key" and ask "is a key stored?"; it can never read one back, and it never
 * sees a provider response header.
 */

/**
 * Emitted by the preload bridge when the stream's MessagePort closes without a
 * terminal chunk — main crashed, or the window is tearing down. Not an AI SDK
 * chunk type; the transport consumes it and never forwards it.
 */
export const PORT_CLOSED = '__mutatis/port-closed';

export type ProviderId = 'deepseek' | 'alibaba';

export interface ModelChoice {
  readonly provider: ProviderId;
  readonly model: string;
}

export interface ProviderInfo {
  readonly id: ProviderId;
  readonly label: string;
  readonly models: readonly { readonly id: string; readonly label: string }[];
  readonly keyUrl: string;
}

export interface ProvidersSnapshot {
  readonly providers: readonly ProviderInfo[];
  readonly defaultModel: ModelChoice;
  /**
   * False when the OS keychain is unavailable. We refuse to store a key at all
   * in that case rather than falling back to plaintext, so the UI has to say so.
   */
  readonly encryptionAvailable: boolean;
}

export type KeyStatus = Record<ProviderId, { readonly configured: boolean; readonly mask: string }>;

export interface SetKeyResult {
  readonly ok: boolean;
  readonly mask?: string;
  readonly error?: string;
}

export interface StoredThread {
  readonly id: string;
  readonly title: string;
  readonly createdAt: number;
  readonly updatedAt: number;
}

export interface StoredMessage {
  readonly id: string;
  readonly threadId: string;
  readonly role: 'user' | 'assistant' | 'system';
  readonly parts: unknown;
  readonly createdAt: number;
}

export interface LoadedThread {
  readonly messages: StoredMessage[];
  /** Rows that could not be parsed. Reported to the user, never silently dropped. */
  readonly skipped: number;
}

export interface ChatTurnRequest {
  readonly messages: readonly UIMessage[];
  readonly model: ModelChoice;
  readonly system?: string;
}

/**
 * One stream chunk. This is the AI SDK's own `UIMessageChunk` union, kept as a
 * loose record here so the type does not have to be re-declared in three places
 * — the renderer transport narrows it back to `UIMessageChunk` at the one point
 * where that matters.
 */
export interface ChatChunk {
  readonly type: string;
  readonly [key: string]: unknown;
}

/**
 * Where an update check has got to.
 *
 * `unsupported` is the DEFAULT and the interesting one: it is not an error
 * state, it means this build has no update path, and `blockers` says why in
 * words a non-developer can act on. An updater that silently does nothing is
 * indistinguishable from an updater that is broken.
 */
export type UpdaterPhase =
  | 'unsupported'
  | 'idle'
  | 'checking'
  | 'up-to-date'
  | 'available'
  | 'downloading'
  | 'downloaded'
  | 'error';

export interface UpdaterState {
  readonly phase: UpdaterPhase;
  /** The version running right now. */
  readonly version: string;
  readonly newVersion?: string;
  /** 0-100 while downloading. */
  readonly percent?: number;
  readonly message?: string;
  /** Plain-language reasons this build cannot update itself. Empty when it can. */
  readonly blockers: readonly string[];
}

export interface MutatisApi {
  readonly platform: string;

  providers(): Promise<ProvidersSnapshot | null>;

  readonly secrets: {
    status(): Promise<KeyStatus | null>;
    set(provider: ProviderId, key: string): Promise<SetKeyResult>;
  };

  readonly threads: {
    list(): Promise<StoredThread[]>;
    load(id: string): Promise<LoadedThread>;
    ensure(id: string, title: string): Promise<boolean>;
    rename(id: string, title: string): Promise<boolean>;
    save(message: StoredMessage): Promise<boolean>;
    remove(id: string): Promise<boolean>;
  };

  /**
   * Application updates.
   *
   * Every one of these is user-initiated. Nothing here runs at startup, nothing
   * polls, and `download` is a separate call from `check` — see
   * src/main/updater.ts for why an app that reads legal documents does not
   * quietly phone home.
   */
  readonly updates: {
    state(): Promise<UpdaterState>;
    check(): Promise<UpdaterState>;
    download(): Promise<UpdaterState>;
    install(): Promise<boolean>;
    /** Subscribe to progress. Returns an unsubscribe function. */
    onState(listener: (state: UpdaterState) => void): () => void;
  };

  /** Opens a provider's own API-key page in the system browser. Allow-listed in main. */
  openKeyPage(url: string): Promise<boolean>;

  /**
   * Start one chat turn. Returns the stop function: calling it closes the
   * MessagePort, which aborts the upstream HTTP request in main. There is no
   * separate cancel message that could be lost or arrive out of order.
   */
  startChat(request: ChatTurnRequest, onChunk: (chunk: ChatChunk) => void): () => void;
}

declare global {
  interface Window {
    readonly mutatis: MutatisApi;
  }
}
