export interface HelperEvent {
    event: string;
    reason?: string;
}
export interface HelperLike {
    call<T = unknown>(cmd: string, args?: Record<string, unknown>, timeoutMs?: number): Promise<T>;
    onEvent(listener: (event: HelperEvent) => void): () => void;
    dispose(): void;
}
/** A file in helper/ at the package root: one level up from lib/ or src/, two from lib/types/. */
export declare function helperAssetPath(name: string): string;
export declare function helperSourcePath(): string;
/** Compile the helper if this source version has no cached exe yet; returns the exe path. */
export declare function ensureHelperExe(source?: string): Promise<string>;
export declare class HelperClient implements HelperLike {
    private readonly log;
    private child;
    private starting;
    private readonly pending;
    private readonly listeners;
    private nextId;
    private disposed;
    constructor(log?: (message: string) => void);
    onEvent(listener: (event: HelperEvent) => void): () => void;
    call<T = unknown>(cmd: string, args?: Record<string, unknown>, timeoutMs?: number): Promise<T>;
    dispose(): void;
    private kill;
    private ensure;
    private start;
    private handle;
}
