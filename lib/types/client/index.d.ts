type RemoteResult<T> = {
    ok: true;
    value: T;
} | {
    ok: false;
    error: {
        code?: string;
        message: string;
    };
};
interface NamespaceView {
    ns: string;
    value: unknown;
    revision: number;
}
interface ClientContext {
    effect(execute: () => (() => void) | void, label?: string): void;
    slots: {
        inject(name: string, register: () => unknown): void;
        register(meta: Record<string, unknown>, render: (props: unknown) => unknown): unknown;
    };
    remote: {
        $on?(event: string, listener: (...args: unknown[]) => void): () => void;
        settings: {
            describe(): Promise<RemoteResult<{
                writable: boolean;
                namespaces: NamespaceView[];
            }>>;
            mutate(ns: string, ops: {
                op: 'set' | 'unset';
                path: string[];
                value?: unknown;
            }[], expectedRevision?: number): Promise<RemoteResult<unknown>>;
        };
    };
}
export declare const inject: string[];
export declare function apply(ctx: ClientContext): void;
export {};
