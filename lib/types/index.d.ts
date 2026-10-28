import type { Context } from '@deepseek-ai/cordis';
import z from '@deepseek-ai/schemastery';
export { DEFAULTS, ENTRY_ID, resolveConfig } from './settings.js';
export declare const name = "@andrewveraside/dsh-wps-sidebar";
export declare const inject: string[];
export declare const STATUS_ROUTE = "/api/dsh-office/status";
export interface Config {
    suite?: string;
    silent?: boolean;
    follow?: boolean;
    typing?: boolean;
    card?: boolean;
    finalCheck?: boolean;
    preferLive?: boolean;
    showOnOpen?: boolean;
    renderWidth?: number;
}
export declare const Config: z<Config>;
export declare function apply(ctx: Context, config?: Config): void;
