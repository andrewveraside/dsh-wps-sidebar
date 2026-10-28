import type { ImageAttachmentRef } from '@deepseek-ai/dsh-attachment';
import { defineTool } from '@deepseek-ai/dsh-tools';
import { type AppKind } from './format.js';
import type { HelperLike } from './helper-client.js';
import type { Settings } from './settings.js';
type ToolDefinition = ReturnType<typeof defineTool>;
export interface ToolHost {
    helper: HelperLike;
    settings(): Settings;
    /** Persist a picture so the model can see it. */
    saveImage(data: Buffer, name: string): Promise<ImageAttachmentRef>;
    /** Whether the conversation model can look at images. */
    vision(exec: unknown): Promise<boolean>;
}
interface Target {
    app?: string;
    doc?: string;
}
/** Which app a call is about: said outright, or told from the file name. */
export declare function appOf(args: Target & {
    path?: string;
}): AppKind;
/** The operations of each app. Not sent with every request: handed over when the app is first used, and by office_help. */
export declare const GUIDES: Record<AppKind, string>;
export declare function createTools(host: ToolHost): ToolDefinition[];
export {};
