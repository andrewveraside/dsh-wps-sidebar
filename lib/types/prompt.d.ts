/**
 * The host's file-based Office skills this plugin stands in for, and the app of each. office-xlsx is left out on
 * purpose: spreadsheet work is mostly data analysis, which code does better, so that skill loads as usual.
 */
export declare const FILE_SKILLS: Record<string, 'word' | 'ppt'>;
/**
 * What the model reads the first time it loads one of those skills: this plugin's own short skill in its place, as
 * the result of the call (not a refusal: nothing went wrong, and the user should not see an error).
 */
export declare function redirectText(skill: string, app: 'word' | 'ppt'): string;
/** System-prompt section: when to reach for the office_ tools, in as few tokens as will do. */
export declare function promptText(): string;
