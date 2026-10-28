/**
 * What the helper reports, turned into the short plain text the model reads.
 * Kept terse: every line here is paid for on each later request of the turn.
 */
export type AppKind = 'word' | 'excel' | 'ppt';
export declare const APP_NAMES: Record<AppKind, string>;
export declare const WPS_NAMES: Record<AppKind, string>;
/** The app as the model should name it: WPS has its own names, the tools and operations are the same. */
export declare const appName: (on: {
    app: AppKind;
    suite?: string;
}) => string;
export interface DocInfo {
    app: AppKind;
    suite?: string;
    name: string;
    path: string | null;
    saved: boolean;
    active: boolean;
    readOnly?: boolean;
    how?: 'attached' | 'opened' | 'created';
    background?: boolean;
}
export interface AppStatus {
    app: AppKind;
    suite?: string;
    installed: boolean;
    running: boolean;
    version?: string;
    error?: string;
    documents: DocInfo[];
}
export interface WordItem {
    i: number;
    text?: string;
    style?: string;
    level?: number;
    to?: number;
    table?: number;
    size?: string;
    open?: boolean;
    inTable?: boolean;
}
export interface WordFootnote {
    i: number;
    label: string | null;
    numberStyle: number;
    para: number;
    offset: number;
    page: number;
    section: number;
    text: string;
}
export interface WordRead {
    paragraphs: number;
    tables: number;
    pages?: number;
    from?: number;
    items?: WordItem[];
    footnotes?: WordFootnote[];
    more?: string;
    table?: number;
    cells?: Array<Array<string | null>>;
}
export interface ExcelRead {
    sheets: Array<{
        name: string;
        used: string;
        charts?: number;
    }>;
    sheet: string;
    range: string;
    values: unknown[][];
    formulas?: Array<{
        cell: string;
        formula: string;
    }>;
    clipped?: string;
    objects?: string[];
    profile?: string[];
}
export interface PptRead {
    slides: number;
    width: number;
    height: number;
    items: Array<{
        slide: number;
        layout?: string;
        plain?: number;
        shapes: Array<{
            name: string;
            id?: number;
            type: string;
            box: number[];
            text?: string;
        }>;
    }>;
}
export interface EditResult {
    done: string[];
    total: number;
    failed?: {
        index: number;
        op: string;
        error: string;
        code: string;
    };
}
export declare function formatStatus(apps: AppStatus[]): string;
export declare function formatOpened(doc: DocInfo): string;
export declare function formatWord(read: WordRead): string;
export declare function formatExcel(read: ExcelRead): string;
export declare function formatPpt(read: PptRead): string;
export declare function formatEdit(result: EditResult): string;
