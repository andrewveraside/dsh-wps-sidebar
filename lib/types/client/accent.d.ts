/**
 * The accent colour shared by the copylee DSH plugins (docs/ui-spec.md).
 *
 * One choice — terracotta orange by default, blue or black — is kept in the
 * browser under a key every plugin reads, and painted as two CSS variables on
 * `<body>`, where the Host keeps its own theme variables. A plugin's styles
 * only ever write `var(--cl-accent)` / `var(--cl-accent-ink)`, so changing the
 * colour in any plugin's settings recolours all of them at once.
 *
 * This file is identical in every plugin; change it in all of them together.
 */
import * as React from 'react';
export type Accent = 'orange' | 'blue' | 'black';
export declare const ACCENTS: Record<Accent, {
    name: string;
    accent: string;
    ink: string;
}>;
export declare const ACCENT_IDS: Accent[];
export declare const DEFAULT_ACCENT: Accent;
/** Use these in styles; the fallbacks cover the moment before the first paint. */
export declare const ACCENT: string;
export declare const ACCENT_INK: string;
export declare function isAccent(value: unknown): value is Accent;
/** The choice made in this browser, or null while none has been made. */
export declare function storedAccent(): Accent | null;
export declare function readAccent(): Accent;
/** Choose the accent for every plugin. */
export declare function writeAccent(accent: Accent): void;
/** Paint the current accent and keep it current. Returns the undo for plugin disposal. */
export declare function installAccent(): () => void;
export declare function useAccent(): [Accent, (accent: Accent) => void];
/** A row of round swatches: one of them is always chosen. */
export declare function AccentPicker({ label, hint, names }: {
    label?: string;
    hint?: string;
    /** Colour names in the page's language; the built-in ones are Chinese. */
    names?: Partial<Record<Accent, string>>;
}): React.ReactElement;
