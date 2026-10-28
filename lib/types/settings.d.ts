/**
 * Plugin settings shared by the host and the settings page: defaults and the
 * tolerant reader for the Loader entry config. No schemastery here so the
 * browser bundle stays small.
 */
/** The Loader entry id (settings namespace); see cordis.patch.yml. */
export declare const ENTRY_ID = "andrewveraside-wps-sidebar";
/** Which office suite documents are opened in: Microsoft Office, WPS Office, or whichever is there (Microsoft first). */
export type Suite = 'auto' | 'office' | 'wps';
export interface Settings {
    /** The suite a file is opened in when it is not open anywhere yet. A document that is open already is worked on where it is. */
    suite: Suite;
    /** Work in the background: no window, no card, text written at once; what was opened is saved and closed when the agent stops. */
    silent: boolean;
    /** Scroll the window to where the agent is working; off = silent, the user's view is left alone. */
    follow: boolean;
    /** Write text a few characters at a time, the way a person types. */
    typing: boolean;
    /** Show the small card with the two switches while the agent edits. */
    card: boolean;
    /** On save, show the model every page as it is now if it has not looked since its last edits. */
    finalCheck: boolean;
    /** Turn the model away from the host's file-based Office skills when the real app is installed. */
    preferLive: boolean;
    /** Bring the document's window to the front when office_open opens it. */
    showOnOpen: boolean;
    /** Width in pixels of the pictures office_render gives the model. */
    renderWidth: number;
}
export declare const DEFAULTS: Settings;
/** Unwrap `.volatile()` refs (`{ get() }`) and fall back to defaults for bad values. */
export declare function resolveConfig(raw: unknown): Settings;
