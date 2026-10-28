/** One formula as an `<m:oMath>` element, or undefined when the LaTeX cannot be read. */
export declare function latexToOmml(tex: string, display: boolean): string | undefined;
/**
 * Every formula written between dollar signs anywhere in a batch of operations, converted: keyed "d:" + source for
 * one on a line of its own and "i:" + source for one inside a line, the source as the helper will read it.
 */
export declare function formulasOf(value: unknown, found?: Record<string, string>): Record<string, string>;
