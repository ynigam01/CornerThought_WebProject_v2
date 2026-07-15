/**
 * Cosine similarity between two equal-length numeric vectors.
 * Returns 0 for empty, mismatched, or zero-magnitude inputs.
 */
export declare function cosineSimilarity(a: number[], b: number[]): number;
/**
 * Normalize a pgvector / JSON embedding value into a number[].
 */
export declare function parseEmbedding(value: unknown): number[] | null;
//# sourceMappingURL=cosineSimilarity.d.ts.map