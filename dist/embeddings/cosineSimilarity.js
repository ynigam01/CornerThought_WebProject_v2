"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.cosineSimilarity = cosineSimilarity;
exports.parseEmbedding = parseEmbedding;
/**
 * Cosine similarity between two equal-length numeric vectors.
 * Returns 0 for empty, mismatched, or zero-magnitude inputs.
 */
function cosineSimilarity(a, b) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length === 0 || a.length !== b.length) {
        return 0;
    }
    let dot = 0;
    let normA = 0;
    let normB = 0;
    for (let i = 0; i < a.length; i += 1) {
        const x = Number(a[i]) || 0;
        const y = Number(b[i]) || 0;
        dot += x * y;
        normA += x * x;
        normB += y * y;
    }
    if (normA <= 0 || normB <= 0)
        return 0;
    return dot / (Math.sqrt(normA) * Math.sqrt(normB));
}
/**
 * Normalize a pgvector / JSON embedding value into a number[].
 */
function parseEmbedding(value) {
    if (value == null)
        return null;
    if (Array.isArray(value)) {
        const nums = value.map((v) => Number(v));
        if (nums.some((n) => !Number.isFinite(n)))
            return null;
        return nums.length ? nums : null;
    }
    if (typeof value === 'string') {
        const trimmed = value.trim();
        if (!trimmed)
            return null;
        try {
            const parsed = JSON.parse(trimmed);
            return parseEmbedding(parsed);
        }
        catch {
            // pgvector sometimes returns "{1,2,3}" style
            if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
                const inner = trimmed.slice(1, -1).trim();
                if (!inner)
                    return null;
                const nums = inner.split(',').map((p) => Number(p.trim()));
                if (nums.some((n) => !Number.isFinite(n)))
                    return null;
                return nums.length ? nums : null;
            }
            return null;
        }
    }
    return null;
}
//# sourceMappingURL=cosineSimilarity.js.map