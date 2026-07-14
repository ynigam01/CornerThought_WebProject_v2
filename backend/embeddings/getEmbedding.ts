import { InferenceClient } from '@huggingface/inference';

const EMBEDDING_MODEL = 'sentence-transformers/all-MiniLM-L6-v2';
const EMBEDDING_DIMS = 384;

let hfClient: InferenceClient | null = null;

function getHfClient(): InferenceClient {
  if (hfClient) return hfClient;
  const token = process.env.HF_API_TOKEN;
  if (!token) {
    throw new Error('HF_API_TOKEN is not set.');
  }
  hfClient = new InferenceClient(token);
  return hfClient;
}

/**
 * Embed text with sentence-transformers/all-MiniLM-L6-v2 (384 dims).
 * Empty input returns a zero vector (same behavior as server.js).
 */
export async function getEmbedding(text: string): Promise<number[]> {
  const trimmed = String(text || '').trim();
  if (!trimmed) {
    return Array(EMBEDDING_DIMS).fill(0);
  }

  const raw = await getHfClient().featureExtraction({
    model: EMBEDDING_MODEL,
    inputs: trimmed,
  });

  const embedding = Array.isArray(raw[0]) ? (raw[0] as number[]) : (raw as number[]);

  if (!Array.isArray(embedding)) {
    throw new Error(`Unexpected embedding format from HF: ${JSON.stringify(raw)}`);
  }

  return embedding;
}
