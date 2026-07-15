import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
export interface EmbedLessonForFindRequest {
    lessonId: Id;
    organizationId: Id;
}
export interface EmbedLessonForFindResult {
    ok: true;
}
/**
 * Awaits full lesson field embeddings before Find Relevant can run.
 * Separate from complete so Complete stays fire-and-forget / non-blocking.
 */
export declare function embedLessonForFind(supabase: SupabaseClient, req: EmbedLessonForFindRequest): Promise<EmbedLessonForFindResult>;
//# sourceMappingURL=embedLessonForFind.d.ts.map