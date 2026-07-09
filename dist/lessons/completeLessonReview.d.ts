import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
export interface CompleteLessonRequest {
    lessonId: Id;
    organizationId: Id;
    userId: Id;
}
export type CompleteLessonResult = {
    ok: true;
} | {
    ok: false;
    reason: 'insufficient';
};
export declare function completeLessonReview(supabase: SupabaseClient, req: CompleteLessonRequest): Promise<CompleteLessonResult>;
//# sourceMappingURL=completeLessonReview.d.ts.map