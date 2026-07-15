import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
export interface EmbedCompletedLessonsForProjectRequest {
    organizationId: Id;
    projectId: Id;
}
export type EmbedProjectLessonProgressEvent = {
    type: 'start';
    total: number;
} | {
    type: 'lesson';
    index: number;
    total: number;
    lessonId: Id;
    status: 'updated' | 'failed';
} | {
    type: 'done';
    processed: number;
    skipped: number;
    failed: number;
    totalCompleted: number;
} | {
    type: 'error';
    message: string;
};
/**
 * Finds completed lessons for a project that are missing any embeddings,
 * then fills only those missing fields (same field set as complete / find).
 * Calls onProgress for each streaming event when provided.
 */
export declare function embedCompletedLessonsForProject(supabase: SupabaseClient, req: EmbedCompletedLessonsForProjectRequest, onProgress?: (event: EmbedProjectLessonProgressEvent) => void): Promise<{
    processed: number;
    skipped: number;
    failed: number;
    totalCompleted: number;
}>;
//# sourceMappingURL=embedCompletedLessonsForProject.d.ts.map