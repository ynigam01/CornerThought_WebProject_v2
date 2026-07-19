import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
export interface RankUpcomingTaskLessonsRequest {
    organizationId: Id;
    userId: Id;
    metadataListId: Id;
}
export interface RankUpcomingTaskLessonResult {
    lessonId: Id;
    projectId: Id | null;
    projectName: string | null;
    title: string | null;
    highLevelTitle: string | null;
    category: string | null;
    metadataSimilarity: number;
    score: number;
    components: {
        projectType: number;
        metadata: number;
        projectParams: number;
    };
}
export interface RankUpcomingTaskLessonsResponse {
    results: RankUpcomingTaskLessonResult[];
}
/**
 * Rank completed org lessons relevant to an upcoming assigned task
 * using metadata-list ↔ lesson-metadata cosine (>= 0.65) and
 * weighted project-type / metadata / project-params scores.
 */
export declare function rankUpcomingTaskLessons(supabase: SupabaseClient, req: RankUpcomingTaskLessonsRequest): Promise<RankUpcomingTaskLessonsResponse>;
//# sourceMappingURL=rankUpcomingTaskLessons.d.ts.map