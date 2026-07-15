import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
export interface RankMatchInput {
    number?: number | string;
    name?: string;
    entry?: string;
}
export interface RankRelevantLessonsRequest {
    lessonId: Id;
    organizationId: Id;
    projectId: Id;
    matches?: RankMatchInput[];
    projectDetailsNumberToIdMap?: Record<string, string | number>;
}
export interface RankRelevantLessonResult {
    lessonId: Id;
    projectId: Id | null;
    title: string | null;
    highLevelTitle: string | null;
    category: string | null;
    highLevelSimilarity: number;
    score: number;
    components: {
        projectType: number;
        metadata: number;
        projectParams: number;
    };
}
export interface RankRelevantLessonsResponse {
    matches: RankMatchInput[];
    results: RankRelevantLessonResult[];
}
/**
 * Rank completed org lessons relevant to the current lesson using
 * high_level shortlist (>= 0.7) and weighted project-type / metadata / params boosts.
 */
export declare function rankRelevantLessons(supabase: SupabaseClient, req: RankRelevantLessonsRequest): Promise<RankRelevantLessonsResponse>;
//# sourceMappingURL=rankRelevantLessons.d.ts.map