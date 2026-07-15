import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
export interface EmbedLessonFieldsOptions {
    /** When true, skip fields that already have a non-null embedding. */
    onlyMissing?: boolean;
}
/**
 * Embeds searchable text fields for a lesson into their vector columns.
 * By default rewrites present text. With onlyMissing, skips fields that already have embeddings.
 */
export declare function embedLessonFields(supabase: SupabaseClient, lessonId: Id, organizationId: Id, options?: EmbedLessonFieldsOptions): Promise<void>;
/**
 * Returns true if this completed lesson has any embeddable text field with a null embedding.
 */
export declare function lessonNeedsMissingEmbeddings(supabase: SupabaseClient, lessonId: Id, organizationId: Id, lessonRow?: {
    title?: unknown;
    high_level_title?: unknown;
    title_search_embedding?: unknown;
    high_level_search_embedding?: unknown;
} | null): Promise<boolean>;
//# sourceMappingURL=embedLessonFields.d.ts.map