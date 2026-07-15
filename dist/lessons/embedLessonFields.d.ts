import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
/**
 * Embeds all searchable text fields for a lesson into their vector columns.
 * Always rewrites present text. Per-row failures are logged and skipped.
 */
export declare function embedLessonFields(supabase: SupabaseClient, lessonId: Id, organizationId: Id): Promise<void>;
//# sourceMappingURL=embedLessonFields.d.ts.map