import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
/**
 * Embeds all searchable text fields for a completed lesson into their
 * respective vector columns. Safe to fire-and-forget; never throws to callers
 * that catch — individual row failures are logged and skipped.
 */
export declare function embedLessonOnComplete(supabase: SupabaseClient, lessonId: Id, organizationId: Id): Promise<void>;
//# sourceMappingURL=embedLessonOnComplete.d.ts.map