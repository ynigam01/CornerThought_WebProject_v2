import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
/**
 * Fire-and-forget friendly wrapper used when a lesson is marked complete.
 * Delegates to the shared embedLessonFields implementation.
 */
export declare function embedLessonOnComplete(supabase: SupabaseClient, lessonId: Id, organizationId: Id): Promise<void>;
//# sourceMappingURL=embedLessonOnComplete.d.ts.map