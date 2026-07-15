import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
import { embedLessonFields } from './embedLessonFields';

/**
 * Fire-and-forget friendly wrapper used when a lesson is marked complete.
 * Delegates to the shared embedLessonFields implementation.
 */
export async function embedLessonOnComplete(
  supabase: SupabaseClient,
  lessonId: Id,
  organizationId: Id,
): Promise<void> {
  await embedLessonFields(supabase, lessonId, organizationId);
}
