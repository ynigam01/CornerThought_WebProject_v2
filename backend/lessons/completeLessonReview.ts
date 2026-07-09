import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
import { updateCompleteness } from './updateCompleteness';

export interface CompleteLessonRequest {
  lessonId: Id;
  organizationId: Id;
  userId: Id;
}

export type CompleteLessonResult =
  | { ok: true }
  | { ok: false; reason: 'insufficient' };

// Recomputes completeness_quality, then marks the lesson as complete
// (review = 'complete', share = 'share with organization') if the quality
// meets or exceeds 'minimum'. Returns { ok: false, reason: 'insufficient' }
// when the lesson does not have enough detail to be completed.
export async function completeLessonReview(
  supabase: SupabaseClient,
  req: CompleteLessonRequest,
): Promise<CompleteLessonResult> {
  const { lessonId, organizationId } = req;

  if (lessonId == null || organizationId == null) {
    throw new Error('Missing lesson or organization.');
  }

  const quality = await updateCompleteness(supabase, lessonId, organizationId);

  if (quality === null) {
    return { ok: false, reason: 'insufficient' };
  }

  const { error } = await supabase
    .from('lessons_learned')
    .update({ review: 'complete', share: 'share with organization' })
    .eq('id', lessonId)
    .eq('organization_id', organizationId);

  if (error) throw new Error(error.message || 'Could not complete lesson.');

  return { ok: true };
}
