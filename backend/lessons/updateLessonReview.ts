import type { SupabaseClient } from '@supabase/supabase-js';
import type { UpdateLessonReviewRequest } from './types';

// Handles the draft editor's "Save Draft" (review = 'draft') and
// "Send for Review" (review = 'for review') toolbar buttons.
// Ported from the review updates in my-projects-lesson-draft-editor.js.
export async function updateLessonReview(
  supabase: SupabaseClient,
  req: UpdateLessonReviewRequest,
): Promise<{ ok: true }> {
  const { lessonId, organizationId, userId, review, forReviewCollaborative } = req;

  if (lessonId == null || organizationId == null) {
    throw new Error('Missing lesson or organization.');
  }
  const reviewForDb = review === 'draft' ? 'draft' : 'for review';

  let query = supabase
    .from('lessons_learned')
    .update({ review: reviewForDb })
    .eq('id', lessonId)
    .eq('organization_id', organizationId);

  if (forReviewCollaborative) {
    query = query.eq('created_by', userId);
  }

  const { error } = await query;
  if (error) throw new Error(error.message || 'Could not update status.');
  return { ok: true };
}
