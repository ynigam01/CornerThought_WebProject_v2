"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.updateLessonReview = updateLessonReview;
const updateCompleteness_1 = require("./updateCompleteness");
// Handles the draft editor's "Save Draft" (review = 'draft') and
// "Send for Review" (review = 'for review') toolbar buttons.
// Ported from the review updates in my-projects-lesson-draft-editor.js.
async function updateLessonReview(supabase, req) {
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
    if (error)
        throw new Error(error.message || 'Could not update status.');
    await (0, updateCompleteness_1.updateCompleteness)(supabase, lessonId, organizationId);
    return { ok: true };
}
//# sourceMappingURL=updateLessonReview.js.map