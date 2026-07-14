"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.completeLessonReview = completeLessonReview;
const updateCompleteness_1 = require("./updateCompleteness");
const embedLessonOnComplete_1 = require("./embedLessonOnComplete");
// Recomputes completeness_quality, then marks the lesson as complete
// (review = 'complete', share = 'share with organization') if the quality
// meets or exceeds 'minimum'. Returns { ok: false, reason: 'insufficient' }
// when the lesson does not have enough detail to be completed.
async function completeLessonReview(supabase, req) {
    const { lessonId, organizationId } = req;
    if (lessonId == null || organizationId == null) {
        throw new Error('Missing lesson or organization.');
    }
    const quality = await (0, updateCompleteness_1.updateCompleteness)(supabase, lessonId, organizationId);
    if (quality === null) {
        return { ok: false, reason: 'insufficient' };
    }
    const { error } = await supabase
        .from('lessons_learned')
        .update({ review: 'complete', share: 'share with organization' })
        .eq('id', lessonId)
        .eq('organization_id', organizationId);
    if (error)
        throw new Error(error.message || 'Could not complete lesson.');
    void (0, embedLessonOnComplete_1.embedLessonOnComplete)(supabase, lessonId, organizationId).catch((err) => {
        console.error('embedLessonOnComplete failed:', err);
    });
    return { ok: true };
}
//# sourceMappingURL=completeLessonReview.js.map