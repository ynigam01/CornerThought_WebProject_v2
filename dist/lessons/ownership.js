"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.effectiveRowCreatedBy = effectiveRowCreatedBy;
exports.canMutateSubRow = canMutateSubRow;
exports.applyOwnerConstraint = applyOwnerConstraint;
// Ported from my-projects-lesson-draft-editor.js so the backend can enforce the
// exact same ownership rules the client uses to gate the editor UI.
function effectiveRowCreatedBy(createdBy, lessonCreatorId) {
    if (createdBy != null && String(createdBy).trim() !== '')
        return String(createdBy);
    if (lessonCreatorId != null)
        return String(lessonCreatorId);
    return '';
}
function canMutateSubRow(createdBy, lessonCreatorId, userId) {
    if (userId == null)
        return false;
    return effectiveRowCreatedBy(createdBy, lessonCreatorId) === String(userId);
}
// Mirrors applyReviewOwnerConstraint from the editor: in collaborative "for review"
// mode, scope a mutation to rows the requesting user created.
function applyOwnerConstraint(query, ctx) {
    if (ctx.forReviewCollaborative && ctx.rowCreatedBy != null) {
        return query.eq('created_by', ctx.userId);
    }
    return query;
}
//# sourceMappingURL=ownership.js.map