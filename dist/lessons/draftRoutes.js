"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.registerDraftLessonRoutes = registerDraftLessonRoutes;
const updateLessonReview_1 = require("./updateLessonReview");
const updateCompleteness_1 = require("./updateCompleteness");
const draftEditor_1 = require("./draftEditor");
// Wraps a handler with uniform error handling, mirroring server.js style.
function run(res, label, fn) {
    return fn()
        .then((result) => {
        res.json(result);
    })
        .catch((err) => {
        console.error(`Draft lesson ${label} failed:`, err);
        res.status(500).json({ error: err?.message || 'Request failed.' });
    });
}
const BASE = '/api/draft-lessons';
function registerDraftLessonRoutes(app, supabase) {
    // Recompute and persist completeness_quality for a lesson
    app.post(`${BASE}/:lessonId/completeness`, (req, res) => run(res, 'completeness', async () => {
        const lessonId = req.params.lessonId;
        const organizationId = req.body?.organizationId;
        if (lessonId == null || organizationId == null) {
            throw new Error('Missing lesson or organization.');
        }
        const completenessQuality = await (0, updateCompleteness_1.updateCompleteness)(supabase, lessonId, organizationId);
        return { ok: true, completenessQuality };
    }));
    // Save Draft / Send for Review status update
    app.patch(`${BASE}/:lessonId/review`, (req, res) => run(res, 'review', () => (0, updateLessonReview_1.updateLessonReview)(supabase, {
        lessonId: req.params.lessonId,
        organizationId: req.body?.organizationId,
        userId: req.body?.userId,
        review: req.body?.review,
        forReviewCollaborative: req.body?.forReviewCollaborative,
    })));
    // Lesson title
    app.patch(`${BASE}/:lessonId/title`, (req, res) => run(res, 'title', () => (0, draftEditor_1.updateTitle)(supabase, {
        lessonId: req.params.lessonId,
        title: req.body?.title,
        organizationId: req.body?.organizationId,
        userId: req.body?.userId,
        forReviewCollaborative: req.body?.forReviewCollaborative,
    })));
    // Create a sub-item (cause/impact/action/fpc/note)
    app.post(`${BASE}/:lessonId/sub-items`, (req, res) => run(res, 'create sub-item', () => (0, draftEditor_1.createSubItem)(supabase, {
        lessonId: req.params.lessonId,
        kind: req.body?.kind,
        value: req.body?.value,
        organizationId: req.body?.organizationId,
        projectId: req.body?.projectId,
        userId: req.body?.userId,
    })));
    // Update a sub-item
    app.patch(`${BASE}/sub-items/:kind/:id`, (req, res) => run(res, 'update sub-item', () => (0, draftEditor_1.updateSubItem)(supabase, {
        kind: req.params.kind,
        id: req.params.id,
        value: req.body?.value,
        organizationId: req.body?.organizationId,
        userId: req.body?.userId,
        forReviewCollaborative: req.body?.forReviewCollaborative,
        rowCreatedBy: req.body?.rowCreatedBy,
    })));
    // Delete a sub-item (cause/impact also unassign dependents)
    app.delete(`${BASE}/sub-items/:kind/:id`, (req, res) => run(res, 'delete sub-item', () => (0, draftEditor_1.deleteSubItem)(supabase, {
        kind: req.params.kind,
        id: req.params.id,
        organizationId: req.body?.organizationId,
        projectId: req.body?.projectId,
        userId: req.body?.userId,
        forReviewCollaborative: req.body?.forReviewCollaborative,
        rowCreatedBy: req.body?.rowCreatedBy,
    })));
    // Drag-drop reassignment of an action/fpc to a cause/impact
    app.patch(`${BASE}/assignments/:kind/:id`, (req, res) => run(res, 'assignment', () => (0, draftEditor_1.reassignItem)(supabase, {
        kind: req.params.kind,
        id: req.params.id,
        causeId: req.body?.causeId ?? null,
        impactId: req.body?.impactId ?? null,
        organizationId: req.body?.organizationId,
        projectId: req.body?.projectId,
        userId: req.body?.userId,
        forReviewCollaborative: req.body?.forReviewCollaborative,
        rowCreatedBy: req.body?.rowCreatedBy,
    })));
    // Upload an attachment (base64)
    app.post(`${BASE}/:lessonId/attachments`, (req, res) => run(res, 'create attachment', () => (0, draftEditor_1.createAttachment)(supabase, {
        lessonId: req.params.lessonId,
        organizationId: req.body?.organizationId,
        projectId: req.body?.projectId,
        userId: req.body?.userId,
        fileName: req.body?.fileName,
        contentType: req.body?.contentType,
        base64: req.body?.base64,
    })));
    // Remove an attachment
    app.delete(`${BASE}/attachments/:id`, (req, res) => run(res, 'delete attachment', () => (0, draftEditor_1.deleteAttachment)(supabase, {
        id: req.params.id,
        organizationId: req.body?.organizationId,
        userId: req.body?.userId,
        forReviewCollaborative: req.body?.forReviewCollaborative,
        rowCreatedBy: req.body?.rowCreatedBy,
    })));
    // Apply selected metadata list ids to the lesson
    app.post(`${BASE}/:lessonId/metadata`, (req, res) => run(res, 'apply metadata', () => (0, draftEditor_1.applyMetadata)(supabase, {
        lessonId: req.params.lessonId,
        organizationId: req.body?.organizationId,
        projectId: req.body?.projectId,
        userId: req.body?.userId,
        projectTypeId: req.body?.projectTypeId ?? null,
        listIds: Array.isArray(req.body?.listIds) ? req.body.listIds : [],
    })));
    // Remove a metadata tag
    app.delete(`${BASE}/metadata/:id`, (req, res) => run(res, 'delete metadata', () => (0, draftEditor_1.deleteMetadata)(supabase, {
        id: req.params.id,
        organizationId: req.body?.organizationId,
        userId: req.body?.userId,
        forReviewCollaborative: req.body?.forReviewCollaborative,
        rowCreatedBy: req.body?.rowCreatedBy,
    })));
}
//# sourceMappingURL=draftRoutes.js.map