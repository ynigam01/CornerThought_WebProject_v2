"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.embedCompletedLessonsForProject = embedCompletedLessonsForProject;
const embedLessonFields_1 = require("./embedLessonFields");
/**
 * Finds completed lessons for a project that are missing any embeddings,
 * then fills only those missing fields (same field set as complete / find).
 * Calls onProgress for each streaming event when provided.
 */
async function embedCompletedLessonsForProject(supabase, req, onProgress) {
    const { organizationId, projectId } = req;
    if (organizationId == null || organizationId === '') {
        throw new Error('organizationId is required');
    }
    if (projectId == null || projectId === '') {
        throw new Error('projectId is required');
    }
    const emit = (event) => {
        if (onProgress)
            onProgress(event);
    };
    const { data: lessons, error: lessonsError } = await supabase
        .from('lessons_learned')
        .select('id, title, high_level_title, title_search_embedding, high_level_search_embedding')
        .eq('organization_id', organizationId)
        .eq('project_id', projectId)
        .eq('review', 'complete');
    if (lessonsError) {
        throw new Error(lessonsError.message || 'Failed to load completed lessons.');
    }
    const allCompleted = lessons || [];
    const totalCompleted = allCompleted.length;
    const needing = [];
    for (const lesson of allCompleted) {
        const needs = await (0, embedLessonFields_1.lessonNeedsMissingEmbeddings)(supabase, lesson.id, organizationId, lesson);
        if (needs)
            needing.push(lesson);
    }
    const total = needing.length;
    emit({ type: 'start', total });
    let processed = 0;
    let failed = 0;
    const skipped = totalCompleted - total;
    for (let i = 0; i < needing.length; i += 1) {
        const lesson = needing[i];
        const index = i + 1;
        try {
            await (0, embedLessonFields_1.embedLessonFields)(supabase, lesson.id, organizationId, {
                onlyMissing: true,
            });
            processed += 1;
            emit({
                type: 'lesson',
                index,
                total,
                lessonId: lesson.id,
                status: 'updated',
            });
        }
        catch (err) {
            failed += 1;
            const message = err instanceof Error ? err.message : String(err);
            console.error(`embedCompletedLessonsForProject: lesson id=${lesson.id} failed:`, message);
            emit({
                type: 'lesson',
                index,
                total,
                lessonId: lesson.id,
                status: 'failed',
            });
        }
    }
    const summary = { processed, skipped, failed, totalCompleted };
    emit({ type: 'done', ...summary });
    return summary;
}
//# sourceMappingURL=embedCompletedLessonsForProject.js.map