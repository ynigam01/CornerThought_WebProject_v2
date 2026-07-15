"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.embedLessonForFind = embedLessonForFind;
const embedLessonFields_1 = require("./embedLessonFields");
/**
 * Awaits full lesson field embeddings before Find Relevant can run.
 * Separate from complete so Complete stays fire-and-forget / non-blocking.
 */
async function embedLessonForFind(supabase, req) {
    const { lessonId, organizationId } = req;
    if (lessonId == null || organizationId == null) {
        throw new Error('Missing lesson or organization.');
    }
    await (0, embedLessonFields_1.embedLessonFields)(supabase, lessonId, organizationId);
    return { ok: true };
}
//# sourceMappingURL=embedLessonForFind.js.map