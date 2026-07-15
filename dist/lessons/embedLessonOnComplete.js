"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.embedLessonOnComplete = embedLessonOnComplete;
const embedLessonFields_1 = require("./embedLessonFields");
/**
 * Fire-and-forget friendly wrapper used when a lesson is marked complete.
 * Delegates to the shared embedLessonFields implementation.
 */
async function embedLessonOnComplete(supabase, lessonId, organizationId) {
    await (0, embedLessonFields_1.embedLessonFields)(supabase, lessonId, organizationId);
}
//# sourceMappingURL=embedLessonOnComplete.js.map