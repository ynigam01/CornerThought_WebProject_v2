"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.saveLessons = saveLessons;
const updateCompleteness_1 = require("./updateCompleteness");
const MAX_ATTACHMENT_FILE_BYTES = 10 * 1024 * 1024; // 10 MB per file
// Convert a decoded file buffer into the Postgres bytea hex format (\x...),
// matching the original browser-side arrayBufferToPgBytea helper.
function bufferToPgBytea(buffer) {
    return `\\x${buffer.toString('hex')}`;
}
function normalizeCategory(category) {
    return String(category || '').toLowerCase() === 'success' ? 'success' : 'issue';
}
function asTrimmedStrings(values) {
    if (!Array.isArray(values))
        return [];
    return values.map((value) => String(value ?? '').trim()).filter(Boolean);
}
// Accept either plain strings or { text, actions, lessons } objects from the client.
// Also recover if a previous buggy save stringified the whole object into the text field.
function normalizeCauseImpactItems(items) {
    if (!Array.isArray(items))
        return [];
    return items
        .map((item) => {
        if (typeof item === 'string') {
            const raw = item.trim();
            if (!raw)
                return null;
            if (raw.startsWith('{')) {
                try {
                    const parsed = JSON.parse(raw);
                    if (parsed && typeof parsed === 'object' && parsed.text != null) {
                        return {
                            text: String(parsed.text || '').trim(),
                            actions: asTrimmedStrings(parsed.actions),
                            lessons: asTrimmedStrings(parsed.lessons),
                        };
                    }
                }
                catch {
                    // Fall through and treat as plain text.
                }
            }
            return { text: raw, actions: [], lessons: [] };
        }
        if (!item || typeof item !== 'object')
            return null;
        const text = String(item.text || '').trim();
        if (!text)
            return null;
        return {
            text,
            actions: asTrimmedStrings(item.actions),
            lessons: asTrimmedStrings(item.lessons),
        };
    })
        .filter((item) => item != null && Boolean(item.text));
}
async function insertLinkedActionsAndLessons(supabase, args) {
    const { lessonId, parentId, parentKind, actions, lessons, userId, organizationId, projectId, } = args;
    const causeId = parentKind === 'cause' ? parentId : null;
    const impactId = parentKind === 'impact' ? parentId : null;
    if (actions.length) {
        const { error } = await supabase.from('action_items').insert(actions.map((actionItem) => ({
            lessons_learned_id: lessonId,
            action_item: actionItem,
            lessons_learned_cause_id: causeId,
            lessons_learned_impact_id: impactId,
            created_by: userId,
            organization_id: organizationId,
            project_id: projectId,
        })));
        if (error)
            throw new Error(error.message || 'Failed to save action items.');
    }
    if (lessons.length) {
        const { error } = await supabase.from('future_project_considerations').insert(lessons.map((fpc) => ({
            lessons_learned_id: lessonId,
            fpc,
            lessons_learned_cause_id: causeId,
            lessons_learned_impact_id: impactId,
            created_by: userId,
            organization_id: organizationId,
            project_id: projectId,
        })));
        if (error)
            throw new Error(error.message || 'Failed to save lessons learned items.');
    }
}
function splitMetadataLabel(label) {
    if (label.includes(':')) {
        return {
            metadataType: label.split(':')[0].trim(),
            metadata: label.split(':').slice(1).join(':').trim(),
        };
    }
    return { metadataType: null, metadata: label };
}
// Persists one Add Data submission (one or more lesson entries) to Supabase.
// Ported from saveLessonsLearned() in frontend/public/js/user-portal.js.
async function saveLessons(supabase, req) {
    const { userId, organizationId, projectId, projectTypeId, review, entries } = req;
    if (userId == null || organizationId == null) {
        throw new Error('Missing user information.');
    }
    if (projectId == null) {
        throw new Error('Select a project before saving.');
    }
    if (!Array.isArray(entries) || entries.length === 0) {
        throw new Error('No entries to save.');
    }
    const reviewForDb = review === 'draft' ? 'draft' : 'for review';
    let savedCount = 0;
    const lessonIds = [];
    for (const entry of entries) {
        const title = entry.title || '';
        const highLevelTitle = entry.highLevelTitle || '';
        const category = normalizeCategory(entry.category);
        const { data: lessonRows, error: lessonErr } = await supabase
            .from('lessons_learned')
            .insert({
            title,
            high_level_title: highLevelTitle || null,
            category,
            review: reviewForDb,
            share: '',
            created_by: userId,
            organization_id: organizationId,
            project_id: projectId,
            project_type_id: projectTypeId,
        })
            .select('id')
            .single();
        if (lessonErr || !lessonRows) {
            throw new Error(lessonErr?.message || 'Failed to save lesson.');
        }
        const lessonId = lessonRows.id;
        savedCount += 1;
        lessonIds.push(lessonId);
        const causes = normalizeCauseImpactItems(entry.causes);
        for (const cause of causes) {
            const { data: causeRow, error } = await supabase
                .from('lessons_learned_causes')
                .insert({
                lessons_learned_id: lessonId,
                cause: cause.text,
                created_by: userId,
                organization_id: organizationId,
                project_id: projectId,
            })
                .select('id')
                .single();
            if (error || !causeRow?.id) {
                throw new Error(error?.message || 'Failed to save causes.');
            }
            await insertLinkedActionsAndLessons(supabase, {
                lessonId,
                parentId: causeRow.id,
                parentKind: 'cause',
                actions: cause.actions,
                lessons: cause.lessons,
                userId,
                organizationId,
                projectId,
            });
        }
        const impacts = normalizeCauseImpactItems(entry.impacts);
        for (const impact of impacts) {
            const { data: impactRow, error } = await supabase
                .from('lessons_learned_impacts')
                .insert({
                lessons_learned_id: lessonId,
                impact: impact.text,
                created_by: userId,
                organization_id: organizationId,
                project_id: projectId,
            })
                .select('id')
                .single();
            if (error || !impactRow?.id) {
                throw new Error(error?.message || 'Failed to save impacts.');
            }
            await insertLinkedActionsAndLessons(supabase, {
                lessonId,
                parentId: impactRow.id,
                parentKind: 'impact',
                actions: impact.actions,
                lessons: impact.lessons,
                userId,
                organizationId,
                projectId,
            });
        }
        const notes = Array.isArray(entry.notes) ? entry.notes : [];
        if (notes.length) {
            const { error } = await supabase.from('lessons_learned_notes').insert(notes.map((notesText) => ({
                lessons_learned_id: lessonId,
                notes: notesText,
                created_by: userId,
                organization_id: organizationId,
                project_id: projectId,
            })));
            if (error)
                throw new Error(error.message || 'Failed to save notes.');
        }
        const metadataItems = Array.isArray(entry.metadataItems) ? entry.metadataItems : [];
        if (metadataItems.length) {
            const { error } = await supabase.from('lessons_learned_metadata').insert(metadataItems.map((item) => {
                const { metadataType, metadata } = splitMetadataLabel(item.label);
                return {
                    lessons_learned_id: lessonId,
                    metadata_type: metadataType,
                    metadata,
                    lessons_learned_metadata_list_id: item.id,
                    created_by: userId,
                    organization_id: organizationId,
                    project_id: projectId,
                    project_type_id: projectTypeId,
                };
            }));
            if (error)
                throw new Error(error.message || 'Failed to save metadata.');
        }
        const attachments = Array.isArray(entry.attachments) ? entry.attachments : [];
        if (attachments.length) {
            for (const file of attachments) {
                const savedAttachmentError = await saveAttachment(supabase, {
                    file,
                    lessonId,
                    projectId,
                    organizationId,
                    userId,
                });
                if (savedAttachmentError) {
                    throw new Error(savedAttachmentError);
                }
            }
        }
        await (0, updateCompleteness_1.updateCompleteness)(supabase, lessonId, organizationId);
    }
    return { savedCount, lessonIds };
}
async function saveAttachment(supabase, { file, lessonId, projectId, organizationId, userId }) {
    const fileName = file && file.fileName ? file.fileName : 'attachment';
    const base64 = file && file.base64 ? file.base64 : '';
    if (!base64) {
        return null;
    }
    const fileBuffer = Buffer.from(base64, 'base64');
    const fileSize = fileBuffer.length;
    if (!fileSize) {
        return null;
    }
    if (fileSize > MAX_ATTACHMENT_FILE_BYTES) {
        return `Attachment "${fileName}" exceeds the 10 MB limit.`;
    }
    const { error } = await supabase.from('lessons_learned_attachments').insert({
        lessons_learned_id: lessonId,
        project_id: projectId,
        organization_id: organizationId,
        created_by: userId,
        file_data: bufferToPgBytea(fileBuffer),
        file_name: fileName,
        content_type: file.contentType || 'application/octet-stream',
    });
    if (error) {
        return error.message || `Failed to save attachment "${fileName}".`;
    }
    return null;
}
//# sourceMappingURL=saveLessons.js.map