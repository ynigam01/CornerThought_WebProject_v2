/**
 * Lessons learned review: populate lessons_learned_review_notifications from metadata tags
 * and project_team_member_assignments.
 *
 * Requires Supabase RLS policies that allow the portal user's role to SELECT the source tables
 * and INSERT into `lessons_learned_review_notifications` (and read assignments/metadata as used below).
 *
 * For the welcome / notifications list UI, the same user needs SELECT on their
 * `lessons_learned_review_notifications` rows plus read access to `lessons_learned` and `projects`
 * as used by fetchReviewNotificationsForUserGrouped.
 */

const IN_CHUNK = 250;

function chunkArray(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) {
        out.push(arr.slice(i, i + size));
    }
    return out;
}

/**
 * @param {{
 *   supabase: import('@supabase/supabase-js').SupabaseClient,
 *   lessonsLearnedId: string | number,
 *   projectId: string | number,
 *   organizationId: string | number,
 * }} args
 * @returns {Promise<{ inserted: number, error: Error | null }>}
 */
export async function createLessonsLearnedReviewNotifications({
    supabase,
    lessonsLearnedId,
    projectId,
    organizationId,
}) {
    if (!supabase || lessonsLearnedId == null || projectId == null || organizationId == null) {
        return {
            inserted: 0,
            error: new Error('Missing supabase client or lesson/project/organization id.'),
        };
    }

    const { data: metaRows, error: metaErr } = await supabase
        .from('lessons_learned_metadata')
        .select('lessons_learned_metadata_list_id')
        .eq('lessons_learned_id', lessonsLearnedId)
        .eq('organization_id', organizationId)
        .eq('project_id', projectId);

    if (metaErr) {
        return { inserted: 0, error: new Error(metaErr.message || 'Failed to load lesson metadata.') };
    }

    const rawListIds = new Set(
        (metaRows || [])
            .map((r) => r && r.lessons_learned_metadata_list_id)
            .filter((v) => v != null)
    );

    if (rawListIds.size === 0) {
        return { inserted: 0, error: null };
    }

    const listIdList = Array.from(rawListIds);
    const validated = new Set();

    for (const chunk of chunkArray(listIdList, IN_CHUNK)) {
        const { data: listRows, error: listErr } = await supabase
            .from('lessons_learned_metadata_list')
            .select('id')
            .eq('organization_id', organizationId)
            .eq('project_id', projectId)
            .in('id', chunk);

        if (listErr) {
            return {
                inserted: 0,
                error: new Error(listErr.message || 'Failed to validate metadata list ids.'),
            };
        }
        (listRows || []).forEach((r) => {
            if (r && r.id != null) validated.add(r.id);
        });
    }

    if (validated.size === 0) {
        return { inserted: 0, error: null };
    }

    const validatedList = Array.from(validated);
    const byUserId = new Map();

    for (const chunk of chunkArray(validatedList, IN_CHUNK)) {
        const { data: assignRows, error: assignErr } = await supabase
            .from('project_team_member_assignments')
            .select('user_id, project_id, organization_id')
            .eq('organization_id', organizationId)
            .eq('project_id', projectId)
            .in('lessons_learned_metadata_list_id', chunk);

        if (assignErr) {
            return {
                inserted: 0,
                error: new Error(assignErr.message || 'Failed to load team assignments.'),
            };
        }

        (assignRows || []).forEach((row) => {
            const uid = row && row.user_id;
            if (uid == null) return;
            const key = String(uid);
            if (!byUserId.has(key)) {
                byUserId.set(key, {
                    user_id: uid,
                    project_id: row.project_id != null ? row.project_id : projectId,
                    organization_id:
                        row.organization_id != null ? row.organization_id : organizationId,
                });
            }
        });
    }

    if (byUserId.size === 0) {
        return { inserted: 0, error: null };
    }

    const now = new Date().toISOString();
    const payload = Array.from(byUserId.values()).map((row) => ({
        lessons_learned_id: lessonsLearnedId,
        user_id: row.user_id,
        project_id: row.project_id,
        organization_id: row.organization_id,
        notified: false,
        created_at: now,
        updated_at: now,
    }));

    let inserted = 0;
    for (const chunk of chunkArray(payload, IN_CHUNK)) {
        const { error: insErr } = await supabase
            .from('lessons_learned_review_notifications')
            .insert(chunk);
        if (insErr) {
            return {
                inserted,
                error: new Error(insErr.message || 'Failed to insert review notifications.'),
            };
        }
        inserted += chunk.length;
    }

    return { inserted, error: null };
}

/**
 * @param {{
 *   supabase: import('@supabase/supabase-js').SupabaseClient,
 *   userId: string | number,
 *   organizationId: string | number,
 * }} args
 * @returns {Promise<{
 *   groups: Array<{ projectId: string | number, projectName: string, lessons: Array<{ id: unknown, title: string }> }>,
 *   error: Error | null,
 * }>}
 */
export async function fetchReviewNotificationsForUserGrouped({
    supabase,
    userId,
    organizationId,
}) {
    if (!supabase || userId == null || organizationId == null) {
        return {
            groups: [],
            error: new Error('Missing supabase client, user id, or organization id.'),
        };
    }

    const { data: notifRows, error: notifErr } = await supabase
        .from('lessons_learned_review_notifications')
        .select('lessons_learned_id, project_id')
        .eq('user_id', userId)
        .eq('organization_id', organizationId);

    if (notifErr) {
        return {
            groups: [],
            error: new Error(notifErr.message || 'Failed to load review notifications.'),
        };
    }

    const lessonIds = [
        ...new Set(
            (notifRows || [])
                .map((r) => r && r.lessons_learned_id)
                .filter((v) => v != null)
        ),
    ];

    if (lessonIds.length === 0) {
        return { groups: [], error: null };
    }

    const lessonById = new Map();
    for (const chunk of chunkArray(lessonIds, IN_CHUNK)) {
        const { data: lessonRows, error: leErr } = await supabase
            .from('lessons_learned')
            .select('id, title, project_id')
            .in('id', chunk)
            .eq('review', 'for review');

        if (leErr) {
            return {
                groups: [],
                error: new Error(leErr.message || 'Failed to load lessons for review.'),
            };
        }
        (lessonRows || []).forEach((row) => {
            if (row && row.id != null) lessonById.set(String(row.id), row);
        });
    }

    /** @type {Map<string, Map<string, { id: unknown, title: string }>>} */
    const projectToLessons = new Map();

    for (const nRow of notifRows || []) {
        const lid = nRow && nRow.lessons_learned_id;
        if (lid == null) continue;
        const lesson = lessonById.get(String(lid));
        if (!lesson) continue;
        const pid =
            lesson.project_id != null
                ? lesson.project_id
                : nRow.project_id != null
                  ? nRow.project_id
                  : null;
        if (pid == null) continue;
        const pidKey = String(pid);
        if (!projectToLessons.has(pidKey)) {
            projectToLessons.set(pidKey, new Map());
        }
        const title = lesson.title != null ? String(lesson.title) : '';
        projectToLessons.get(pidKey).set(String(lid), { id: lesson.id, title });
    }

    const projectIdList = Array.from(projectToLessons.keys()).map((k) => {
        const n = Number(k);
        return Number.isFinite(n) ? n : k;
    });

    const projectNameById = new Map();
    for (const chunk of chunkArray(projectIdList, IN_CHUNK)) {
        const { data: projectRows, error: projErr } = await supabase
            .from('projects')
            .select('project_id, project_name')
            .eq('organization_id', organizationId)
            .in('project_id', chunk);

        if (projErr) {
            return {
                groups: [],
                error: new Error(projErr.message || 'Failed to load projects.'),
            };
        }
        (projectRows || []).forEach((row) => {
            if (row && row.project_id != null) {
                projectNameById.set(
                    String(row.project_id),
                    row.project_name ? String(row.project_name) : `Project ${row.project_id}`
                );
            }
        });
    }

    const groups = [];
    for (const pidKey of projectToLessons.keys()) {
        const lessonsMap = projectToLessons.get(pidKey);
        const lessons = Array.from(lessonsMap.values()).sort((a, b) =>
            String(a.title).localeCompare(String(b.title))
        );
        const n = Number(pidKey);
        groups.push({
            projectId: Number.isFinite(n) ? n : pidKey,
            projectName: projectNameById.get(pidKey) || `Project ${pidKey}`,
            lessons,
        });
    }
    groups.sort((a, b) => String(a.projectName).localeCompare(String(b.projectName)));

    return { groups, error: null };
}
