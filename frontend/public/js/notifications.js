/**
 * Lessons learned review: populate lessons_learned_review_notifications from metadata tags
 * and project_team_member_assignments.
 *
 * Requires Supabase RLS policies that allow the portal user's role to SELECT the source tables
 * and INSERT into `lessons_learned_review_notifications` (and read assignments/metadata as used below).
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
