// ms-project-txt-parser.js
// Placeholder parser for MS Project TXT exports used in the
// Lessons Learned Metadata sub-module of the user portal.
// 
// You will provide the actual parsing logic later. For now, this file
// simply exposes an async function that reads the file text so that
// wiring and imports can be verified end-to-end.

function extractSection(text, tagName) {
    if (!text || !tagName) return null;
    const re = new RegExp(`<${tagName}>([\\s\\S]*?)<\\/${tagName}>`, 'i');
    const match = text.match(re);
    return match ? match[1] : null;
}

function toNullIfEmpty(value) {
    const s = value != null ? String(value).trim() : '';
    return s ? s : null;
}

function parseIntOrNull(value) {
    const s = toNullIfEmpty(value);
    if (s == null) return null;
    const n = Number.parseInt(s, 10);
    return Number.isFinite(n) ? n : null;
}

function parseFloatOrNull(value) {
    const s = toNullIfEmpty(value);
    if (s == null) return null;
    const n = Number.parseFloat(s);
    return Number.isFinite(n) ? n : null;
}

function computeWbsParent(wbs) {
    const s = toNullIfEmpty(wbs);
    if (!s) return null;
    const parts = s.split('.').map(p => p.trim()).filter(Boolean);
    if (parts.length <= 1) return null;
    return parts.slice(0, -1).join('.');
}

// Convert ISO-8601 duration (e.g. PT240H0M0S, P2DT3H) to a Postgres interval string.
// This is intentionally conservative and supports common MS Project duration shapes.
function isoDurationToPgInterval(iso) {
    const s = toNullIfEmpty(iso);
    if (!s) return null;
    const re = /^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/i;
    const m = s.match(re);
    if (!m) {
        // If unknown, return raw string; Postgres may still parse it depending on settings.
        return s;
    }
    const days = m[1] ? Number(m[1]) : 0;
    const hours = m[2] ? Number(m[2]) : 0;
    const minutes = m[3] ? Number(m[3]) : 0;
    const seconds = m[4] ? Number(m[4]) : 0;

    const parts = [];
    if (days) parts.push(`${days} days`);
    if (hours) parts.push(`${hours} hours`);
    if (minutes) parts.push(`${minutes} minutes`);
    if (seconds) parts.push(`${seconds} seconds`);

    // If everything is 0, still return 0 seconds so Postgres has something parseable.
    if (parts.length === 0) return '0 seconds';
    return parts.join(' ');
}

function getFirstText(parent, tagName) {
    if (!parent) return null;
    const el = parent.getElementsByTagName(tagName)[0];
    return el && el.textContent != null ? el.textContent : null;
}

function parseTaskXml(taskXml) {
    // Wrap with a single root to ensure valid XML parsing.
    const wrapped = `<root>${taskXml}</root>`;
    const parser = new DOMParser();
    const doc = parser.parseFromString(wrapped, 'text/xml');
    // Detect parser errors (browser dependent; best-effort)
    if (doc.getElementsByTagName('parsererror').length) {
        return null;
    }
    const taskEl = doc.getElementsByTagName('Task')[0];
    if (!taskEl) return null;

    const uid = toNullIfEmpty(getFirstText(taskEl, 'UID'));
    const name = toNullIfEmpty(getFirstText(taskEl, 'Name'));
    const wbs = toNullIfEmpty(getFirstText(taskEl, 'WBS'));
    const outlineLevel = parseIntOrNull(getFirstText(taskEl, 'OutlineLevel'));

    const start = toNullIfEmpty(getFirstText(taskEl, 'Start'));
    const finish = toNullIfEmpty(getFirstText(taskEl, 'Finish'));
    const durationRaw = toNullIfEmpty(getFirstText(taskEl, 'Duration'));
    const duration = isoDurationToPgInterval(durationRaw);
    const percentComplete = parseFloatOrNull(getFirstText(taskEl, 'PercentComplete'));
    const actualStart = toNullIfEmpty(getFirstText(taskEl, 'ActualStart'));
    const actualFinish = toNullIfEmpty(getFirstText(taskEl, 'ActualFinish'));
    const fixedCost = parseFloatOrNull(getFirstText(taskEl, 'FixedCost'));
    const notes = toNullIfEmpty(getFirstText(taskEl, 'Notes'));

    // Baseline: choose Number 0
    let baselineStart = null;
    let baselineFinish = null;
    const baselines = Array.from(taskEl.getElementsByTagName('Baseline'));
    for (const b of baselines) {
        const num = toNullIfEmpty(getFirstText(b, 'Number'));
        if (num === '0') {
            baselineStart = toNullIfEmpty(getFirstText(b, 'Start'));
            baselineFinish = toNullIfEmpty(getFirstText(b, 'Finish'));
            break;
        }
    }

    const predecessorLinks = Array.from(taskEl.getElementsByTagName('PredecessorLink')).map(link => {
        const predecessorUid = toNullIfEmpty(getFirstText(link, 'PredecessorUID'));
        const type = toNullIfEmpty(getFirstText(link, 'Type'));
        if (!predecessorUid && !type) return null;
        return {
            predecessor_uid: predecessorUid,
            predecessor_type: type
        };
    }).filter(Boolean);

    return {
        uid,
        task_name: name,
        wbs,
        wbs_parent: computeWbsParent(wbs),
        outline_level: outlineLevel,
        start,
        finish,
        duration,
        duration_raw: durationRaw,
        percent_complete: percentComplete,
        actual_start: actualStart,
        actual_finish: actualFinish,
        fixed_cost: fixedCost,
        notes,
        predecessors: predecessorLinks.length > 0,
        baseline_start: baselineStart,
        baseline_finish: baselineFinish,
        predecessor_links: predecessorLinks
    };
}

function parseResourceXml(resourceXml) {
    const wrapped = `<root>${resourceXml}</root>`;
    const parser = new DOMParser();
    const doc = parser.parseFromString(wrapped, 'text/xml');
    if (doc.getElementsByTagName('parsererror').length) {
        return null;
    }
    const el = doc.getElementsByTagName('Resource')[0];
    if (!el) return null;

    const uid = toNullIfEmpty(getFirstText(el, 'UID'));
    const rowId = toNullIfEmpty(getFirstText(el, 'ID'));
    const name = toNullIfEmpty(getFirstText(el, 'Name'));
    const type = toNullIfEmpty(getFirstText(el, 'Type'));
    const maxUnits = parseIntOrNull(getFirstText(el, 'MaxUnits'));
    const standardRate = parseFloatOrNull(getFirstText(el, 'StandardRate'));

    return {
        uid,
        row_id: rowId,
        resource_name: name,
        type,
        max_units: maxUnits,
        standard_rate: standardRate
    };
}

function parseAssignmentXml(assignmentXml) {
    const wrapped = `<root>${assignmentXml}</root>`;
    const parser = new DOMParser();
    const doc = parser.parseFromString(wrapped, 'text/xml');
    if (doc.getElementsByTagName('parsererror').length) {
        return null;
    }
    const el = doc.getElementsByTagName('Assignment')[0];
    if (!el) return null;

    const uid = toNullIfEmpty(getFirstText(el, 'UID'));
    const taskUid = toNullIfEmpty(getFirstText(el, 'TaskUID'));
    const resourceUid = toNullIfEmpty(getFirstText(el, 'ResourceUID'));

    if (!uid && !taskUid && !resourceUid) return null;

    return {
        uid,
        task_uid: taskUid,
        resource_uid: resourceUid
    };
}

/**
 * Parse a TXT file that was exported from MS Project.
 *
 * Responsibilities:
 * - Detect which major sections exist (<Resources>, <Tasks>, <Assignments>)
 * - Extract <Task>, <Resource>, and <Assignment> entries and normalize them into JS objects
 *
 * @param {File} file - The TXT file selected by the user.
 * @returns {Promise<object>} Summary info including which sections are present and parsed entities.
 */
export async function parseMsProjectTxt(file) {
    if (!file) {
        throw new Error('No file provided to parseMsProjectTxt.');
    }

    const text = await file.text();

    const hasResources = text.includes('<Resources>') && text.includes('</Resources>');
    const hasTasks = text.includes('<Tasks>') && text.includes('</Tasks>');
    const hasAssignments = text.includes('<Assignments>') && text.includes('</Assignments>');

    const presentSections = [];
    if (hasResources) presentSections.push('Resources');
    if (hasTasks) presentSections.push('Tasks');
    if (hasAssignments) presentSections.push('Assignments');

    const tasksSection = hasTasks ? extractSection(text, 'Tasks') : null;
    const tasks = [];

    if (tasksSection) {
        const taskRe = /<Task\b[^>]*>[\s\S]*?<\/Task>/gi;
        const matches = tasksSection.match(taskRe) || [];
        for (const taskXml of matches) {
            const parsed = parseTaskXml(taskXml);
            if (parsed) tasks.push(parsed);
        }
    }

    const predecessorLinkCount = tasks.reduce((sum, t) => sum + (t.predecessor_links ? t.predecessor_links.length : 0), 0);

    const resourcesSection = hasResources ? extractSection(text, 'Resources') : null;
    const resources = [];
    if (resourcesSection) {
        const re = /<Resource\b[^>]*>[\s\S]*?<\/Resource>/gi;
        const matches = resourcesSection.match(re) || [];
        for (const xml of matches) {
            const parsed = parseResourceXml(xml);
            if (parsed) resources.push(parsed);
        }
    }

    const assignmentsSection = hasAssignments ? extractSection(text, 'Assignments') : null;
    const assignments = [];
    if (assignmentsSection) {
        const re = /<Assignment\b[^>]*>[\s\S]*?<\/Assignment>/gi;
        const matches = assignmentsSection.match(re) || [];
        for (const xml of matches) {
            const parsed = parseAssignmentXml(xml);
            if (parsed) assignments.push(parsed);
        }
    }

    console.log('[ms-project-txt-parser] Parse summary:', {
        fileName: file.name,
        hasResources,
        hasTasks,
        hasAssignments,
        presentSections,
        taskCount: tasks.length,
        predecessorLinkCount,
        resourceCount: resources.length,
        assignmentCount: assignments.length
    });

    return {
        fileName: file.name,
        size: file.size,
        hasResources,
        hasTasks,
        hasAssignments,
        presentSections,
        presentCount: presentSections.length,
        tasks,
        taskCount: tasks.length,
        predecessorLinkCount,
        resources,
        resourceCount: resources.length,
        assignments,
        assignmentCount: assignments.length
    };
}

function chunkArray(arr, size) {
    if (!Array.isArray(arr) || arr.length === 0) return [];
    const n = Math.max(1, Number(size) || 1);
    const chunks = [];
    for (let i = 0; i < arr.length; i += n) {
        chunks.push(arr.slice(i, i + n));
    }
    return chunks;
}

async function insertChunked({ supabase, table, rows, select, chunkSize, onProgress, label }) {
    if (!rows || rows.length === 0) return [];
    const chunks = chunkArray(rows, chunkSize);
    const inserted = [];

    for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i];
        if (typeof onProgress === 'function') {
            onProgress(`${label} (${i + 1}/${chunks.length})...`);
        }

        const query = supabase
            .from(table)
            .insert(chunk);

        const { data, error } = select
            ? await query.select(select)
            : await query.select();

        if (error) {
            throw new Error(`Failed inserting into ${table}: ${error.message || String(error)}`);
        }
        if (data && data.length) {
            inserted.push(...data);
        }
    }

    return inserted;
}

/**
 * Parse an MS Project TXT file and persist Tasks (+ predecessor links) into Supabase.
 *
 * Tables written:
 * - lessons_learned_metadata_list (one row per task; metadata = task_name; metadata_type = "task")
 * - msproject_task_details (one row per task; links back to lessons_learned_metadata_list)
 * - msproject_task_predecessors (one row per predecessor link; links back to task_details + metadata_list)
 *
 * @param {object} params
 * @param {any} params.supabase - Supabase client (from supabase-js)
 * @param {File} params.file - Uploaded TXT file
 * @param {object} params.context - Required context for inserts
 * @param {number} params.context.organization_id
 * @param {number} params.context.created_by
 * @param {number|string} params.context.project_id
 * @param {number|string} params.context.project_type_id
 * @param {(msg: string) => void} [params.onProgress]
 * @param {number} [params.chunkSize]
 * @returns {Promise<object>} Summary counts and parse info
 */
export async function importMsProjectTxtToSupabase({
    supabase,
    file,
    context,
    onProgress,
    chunkSize = 250
}) {
    if (!supabase) throw new Error('Supabase client is required.');
    if (!file) throw new Error('File is required.');
    if (!context) throw new Error('Context is required.');

    const organization_id = context.organization_id;
    const created_by = context.created_by;
    const project_id = context.project_id;
    const project_type_id = context.project_type_id;

    if (organization_id == null || created_by == null || project_id == null || project_type_id == null) {
        throw new Error('Missing required context: organization_id, created_by, project_id, project_type_id.');
    }

    if (typeof onProgress === 'function') {
        onProgress('Parsing TXT file...');
    }

    const parsed = await parseMsProjectTxt(file);
    const tasks = parsed.tasks || [];
    const resources = parsed.resources || [];
    const assignments = parsed.assignments || [];

    if (!tasks.length) {
        throw new Error('No <Task> entries were found in the <Tasks> section.');
    }

    // 1) Insert lessons_learned_metadata_list rows for TASKS (one per task)
    const taskMetadataRows = tasks.map(t => ({
        metadata_source: 'ms project',
        metadata: t.task_name || null,
        metadata_type: 'task',
        created_by,
        organization_id,
        project_id,
        project_type_id
    }));

    const insertedMetadata = await insertChunked({
        supabase,
        table: 'lessons_learned_metadata_list',
        rows: taskMetadataRows,
        select: 'id',
        chunkSize,
        onProgress,
        label: 'Saving task metadata'
    });

    if (insertedMetadata.length !== taskMetadataRows.length) {
        throw new Error('Unexpected mismatch inserting lessons_learned_metadata_list rows.');
    }

    const taskLessonsIds = insertedMetadata.map(r => r.id);

    // 2) Insert msproject_task_details rows (one per task)
    const taskDetailRows = tasks.map((t, idx) => ({
        lessons_learned_metadata_list_id: taskLessonsIds[idx],
        organization_id,
        project_id,
        project_type_id,
        created_by,

        task_name: t.task_name || null,
        uid: t.uid || null,
        wbs: t.wbs || null,
        wbs_parent: t.wbs_parent || null,
        outline_level: t.outline_level,

        start: t.start || null,
        finish: t.finish || null,
        duration: t.duration || null,
        percent_complete: t.percent_complete,
        actual_start: t.actual_start || null,
        actual_finish: t.actual_finish || null,
        fixed_cost: t.fixed_cost,
        notes: t.notes || null,
        predecessors: !!t.predecessors,
        baseline_start: t.baseline_start || null,
        baseline_finish: t.baseline_finish || null
    }));

    const insertedTaskDetails = await insertChunked({
        supabase,
        table: 'msproject_task_details',
        rows: taskDetailRows,
        select: 'id, lessons_learned_metadata_list_id',
        chunkSize,
        onProgress,
        label: 'Saving task details'
    });

    if (insertedTaskDetails.length !== taskDetailRows.length) {
        throw new Error('Unexpected mismatch inserting msproject_task_details rows.');
    }

    const taskDetailsIdByLessonsId = new Map();
    insertedTaskDetails.forEach(row => {
        taskDetailsIdByLessonsId.set(String(row.lessons_learned_metadata_list_id), row.id);
    });

    // 3) Insert predecessor links
    const predecessorRows = [];
    tasks.forEach((t, idx) => {
        const lessonsId = taskLessonsIds[idx];
        const taskDetailsId = taskDetailsIdByLessonsId.get(String(lessonsId));
        const links = t.predecessor_links || [];
        links.forEach(link => {
            predecessorRows.push({
                msproject_task_details_id: taskDetailsId,
                lessons_learned_metadata_list_id: lessonsId,
                organization_id,
                project_id,
                project_type_id,
                predecessor_uid: link.predecessor_uid || null,
                predecessor_type: link.predecessor_type || null
            });
        });
    });

    let insertedPredecessors = [];
    if (predecessorRows.length) {
        insertedPredecessors = await insertChunked({
            supabase,
            table: 'msproject_task_predecessors',
            rows: predecessorRows,
            select: 'id',
            chunkSize,
            onProgress,
            label: 'Saving predecessor links'
        });
    } else if (typeof onProgress === 'function') {
        onProgress('No predecessor links found. Skipping predecessor import.');
    }

    // 4) Insert lessons_learned_metadata_list rows for RESOURCES (one per resource)
    let insertedResourceMetadata = [];
    let insertedResourceDetails = [];
    if (resources.length) {
        const resourceMetadataRows = resources.map(r => ({
            metadata_source: 'ms project',
            metadata: r.resource_name || null,
            metadata_type: 'resource',
            created_by,
            organization_id,
            project_id,
            project_type_id
        }));

        insertedResourceMetadata = await insertChunked({
            supabase,
            table: 'lessons_learned_metadata_list',
            rows: resourceMetadataRows,
            select: 'id',
            chunkSize,
            onProgress,
            label: 'Saving resource metadata'
        });

        if (insertedResourceMetadata.length !== resourceMetadataRows.length) {
            throw new Error('Unexpected mismatch inserting resource metadata rows.');
        }

        const resourceLessonsIds = insertedResourceMetadata.map(r => r.id);

        // 5) Insert msproject_resource_details (one per resource; linked via resource metadata id)
        const resourceDetailRows = resources.map((r, idx) => ({
            lessons_learned_metadata_list_id: resourceLessonsIds[idx],
            organization_id,
            project_id,
            project_type_id,
            created_by,
            resource_name: r.resource_name || null,
            uid: r.uid || null,
            row_id: r.row_id || null,
            type: r.type || null,
            max_units: r.max_units,
            standard_rate: r.standard_rate
        }));

        insertedResourceDetails = await insertChunked({
            supabase,
            table: 'msproject_resource_details',
            rows: resourceDetailRows,
            select: 'id',
            chunkSize,
            onProgress,
            label: 'Saving resource details'
        });
    } else if (typeof onProgress === 'function') {
        onProgress('No resources found. Skipping resource import.');
    }

    // 6) Insert msproject_assignments (link rows only)
    let insertedAssignments = [];
    if (assignments.length) {
        const assignmentRows = assignments.map(a => ({
            organization_id,
            project_id,
            project_type_id,
            created_by,
            uid: a.uid || null,
            task_uid: a.task_uid || null,
            resource_uid: a.resource_uid || null
        }));

        insertedAssignments = await insertChunked({
            supabase,
            table: 'msproject_assignments',
            rows: assignmentRows,
            select: 'id',
            chunkSize,
            onProgress,
            label: 'Saving assignments'
        });
    } else if (typeof onProgress === 'function') {
        onProgress('No assignments found. Skipping assignment import.');
    }

    return {
        ...parsed,
        inserted: {
            lessons_learned_metadata_list: insertedMetadata.length + insertedResourceMetadata.length,
            msproject_task_details: insertedTaskDetails.length,
            msproject_task_predecessors: insertedPredecessors.length,
            msproject_resource_details: insertedResourceDetails.length,
            msproject_assignments: insertedAssignments.length
        }
    };
}


