// ms-project-xml-update.js
// Update importer for MS Project XML (Lessons Learned Metadata sub-module).

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
function isoDurationToPgInterval(iso) {
    const s = toNullIfEmpty(iso);
    if (!s) return null;
    const re = /^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/i;
    const m = s.match(re);
    if (!m) return s;
    const days = m[1] ? Number(m[1]) : 0;
    const hours = m[2] ? Number(m[2]) : 0;
    const minutes = m[3] ? Number(m[3]) : 0;
    const seconds = m[4] ? Number(m[4]) : 0;

    const parts = [];
    if (days) parts.push(`${days} days`);
    if (hours) parts.push(`${hours} hours`);
    if (minutes) parts.push(`${minutes} minutes`);
    if (seconds) parts.push(`${seconds} seconds`);
    if (parts.length === 0) return '0 seconds';
    return parts.join(' ');
}

function getFirstText(parent, tagName) {
    if (!parent) return null;
    const el = parent.getElementsByTagName(tagName)[0];
    return el && el.textContent != null ? el.textContent : null;
}

function parseTaskElement(taskEl) {
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

function parseResourceElement(resourceEl) {
    if (!resourceEl) return null;
    const uid = toNullIfEmpty(getFirstText(resourceEl, 'UID'));
    const rowId = toNullIfEmpty(getFirstText(resourceEl, 'ID'));
    const name = toNullIfEmpty(getFirstText(resourceEl, 'Name'));
    const type = toNullIfEmpty(getFirstText(resourceEl, 'Type'));
    const maxUnits = parseIntOrNull(getFirstText(resourceEl, 'MaxUnits'));
    const standardRate = parseFloatOrNull(getFirstText(resourceEl, 'StandardRate'));
    return {
        uid,
        row_id: rowId,
        resource_name: name,
        type,
        max_units: maxUnits,
        standard_rate: standardRate
    };
}

function parseAssignmentElement(assignmentEl) {
    if (!assignmentEl) return null;
    const uid = toNullIfEmpty(getFirstText(assignmentEl, 'UID'));
    const taskUid = toNullIfEmpty(getFirstText(assignmentEl, 'TaskUID'));
    const resourceUid = toNullIfEmpty(getFirstText(assignmentEl, 'ResourceUID'));
    if (!uid && !taskUid && !resourceUid) return null;
    return {
        uid,
        task_uid: taskUid,
        resource_uid: resourceUid
    };
}

export async function parseMsProjectXml(file) {
    if (!file) {
        throw new Error('No file provided to parseMsProjectXml.');
    }
    const text = await file.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, 'text/xml');
    if (doc.getElementsByTagName('parsererror').length) {
        throw new Error('Invalid XML. Please upload a valid MS Project XML file.');
    }

    const tasks = Array.from(doc.getElementsByTagName('Task'))
        .map(parseTaskElement)
        .filter(Boolean);
    const resources = Array.from(doc.getElementsByTagName('Resource'))
        .map(parseResourceElement)
        .filter(Boolean);
    const assignments = Array.from(doc.getElementsByTagName('Assignment'))
        .map(parseAssignmentElement)
        .filter(Boolean);

    return {
        fileName: file.name,
        size: file.size,
        tasks,
        taskCount: tasks.length,
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

function coerceComparable(value) {
    if (value == null || value === '') return null;
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (typeof value === 'boolean') return value;
    return String(value);
}

function valuesEqual(a, b) {
    return coerceComparable(a) === coerceComparable(b);
}

function normalizeLinkList(links) {
    const list = Array.isArray(links) ? links : [];
    return list.map(l => `${l.predecessor_uid || ''}|${l.predecessor_type || ''}`).sort();
}

function filterNewUnique(items, existingMap) {
    const out = [];
    const seen = new Set();
    (items || []).forEach(item => {
        if (!item || !item.uid) return;
        const uid = String(item.uid);
        if (existingMap && existingMap.has(uid)) return;
        if (seen.has(uid)) return;
        seen.add(uid);
        out.push(item);
    });
    return out;
}

function formatValue(value) {
    if (value == null || value === '') return 'null';
    if (typeof value === 'boolean') return value ? 'true' : 'false';
    if (typeof value === 'number') return String(value);
    return String(value);
}

function formatChanges(changes) {
    const parts = [];
    Object.keys(changes).forEach(field => {
        const { from, to } = changes[field];
        parts.push(`${field}: ${formatValue(from)} -> ${formatValue(to)}`);
    });
    return parts.join('; ');
}

function formatList(items, limit, formatItem) {
    const slice = items.slice(0, limit);
    const lines = slice.map(formatItem);
    if (items.length > limit) {
        lines.push(`...and ${items.length - limit} more`);
    }
    return lines;
}

async function fetchMetadataRowsByIds({ supabase, ids, chunkSize }) {
    const rows = [];
    const chunks = chunkArray(ids, chunkSize);
    for (const chunk of chunks) {
        const { data, error } = await supabase
            .from('lessons_learned_metadata_list')
            .select('id, metadata, metadata_type, metadata_source')
            .in('id', chunk);
        if (error) throw new Error(error.message || 'Failed loading metadata list rows.');
        rows.push(...(data || []));
    }
    return rows;
}

async function fetchUsedMetadataListIds({ supabase, ids, organization_id, project_id, chunkSize }) {
    const used = new Set();
    const chunks = chunkArray(ids, chunkSize);
    for (const chunk of chunks) {
        const { data, error } = await supabase
            .from('lessons_learned_metadata')
            .select('lessons_learned_metadata_list_id')
            .eq('organization_id', organization_id)
            .eq('project_id', project_id)
            .in('lessons_learned_metadata_list_id', chunk);
        if (error) throw new Error(error.message || 'Failed checking metadata usage.');
        (data || []).forEach(row => {
            if (row && row.lessons_learned_metadata_list_id != null) {
                used.add(String(row.lessons_learned_metadata_list_id));
            }
        });
    }
    return used;
}

export async function updateMsProjectXmlToSupabase({
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
        onProgress('Parsing XML file...');
    }
    const parsed = await parseMsProjectXml(file);

    const tasks = parsed.tasks || [];
    const resources = parsed.resources || [];
    const assignments = parsed.assignments || [];

    if (typeof onProgress === 'function') {
        onProgress('Loading existing MS Project data...');
    }

    const { data: taskDetailsRows, error: taskErr } = await supabase
        .from('msproject_task_details')
        .select('id, lessons_learned_metadata_list_id, uid, task_name, wbs, wbs_parent, outline_level, start, finish, duration, percent_complete, actual_start, actual_finish, fixed_cost, notes, predecessors, baseline_start, baseline_finish')
        .eq('organization_id', organization_id)
        .eq('project_id', project_id);
    if (taskErr) throw new Error(taskErr.message || 'Failed loading task details.');

    const { data: resourceDetailsRows, error: resourceErr } = await supabase
        .from('msproject_resource_details')
        .select('id, lessons_learned_metadata_list_id, uid, row_id, resource_name, type, max_units, standard_rate')
        .eq('organization_id', organization_id)
        .eq('project_id', project_id);
    if (resourceErr) throw new Error(resourceErr.message || 'Failed loading resource details.');

    const { data: assignmentRows, error: assignmentErr } = await supabase
        .from('msproject_assignments')
        .select('id, uid, task_uid, resource_uid')
        .eq('organization_id', organization_id)
        .eq('project_id', project_id);
    if (assignmentErr) throw new Error(assignmentErr.message || 'Failed loading assignments.');

    const taskDetails = taskDetailsRows || [];
    const resourceDetails = resourceDetailsRows || [];
    const existingAssignments = assignmentRows || [];

    const metadataIds = [
        ...taskDetails.map(r => r && r.lessons_learned_metadata_list_id).filter(v => v != null),
        ...resourceDetails.map(r => r && r.lessons_learned_metadata_list_id).filter(v => v != null)
    ];
    const metadataRows = metadataIds.length
        ? await fetchMetadataRowsByIds({ supabase, ids: metadataIds, chunkSize })
        : [];
    const metadataById = new Map(metadataRows.map(r => [String(r.id), r]));

    const { data: predecessorRows, error: predErr } = await supabase
        .from('msproject_task_predecessors')
        .select('id, lessons_learned_metadata_list_id, predecessor_uid, predecessor_type')
        .eq('organization_id', organization_id)
        .eq('project_id', project_id);
    if (predErr) throw new Error(predErr.message || 'Failed loading predecessor links.');

    const predecessorsByListId = new Map();
    (predecessorRows || []).forEach(row => {
        const id = row && row.lessons_learned_metadata_list_id != null ? String(row.lessons_learned_metadata_list_id) : null;
        if (!id) return;
        if (!predecessorsByListId.has(id)) predecessorsByListId.set(id, []);
        predecessorsByListId.get(id).push({
            predecessor_uid: row.predecessor_uid || null,
            predecessor_type: row.predecessor_type || null
        });
    });

    const tasksByUid = new Map();
    taskDetails.forEach(row => {
        if (!row || !row.uid) return;
        const meta = metadataById.get(String(row.lessons_learned_metadata_list_id || '')) || null;
        tasksByUid.set(String(row.uid), { detail: row, metadata: meta });
    });
    const resourcesByUid = new Map();
    resourceDetails.forEach(row => {
        if (!row || !row.uid) return;
        const meta = metadataById.get(String(row.lessons_learned_metadata_list_id || '')) || null;
        resourcesByUid.set(String(row.uid), { detail: row, metadata: meta });
    });
    const assignmentsByUid = new Map();
    const assignmentsNoUid = [];
    existingAssignments.forEach(row => {
        if (!row || !row.uid) {
            assignmentsNoUid.push(row);
            return;
        }
        assignmentsByUid.set(String(row.uid), row);
    });

    const summary = {
        updatedTasks: [],
        insertedTasks: [],
        deletedTasks: [],
        keptOldTasks: [],
        skippedTasks: [],
        updatedResources: [],
        insertedResources: [],
        deletedResources: [],
        keptOldResources: [],
        skippedResources: [],
        updatedAssignments: [],
        insertedAssignments: [],
        deletedAssignments: [],
        skippedAssignments: [],
        notes: []
    };

    const parsedTaskUids = new Set();
    const parsedResourceUids = new Set();
    const parsedAssignmentUids = new Set();

    if (typeof onProgress === 'function') onProgress('Updating task records...');

    for (const task of tasks) {
        if (!task || !task.uid) {
            summary.skippedTasks.push({ reason: 'missing uid', task_name: task && task.task_name });
            continue;
        }
        const uid = String(task.uid);
        if (parsedTaskUids.has(uid)) {
            summary.skippedTasks.push({ uid, reason: 'duplicate uid' });
            continue;
        }
        parsedTaskUids.add(uid);
        const existing = tasksByUid.get(uid);
        if (!existing) continue;

        const detail = existing.detail;
        const updatePayload = {};
        const changes = {};

        const fields = {
            task_name: task.task_name || null,
            wbs: task.wbs || null,
            wbs_parent: task.wbs_parent || null,
            outline_level: task.outline_level,
            start: task.start || null,
            finish: task.finish || null,
            duration: task.duration || null,
            percent_complete: task.percent_complete,
            actual_start: task.actual_start || null,
            actual_finish: task.actual_finish || null,
            fixed_cost: task.fixed_cost,
            notes: task.notes || null,
            predecessors: !!task.predecessors,
            baseline_start: task.baseline_start || null,
            baseline_finish: task.baseline_finish || null
        };

        Object.keys(fields).forEach(field => {
            const nextValue = fields[field];
            const prevValue = detail[field];
            if (!valuesEqual(prevValue, nextValue)) {
                updatePayload[field] = nextValue;
                changes[field] = { from: prevValue, to: nextValue };
            }
        });

        const listId = detail.lessons_learned_metadata_list_id;
        const existingLinks = predecessorsByListId.get(String(listId)) || [];
        const nextLinks = task.predecessor_links || [];
        if (normalizeLinkList(existingLinks).join(',') !== normalizeLinkList(nextLinks).join(',')) {
            changes.predecessor_links = {
                from: existingLinks.map(l => `${l.predecessor_uid || ''}:${l.predecessor_type || ''}`),
                to: nextLinks.map(l => `${l.predecessor_uid || ''}:${l.predecessor_type || ''}`)
            };
            updatePayload.predecessors = nextLinks.length > 0;
        }

        if (Object.keys(updatePayload).length) {
            const { error } = await supabase
                .from('msproject_task_details')
                .update(updatePayload)
                .eq('id', detail.id);
            if (error) throw new Error(error.message || 'Failed updating task details.');
        }

        if (changes.predecessor_links) {
            const { error: deleteErr } = await supabase
                .from('msproject_task_predecessors')
                .delete()
                .eq('lessons_learned_metadata_list_id', listId);
            if (deleteErr) throw new Error(deleteErr.message || 'Failed deleting predecessor links.');

            if (nextLinks.length) {
                const insertRows = nextLinks.map(link => ({
                    msproject_task_details_id: detail.id,
                    lessons_learned_metadata_list_id: listId,
                    organization_id,
                    project_id,
                    project_type_id,
                    predecessor_uid: link.predecessor_uid || null,
                    predecessor_type: link.predecessor_type || null
                }));
                await insertChunked({
                    supabase,
                    table: 'msproject_task_predecessors',
                    rows: insertRows,
                    select: 'id',
                    chunkSize,
                    onProgress,
                    label: 'Saving predecessor links'
                });
            }
        }

        const meta = existing.metadata;
        if (meta) {
            const metaUpdate = {};
            if (!valuesEqual(meta.metadata, task.task_name || null)) {
                metaUpdate.metadata = task.task_name || null;
            }
            if (meta.metadata_source === 'ms project - old') {
                metaUpdate.metadata_source = 'ms project';
            }
            if (Object.keys(metaUpdate).length) {
                const { error } = await supabase
                    .from('lessons_learned_metadata_list')
                    .update(metaUpdate)
                    .eq('id', meta.id);
                if (error) throw new Error(error.message || 'Failed updating task metadata list row.');
                if (metaUpdate.metadata_source) {
                    changes.metadata_source = { from: meta.metadata_source, to: metaUpdate.metadata_source };
                }
                if (metaUpdate.metadata) {
                    changes.metadata = { from: meta.metadata, to: metaUpdate.metadata };
                }
            }
        }

        if (Object.keys(changes).length) {
            summary.updatedTasks.push({
                uid,
                task_name: task.task_name || detail.task_name || null,
                changes
            });
        }
    }

    if (typeof onProgress === 'function') onProgress('Updating resource records...');

    for (const resource of resources) {
        if (!resource || !resource.uid) {
            summary.skippedResources.push({ reason: 'missing uid', resource_name: resource && resource.resource_name });
            continue;
        }
        const uid = String(resource.uid);
        if (parsedResourceUids.has(uid)) {
            summary.skippedResources.push({ uid, reason: 'duplicate uid' });
            continue;
        }
        parsedResourceUids.add(uid);
        const existing = resourcesByUid.get(uid);
        if (!existing) continue;

        const detail = existing.detail;
        const updatePayload = {};
        const changes = {};

        const fields = {
            row_id: resource.row_id || null,
            resource_name: resource.resource_name || null,
            type: resource.type || null,
            max_units: resource.max_units,
            standard_rate: resource.standard_rate
        };

        Object.keys(fields).forEach(field => {
            const nextValue = fields[field];
            const prevValue = detail[field];
            if (!valuesEqual(prevValue, nextValue)) {
                updatePayload[field] = nextValue;
                changes[field] = { from: prevValue, to: nextValue };
            }
        });

        if (Object.keys(updatePayload).length) {
            const { error } = await supabase
                .from('msproject_resource_details')
                .update(updatePayload)
                .eq('id', detail.id);
            if (error) throw new Error(error.message || 'Failed updating resource details.');
        }

        const meta = existing.metadata;
        if (meta) {
            const metaUpdate = {};
            if (!valuesEqual(meta.metadata, resource.resource_name || null)) {
                metaUpdate.metadata = resource.resource_name || null;
            }
            if (meta.metadata_source === 'ms project - old') {
                metaUpdate.metadata_source = 'ms project';
            }
            if (Object.keys(metaUpdate).length) {
                const { error } = await supabase
                    .from('lessons_learned_metadata_list')
                    .update(metaUpdate)
                    .eq('id', meta.id);
                if (error) throw new Error(error.message || 'Failed updating resource metadata list row.');
                if (metaUpdate.metadata_source) {
                    changes.metadata_source = { from: meta.metadata_source, to: metaUpdate.metadata_source };
                }
                if (metaUpdate.metadata) {
                    changes.metadata = { from: meta.metadata, to: metaUpdate.metadata };
                }
            }
        }

        if (Object.keys(changes).length) {
            summary.updatedResources.push({
                uid,
                resource_name: resource.resource_name || detail.resource_name || null,
                changes
            });
        }
    }

    if (typeof onProgress === 'function') onProgress('Updating assignment records...');

    for (const assignment of assignments) {
        if (!assignment || !assignment.uid) {
            summary.skippedAssignments.push({ reason: 'missing uid' });
            continue;
        }
        const uid = String(assignment.uid);
        if (parsedAssignmentUids.has(uid)) {
            summary.skippedAssignments.push({ uid, reason: 'duplicate uid' });
            continue;
        }
        parsedAssignmentUids.add(uid);
        const existing = assignmentsByUid.get(uid);
        if (!existing) continue;

        const updatePayload = {};
        const changes = {};

        if (!valuesEqual(existing.task_uid, assignment.task_uid || null)) {
            updatePayload.task_uid = assignment.task_uid || null;
            changes.task_uid = { from: existing.task_uid, to: assignment.task_uid || null };
        }
        if (!valuesEqual(existing.resource_uid, assignment.resource_uid || null)) {
            updatePayload.resource_uid = assignment.resource_uid || null;
            changes.resource_uid = { from: existing.resource_uid, to: assignment.resource_uid || null };
        }

        if (Object.keys(updatePayload).length) {
            const { error } = await supabase
                .from('msproject_assignments')
                .update(updatePayload)
                .eq('id', existing.id);
            if (error) throw new Error(error.message || 'Failed updating assignment.');
            summary.updatedAssignments.push({ uid, changes });
        }
    }

    if (typeof onProgress === 'function') onProgress('Inserting new items...');

    const newTasks = filterNewUnique(tasks, tasksByUid);
    const newResources = filterNewUnique(resources, resourcesByUid);
    const newAssignments = filterNewUnique(assignments, assignmentsByUid);

    if (newTasks.length) {
        const taskMetaRows = newTasks.map(t => ({
            metadata_source: 'ms project',
            metadata: t.task_name || null,
            metadata_type: 'task',
            created_by,
            organization_id,
            project_id,
            project_type_id
        }));

        const insertedMeta = await insertChunked({
            supabase,
            table: 'lessons_learned_metadata_list',
            rows: taskMetaRows,
            select: 'id',
            chunkSize,
            onProgress,
            label: 'Saving task metadata'
        });

        if (insertedMeta.length !== taskMetaRows.length) {
            throw new Error('Unexpected mismatch inserting task metadata rows.');
        }

        const taskDetailRows = newTasks.map((t, idx) => ({
            lessons_learned_metadata_list_id: insertedMeta[idx].id,
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

        const insertedDetails = await insertChunked({
            supabase,
            table: 'msproject_task_details',
            rows: taskDetailRows,
            select: 'id, lessons_learned_metadata_list_id',
            chunkSize,
            onProgress,
            label: 'Saving task details'
        });

        if (insertedDetails.length !== taskDetailRows.length) {
            throw new Error('Unexpected mismatch inserting task details rows.');
        }

        const detailsByLessonsId = new Map();
        insertedDetails.forEach(row => {
            detailsByLessonsId.set(String(row.lessons_learned_metadata_list_id), row.id);
        });

        const predecessorRows = [];
        newTasks.forEach((t, idx) => {
            const listId = insertedMeta[idx].id;
            const taskDetailsId = detailsByLessonsId.get(String(listId));
            const links = t.predecessor_links || [];
            links.forEach(link => {
                predecessorRows.push({
                    msproject_task_details_id: taskDetailsId,
                    lessons_learned_metadata_list_id: listId,
                    organization_id,
                    project_id,
                    project_type_id,
                    predecessor_uid: link.predecessor_uid || null,
                    predecessor_type: link.predecessor_type || null
                });
            });
        });
        if (predecessorRows.length) {
            await insertChunked({
                supabase,
                table: 'msproject_task_predecessors',
                rows: predecessorRows,
                select: 'id',
                chunkSize,
                onProgress,
                label: 'Saving predecessor links'
            });
        }

        newTasks.forEach(t => {
            summary.insertedTasks.push({ uid: String(t.uid), task_name: t.task_name || null });
        });
    }

    if (newResources.length) {
        const resourceMetaRows = newResources.map(r => ({
            metadata_source: 'ms project',
            metadata: r.resource_name || null,
            metadata_type: 'resource',
            created_by,
            organization_id,
            project_id,
            project_type_id
        }));

        const insertedMeta = await insertChunked({
            supabase,
            table: 'lessons_learned_metadata_list',
            rows: resourceMetaRows,
            select: 'id',
            chunkSize,
            onProgress,
            label: 'Saving resource metadata'
        });

        if (insertedMeta.length !== resourceMetaRows.length) {
            throw new Error('Unexpected mismatch inserting resource metadata rows.');
        }

        const resourceDetailRows = newResources.map((r, idx) => ({
            lessons_learned_metadata_list_id: insertedMeta[idx].id,
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

        await insertChunked({
            supabase,
            table: 'msproject_resource_details',
            rows: resourceDetailRows,
            select: 'id',
            chunkSize,
            onProgress,
            label: 'Saving resource details'
        });

        newResources.forEach(r => {
            summary.insertedResources.push({ uid: String(r.uid), resource_name: r.resource_name || null });
        });
    }

    if (newAssignments.length) {
        const assignmentRows = newAssignments.map(a => ({
            organization_id,
            project_id,
            project_type_id,
            created_by,
            uid: a.uid || null,
            task_uid: a.task_uid || null,
            resource_uid: a.resource_uid || null
        }));
        await insertChunked({
            supabase,
            table: 'msproject_assignments',
            rows: assignmentRows,
            select: 'id',
            chunkSize,
            onProgress,
            label: 'Saving assignments'
        });
        newAssignments.forEach(a => {
            summary.insertedAssignments.push({ uid: String(a.uid), task_uid: a.task_uid || null, resource_uid: a.resource_uid || null });
        });
    }

    if (typeof onProgress === 'function') onProgress('Deleting missing items...');

    const missingTasks = taskDetails.filter(row => row && row.uid && !parsedTaskUids.has(String(row.uid)));
    const missingResources = resourceDetails.filter(row => row && row.uid && !parsedResourceUids.has(String(row.uid)));
    const missingAssignments = existingAssignments.filter(row => row && row.uid && !parsedAssignmentUids.has(String(row.uid)));

    const missingTaskListIds = missingTasks.map(r => r.lessons_learned_metadata_list_id).filter(v => v != null);
    const missingResourceListIds = missingResources.map(r => r.lessons_learned_metadata_list_id).filter(v => v != null);

    const usedIds = await fetchUsedMetadataListIds({
        supabase,
        ids: [...missingTaskListIds, ...missingResourceListIds].map(v => String(v)),
        organization_id,
        project_id,
        chunkSize
    });

    const deletableTasks = [];
    const protectedTasks = [];
    missingTasks.forEach(row => {
        const listId = row.lessons_learned_metadata_list_id != null ? String(row.lessons_learned_metadata_list_id) : null;
        if (listId && usedIds.has(listId)) {
            protectedTasks.push(row);
        } else {
            deletableTasks.push(row);
        }
    });

    const deletableResources = [];
    const protectedResources = [];
    missingResources.forEach(row => {
        const listId = row.lessons_learned_metadata_list_id != null ? String(row.lessons_learned_metadata_list_id) : null;
        if (listId && usedIds.has(listId)) {
            protectedResources.push(row);
        } else {
            deletableResources.push(row);
        }
    });

    if (protectedTasks.length || protectedResources.length) {
        const protectIds = [
            ...protectedTasks.map(r => r.lessons_learned_metadata_list_id),
            ...protectedResources.map(r => r.lessons_learned_metadata_list_id)
        ].filter(v => v != null);
        const protectChunks = chunkArray(protectIds, chunkSize);
        for (const chunk of protectChunks) {
            const { error } = await supabase
                .from('lessons_learned_metadata_list')
                .update({ metadata_source: 'ms project - old' })
                .in('id', chunk);
            if (error) throw new Error(error.message || 'Failed marking metadata as old.');
        }
        protectedTasks.forEach(row => {
            summary.keptOldTasks.push({ uid: String(row.uid), task_name: row.task_name || null });
        });
        protectedResources.forEach(row => {
            summary.keptOldResources.push({ uid: String(row.uid), resource_name: row.resource_name || null });
        });
    }

    if (deletableTasks.length) {
        const listIds = deletableTasks.map(r => r.lessons_learned_metadata_list_id).filter(v => v != null);
        const chunks = chunkArray(listIds, chunkSize);
        for (const chunk of chunks) {
            const { error: predDeleteErr } = await supabase
                .from('msproject_task_predecessors')
                .delete()
                .in('lessons_learned_metadata_list_id', chunk);
            if (predDeleteErr) throw new Error(predDeleteErr.message || 'Failed deleting predecessor links.');

            const { error: detailsErr } = await supabase
                .from('msproject_task_details')
                .delete()
                .in('lessons_learned_metadata_list_id', chunk);
            if (detailsErr) throw new Error(detailsErr.message || 'Failed deleting task details.');

            const { error: metaErr } = await supabase
                .from('lessons_learned_metadata_list')
                .delete()
                .in('id', chunk);
            if (metaErr) throw new Error(metaErr.message || 'Failed deleting task metadata.');
        }
        deletableTasks.forEach(row => {
            summary.deletedTasks.push({ uid: String(row.uid), task_name: row.task_name || null });
        });
    }

    if (deletableResources.length) {
        const listIds = deletableResources.map(r => r.lessons_learned_metadata_list_id).filter(v => v != null);
        const chunks = chunkArray(listIds, chunkSize);
        for (const chunk of chunks) {
            const { error: detailsErr } = await supabase
                .from('msproject_resource_details')
                .delete()
                .in('lessons_learned_metadata_list_id', chunk);
            if (detailsErr) throw new Error(detailsErr.message || 'Failed deleting resource details.');

            const { error: metaErr } = await supabase
                .from('lessons_learned_metadata_list')
                .delete()
                .in('id', chunk);
            if (metaErr) throw new Error(metaErr.message || 'Failed deleting resource metadata.');
        }
        deletableResources.forEach(row => {
            summary.deletedResources.push({ uid: String(row.uid), resource_name: row.resource_name || null });
        });
    }

    if (missingAssignments.length) {
        const ids = missingAssignments.map(r => r.id).filter(v => v != null);
        const chunks = chunkArray(ids, chunkSize);
        let deletedAssignments = 0;
        for (const chunk of chunks) {
            const { data, error } = await supabase
                .from('msproject_assignments')
                .delete()
                .in('id', chunk)
                .select('id');
            if (error) throw new Error(error.message || 'Failed deleting assignments.');
            deletedAssignments += (data || []).length;
        }
        missingAssignments.forEach(row => {
            summary.deletedAssignments.push({ uid: String(row.uid) });
        });
        if (deletedAssignments !== missingAssignments.length) {
            summary.notes.push('Some assignments could not be deleted (row count mismatch).');
        }
    }

    if (assignmentsNoUid.length) {
        summary.notes.push(`${assignmentsNoUid.length} existing assignment rows had no UID and were left unchanged.`);
    }

    const summaryLines = [];
    summaryLines.push(`Updated "${parsed.fileName}".`);
    summaryLines.push(
        `Tasks: updated ${summary.updatedTasks.length}, added ${summary.insertedTasks.length}, ` +
        `deleted ${summary.deletedTasks.length}, kept old ${summary.keptOldTasks.length}.`
    );
    summaryLines.push(
        `Resources: updated ${summary.updatedResources.length}, added ${summary.insertedResources.length}, ` +
        `deleted ${summary.deletedResources.length}, kept old ${summary.keptOldResources.length}.`
    );
    summaryLines.push(
        `Assignments: updated ${summary.updatedAssignments.length}, added ${summary.insertedAssignments.length}, deleted ${summary.deletedAssignments.length}.`
    );

    const detailLimit = 25;
    if (summary.updatedTasks.length) {
        summaryLines.push('Updated tasks (fields):');
        summaryLines.push(...formatList(summary.updatedTasks, detailLimit, item =>
            `- ${item.uid}${item.task_name ? ` (${item.task_name})` : ''}: ${formatChanges(item.changes)}`
        ));
    }
    if (summary.insertedTasks.length) {
        summaryLines.push('New tasks:');
        summaryLines.push(...formatList(summary.insertedTasks, detailLimit, item =>
            `- ${item.uid}${item.task_name ? ` (${item.task_name})` : ''}`
        ));
    }
    if (summary.deletedTasks.length) {
        summaryLines.push('Deleted tasks:');
        summaryLines.push(...formatList(summary.deletedTasks, detailLimit, item =>
            `- ${item.uid}${item.task_name ? ` (${item.task_name})` : ''}`
        ));
    }
    if (summary.keptOldTasks.length) {
        summaryLines.push('Tasks kept (tagged in lessons learned):');
        summaryLines.push(...formatList(summary.keptOldTasks, detailLimit, item =>
            `- ${item.uid}${item.task_name ? ` (${item.task_name})` : ''} -> metadata_source set to "ms project - old"`
        ));
    }
    if (summary.updatedResources.length) {
        summaryLines.push('Updated resources (fields):');
        summaryLines.push(...formatList(summary.updatedResources, detailLimit, item =>
            `- ${item.uid}${item.resource_name ? ` (${item.resource_name})` : ''}: ${formatChanges(item.changes)}`
        ));
    }
    if (summary.insertedResources.length) {
        summaryLines.push('New resources:');
        summaryLines.push(...formatList(summary.insertedResources, detailLimit, item =>
            `- ${item.uid}${item.resource_name ? ` (${item.resource_name})` : ''}`
        ));
    }
    if (summary.deletedResources.length) {
        summaryLines.push('Deleted resources:');
        summaryLines.push(...formatList(summary.deletedResources, detailLimit, item =>
            `- ${item.uid}${item.resource_name ? ` (${item.resource_name})` : ''}`
        ));
    }
    if (summary.keptOldResources.length) {
        summaryLines.push('Resources kept (tagged in lessons learned):');
        summaryLines.push(...formatList(summary.keptOldResources, detailLimit, item =>
            `- ${item.uid}${item.resource_name ? ` (${item.resource_name})` : ''} -> metadata_source set to "ms project - old"`
        ));
    }
    if (summary.updatedAssignments.length) {
        summaryLines.push('Updated assignments:');
        summaryLines.push(...formatList(summary.updatedAssignments, detailLimit, item =>
            `- ${item.uid}: ${formatChanges(item.changes)}`
        ));
    }
    if (summary.insertedAssignments.length) {
        summaryLines.push('New assignments:');
        summaryLines.push(...formatList(summary.insertedAssignments, detailLimit, item =>
            `- ${item.uid}`
        ));
    }
    if (summary.deletedAssignments.length) {
        summaryLines.push('Deleted assignments:');
        summaryLines.push(...formatList(summary.deletedAssignments, detailLimit, item =>
            `- ${item.uid}`
        ));
    }
    if (summary.notes.length) {
        summaryLines.push('Notes:');
        summaryLines.push(...summary.notes.map(note => `- ${note}`));
    }

    return {
        parsed,
        summary,
        summaryLines
    };
}
