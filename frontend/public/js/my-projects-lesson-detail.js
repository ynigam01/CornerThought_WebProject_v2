/**
 * Organizational lesson learned: full structure view (My Projects).
 * Fetch + render only; drag/drop is UI preview (not persisted).
 */

let dragSourceEl = null;

export function pgByteaToUint8Array(value) {
    if (value == null) return new Uint8Array(0);
    if (value instanceof ArrayBuffer) return new Uint8Array(value);
    if (ArrayBuffer.isView(value)) {
        return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
    }
    const s = String(value).trim();
    const m = s.match(/^\\x([0-9a-fA-F]*)$/);
    if (m) {
        const hex = m[1];
        if (hex.length % 2 !== 0) return new Uint8Array(0);
        const out = new Uint8Array(hex.length / 2);
        for (let i = 0; i < out.length; i += 1) {
            out[i] = parseInt(hex.substr(i * 2, 2), 16);
        }
        return out;
    }
    return new Uint8Array(0);
}

export function formatMetadataRows(rows) {
    const safe = Array.isArray(rows) ? rows : [];
    return safe
        .map((row) => {
            const type = row && row.metadata_type ? String(row.metadata_type).trim() : '';
            let meta = row && row.metadata;
            if (meta && typeof meta === 'object') {
                try {
                    meta = JSON.stringify(meta);
                } catch (_) {
                    meta = String(meta);
                }
            } else if (meta != null) {
                meta = String(meta);
            } else {
                meta = '';
            }
            if (type && meta) return `${type}: ${meta}`;
            return type || meta || '';
        })
        .filter(Boolean);
}

/**
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {{ organizationId: string|number|null, projectId: string|number|null, lessonId: string|number|null }} params
 */
export async function fetchLessonStructure(supabase, { organizationId, projectId, lessonId }) {
    const orgId = organizationId;
    const pid = projectId;
    const lid = lessonId;
    if (orgId == null || pid == null || lid == null) {
        throw new Error('Missing organization, project, or lesson.');
    }

    const filterLesson = (q) =>
        q
            .eq('lessons_learned_id', lid)
            .eq('organization_id', orgId)
            .eq('project_id', pid);

    const [
        causesResp,
        impactsResp,
        actionsResp,
        fpcResp,
        notesResp,
        metaResp,
        attachResp,
    ] = await Promise.all([
        filterLesson(supabase.from('lessons_learned_causes').select('id, cause')),
        filterLesson(supabase.from('lessons_learned_impacts').select('id, impact')),
        filterLesson(
            supabase
                .from('action_items')
                .select('id, action_item, lessons_learned_cause_id, lessons_learned_impact_id')
        ),
        filterLesson(
            supabase
                .from('future_project_considerations')
                .select('id, fpc, lessons_learned_cause_id, lessons_learned_impact_id')
        ),
        filterLesson(supabase.from('lessons_learned_notes').select('id, notes')),
        filterLesson(
            supabase
                .from('lessons_learned_metadata')
                .select('id, metadata, metadata_type, lessons_learned_metadata_list_id')
        ),
        filterLesson(supabase.from('lessons_learned_attachments').select('id, file_name, content_type')),
    ]);

    const responses = [
        causesResp,
        impactsResp,
        actionsResp,
        fpcResp,
        notesResp,
        metaResp,
        attachResp,
    ];
    const failed = responses.find((r) => r.error);
    if (failed && failed.error) {
        throw new Error(failed.error.message || 'Failed to load lesson details.');
    }

    return {
        causes: Array.isArray(causesResp.data) ? causesResp.data : [],
        impacts: Array.isArray(impactsResp.data) ? impactsResp.data : [],
        actions: Array.isArray(actionsResp.data) ? actionsResp.data : [],
        fpcs: Array.isArray(fpcResp.data) ? fpcResp.data : [],
        notes: Array.isArray(notesResp.data) ? notesResp.data : [],
        metadata: Array.isArray(metaResp.data) ? metaResp.data : [],
        attachments: Array.isArray(attachResp.data) ? attachResp.data : [],
    };
}

function wireDraggable(el) {
    el.setAttribute('draggable', 'true');
    el.addEventListener('dragstart', (e) => {
        dragSourceEl = el;
        e.dataTransfer.effectAllowed = 'move';
        try {
            e.dataTransfer.setData('text/plain', el.dataset.dragKind || 'item');
        } catch (_) {
            // IE / legacy
        }
    });
    el.addEventListener('dragend', () => {
        dragSourceEl = null;
        document.querySelectorAll('.is-org-lesson-drag-over').forEach((n) => {
            n.classList.remove('is-org-lesson-drag-over');
        });
    });
}

function applyAssignedCardStyle(el) {
    el.classList.remove('org-lesson-draggable-card--nested', 'org-lesson-draggable-card--fpc');
    el.classList.add('org-lesson-draggable-card--nested');
    if (el.dataset.dragKind === 'fpc') {
        el.classList.add('org-lesson-draggable-card--fpc');
    }
}

function applyUnassignedCardStyle(el) {
    el.classList.remove('org-lesson-draggable-card--nested', 'org-lesson-draggable-card--fpc');
}

/**
 * @param {HTMLElement} zoneEl
 * @param {HTMLElement} appendRoot
 */
function wireDropZone(zoneEl, appendRoot) {
    const isUnassignPool = appendRoot.classList.contains('org-lesson-unassigned-card-grid');

    zoneEl.addEventListener('dragover', (e) => {
        if (!dragSourceEl) return;
        e.preventDefault();
        zoneEl.classList.add('is-org-lesson-drag-over');
    });
    zoneEl.addEventListener('dragleave', (e) => {
        if (e.currentTarget.contains(e.relatedTarget)) return;
        zoneEl.classList.remove('is-org-lesson-drag-over');
    });
    zoneEl.addEventListener('drop', (e) => {
        e.preventDefault();
        zoneEl.classList.remove('is-org-lesson-drag-over');
        if (!dragSourceEl) return;
        if (isUnassignPool) {
            applyUnassignedCardStyle(dragSourceEl);
        } else {
            applyAssignedCardStyle(dragSourceEl);
        }
        appendRoot.appendChild(dragSourceEl);
    });
}

function applyExpandState(headerBtn, bodyEl, toggleEl, expanded) {
    bodyEl.style.display = expanded ? 'block' : 'none';
    if (toggleEl) {
        toggleEl.innerHTML = expanded ? '&#9650;' : '&#9660;';
    }
    headerBtn.classList.toggle('is-expanded', !!expanded);
}

function buildDraggableItem(kind, text) {
    const card = document.createElement('div');
    card.className = 'org-lesson-draggable-card';
    card.dataset.dragKind = kind;
    const badge = document.createElement('span');
    badge.className = 'org-lesson-kind-badge';
    badge.textContent = kind === 'action' ? 'Action item' : 'Future consideration';
    const body = document.createElement('div');
    body.textContent = text || '';
    card.appendChild(badge);
    card.appendChild(body);
    wireDraggable(card);
    return card;
}

function buildCauseImpactSection({ title, sectionClass, items, textKey, idKey, actions, fpcs, linkKey }) {
    const section = document.createElement('section');
    section.className = `lesson-detail-section ${sectionClass}`;

    const header = document.createElement('div');
    header.className = 'lesson-detail-section-header';
    header.textContent = title;
    section.appendChild(header);

    items.forEach((item) => {
        const itemId = item[idKey];
        const row = document.createElement('div');
        row.className = 'lesson-detail-item';

        const hdr = document.createElement('button');
        hdr.type = 'button';
        hdr.className = 'lesson-detail-item-header';

        const lbl = document.createElement('span');
        lbl.className = 'lesson-detail-item-label';
        lbl.textContent = item[textKey] || '';

        const tg = document.createElement('span');
        tg.className = 'lesson-detail-toggle';

        hdr.appendChild(lbl);
        hdr.appendChild(tg);

        const body = document.createElement('div');
        body.className = 'lesson-detail-item-body org-lesson-item-body';

        const list = document.createElement('div');
        list.className = 'org-lesson-assigned-list';

        const assignedActions = actions.filter((a) => a && a[linkKey] === itemId);
        const assignedFpcs = fpcs.filter((f) => f && f[linkKey] === itemId);
        assignedActions.forEach((a) => {
            list.appendChild(buildDraggableItem('action', a.action_item));
            applyAssignedCardStyle(list.lastChild);
        });
        assignedFpcs.forEach((f) => {
            list.appendChild(buildDraggableItem('fpc', f.fpc));
            applyAssignedCardStyle(list.lastChild);
        });

        body.appendChild(list);
        const hint = document.createElement('div');
        hint.className = 'org-lesson-drop-hint';
        hint.textContent =
            'Drop action items or future considerations here. Preview only — not saved to the database.';
        body.appendChild(hint);

        wireDropZone(body, list);

        let expanded = assignedActions.length > 0 || assignedFpcs.length > 0;
        applyExpandState(hdr, body, tg, expanded);
        hdr.addEventListener('click', () => {
            expanded = body.style.display !== 'block';
            applyExpandState(hdr, body, tg, expanded);
        });

        row.appendChild(hdr);
        row.appendChild(body);
        section.appendChild(row);
    });

    return section;
}

function splitUnassigned(rows, causeKey, impactKey) {
    return rows.filter((r) => r && !r[causeKey] && !r[impactKey]);
}

/**
 * @param {HTMLElement} container
 * @param {Awaited<ReturnType<typeof fetchLessonStructure>>} detail
 * @param {{ supabase: import('@supabase/supabase-js').SupabaseClient, organizationId: string|number|null, projectId: string|number|null }} ctx
 */
export function renderLessonStructureInto(container, detail, ctx) {
    const { supabase, organizationId, projectId } = ctx;
    container.innerHTML = '';

    const shell = document.createElement('div');
    shell.className = 'org-lesson-detail-shell';

    const causes = detail.causes || [];
    const impacts = detail.impacts || [];
    const actions = detail.actions || [];
    const fpcs = detail.fpcs || [];

    const unassignedActions = splitUnassigned(
        actions,
        'lessons_learned_cause_id',
        'lessons_learned_impact_id'
    );
    const unassignedFpcs = splitUnassigned(
        fpcs,
        'lessons_learned_cause_id',
        'lessons_learned_impact_id'
    );

    if (causes.length) {
        shell.appendChild(
            buildCauseImpactSection({
                title: 'Cause(s):',
                sectionClass: 'lesson-detail-cause-card',
                items: causes,
                textKey: 'cause',
                idKey: 'id',
                actions,
                fpcs,
                linkKey: 'lessons_learned_cause_id',
            })
        );
    }

    if (impacts.length) {
        shell.appendChild(
            buildCauseImpactSection({
                title: 'Impact(s):',
                sectionClass: 'lesson-detail-impact-card',
                items: impacts,
                textKey: 'impact',
                idKey: 'id',
                actions,
                fpcs,
                linkKey: 'lessons_learned_impact_id',
            })
        );
    }

    if (unassignedActions.length || unassignedFpcs.length) {
        const pool = document.createElement('section');
        pool.className = 'org-lesson-unassigned-section';
        const ph = document.createElement('h4');
        ph.textContent = 'Unassigned action items & considerations';
        pool.appendChild(ph);
        const sub = document.createElement('p');
        sub.className = 'org-lesson-unassigned-hint';
        sub.textContent =
            'These are not linked to a cause or impact. Drag them onto a cause or impact above, or use the zone below to unassign. Preview only — not saved.';
        pool.appendChild(sub);
        const cardGrid = document.createElement('div');
        cardGrid.className = 'org-lesson-unassigned-grid org-lesson-unassigned-card-grid';
        unassignedActions.forEach((a) => cardGrid.appendChild(buildDraggableItem('action', a.action_item)));
        unassignedFpcs.forEach((f) => cardGrid.appendChild(buildDraggableItem('fpc', f.fpc)));
        pool.appendChild(cardGrid);

        const unassignReceiver = document.createElement('div');
        unassignReceiver.className = 'org-lesson-unassign-receiver';
        const dh = document.createElement('p');
        dh.className = 'org-lesson-drop-hint';
        dh.style.margin = '0 0 8px';
        dh.textContent = 'Drop here to move an item back to unassigned (preview only).';
        unassignReceiver.appendChild(dh);
        wireDropZone(unassignReceiver, cardGrid);
        pool.appendChild(unassignReceiver);

        shell.appendChild(pool);
    }

    const notes = detail.notes || [];
    if (notes.length) {
        const ns = document.createElement('section');
        ns.className = 'org-lesson-notes-section';
        const nh = document.createElement('h4');
        nh.textContent = 'Notes';
        ns.appendChild(nh);
        const ul = document.createElement('ul');
        ul.className = 'org-lesson-notes-list';
        notes.forEach((n) => {
            const li = document.createElement('li');
            li.textContent = n.notes || '';
            ul.appendChild(li);
        });
        ns.appendChild(ul);
        shell.appendChild(ns);
    }

    const metaLines = formatMetadataRows(detail.metadata);
    const attachments = detail.attachments || [];

    if (metaLines.length || attachments.length) {
        const footer = document.createElement('div');
        footer.className = 'org-lesson-footer';

        if (metaLines.length) {
            const mh = document.createElement('h4');
            mh.textContent = 'Metadata';
            footer.appendChild(mh);
            const row = document.createElement('div');
            row.className = 'org-lesson-footer-row';
            const lab = document.createElement('span');
            lab.className = 'org-lesson-footer-label';
            lab.textContent = 'Entries:';
            row.appendChild(lab);
            row.appendChild(document.createTextNode(` ${metaLines.join(' · ')}`));
            footer.appendChild(row);
        }

        if (attachments.length) {
            const ah = document.createElement('h4');
            ah.textContent = 'Attachments';
            footer.appendChild(ah);
            attachments.forEach((att) => {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'org-lesson-attachment-link';
                btn.textContent = att.file_name || 'Download';
                btn.addEventListener('click', async () => {
                    try {
                        const attachId = att && att.id;
                        if (attachId == null) return;
                        const { data, error } = await supabase
                            .from('lessons_learned_attachments')
                            .select('file_data, file_name, content_type')
                            .eq('id', attachId)
                            .eq('organization_id', organizationId)
                            .eq('project_id', projectId)
                            .maybeSingle();
                        if (error) throw error;
                        if (!data || data.file_data == null) {
                            throw new Error('No file data returned.');
                        }
                        const bytes = pgByteaToUint8Array(data.file_data);
                        const blob = new Blob([bytes], {
                            type: data.content_type || 'application/octet-stream',
                        });
                        const url = URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = data.file_name || att.file_name || 'download';
                        document.body.appendChild(a);
                        a.click();
                        a.remove();
                        URL.revokeObjectURL(url);
                    } catch (err) {
                        console.error(err);
                        alert(err.message || 'Download failed.');
                    }
                });
                footer.appendChild(btn);
            });
        }

        shell.appendChild(footer);
    }

    if (!shell.childNodes.length) {
        const empty = document.createElement('p');
        empty.className = 'org-lesson-detail-loading';
        empty.textContent =
            'No causes, impacts, unassigned items, notes, metadata, or attachments for this lesson yet.';
        shell.appendChild(empty);
    }

    container.appendChild(shell);
}

/**
 * Primary headline (Issue / Success + title) for full-page lesson view.
 * @param {{ category?: unknown, title?: unknown }} row
 */
export function buildLessonPrimaryTitle(row) {
    const titleDiv = document.createElement('div');
    titleDiv.className = 'lesson-detail-title';

    const categoryRaw = row && row.category ? String(row.category).trim() : '';
    const categoryLower = categoryRaw.toLowerCase();
    const categoryLabel =
        categoryLower === 'success' ? 'Success' : categoryLower === 'issue' ? 'Issue' : 'Lesson';
    const title = row && row.title ? String(row.title).trim() : '(Untitled)';

    const strong = document.createElement('strong');
    strong.textContent = `${categoryLabel}: `;
    titleDiv.appendChild(strong);
    titleDiv.appendChild(document.createTextNode(title));
    return titleDiv;
}

export function isLessonDraftForEditing(row) {
    return (
        String(row && row.review ? row.review : '')
            .trim()
            .toLowerCase()
            .replace(/_/g, ' ') === 'draft'
    );
}

function normalizeLessonReviewValue(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/_/g, ' ');
}

function isLessonForReviewMyProjects(row) {
    return normalizeLessonReviewValue(row && row.review) === 'for review';
}

/**
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string|number|null|undefined} organizationId
 * @param {string|number|null|undefined} projectId
 * @param {string|number|null|undefined} userId
 * @returns {Promise<Set<string>>}
 */
async function loadAssignedMetadataListIdsForMyProjects(supabase, organizationId, projectId, userId) {
    const set = new Set();
    if (organizationId == null || projectId == null || userId == null) return set;
    const { data, error } = await supabase
        .from('project_team_member_assignments')
        .select('lessons_learned_metadata_list_id')
        .eq('organization_id', organizationId)
        .eq('project_id', projectId)
        .eq('user_id', userId)
        .limit(5000);
    if (error) {
        throw new Error(error.message || 'Failed to verify lesson access.');
    }
    (Array.isArray(data) ? data : []).forEach((r) => {
        const id = r && r.lessons_learned_metadata_list_id;
        if (id != null) set.add(String(id));
    });
    return set;
}

/**
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string|number|null|undefined} organizationId
 * @param {string|number|null|undefined} projectId
 * @param {unknown} lessonId
 * @returns {Promise<string[]>}
 */
async function loadLessonMetadataListIdsForProject(
    supabase,
    organizationId,
    projectId,
    lessonId
) {
    if (organizationId == null || projectId == null || lessonId == null) return [];
    const { data, error } = await supabase
        .from('lessons_learned_metadata')
        .select('lessons_learned_metadata_list_id')
        .eq('organization_id', organizationId)
        .eq('project_id', projectId)
        .eq('lessons_learned_id', lessonId)
        .limit(5000);
    if (error) {
        throw new Error(error.message || 'Failed to verify lesson access.');
    }
    return Array.from(
        new Set(
            (Array.isArray(data) ? data : [])
                .map((r) => (r && r.lessons_learned_metadata_list_id != null ? String(r.lessons_learned_metadata_list_id) : ''))
                .filter(Boolean)
        )
    );
}

/**
 * @param {{ id?: unknown, review?: unknown, created_by?: unknown }} row
 * @param {{ project_id?: unknown }} project
 * @param {{ supabase: import('@supabase/supabase-js').SupabaseClient, organizationId: string|number|null, projectId?: string|number|null, userId?: string|number|null, isProjectManager?: boolean, assignedMetadataListIds?: Set<string>|null|undefined }} ctx
 */
async function assertMyProjectsForReviewLessonAccess(row, project, ctx) {
    if (!isLessonForReviewMyProjects(row)) return;
    const { supabase, organizationId, projectId: ctxPid, userId, isProjectManager, assignedMetadataListIds } = ctx;
    const pid = project && project.project_id != null ? project.project_id : ctxPid;
    if (isProjectManager === true) return;
    if (userId != null && String(row && row.created_by) === String(userId)) return;

    let assigned =
        assignedMetadataListIds instanceof Set ? assignedMetadataListIds : null;
    if (!assigned) {
        assigned = await loadAssignedMetadataListIdsForMyProjects(supabase, organizationId, pid, userId);
    }

    const metaIds = await loadLessonMetadataListIdsForProject(supabase, organizationId, pid, row.id);
    if (metaIds.some((id) => assigned.has(id))) return;

    throw new Error('You do not have permission to view this lesson.');
}

/**
 * Renders the full lesson into mountEl (caller provides back navigation outside this tree).
 * @param {HTMLElement} mountEl
 * @param {{ id?: unknown, category?: unknown, title?: unknown, review?: unknown }} row
 * @param {{ project_id?: unknown, project_type_id?: unknown }} project
 * @param {{ supabase: import('@supabase/supabase-js').SupabaseClient, organizationId: string|number|null, projectId: string|number|null, userId?: string|number|null, projectTypeId?: string|number|null, onLessonReviewSaved?: () => void, isProjectManager?: boolean, assignedMetadataListIds?: Set<string>|null }} ctx
 */
export async function mountLessonFullPage(mountEl, row, project, ctx) {
    const { supabase, organizationId, projectId } = ctx;

    if (isLessonDraftForEditing(row)) {
        const { mountDraftLessonEditor } = await import('./my-projects-lesson-draft-editor.js');
        return mountDraftLessonEditor(mountEl, row, project, ctx);
    }

    mountEl.innerHTML = '';

    const card = document.createElement('article');
    card.className = 'lesson-detail-card org-lesson-full-page-card';
    card.appendChild(buildLessonPrimaryTitle(row));

    const detailHost = document.createElement('div');
    card.appendChild(detailHost);
    mountEl.appendChild(card);

    try {
        await assertMyProjectsForReviewLessonAccess(row, project, ctx);
    } catch (accessErr) {
        const errEl = document.createElement('div');
        errEl.className = 'upload-message upload-message--error';
        errEl.textContent =
            (accessErr && accessErr.message) || 'You do not have permission to view this lesson.';
        detailHost.appendChild(errEl);
        return;
    }

    const loading = document.createElement('div');
    loading.className = 'org-lesson-detail-loading';
    loading.textContent = 'Loading lesson structure…';
    detailHost.appendChild(loading);

    try {
        const detail = await fetchLessonStructure(supabase, {
            organizationId,
            projectId: project.project_id,
            lessonId: row.id,
        });
        detailHost.innerHTML = '';
        renderLessonStructureInto(detailHost, detail, {
            supabase,
            organizationId,
            projectId,
        });
    } catch (err) {
        detailHost.innerHTML = '';
        const errEl = document.createElement('div');
        errEl.className = 'upload-message upload-message--error';
        errEl.textContent = (err && err.message) || 'Failed to load lesson details.';
        detailHost.appendChild(errEl);
    }
}

/**
 * @param {{ id?: unknown, category?: unknown, title?: unknown }} row
 * @param {{ project_id?: unknown }} project
 * @param {{ onOpenLesson?: (args: { row: typeof row, project: typeof project }) => void }} deps
 */
export function createMyProjectsLessonWrap(row, project, deps) {
    const { onOpenLesson } = deps;
    const wrap = document.createElement('div');
    wrap.className = 'my-projects-lesson-wrap';

    const categoryRaw = row && row.category ? String(row.category).trim() : '';
    const categoryLower = categoryRaw.toLowerCase();
    const categoryDisplay = categoryRaw
        ? categoryRaw.charAt(0).toUpperCase() + categoryRaw.slice(1)
        : 'Issue';
    const title = row && row.title ? String(row.title).trim() : '(Untitled)';

    const card = document.createElement('div');
    card.className = 'my-projects-lesson-card my-projects-lesson-card--clickable';
    if (categoryLower === 'success') {
        card.classList.add('my-projects-lesson-card--success');
    } else {
        card.classList.add('my-projects-lesson-card--issue');
    }
    const label = document.createElement('strong');
    label.textContent = `${categoryDisplay}: `;
    card.appendChild(label);
    card.appendChild(document.createTextNode(title));

    card.setAttribute('role', 'button');
    card.tabIndex = 0;
    card.setAttribute(
        'aria-label',
        `Open full lesson, ${categoryDisplay}: ${title}`
    );

    const open = () => {
        if (typeof onOpenLesson === 'function') {
            onOpenLesson({ row, project });
        }
    };
    card.addEventListener('click', open);
    card.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            open();
        }
    });

    wrap.appendChild(card);
    return wrap;
}
