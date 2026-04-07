/**
 * Draft lesson editor: My Projects full-page view for lessons with review = draft.
 * Persists CRUD + drag-drop assignments immediately; Save Draft / Send for Review update review status.
 */

import {
    fetchLessonStructure,
    buildLessonPrimaryTitle,
    pgByteaToUint8Array,
    formatMetadataRows,
} from './my-projects-lesson-detail.js';

const MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024;

function arrayBufferToPgBytea(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    let hex = '';
    for (let i = 0; i < bytes.length; i += 1) {
        hex += bytes[i].toString(16).padStart(2, '0');
    }
    return `\\x${hex}`;
}

function formatListMetadataLabel(row) {
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
    return type || meta || `Tag ${row && row.id != null ? row.id : ''}`;
}

/**
 * @param {{
 *   title: string,
 *   label: string,
 *   initialValue?: string,
 *   mode: 'create'|'edit',
 *   primaryCreateLabel?: string,
 *   primaryEditLabel?: string,
 *   showDelete?: boolean,
 * }} opts
 */
function openLessonDraftTextModal(opts) {
    const {
        title,
        label,
        initialValue = '',
        mode,
        primaryCreateLabel = 'Create',
        primaryEditLabel = 'Make Change',
        showDelete = true,
    } = opts;

    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal show lesson-draft-dialog';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');

        const content = document.createElement('div');
        content.className = 'modal-content lesson-draft-dialog-content';

        const head = document.createElement('div');
        head.className = 'lesson-draft-dialog-header';
        const h = document.createElement('h3');
        h.textContent = title;
        const closeBtn = document.createElement('button');
        closeBtn.type = 'button';
        closeBtn.className = 'lesson-draft-dialog-close';
        closeBtn.setAttribute('aria-label', 'Close');
        closeBtn.innerHTML = '&times;';
        head.appendChild(h);
        head.appendChild(closeBtn);

        const body = document.createElement('div');
        body.className = 'lesson-draft-dialog-body';
        const lab = document.createElement('label');
        lab.textContent = label;
        const ta = document.createElement('textarea');
        ta.className = 'lesson-draft-textarea';
        ta.value = initialValue;
        ta.rows = 4;
        body.appendChild(lab);
        body.appendChild(ta);

        const actions = document.createElement('div');
        actions.className = 'lesson-draft-dialog-actions';

        const cancelBtn = document.createElement('button');
        cancelBtn.type = 'button';
        cancelBtn.className = 'secondary-button';
        cancelBtn.textContent = 'Cancel';

        const primaryBtn = document.createElement('button');
        primaryBtn.type = 'button';
        primaryBtn.className = 'save-lessons-button';
        primaryBtn.textContent = mode === 'edit' ? primaryEditLabel : primaryCreateLabel;

        let deleteBtn = null;
        if (mode === 'edit' && showDelete) {
            deleteBtn = document.createElement('button');
            deleteBtn.type = 'button';
            deleteBtn.className = 'logout-button';
            deleteBtn.style.padding = '8px 14px';
            deleteBtn.textContent = 'Delete';
        }

        actions.appendChild(cancelBtn);
        if (deleteBtn) actions.appendChild(deleteBtn);
        actions.appendChild(primaryBtn);

        content.appendChild(head);
        content.appendChild(body);
        content.appendChild(actions);
        overlay.appendChild(content);
        document.body.appendChild(overlay);

        function cleanup() {
            overlay.remove();
        }

        function finish(result) {
            cleanup();
            resolve(result);
        }

        closeBtn.addEventListener('click', () => finish({ action: 'cancel' }));
        cancelBtn.addEventListener('click', () => finish({ action: 'cancel' }));
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) finish({ action: 'cancel' });
        });

        primaryBtn.addEventListener('click', () => {
            finish({ action: 'save', value: ta.value.trim() });
        });

        if (deleteBtn) {
            deleteBtn.addEventListener('click', () => {
                finish({ action: 'delete' });
            });
        }

        setTimeout(() => ta.focus(), 0);
    });
}

async function unassignCauseDependencies(supabase, orgId, pid, causeId) {
    await supabase
        .from('action_items')
        .update({ lessons_learned_cause_id: null })
        .eq('organization_id', orgId)
        .eq('project_id', pid)
        .eq('lessons_learned_cause_id', causeId);
    await supabase
        .from('future_project_considerations')
        .update({ lessons_learned_cause_id: null })
        .eq('organization_id', orgId)
        .eq('project_id', pid)
        .eq('lessons_learned_cause_id', causeId);
}

async function unassignImpactDependencies(supabase, orgId, pid, impactId) {
    await supabase
        .from('action_items')
        .update({ lessons_learned_impact_id: null })
        .eq('organization_id', orgId)
        .eq('project_id', pid)
        .eq('lessons_learned_impact_id', impactId);
    await supabase
        .from('future_project_considerations')
        .update({ lessons_learned_impact_id: null })
        .eq('organization_id', orgId)
        .eq('project_id', pid)
        .eq('lessons_learned_impact_id', impactId);
}

/**
 * @param {HTMLElement} mountEl
 * @param {object} row
 * @param {object} project
 * @param {object} ctx
 */
export async function mountDraftLessonEditor(mountEl, row, project, ctx) {
    const {
        supabase,
        organizationId: orgId,
        projectId: ctxPid,
        userId,
        projectTypeId: ctxTypeId,
        onLessonReviewSaved,
    } = ctx;

    const lessonId = row && row.id;
    const pid = ctxPid != null ? ctxPid : project && project.project_id;
    const projectTypeId =
        ctxTypeId != null ? ctxTypeId : project && project.project_type_id != null
            ? project.project_type_id
            : null;

    if (lessonId == null || orgId == null || pid == null || userId == null) {
        mountEl.innerHTML = '';
        const err = document.createElement('div');
        err.className = 'upload-message upload-message--error';
        err.textContent = 'Cannot edit draft: missing lesson, project, or user context.';
        mountEl.appendChild(err);
        return;
    }

    let lessonRowState = {
        id: lessonId,
        title: row.title,
        category: row.category,
        review: row.review,
    };

    mountEl.innerHTML = '';
    const card = document.createElement('article');
    card.className = 'lesson-detail-card org-lesson-full-page-card org-lesson-draft-editor';

    const toolbar = document.createElement('div');
    toolbar.className = 'lesson-draft-toolbar';
    const btnSaveDraft = document.createElement('button');
    btnSaveDraft.type = 'button';
    btnSaveDraft.className = 'save-lessons-button';
    btnSaveDraft.textContent = 'Save Draft';
    const btnSendReview = document.createElement('button');
    btnSendReview.type = 'button';
    btnSendReview.className = 'save-lessons-button';
    btnSendReview.textContent = 'Send for Review';
    const toolbarStatus = document.createElement('div');
    toolbarStatus.className = 'lesson-draft-toolbar-status upload-message';
    toolbarStatus.setAttribute('aria-live', 'polite');
    toolbar.appendChild(btnSaveDraft);
    toolbar.appendChild(btnSendReview);
    toolbar.appendChild(toolbarStatus);

    const titleRow = document.createElement('div');
    titleRow.className = 'lesson-draft-title-row';
    const titleHost = document.createElement('div');
    titleHost.className = 'lesson-draft-title-host';
    const editTitleBtn = document.createElement('button');
    editTitleBtn.type = 'button';
    editTitleBtn.className = 'lesson-draft-section-add';
    editTitleBtn.textContent = 'Edit';
    titleRow.appendChild(titleHost);
    titleRow.appendChild(editTitleBtn);

    function renderTitleIntoHost() {
        titleHost.innerHTML = '';
        titleHost.appendChild(buildLessonPrimaryTitle(lessonRowState));
    }
    renderTitleIntoHost();

    const detailHost = document.createElement('div');
    detailHost.className = 'lesson-draft-detail-host';

    card.appendChild(toolbar);
    card.appendChild(titleRow);
    card.appendChild(detailHost);
    mountEl.appendChild(card);

    function setToolbarStatus(msg, isError = false) {
        toolbarStatus.textContent = msg || '';
        toolbarStatus.classList.remove('upload-message--success', 'upload-message--error');
        if (!msg) return;
        toolbarStatus.classList.add(isError ? 'upload-message--error' : 'upload-message--success');
    }

    let editorDragEl = null;

    function wireDraggableEditor(el) {
        el.setAttribute('draggable', 'true');
        el.addEventListener('dragstart', (e) => {
            editorDragEl = el;
            e.dataTransfer.effectAllowed = 'move';
            try {
                e.dataTransfer.setData('text/plain', el.dataset.dragKind || 'item');
            } catch (_) {}
        });
        el.addEventListener('dragend', () => {
            editorDragEl = null;
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

    async function persistCardAssignment(card, listEl) {
        const table = card.dataset.itemTable;
        const id = card.dataset.itemId;
        if (!table || !id) return;

        let causeId = null;
        let impactId = null;
        const c = listEl.getAttribute('data-assign-cause-id');
        const i = listEl.getAttribute('data-assign-impact-id');
        if (c) causeId = c;
        else if (i) impactId = i;
        else {
            const inUnassigned = listEl.closest('.org-lesson-unassigned-card-grid');
            if (!inUnassigned) return;
        }

        const patch = {
            lessons_learned_cause_id: causeId,
            lessons_learned_impact_id: impactId,
        };

        const { error } = await supabase.from(table).update(patch).eq('id', id).eq('organization_id', orgId).eq('project_id', pid);
        if (error) throw error;
    }

    function wireDropZoneEditor(zoneEl, appendRoot) {
        const isUnassignPool = appendRoot.classList.contains('org-lesson-unassigned-card-grid');

        zoneEl.addEventListener('dragover', (e) => {
            if (!editorDragEl) return;
            e.preventDefault();
            zoneEl.classList.add('is-org-lesson-drag-over');
        });
        zoneEl.addEventListener('dragleave', (e) => {
            if (e.currentTarget.contains(e.relatedTarget)) return;
            zoneEl.classList.remove('is-org-lesson-drag-over');
        });
        zoneEl.addEventListener('drop', async (e) => {
            e.preventDefault();
            zoneEl.classList.remove('is-org-lesson-drag-over');
            if (!editorDragEl) return;
            if (isUnassignPool) {
                applyUnassignedCardStyle(editorDragEl);
            } else {
                applyAssignedCardStyle(editorDragEl);
            }
            appendRoot.appendChild(editorDragEl);
            const dragged = editorDragEl;
            try {
                await persistCardAssignment(dragged, appendRoot);
                setToolbarStatus('Assignment updated.', false);
            } catch (err) {
                console.error(err);
                setToolbarStatus(err.message || 'Failed to save assignment.', true);
                await refreshDetail();
            }
        });
    }

    async function refreshDetail() {
        detailHost.innerHTML = '';
        const loading = document.createElement('div');
        loading.className = 'org-lesson-detail-loading';
        loading.textContent = 'Loading…';
        detailHost.appendChild(loading);
        try {
            const detail = await fetchLessonStructure(supabase, {
                organizationId: orgId,
                projectId: pid,
                lessonId,
            });
            detailHost.innerHTML = '';
            renderDraftDetail(detailHost, detail);
        } catch (err) {
            detailHost.innerHTML = '';
            const errEl = document.createElement('div');
            errEl.className = 'upload-message upload-message--error';
            errEl.textContent = (err && err.message) || 'Failed to load lesson.';
            detailHost.appendChild(errEl);
        }
    }

    function buildEditorDraggableItem(kind, text, dbRow, tableName, onEditClick) {
        const card = document.createElement('div');
        card.className = 'org-lesson-draggable-card';
        card.dataset.dragKind = kind;
        card.dataset.itemTable = tableName;
        card.dataset.itemId = String(dbRow.id);
        const top = document.createElement('div');
        top.className = 'lesson-draft-card-head';
        const badge = document.createElement('span');
        badge.className = 'org-lesson-kind-badge';
        badge.textContent = kind === 'action' ? 'Action item' : 'Future consideration';
        const editBtn = document.createElement('button');
        editBtn.type = 'button';
        editBtn.className = 'lesson-draft-section-add lesson-draft-card-edit';
        editBtn.textContent = 'Edit';
        editBtn.addEventListener('click', (ev) => {
            ev.stopPropagation();
            onEditClick(dbRow, card);
        });
        top.appendChild(badge);
        top.appendChild(editBtn);
        const body = document.createElement('div');
        body.textContent = text || '';
        card.appendChild(top);
        card.appendChild(body);
        wireDraggableEditor(card);
        return card;
    }

    function applyExpandState(headerBtn, bodyEl, toggleEl, expanded) {
        bodyEl.style.display = expanded ? 'block' : 'none';
        if (toggleEl) {
            toggleEl.innerHTML = expanded ? '&#9650;' : '&#9660;';
        }
        headerBtn.classList.toggle('is-expanded', !!expanded);
    }

    function buildCauseImpactSectionEditable({ title, sectionClass, items, textKey, idKey, actions, fpcs, linkKey, kind }) {
        const section = document.createElement('section');
        section.className = `lesson-detail-section ${sectionClass}`;

        const headRow = document.createElement('div');
        headRow.className = 'lesson-draft-section-heading-row';
        const sh = document.createElement('div');
        sh.className = 'lesson-detail-section-header lesson-draft-section-title';
        sh.textContent = title;
        const addBtn = document.createElement('button');
        addBtn.type = 'button';
        addBtn.className = 'lesson-draft-section-add';
        addBtn.textContent = kind === 'cause' ? 'Add cause' : 'Add impact';
        addBtn.addEventListener('click', async () => {
            const r = await openLessonDraftTextModal({
                title: kind === 'cause' ? 'Add cause' : 'Add impact',
                label: kind === 'cause' ? 'Cause' : 'Impact',
                mode: 'create',
            });
            if (r.action !== 'save' || !r.value) return;
            try {
                const table = kind === 'cause' ? 'lessons_learned_causes' : 'lessons_learned_impacts';
                const col = kind === 'cause' ? 'cause' : 'impact';
                const { error } = await supabase.from(table).insert({
                    lessons_learned_id: lessonId,
                    [col]: r.value,
                    created_by: userId,
                    organization_id: orgId,
                    project_id: pid,
                });
                if (error) throw error;
                setToolbarStatus('Added.', false);
                await refreshDetail();
            } catch (err) {
                console.error(err);
                setToolbarStatus(err.message || 'Save failed.', true);
            }
        });
        headRow.appendChild(sh);
        headRow.appendChild(addBtn);
        section.appendChild(headRow);

        if (!items.length) {
            const empty = document.createElement('p');
            empty.className = 'lesson-draft-empty-hint';
            empty.textContent = `No ${kind === 'cause' ? 'causes' : 'impacts'} yet. Use Add to create one.`;
            section.appendChild(empty);
            return section;
        }

        items.forEach((item) => {
            const itemId = item[idKey];
            const rowWrap = document.createElement('div');
            rowWrap.className = 'lesson-detail-item';

            const hdr = document.createElement('div');
            hdr.className = 'lesson-detail-item-header lesson-draft-cause-impact-header';

            const hdrBtn = document.createElement('button');
            hdrBtn.type = 'button';
            hdrBtn.className = 'lesson-detail-item-header lesson-draft-item-expand-btn';
            const lbl = document.createElement('span');
            lbl.className = 'lesson-detail-item-label';
            lbl.textContent = item[textKey] || '';
            const tg = document.createElement('span');
            tg.className = 'lesson-detail-toggle';
            hdrBtn.appendChild(lbl);
            hdrBtn.appendChild(tg);

            const editHdr = document.createElement('button');
            editHdr.type = 'button';
            editHdr.className = 'lesson-draft-section-add';
            editHdr.textContent = 'Edit';
            editHdr.addEventListener('click', async (e) => {
                e.stopPropagation();
                const r = await openLessonDraftTextModal({
                    title: kind === 'cause' ? 'Edit cause' : 'Edit impact',
                    label: kind === 'cause' ? 'Cause' : 'Impact',
                    initialValue: item[textKey] || '',
                    mode: 'edit',
                });
                const table = kind === 'cause' ? 'lessons_learned_causes' : 'lessons_learned_impacts';
                const col = kind === 'cause' ? 'cause' : 'impact';
                try {
                    if (r.action === 'delete') {
                        if (kind === 'cause') await unassignCauseDependencies(supabase, orgId, pid, itemId);
                        else await unassignImpactDependencies(supabase, orgId, pid, itemId);
                        const { error } = await supabase.from(table).delete().eq('id', itemId).eq('organization_id', orgId);
                        if (error) throw error;
                    } else if (r.action === 'save' && r.value) {
                        const { error } = await supabase
                            .from(table)
                            .update({ [col]: r.value })
                            .eq('id', itemId)
                            .eq('organization_id', orgId);
                        if (error) throw error;
                    } else return;
                    setToolbarStatus('Updated.', false);
                    await refreshDetail();
                } catch (err) {
                    console.error(err);
                    setToolbarStatus(err.message || 'Update failed.', true);
                }
            });

            hdr.appendChild(hdrBtn);
            hdr.appendChild(editHdr);

            const body = document.createElement('div');
            body.className = 'lesson-detail-item-body org-lesson-item-body';

            const list = document.createElement('div');
            list.className = 'org-lesson-assigned-list';
            if (kind === 'cause') list.setAttribute('data-assign-cause-id', String(itemId));
            else list.setAttribute('data-assign-impact-id', String(itemId));

            const assignedActions = actions.filter((a) => a && a[linkKey] === itemId);
            const assignedFpcs = fpcs.filter((f) => f && f[linkKey] === itemId);

            assignedActions.forEach((a) => {
                list.appendChild(
                    buildEditorDraggableItem('action', a.action_item, a, 'action_items', async (db) => {
                        const r = await openLessonDraftTextModal({
                            title: 'Edit action item',
                            label: 'Action item',
                            initialValue: db.action_item || '',
                            mode: 'edit',
                        });
                        try {
                            if (r.action === 'delete') {
                                const { error } = await supabase
                                    .from('action_items')
                                    .delete()
                                    .eq('id', db.id)
                                    .eq('organization_id', orgId);
                                if (error) throw error;
                            } else if (r.action === 'save') {
                                const { error } = await supabase
                                    .from('action_items')
                                    .update({ action_item: r.value })
                                    .eq('id', db.id)
                                    .eq('organization_id', orgId);
                                if (error) throw error;
                            } else return;
                            setToolbarStatus('Updated.', false);
                            await refreshDetail();
                        } catch (err) {
                            console.error(err);
                            setToolbarStatus(err.message || 'Update failed.', true);
                        }
                    })
                );
                applyAssignedCardStyle(list.lastChild);
            });
            assignedFpcs.forEach((f) => {
                list.appendChild(
                    buildEditorDraggableItem('fpc', f.fpc, f, 'future_project_considerations', async (db) => {
                        const r = await openLessonDraftTextModal({
                            title: 'Edit future consideration',
                            label: 'Consideration',
                            initialValue: db.fpc || '',
                            mode: 'edit',
                        });
                        try {
                            if (r.action === 'delete') {
                                const { error } = await supabase
                                    .from('future_project_considerations')
                                    .delete()
                                    .eq('id', db.id)
                                    .eq('organization_id', orgId);
                                if (error) throw error;
                            } else if (r.action === 'save') {
                                const { error } = await supabase
                                    .from('future_project_considerations')
                                    .update({ fpc: r.value })
                                    .eq('id', db.id)
                                    .eq('organization_id', orgId);
                                if (error) throw error;
                            } else return;
                            setToolbarStatus('Updated.', false);
                            await refreshDetail();
                        } catch (err) {
                            console.error(err);
                            setToolbarStatus(err.message || 'Update failed.', true);
                        }
                    })
                );
                applyAssignedCardStyle(list.lastChild);
            });

            body.appendChild(list);
            const hint = document.createElement('div');
            hint.className = 'org-lesson-drop-hint';
            hint.textContent = 'Drop action items or future considerations here to assign them.';
            body.appendChild(hint);
            wireDropZoneEditor(body, list);

            let expanded = assignedActions.length > 0 || assignedFpcs.length > 0;
            applyExpandState(hdrBtn, body, tg, expanded);
            hdrBtn.addEventListener('click', () => {
                expanded = body.style.display !== 'block';
                applyExpandState(hdrBtn, body, tg, expanded);
            });

            rowWrap.appendChild(hdr);
            rowWrap.appendChild(body);
            section.appendChild(rowWrap);
        });

        return section;
    }

    function renderDraftDetail(shell, detail) {
        const causes = detail.causes || [];
        const impacts = detail.impacts || [];
        const actions = detail.actions || [];
        const fpcs = detail.fpcs || [];

        const unassignedActions = actions.filter((r) => r && !r.lessons_learned_cause_id && !r.lessons_learned_impact_id);
        const unassignedFpcs = fpcs.filter((r) => r && !r.lessons_learned_cause_id && !r.lessons_learned_impact_id);

        const root = document.createElement('div');
        root.className = 'org-lesson-detail-shell';

        root.appendChild(
            buildCauseImpactSectionEditable({
                title: 'Cause(s):',
                sectionClass: 'lesson-detail-cause-card',
                items: causes,
                textKey: 'cause',
                idKey: 'id',
                actions,
                fpcs,
                linkKey: 'lessons_learned_cause_id',
                kind: 'cause',
            })
        );

        root.appendChild(
            buildCauseImpactSectionEditable({
                title: 'Impact(s):',
                sectionClass: 'lesson-detail-impact-card',
                items: impacts,
                textKey: 'impact',
                idKey: 'id',
                actions,
                fpcs,
                linkKey: 'lessons_learned_impact_id',
                kind: 'impact',
            })
        );

        const pool = document.createElement('section');
        pool.className = 'org-lesson-unassigned-section';
        const poolHead = document.createElement('div');
        poolHead.className = 'lesson-draft-section-heading-row';
        const ph = document.createElement('h4');
        ph.textContent = 'Unassigned action items & considerations';
        const poolBtns = document.createElement('div');
        poolBtns.className = 'lesson-draft-inline-actions';
        const addUnA = document.createElement('button');
        addUnA.type = 'button';
        addUnA.className = 'lesson-draft-section-add';
        addUnA.textContent = 'Add action item';
        const addUnF = document.createElement('button');
        addUnF.type = 'button';
        addUnF.className = 'lesson-draft-section-add';
        addUnF.textContent = 'Add consideration';
        poolBtns.appendChild(addUnA);
        poolBtns.appendChild(addUnF);
        poolHead.appendChild(ph);
        poolHead.appendChild(poolBtns);
        pool.appendChild(poolHead);

        addUnA.addEventListener('click', async () => {
            const r = await openLessonDraftTextModal({
                title: 'Add action item',
                label: 'Action item',
                mode: 'create',
            });
            if (r.action !== 'save' || !r.value) return;
            try {
                const { error } = await supabase.from('action_items').insert({
                    lessons_learned_id: lessonId,
                    action_item: r.value,
                    lessons_learned_cause_id: null,
                    lessons_learned_impact_id: null,
                    created_by: userId,
                    organization_id: orgId,
                    project_id: pid,
                });
                if (error) throw error;
                setToolbarStatus('Added.', false);
                await refreshDetail();
            } catch (err) {
                console.error(err);
                setToolbarStatus(err.message || 'Save failed.', true);
            }
        });

        addUnF.addEventListener('click', async () => {
            const r = await openLessonDraftTextModal({
                title: 'Add future consideration',
                label: 'Consideration',
                mode: 'create',
            });
            if (r.action !== 'save' || !r.value) return;
            try {
                const { error } = await supabase.from('future_project_considerations').insert({
                    lessons_learned_id: lessonId,
                    fpc: r.value,
                    lessons_learned_cause_id: null,
                    lessons_learned_impact_id: null,
                    created_by: userId,
                    organization_id: orgId,
                    project_id: pid,
                });
                if (error) throw error;
                setToolbarStatus('Added.', false);
                await refreshDetail();
            } catch (err) {
                console.error(err);
                setToolbarStatus(err.message || 'Save failed.', true);
            }
        });

        const sub = document.createElement('p');
        sub.className = 'org-lesson-unassigned-hint';
        sub.textContent =
            'Drag items onto a cause or impact above. Drop here to unassign.';
        pool.appendChild(sub);
        const cardGrid = document.createElement('div');
        cardGrid.className = 'org-lesson-unassigned-grid org-lesson-unassigned-card-grid';
        unassignedActions.forEach((a) => {
            cardGrid.appendChild(
                buildEditorDraggableItem('action', a.action_item, a, 'action_items', async (db) => {
                    const r = await openLessonDraftTextModal({
                        title: 'Edit action item',
                        label: 'Action item',
                        initialValue: db.action_item || '',
                        mode: 'edit',
                    });
                    try {
                        if (r.action === 'delete') {
                            const { error } = await supabase
                                .from('action_items')
                                .delete()
                                .eq('id', db.id)
                                .eq('organization_id', orgId);
                            if (error) throw error;
                        } else if (r.action === 'save') {
                            const { error } = await supabase
                                .from('action_items')
                                .update({ action_item: r.value })
                                .eq('id', db.id)
                                .eq('organization_id', orgId);
                            if (error) throw error;
                        } else return;
                        setToolbarStatus('Updated.', false);
                        await refreshDetail();
                    } catch (err) {
                        console.error(err);
                        setToolbarStatus(err.message || 'Update failed.', true);
                    }
                })
            );
            applyUnassignedCardStyle(cardGrid.lastChild);
        });
        unassignedFpcs.forEach((f) => {
            cardGrid.appendChild(
                buildEditorDraggableItem('fpc', f.fpc, f, 'future_project_considerations', async (db) => {
                    const r = await openLessonDraftTextModal({
                        title: 'Edit future consideration',
                        label: 'Consideration',
                        initialValue: db.fpc || '',
                        mode: 'edit',
                    });
                    try {
                        if (r.action === 'delete') {
                            const { error } = await supabase
                                .from('future_project_considerations')
                                .delete()
                                .eq('id', db.id)
                                .eq('organization_id', orgId);
                            if (error) throw error;
                        } else if (r.action === 'save') {
                            const { error } = await supabase
                                .from('future_project_considerations')
                                .update({ fpc: r.value })
                                .eq('id', db.id)
                                .eq('organization_id', orgId);
                            if (error) throw error;
                        } else return;
                        setToolbarStatus('Updated.', false);
                        await refreshDetail();
                    } catch (err) {
                        console.error(err);
                        setToolbarStatus(err.message || 'Update failed.', true);
                    }
                })
            );
            applyUnassignedCardStyle(cardGrid.lastChild);
        });
        pool.appendChild(cardGrid);

        const unassignReceiver = document.createElement('div');
        unassignReceiver.className = 'org-lesson-unassign-receiver';
        const dh = document.createElement('p');
        dh.className = 'org-lesson-drop-hint';
        dh.style.margin = '0 0 8px';
        dh.textContent = 'Drop here to move an item back to unassigned.';
        unassignReceiver.appendChild(dh);
        wireDropZoneEditor(unassignReceiver, cardGrid);
        pool.appendChild(unassignReceiver);
        root.appendChild(pool);

        const ns = document.createElement('section');
        ns.className = 'org-lesson-notes-section';
        const notesHead = document.createElement('div');
        notesHead.className = 'lesson-draft-section-heading-row';
        const nh = document.createElement('h4');
        nh.textContent = 'Notes';
        const addNoteBtn = document.createElement('button');
        addNoteBtn.type = 'button';
        addNoteBtn.className = 'lesson-draft-section-add';
        addNoteBtn.textContent = 'Add note';
        notesHead.appendChild(nh);
        notesHead.appendChild(addNoteBtn);
        ns.appendChild(notesHead);

        addNoteBtn.addEventListener('click', async () => {
            const r = await openLessonDraftTextModal({
                title: 'Add note',
                label: 'Note',
                mode: 'create',
            });
            if (r.action !== 'save' || !r.value) return;
            try {
                const { error } = await supabase.from('lessons_learned_notes').insert({
                    lessons_learned_id: lessonId,
                    notes: r.value,
                    created_by: userId,
                    organization_id: orgId,
                    project_id: pid,
                });
                if (error) throw error;
                setToolbarStatus('Added.', false);
                await refreshDetail();
            } catch (err) {
                console.error(err);
                setToolbarStatus(err.message || 'Save failed.', true);
            }
        });

        const notes = detail.notes || [];
        if (!notes.length) {
            const ne = document.createElement('p');
            ne.className = 'lesson-draft-empty-hint';
            ne.textContent = 'No notes yet.';
            ns.appendChild(ne);
        } else {
            const ul = document.createElement('ul');
            ul.className = 'org-lesson-notes-list lesson-draft-notes-editable';
            notes.forEach((n) => {
                const li = document.createElement('li');
                li.className = 'lesson-draft-note-row';
                const span = document.createElement('span');
                span.textContent = n.notes || '';
                const eb = document.createElement('button');
                eb.type = 'button';
                eb.className = 'lesson-draft-section-add';
                eb.textContent = 'Edit';
                eb.addEventListener('click', async () => {
                    const r = await openLessonDraftTextModal({
                        title: 'Edit note',
                        label: 'Note',
                        initialValue: n.notes || '',
                        mode: 'edit',
                    });
                    try {
                        if (r.action === 'delete') {
                            const { error } = await supabase
                                .from('lessons_learned_notes')
                                .delete()
                                .eq('id', n.id)
                                .eq('organization_id', orgId);
                            if (error) throw error;
                        } else if (r.action === 'save') {
                            const { error } = await supabase
                                .from('lessons_learned_notes')
                                .update({ notes: r.value })
                                .eq('id', n.id)
                                .eq('organization_id', orgId);
                            if (error) throw error;
                        } else return;
                        setToolbarStatus('Updated.', false);
                        await refreshDetail();
                    } catch (err) {
                        console.error(err);
                        setToolbarStatus(err.message || 'Update failed.', true);
                    }
                });
                li.appendChild(span);
                li.appendChild(eb);
                ul.appendChild(li);
            });
            ns.appendChild(ul);
        }
        root.appendChild(ns);

        const footer = document.createElement('div');
        footer.className = 'org-lesson-footer lesson-draft-footer-always';

        const metaHead = document.createElement('div');
        metaHead.className = 'lesson-draft-section-heading-row';
        const mh = document.createElement('h4');
        mh.textContent = 'Metadata';
        const metaBtn = document.createElement('button');
        metaBtn.type = 'button';
        metaBtn.className = 'lesson-draft-section-add';
        metaBtn.textContent = 'Add/Remove Metadata';
        metaHead.appendChild(mh);
        metaHead.appendChild(metaBtn);
        footer.appendChild(metaHead);

        const metaRows = detail.metadata || [];
        const metaSummary = document.createElement('div');
        metaSummary.className = 'org-lesson-footer-row lesson-draft-meta-summary';
        const lines = formatMetadataRows(metaRows);
        metaSummary.textContent = lines.length ? lines.join(' · ') : 'No metadata tags yet.';
        footer.appendChild(metaSummary);

        metaBtn.addEventListener('click', () => {
            openDraftMetadataModal({
                supabase,
                orgId,
                pid,
                lessonId,
                userId,
                projectTypeId,
                currentLinks: metaRows,
                onChanged: async () => {
                    setToolbarStatus('Metadata updated.', false);
                    await refreshDetail();
                },
                setToolbarStatus,
            });
        });

        const ah = document.createElement('h4');
        ah.textContent = 'Attachments';
        footer.appendChild(ah);

        const attachRow = document.createElement('div');
        attachRow.className = 'lesson-draft-attach-row';
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.className = 'lesson-draft-file-input';
        fileInput.multiple = true;
        attachRow.appendChild(fileInput);
        footer.appendChild(attachRow);

        fileInput.addEventListener('change', async () => {
            const files = Array.from(fileInput.files || []);
            fileInput.value = '';
            if (!files.length) return;
            for (const file of files) {
                if (!file.size) continue;
                if (file.size > MAX_ATTACHMENT_BYTES) {
                    setToolbarStatus(`"${file.name}" exceeds 10 MB.`, true);
                    continue;
                }
                try {
                    const buf = await file.arrayBuffer();
                    const { error } = await supabase.from('lessons_learned_attachments').insert({
                        lessons_learned_id: lessonId,
                        project_id: pid,
                        organization_id: orgId,
                        created_by: userId,
                        file_data: arrayBufferToPgBytea(buf),
                        file_name: file.name || 'attachment',
                        content_type: file.type || 'application/octet-stream',
                    });
                    if (error) throw error;
                } catch (err) {
                    console.error(err);
                    setToolbarStatus(err.message || 'Attachment upload failed.', true);
                    return;
                }
            }
            setToolbarStatus('Attachment(s) added.', false);
            await refreshDetail();
        });

        const attachments = detail.attachments || [];
        if (!attachments.length) {
            const ae = document.createElement('p');
            ae.className = 'lesson-draft-empty-hint';
            ae.textContent = 'No attachments yet.';
            footer.appendChild(ae);
        } else {
            attachments.forEach((att) => {
                const row = document.createElement('div');
                row.className = 'lesson-draft-attachment-row';
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
                            .eq('organization_id', orgId)
                            .eq('project_id', pid)
                            .maybeSingle();
                        if (error) throw error;
                        if (!data || data.file_data == null) throw new Error('No file data returned.');
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
                const rm = document.createElement('button');
                rm.type = 'button';
                rm.className = 'lesson-draft-section-add lesson-draft-remove-attach';
                rm.textContent = 'Remove';
                rm.addEventListener('click', async () => {
                    if (!confirm('Remove this attachment?')) return;
                    try {
                        const { error } = await supabase
                            .from('lessons_learned_attachments')
                            .delete()
                            .eq('id', att.id)
                            .eq('organization_id', orgId);
                        if (error) throw error;
                        setToolbarStatus('Attachment removed.', false);
                        await refreshDetail();
                    } catch (err) {
                        console.error(err);
                        setToolbarStatus(err.message || 'Remove failed.', true);
                    }
                });
                row.appendChild(btn);
                row.appendChild(rm);
                footer.appendChild(row);
            });
        }

        root.appendChild(footer);
        shell.appendChild(root);
    }

    editTitleBtn.addEventListener('click', async () => {
        const r = await openLessonDraftTextModal({
            title: 'Edit lesson title',
            label: 'Title',
            initialValue: lessonRowState.title ? String(lessonRowState.title) : '',
            mode: 'edit',
            primaryEditLabel: 'Make Change',
            showDelete: false,
        });
        if (r.action !== 'save' || !r.value) return;
        try {
            const { error } = await supabase
                .from('lessons_learned')
                .update({ title: r.value })
                .eq('id', lessonId)
                .eq('organization_id', orgId);
            if (error) throw error;
            lessonRowState.title = r.value;
            renderTitleIntoHost();
            setToolbarStatus('Title updated.', false);
        } catch (err) {
            console.error(err);
            setToolbarStatus(err.message || 'Update failed.', true);
        }
    });

    btnSaveDraft.addEventListener('click', async () => {
        try {
            btnSaveDraft.disabled = true;
            btnSendReview.disabled = true;
            const { error } = await supabase
                .from('lessons_learned')
                .update({ review: 'draft' })
                .eq('id', lessonId)
                .eq('organization_id', orgId);
            if (error) throw error;
            lessonRowState.review = 'draft';
            setToolbarStatus('Draft status saved.', false);
            if (typeof onLessonReviewSaved === 'function') onLessonReviewSaved();
        } catch (err) {
            console.error(err);
            setToolbarStatus(err.message || 'Could not update status.', true);
        } finally {
            btnSaveDraft.disabled = false;
            btnSendReview.disabled = false;
        }
    });

    btnSendReview.addEventListener('click', async () => {
        try {
            btnSaveDraft.disabled = true;
            btnSendReview.disabled = true;
            const { error } = await supabase
                .from('lessons_learned')
                .update({ review: 'for review' })
                .eq('id', lessonId)
                .eq('organization_id', orgId);
            if (error) throw error;
            lessonRowState.review = 'for review';
            if (typeof onLessonReviewSaved === 'function') onLessonReviewSaved();
            const mod = await import('./my-projects-lesson-detail.js');
            mountEl.innerHTML = '';
            await mod.mountLessonFullPage(
                mountEl,
                { ...lessonRowState, review: 'for review' },
                project,
                ctx
            );
        } catch (err) {
            console.error(err);
            setToolbarStatus(err.message || 'Could not send for review.', true);
            btnSaveDraft.disabled = false;
            btnSendReview.disabled = false;
        }
    });

    await refreshDetail();
}

async function openDraftMetadataModal({
    supabase,
    orgId,
    pid,
    lessonId,
    userId,
    projectTypeId,
    currentLinks,
    onChanged,
    setToolbarStatus,
}) {
    const assignedSet = new Set(
        (currentLinks || [])
            .map((r) => r && r.lessons_learned_metadata_list_id)
            .filter((v) => v != null)
            .map((v) => String(v))
    );

    const overlay = document.createElement('div');
    overlay.className = 'modal show lesson-draft-dialog lesson-draft-metadata-modal';
    const content = document.createElement('div');
    content.className = 'modal-content add-data-metadata-modal-content lesson-draft-metadata-content';

    const head = document.createElement('div');
    head.className = 'modal-header-row';
    const h = document.createElement('h3');
    h.textContent = 'Add/Remove Metadata';
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'close-button';
    closeBtn.innerHTML = '&times;';
    head.appendChild(h);
    head.appendChild(closeBtn);

    const sub = document.createElement('p');
    sub.className = 'subtitle';
    sub.textContent = 'Tag this lesson learned with metadata for the selected project.';

    const currentBlock = document.createElement('div');
    currentBlock.className = 'lesson-draft-metadata-current';
    const curLabel = document.createElement('label');
    curLabel.textContent = 'Current tags';
    const curList = document.createElement('ul');
    curList.className = 'lesson-draft-metadata-tag-list';

    function rebuildCurrentList() {
        curList.innerHTML = '';
        (currentLinks || []).forEach((link) => {
            if (!link || link.id == null) return;
            const li = document.createElement('li');
            li.className = 'lesson-draft-metadata-tag-item';
            const t = document.createElement('span');
            t.textContent = formatListMetadataLabel(link);
            const rb = document.createElement('button');
            rb.type = 'button';
            rb.className = 'lesson-draft-section-add';
            rb.textContent = 'Remove';
            rb.addEventListener('click', async () => {
                try {
                    const { error } = await supabase
                        .from('lessons_learned_metadata')
                        .delete()
                        .eq('id', link.id)
                        .eq('organization_id', orgId);
                    if (error) throw error;
                    const idx = currentLinks.indexOf(link);
                    if (idx >= 0) currentLinks.splice(idx, 1);
                    assignedSet.delete(
                        String(link.lessons_learned_metadata_list_id || '')
                    );
                    rebuildCurrentList();
                    setToolbarStatus('Metadata tag removed.', false);
                    await onChanged();
                } catch (err) {
                    console.error(err);
                    setToolbarStatus(err.message || 'Remove failed.', true);
                }
            });
            li.appendChild(t);
            li.appendChild(rb);
            curList.appendChild(li);
        });
        if (!curList.childNodes.length) {
            const empty = document.createElement('li');
            empty.className = 'lesson-draft-empty-hint';
            empty.textContent = 'No tags yet.';
            curList.appendChild(empty);
        }
    }

    const selectWrap = document.createElement('div');
    selectWrap.className = 'input-group';
    const metaLab = document.createElement('label');
    metaLab.setAttribute('for', 'lessonDraftMetadataSelect');
    metaLab.textContent = 'Metadata';
    const select = document.createElement('select');
    select.id = 'lessonDraftMetadataSelect';
    select.multiple = true;
    selectWrap.appendChild(metaLab);
    selectWrap.appendChild(select);

    const statusEl = document.createElement('div');
    statusEl.className = 'upload-message';
    statusEl.setAttribute('aria-live', 'polite');

    const actions = document.createElement('div');
    actions.className = 'modal-actions';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Cancel';
    const applyBtn = document.createElement('button');
    applyBtn.type = 'button';
    applyBtn.textContent = 'Apply';
    actions.appendChild(cancelBtn);
    actions.appendChild(applyBtn);

    currentBlock.appendChild(curLabel);
    currentBlock.appendChild(curList);

    content.appendChild(head);
    content.appendChild(sub);
    content.appendChild(currentBlock);
    content.appendChild(selectWrap);
    content.appendChild(statusEl);
    content.appendChild(actions);
    overlay.appendChild(content);
    document.body.appendChild(overlay);

    let choicesInst = null;

    function closeModal() {
        try {
            if (choicesInst && typeof choicesInst.destroy === 'function') {
                choicesInst.destroy();
            }
        } catch (_) {}
        overlay.remove();
    }

    closeBtn.addEventListener('click', closeModal);
    cancelBtn.addEventListener('click', closeModal);
    overlay.addEventListener('click', (e) => {
        if (e.target === overlay) closeModal();
    });

    rebuildCurrentList();

    statusEl.textContent = 'Loading metadata options…';

    let metadataListOptions = [];

    try {
        const { data: assignedRows } = await supabase
            .from('project_team_member_assignments')
            .select('lessons_learned_metadata_list_id')
            .eq('organization_id', orgId)
            .eq('project_id', pid)
            .eq('user_id', userId)
            .limit(5000);

        const assignedIds = new Set(
            (assignedRows || [])
                .map((r) => r && r.lessons_learned_metadata_list_id)
                .filter((v) => v != null)
                .map((v) => String(v))
        );

        const { data: metadataRows, error: listErr } = await supabase
            .from('lessons_learned_metadata_list')
            .select('id, metadata_type, metadata')
            .eq('organization_id', orgId)
            .eq('project_id', pid)
            .order('id', { ascending: true })
            .limit(5000);

        if (listErr) throw listErr;
        const rows = Array.isArray(metadataRows) ? metadataRows : [];
        const assigned = rows.filter((r) => assignedIds.has(String(r.id)));
        const unassigned = rows.filter((r) => !assignedIds.has(String(r.id)));
        metadataListOptions = [...assigned, ...unassigned];

        select.innerHTML = '';
        metadataListOptions.forEach((r) => {
            const opt = document.createElement('option');
            opt.value = String(r.id);
            opt.textContent = formatListMetadataLabel(r);
            select.appendChild(opt);
        });

        if (typeof window.Choices === 'function') {
            choicesInst = new window.Choices(select, {
                removeItemButton: true,
                searchEnabled: true,
                shouldSort: false,
                placeholderValue: 'Select metadata',
            });
        }

        const choiceData = metadataListOptions.map((r) => ({
            value: String(r.id),
            label: formatListMetadataLabel(r),
            selected: assignedSet.has(String(r.id)),
        }));

        if (choicesInst && typeof choicesInst.setChoices === 'function') {
            choicesInst.clearStore();
            choicesInst.setChoices(choiceData, 'value', 'label', true);
        } else {
            Array.from(select.options).forEach((o) => {
                o.selected = assignedSet.has(o.value);
            });
        }

        statusEl.textContent = '';
    } catch (err) {
        console.error(err);
        statusEl.textContent = err.message || 'Failed to load metadata.';
        statusEl.classList.add('upload-message--error');
    }

    applyBtn.addEventListener('click', async () => {
        let selectedIds = [];
        if (choicesInst && typeof choicesInst.getValue === 'function') {
            const val = choicesInst.getValue(true);
            if (Array.isArray(val)) selectedIds = val.map((v) => String(v));
            else if (val) selectedIds = [String(val)];
        } else {
            selectedIds = Array.from(select.selectedOptions).map((o) => String(o.value));
        }

        const toInsert = selectedIds.filter((id) => id && !assignedSet.has(id));
        if (!toInsert.length) {
            closeModal();
            return;
        }

        applyBtn.disabled = true;
        statusEl.classList.remove('upload-message--error');
        statusEl.textContent = 'Applying…';

        try {
            for (const listIdStr of toInsert) {
                const row = metadataListOptions.find((r) => String(r.id) === listIdStr);
                if (!row) continue;
                const metaVal =
                    row.metadata && typeof row.metadata === 'object' ? row.metadata : row.metadata;
                const { data: ins, error } = await supabase
                    .from('lessons_learned_metadata')
                    .insert({
                        lessons_learned_id: lessonId,
                        metadata_type: row.metadata_type || null,
                        metadata: metaVal,
                        lessons_learned_metadata_list_id: row.id,
                        created_by: userId,
                        organization_id: orgId,
                        project_id: pid,
                        project_type_id: projectTypeId,
                    })
                    .select('id, metadata, metadata_type, lessons_learned_metadata_list_id')
                    .single();
                if (error) throw error;
                if (ins) {
                    currentLinks.push(ins);
                    assignedSet.add(listIdStr);
                }
            }
            setToolbarStatus('Metadata applied.', false);
            await onChanged();
            closeModal();
        } catch (err) {
            console.error(err);
            statusEl.textContent = err.message || 'Apply failed.';
            statusEl.classList.add('upload-message--error');
        } finally {
            applyBtn.disabled = false;
        }
    });
}
