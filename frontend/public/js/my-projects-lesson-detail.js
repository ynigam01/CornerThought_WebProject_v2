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
        filterLesson(supabase.from('lessons_learned_causes').select('id, cause, created_by')),
        filterLesson(supabase.from('lessons_learned_impacts').select('id, impact, created_by')),
        filterLesson(
            supabase
                .from('action_items')
                .select(
                    'id, action_item, status, lessons_learned_cause_id, lessons_learned_impact_id, created_by'
                )
        ),
        filterLesson(
            supabase
                .from('future_project_considerations')
                .select(
                    'id, fpc, lessons_learned_cause_id, lessons_learned_impact_id, created_by'
                )
        ),
        filterLesson(supabase.from('lessons_learned_notes').select('id, notes, created_by')),
        filterLesson(
            supabase
                .from('lessons_learned_metadata')
                .select('id, metadata, metadata_type, lessons_learned_metadata_list_id, created_by')
        ),
        filterLesson(
            supabase.from('lessons_learned_attachments').select('id, file_name, content_type, created_by')
        ),
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

function sameId(a, b) {
    if (a == null || b == null) return false;
    return String(a) === String(b);
}

function actionPrefixForLlm(status) {
    return String(status || '')
        .trim()
        .toLowerCase() === 'recommended'
        ? 'Recommended Action'
        : 'Action Taken';
}

function appendLinkedItemsForLlm(lines, parentId, actions, fpcs, linkKey) {
    const linkedActions = (actions || []).filter((a) => a && sameId(a[linkKey], parentId));
    const linkedFpcs = (fpcs || []).filter((f) => f && sameId(f[linkKey], parentId));
    linkedActions.forEach((a) => {
        lines.push(`-\t${actionPrefixForLlm(a.status)}: ${a.action_item || ''}`);
    });
    linkedFpcs.forEach((f) => {
        lines.push(`-\tLesson: ${f.fpc || ''}`);
    });
}

/**
 * Builds plain text for a lesson learned (issue/success + nested structure)
 * suitable for sending to an LLM.
 * @param {{ category?: unknown, title?: unknown }} lessonRow
 * @param {{
 *   causes?: Array<{ id?: unknown, cause?: unknown }>,
 *   impacts?: Array<{ id?: unknown, impact?: unknown }>,
 *   actions?: Array<{
 *     action_item?: unknown,
 *     status?: unknown,
 *     lessons_learned_cause_id?: unknown,
 *     lessons_learned_impact_id?: unknown,
 *   }>,
 *   fpcs?: Array<{
 *     fpc?: unknown,
 *     lessons_learned_cause_id?: unknown,
 *     lessons_learned_impact_id?: unknown,
 *   }>,
 *   notes?: Array<{ notes?: unknown }>,
 *   metadata?: Array<{ metadata_type?: unknown, metadata?: unknown }>,
 * }} detail
 * @returns {string}
 */
export function formatLessonForLlm(lessonRow, detail) {
    const lines = [];
    const categoryRaw = lessonRow && lessonRow.category ? String(lessonRow.category).trim() : '';
    const categoryLower = categoryRaw.toLowerCase();
    const categoryLabel =
        categoryLower === 'success' ? 'Success' : categoryLower === 'issue' ? 'Issue' : 'Lesson';
    const title =
        lessonRow && lessonRow.title != null && String(lessonRow.title).trim() !== ''
            ? String(lessonRow.title).trim()
            : '(Untitled)';
    lines.push(`${categoryLabel}: ${title}`);

    const causes = detail && Array.isArray(detail.causes) ? detail.causes : [];
    const impacts = detail && Array.isArray(detail.impacts) ? detail.impacts : [];
    const actions = detail && Array.isArray(detail.actions) ? detail.actions : [];
    const fpcs = detail && Array.isArray(detail.fpcs) ? detail.fpcs : [];
    const notes = detail && Array.isArray(detail.notes) ? detail.notes : [];
    const metadata = detail && Array.isArray(detail.metadata) ? detail.metadata : [];

    causes.forEach((cause, i) => {
        lines.push(`Cause ${i + 1}: ${cause && cause.cause != null ? String(cause.cause) : ''}`);
        appendLinkedItemsForLlm(lines, cause && cause.id, actions, fpcs, 'lessons_learned_cause_id');
    });

    impacts.forEach((impact, i) => {
        lines.push(`Impact ${i + 1}: ${impact && impact.impact != null ? String(impact.impact) : ''}`);
        appendLinkedItemsForLlm(
            lines,
            impact && impact.id,
            actions,
            fpcs,
            'lessons_learned_impact_id'
        );
    });

    const unassignedActions = actions.filter(
        (a) => a && !a.lessons_learned_cause_id && !a.lessons_learned_impact_id
    );
    const unassignedFpcs = fpcs.filter(
        (f) => f && !f.lessons_learned_cause_id && !f.lessons_learned_impact_id
    );
    if (unassignedActions.length || unassignedFpcs.length) {
        lines.push('Unassigned:');
        unassignedActions.forEach((a) => {
            lines.push(`-\t${actionPrefixForLlm(a.status)}: ${a.action_item || ''}`);
        });
        unassignedFpcs.forEach((f) => {
            lines.push(`-\tLesson: ${f.fpc || ''}`);
        });
    }

    notes.forEach((note, i) => {
        lines.push(`Note ${i + 1}: ${note && note.notes != null ? String(note.notes) : ''}`);
    });

    const metaLabels = formatMetadataRows(metadata);
    if (metaLabels.length) {
        lines.push(`Metadata: ${metaLabels.join(', ')}`);
    }

    return lines.join('\n');
}

/** @type {Record<number, string>} */
let lastProjectDetailsNumberToIdMap = {};

/**
 * Temporary number → project_details.id map from the last Find Relevant run.
 * @returns {Record<number, string>}
 */
export function getLastProjectDetailsNumberToIdMap() {
    return { ...lastProjectDetailsNumberToIdMap };
}

/**
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {{ organizationId: string|number|null, projectId: string|number|null }} params
 * @returns {Promise<Array<{ id?: unknown, parameter_name?: unknown, parameter_entry?: unknown }>>}
 */
export async function fetchProjectDetailsForLlm(supabase, { organizationId, projectId }) {
    const orgId = organizationId;
    const pid = projectId;
    if (orgId == null || pid == null) {
        throw new Error('Missing organization or project.');
    }

    const { data: rows, error } = await supabase
        .from('project_details')
        .select('id, parameter_name, parameter_entry')
        .eq('organization_id', orgId)
        .eq('project_id', pid)
        .order('parameter_name', { ascending: true })
        .limit(5000);

    if (error) {
        throw new Error(error.message || 'Failed to load project details.');
    }

    return (Array.isArray(rows) ? rows : []).filter((row) => {
        if (!row) return false;
        const name = row.parameter_name != null ? String(row.parameter_name).trim() : '';
        const entry = row.parameter_entry != null ? String(row.parameter_entry).trim() : '';
        return Boolean(name || entry);
    });
}

/**
 * Builds numbered project-details text and a temporary number → id map.
 * @param {Array<{ id?: unknown, parameter_name?: unknown, parameter_entry?: unknown }>} rows
 * @returns {{ text: string, numberToIdMap: Record<number, string> }}
 */
export function formatProjectDetailsForLlm(rows) {
    const numberToIdMap = {};
    const parts = [];
    const list = Array.isArray(rows) ? rows : [];

    list.forEach((row, index) => {
        const n = index + 1;
        const name = row && row.parameter_name != null ? String(row.parameter_name).trim() : '';
        const entry = row && row.parameter_entry != null ? String(row.parameter_entry).trim() : '';
        parts.push(`${n}. ${name}: ${entry}`);
        if (row && row.id != null) {
            numberToIdMap[n] = String(row.id);
        }
    });

    return {
        text: parts.join('\n'),
        numberToIdMap,
    };
}

/**
 * Shows formatted lesson text in a centered read-only popup.
 * Includes a Find button that sends the text to Meta Llama, ranks lessons, and
 * reports only a count in the popup. Ranked results are passed to onRankedResults.
 * When embedBeforeFind is set (draft / for-review), Find stays disabled until
 * lesson embeddings finish.
 * @param {string} text
 * @param {{
 *   embedBeforeFind?: {
 *     lessonId: string|number,
 *     organizationId: string|number,
 *   } | null,
 *   rankContext?: {
 *     lessonId: string|number,
 *     organizationId: string|number,
 *     projectId: string|number,
 *   } | null,
 *   onRankedResults?: (results: Array<object>) => void,
 * }} [options]
 */
export function openLessonLlmTextPopup(text, options = {}) {
    const sourceText = String(text || '');
    const embedBeforeFind = options && options.embedBeforeFind ? options.embedBeforeFind : null;
    const rankContext = options && options.rankContext ? options.rankContext : null;
    const onRankedResults =
        options && typeof options.onRankedResults === 'function' ? options.onRankedResults : null;
    const needsEmbed =
        embedBeforeFind &&
        embedBeforeFind.lessonId != null &&
        embedBeforeFind.organizationId != null;

    const overlay = document.createElement('div');
    overlay.className = 'modal show modal--center lesson-draft-dialog lesson-llm-text-modal';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');

    const content = document.createElement('div');
    content.className = 'modal-content lesson-draft-dialog-content lesson-llm-text-modal-content';

    const head = document.createElement('div');
    head.className = 'lesson-draft-dialog-header';
    const h = document.createElement('h3');
    h.textContent = 'Find Relevant Lessons Learned';
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'lesson-draft-dialog-close';
    closeBtn.setAttribute('aria-label', 'Close');
    closeBtn.innerHTML = '&times;';
    head.appendChild(h);
    head.appendChild(closeBtn);

    const body = document.createElement('div');
    body.className = 'lesson-draft-dialog-body';
    const pre = document.createElement('pre');
    pre.className = 'lesson-llm-text-pre';
    pre.textContent = sourceText;
    body.appendChild(pre);

    const embedStatus = document.createElement('div');
    embedStatus.className = 'upload-message lesson-find-embed-status';
    embedStatus.setAttribute('aria-live', 'polite');
    embedStatus.style.display = 'none';
    body.appendChild(embedStatus);

    const actions = document.createElement('div');
    actions.className = 'lesson-draft-dialog-actions';
    const closeAction = document.createElement('button');
    closeAction.type = 'button';
    closeAction.className = 'secondary-button';
    closeAction.textContent = 'Close';
    const findAction = document.createElement('button');
    findAction.type = 'button';
    findAction.className = 'analyze-parse-button';
    findAction.textContent = 'Find';
    actions.appendChild(closeAction);
    actions.appendChild(findAction);

    content.appendChild(head);
    content.appendChild(body);
    content.appendChild(actions);
    overlay.appendChild(content);
    document.body.appendChild(overlay);

    function setEmbedStatus(message, kind) {
        embedStatus.classList.remove('upload-message--success', 'upload-message--error');
        if (!message) {
            embedStatus.style.display = 'none';
            embedStatus.textContent = '';
            return;
        }
        embedStatus.style.display = '';
        embedStatus.textContent = message;
        if (kind === 'error') {
            embedStatus.classList.add('upload-message--error');
        } else if (kind === 'success') {
            embedStatus.classList.add('upload-message--success');
        }
    }

    function cleanup() {
        overlay.remove();
    }

    closeBtn.addEventListener('click', cleanup);
    closeAction.addEventListener('click', cleanup);
    overlay.addEventListener('click', (e) => {
        if (e.target === overlay) cleanup();
    });

    findAction.addEventListener('click', async () => {
        try {
            findAction.disabled = true;
            closeAction.disabled = true;
            findAction.textContent = 'Finding...';
            setEmbedStatus('Finding relevant project parameters…', null);

            const { findRelevantLessons } = await import('./find-relevant-lessons.js');
            const responseText = await findRelevantLessons(sourceText);

            let matches = [];
            try {
                matches = parseLlamaMatchesJson(responseText);
            } catch (parseErr) {
                console.error(parseErr);
                pre.textContent =
                    parseErr && parseErr.message
                        ? parseErr.message
                        : 'Could not parse parameter matches from the model.';
                setEmbedStatus('', null);
                return;
            }

            if (
                !rankContext ||
                rankContext.lessonId == null ||
                rankContext.organizationId == null ||
                rankContext.projectId == null
            ) {
                pre.textContent = 'Found 0 relevant lessons learned.';
                if (onRankedResults) onRankedResults([]);
                setEmbedStatus('', null);
                return;
            }

            setEmbedStatus('Ranking relevant lessons…', null);
            const rankResponse = await fetch('/api/find-relevant-lessons/rank', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    lessonId: rankContext.lessonId,
                    organizationId: rankContext.organizationId,
                    projectId: rankContext.projectId,
                    matches,
                    projectDetailsNumberToIdMap: getLastProjectDetailsNumberToIdMap(),
                }),
            });

            let rankData = null;
            try {
                rankData = await rankResponse.json();
            } catch (_) {
                rankData = null;
            }

            if (!rankResponse.ok) {
                throw new Error(
                    (rankData && rankData.error) ||
                        `Failed to rank relevant lessons (${rankResponse.status}).`,
                );
            }

            const results = Array.isArray(rankData && rankData.results) ? rankData.results : [];
            const n = results.length;
            pre.textContent = `Found ${n} relevant lesson${n === 1 ? '' : 's'} learned.`;
            setEmbedStatus('', null);
            if (onRankedResults) onRankedResults(results);
        } catch (err) {
            console.error(err);
            pre.textContent = err.message || 'Find Relevant Lessons Learned failed.';
            setEmbedStatus('', null);
            alert(err.message || 'Find Relevant Lessons Learned failed.');
        } finally {
            findAction.disabled = false;
            closeAction.disabled = false;
            findAction.textContent = 'Find';
        }
    });

    if (needsEmbed) {
        findAction.disabled = true;
        setEmbedStatus(
            'Embedding this lesson learned… Find is not ready yet. Please wait.',
            null,
        );

        (async () => {
            try {
                const response = await fetch(
                    `/api/draft-lessons/${encodeURIComponent(String(embedBeforeFind.lessonId))}/embed`,
                    {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            organizationId: embedBeforeFind.organizationId,
                        }),
                    },
                );
                let data = null;
                try {
                    data = await response.json();
                } catch (_) {
                    data = null;
                }
                if (!response.ok) {
                    throw new Error(
                        (data && data.error) ||
                            `Failed to embed lesson (${response.status}).`,
                    );
                }
                setEmbedStatus(
                    'Embeddings ready. You can click Find to run Find Relevant Lessons Learned.',
                    'success',
                );
                findAction.disabled = false;
            } catch (err) {
                console.error(err);
                setEmbedStatus(
                    (err && err.message) ||
                        'Embedding failed. Find is not available until embeddings succeed.',
                    'error',
                );
                findAction.disabled = true;
            }
        })();
    }
}

/**
 * Parse Llama JSON for Find Relevant parameter matches.
 * @param {string} raw
 * @returns {Array<{ number?: number, name?: string, entry?: string }>}
 */
export function parseLlamaMatchesJson(raw) {
    let text = String(raw || '').trim();
    if (!text) throw new Error('Empty model response.');

    const fence = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fence && fence[1]) text = fence[1].trim();

    const start = text.indexOf('{');
    const end = text.lastIndexOf('}');
    if (start >= 0 && end > start) {
        text = text.slice(start, end + 1);
    }

    let parsed = null;
    try {
        parsed = JSON.parse(text);
    } catch (_) {
        throw new Error('Model response was not valid JSON.');
    }

    const matches = parsed && Array.isArray(parsed.matches) ? parsed.matches : null;
    if (!matches) {
        throw new Error('Model JSON must include a "matches" array.');
    }
    return matches;
}

/**
 * Renders ranked relevant-lesson peach cards into a host element.
 * @param {HTMLElement|null} hostEl
 * @param {Array<{ lessonId?: unknown, title?: unknown, highLevelTitle?: unknown, category?: unknown }>} results
 * @param {{ onOpen?: (lessonId: string|number, projectId?: string|number|null) => void }} [opts]
 */
export function renderRelevantLessonCards(hostEl, results, opts = {}) {
    if (!hostEl) return;
    hostEl.innerHTML = '';
    hostEl.className = 'lesson-relevant-results';
    const list = Array.isArray(results) ? results : [];
    if (!list.length) {
        hostEl.hidden = true;
        return;
    }
    hostEl.hidden = false;

    list.forEach((item) => {
        if (!item || item.lessonId == null) return;
        const card = document.createElement('div');
        card.className = 'lesson-relevant-card';
        card.setAttribute('role', 'button');
        card.tabIndex = 0;
        card.dataset.lessonId = String(item.lessonId);
        if (item.projectId != null) card.dataset.projectId = String(item.projectId);

        const dismiss = document.createElement('button');
        dismiss.type = 'button';
        dismiss.className = 'lesson-relevant-card-dismiss';
        dismiss.setAttribute('aria-label', 'Dismiss');
        dismiss.innerHTML = '&times;';
        dismiss.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            card.remove();
            if (!hostEl.querySelector('.lesson-relevant-card')) {
                hostEl.hidden = true;
                hostEl.innerHTML = '';
            }
        });

        const categoryRaw = item.category != null ? String(item.category).trim().toLowerCase() : '';
        const categoryLabel =
            categoryRaw === 'success' ? 'Success' : categoryRaw === 'issue' ? 'Issue' : 'Lesson';
        const labelText =
            item.highLevelTitle != null && String(item.highLevelTitle).trim()
                ? String(item.highLevelTitle).trim()
                : item.title != null && String(item.title).trim()
                  ? String(item.title).trim()
                  : '(Untitled)';

        const textWrap = document.createElement('div');
        textWrap.className = 'lesson-relevant-card-text';
        const strong = document.createElement('strong');
        strong.textContent = `${categoryLabel}: `;
        textWrap.appendChild(strong);
        textWrap.appendChild(document.createTextNode(labelText));

        card.appendChild(dismiss);
        card.appendChild(textWrap);

        const open = () => {
            if (typeof opts.onOpen === 'function') {
                opts.onOpen(item.lessonId, item.projectId != null ? item.projectId : null);
            }
        };
        card.addEventListener('click', open);
        card.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                open();
            }
        });

        hostEl.appendChild(card);
    });
}

/**
 * Loads lesson structure + project details and shows the LLM-formatted text popup.
 * For draft / for-review lessons, embeddings are refreshed before Find is enabled.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {{ id?: unknown, category?: unknown, title?: unknown, review?: unknown }} lessonRow
 * @param {{
 *   organizationId: string|number|null,
 *   projectId: string|number|null,
 *   onRankedResults?: (results: Array<object>) => void,
 * }} ctx
 */
export async function showFindRelevantLessonText(supabase, lessonRow, ctx) {
    const [detail, projectDetailRows] = await Promise.all([
        fetchLessonStructure(supabase, {
            organizationId: ctx.organizationId,
            projectId: ctx.projectId,
            lessonId: lessonRow && lessonRow.id,
        }),
        fetchProjectDetailsForLlm(supabase, {
            organizationId: ctx.organizationId,
            projectId: ctx.projectId,
        }),
    ]);

    const lessonText = formatLessonForLlm(lessonRow, detail);
    const { text: projectDetailsText, numberToIdMap } =
        formatProjectDetailsForLlm(projectDetailRows);
    lastProjectDetailsNumberToIdMap = { ...numberToIdMap };

    let text = lessonText;
    if (projectDetailsText) {
        text = `${lessonText}\n\nProject Details:\n${projectDetailsText}`;
    }

    const review = String((lessonRow && lessonRow.review) || '')
        .trim()
        .toLowerCase();
    const needsEmbed = review === 'draft' || review === 'for review';
    const lessonId = lessonRow && lessonRow.id != null ? lessonRow.id : null;

    openLessonLlmTextPopup(text, {
        embedBeforeFind:
            needsEmbed && lessonId != null && ctx.organizationId != null
                ? { lessonId, organizationId: ctx.organizationId }
                : null,
        rankContext:
            lessonId != null && ctx.organizationId != null && ctx.projectId != null
                ? {
                      lessonId,
                      organizationId: ctx.organizationId,
                      projectId: ctx.projectId,
                  }
                : null,
        onRankedResults: ctx.onRankedResults,
    });
}

/**
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {Iterable<string|number>} userIds
 * @returns {Promise<Map<string, string>>}
 */
export async function fetchUserNamesById(supabase, userIds) {
    const ids = Array.from(
        new Set(
            Array.from(userIds || [])
                .map((id) => (id != null ? String(id).trim() : ''))
                .filter(Boolean)
        )
    );
    const map = new Map();
    if (!ids.length) return map;
    const numericIds = ids.map((s) => {
        const n = Number(s);
        return Number.isFinite(n) ? n : s;
    });
    const { data, error } = await supabase.from('users').select('id, name').in('id', numericIds);
    if (error) {
        console.warn('fetchUserNamesById:', error.message || error);
        return map;
    }
    (Array.isArray(data) ? data : []).forEach((row) => {
        if (row && row.id != null) map.set(String(row.id), row.name ? String(row.name).trim() : '');
    });
    return map;
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
 * @param {{ skipNotes?: boolean }} [options]
 */
export function renderLessonStructureInto(container, detail, ctx, options = {}) {
    const { supabase, organizationId, projectId } = ctx;
    const skipNotes = options && options.skipNotes === true;
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
    if (!skipNotes && notes.length) {
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
 * Primary headline (Issue / Success + high-level title + full title) for full-page lesson view.
 * @param {{ category?: unknown, title?: unknown, high_level_title?: unknown }} row
 */
export function buildLessonPrimaryTitle(row) {
    const titleDiv = document.createElement('div');
    titleDiv.className = 'lesson-detail-title';

    const categoryRaw = row && row.category ? String(row.category).trim() : '';
    const categoryLower = categoryRaw.toLowerCase();
    const categoryLabel =
        categoryLower === 'success' ? 'Success' : categoryLower === 'issue' ? 'Issue' : 'Lesson';
    const title = row && row.title ? String(row.title).trim() : '(Untitled)';
    const highLevelTitle =
        row && row.high_level_title != null ? String(row.high_level_title).trim() : '';

    const strong = document.createElement('strong');
    strong.textContent = `${categoryLabel}: `;

    if (highLevelTitle) {
        const highLevelLine = document.createElement('div');
        highLevelLine.className = 'lesson-detail-high-level';
        highLevelLine.appendChild(strong);
        highLevelLine.appendChild(document.createTextNode(highLevelTitle));
        titleDiv.appendChild(highLevelLine);

        const bodyLine = document.createElement('div');
        bodyLine.className = 'lesson-detail-title-body';
        bodyLine.textContent = title;
        titleDiv.appendChild(bodyLine);
    } else {
        titleDiv.appendChild(strong);
        titleDiv.appendChild(document.createTextNode(title));
    }

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
 * @param {{ supabase: import('@supabase/supabase-js').SupabaseClient, organizationId: string|number|null, projectId?: string|number|null, userId?: string|number|null, isLessonModerator?: boolean, assignedMetadataListIds?: Set<string>|null|undefined }} ctx
 */
async function assertMyProjectsForReviewLessonAccess(row, project, ctx) {
    if (!isLessonForReviewMyProjects(row)) return;
    const { supabase, organizationId, projectId: ctxPid, userId, isLessonModerator, assignedMetadataListIds } = ctx;
    const pid = project && project.project_id != null ? project.project_id : ctxPid;
    const isCreator = userId != null && String(row && row.created_by) === String(userId);

    const metaIds = await loadLessonMetadataListIdsForProject(supabase, organizationId, pid, row.id);
    // Unassigned (no metadata for this org/project): creator only.
    if (metaIds.length === 0) {
        if (isCreator) return;
        throw new Error('You do not have permission to view this lesson.');
    }

    if (isLessonModerator === true) return;
    if (isCreator) return;

    let assigned =
        assignedMetadataListIds instanceof Set ? assignedMetadataListIds : null;
    if (!assigned) {
        assigned = await loadAssignedMetadataListIdsForMyProjects(supabase, organizationId, pid, userId);
    }

    if (metaIds.some((id) => assigned.has(id))) return;

    throw new Error('You do not have permission to view this lesson.');
}

/**
 * Renders the full lesson into mountEl (caller provides back navigation outside this tree).
 * @param {HTMLElement} mountEl
 * @param {{ id?: unknown, category?: unknown, title?: unknown, review?: unknown }} row
 * @param {{ project_id?: unknown, project_type_id?: unknown }} project
 * @param {{ supabase: import('@supabase/supabase-js').SupabaseClient, organizationId: string|number|null, projectId: string|number|null, userId?: string|number|null, projectTypeId?: string|number|null, onLessonReviewSaved?: () => void, isLessonModerator?: boolean, assignedMetadataListIds?: Set<string>|null }} ctx
 */
export async function mountLessonFullPage(mountEl, row, project, ctx) {
    const { supabase, organizationId, projectId, userId } = ctx;

    if (isLessonDraftForEditing(row)) {
        const isCreator =
            userId != null && String(row && row.created_by) === String(userId);
        if (!isCreator) {
            mountEl.innerHTML = '';
            const card = document.createElement('article');
            card.className = 'lesson-detail-card org-lesson-full-page-card';
            const err = document.createElement('div');
            err.className = 'upload-message upload-message--error';
            err.textContent = 'You do not have permission to view this lesson.';
            card.appendChild(err);
            mountEl.appendChild(card);
            return;
        }
        const { mountDraftLessonEditor } = await import('./my-projects-lesson-draft-editor.js');
        return mountDraftLessonEditor(mountEl, row, project, ctx);
    }

    if (isLessonForReviewMyProjects(row)) {
        try {
            await assertMyProjectsForReviewLessonAccess(row, project, ctx);
        } catch (accessErr) {
            mountEl.innerHTML = '';
            const card = document.createElement('article');
            card.className = 'lesson-detail-card org-lesson-full-page-card';
            card.appendChild(buildLessonPrimaryTitle(row));
            const detailHost = document.createElement('div');
            const errEl = document.createElement('div');
            errEl.className = 'upload-message upload-message--error';
            errEl.textContent =
                (accessErr && accessErr.message) || 'You do not have permission to view this lesson.';
            detailHost.appendChild(errEl);
            card.appendChild(detailHost);
            mountEl.appendChild(card);
            return;
        }

        const isModerator = ctx.isLessonModerator === true;
        const isCreator = userId != null && String(row && row.created_by) === String(userId);
        const canCollaborate = isModerator || isCreator;

        if (canCollaborate) {
            const { mountDraftLessonEditor } = await import('./my-projects-lesson-draft-editor.js');
            return mountDraftLessonEditor(mountEl, row, project, {
                ...ctx,
                forReviewCollaborative: true,
                lessonCreatorId: row && row.created_by != null ? row.created_by : null,
            });
        }

        return mountForReviewNotesOnlyLesson(mountEl, row, project, ctx);
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
 * For Review: read-only lesson body; current user may add/edit/delete only their own notes.
 * @param {HTMLElement} mountEl
 * @param {{ id?: unknown, category?: unknown, title?: unknown, review?: unknown, created_by?: unknown }} row
 * @param {{ project_id?: unknown, project_type_id?: unknown }} project
 * @param {{ supabase: import('@supabase/supabase-js').SupabaseClient, organizationId: string|number|null, projectId: string|number|null, userId?: string|number|null }} ctx
 */
async function mountForReviewNotesOnlyLesson(mountEl, row, project, ctx) {
    const { supabase, organizationId, projectId, userId } = ctx;
    const lessonId = row && row.id;
    const pid = project && project.project_id != null ? project.project_id : projectId;
    const lessonCreatorId = row && row.created_by != null ? row.created_by : null;

    mountEl.innerHTML = '';
    const card = document.createElement('article');
    card.className =
        'lesson-detail-card org-lesson-full-page-card org-lesson-for-review-notes-only org-lesson-draft-editor';

    const toolbar = document.createElement('div');
    toolbar.className = 'lesson-draft-toolbar';
    const toolbarMain = document.createElement('div');
    toolbarMain.className = 'lesson-draft-toolbar-main';
    const toolbarLeft = document.createElement('div');
    toolbarLeft.className = 'lesson-draft-toolbar-left';
    const forReviewLabel = document.createElement('div');
    forReviewLabel.className = 'lesson-for-review-toolbar-label';
    forReviewLabel.textContent = 'For review';
    const btnSaveCompleteness = document.createElement('button');
    btnSaveCompleteness.type = 'button';
    btnSaveCompleteness.className = 'save-lessons-button';
    btnSaveCompleteness.textContent = 'Save';
    const btnFindRelevant = document.createElement('button');
    btnFindRelevant.type = 'button';
    btnFindRelevant.className = 'analyze-parse-button lesson-draft-find-relevant-btn';
    btnFindRelevant.textContent = 'Find Relevant Lessons Learned';
    const toolbarStatus = document.createElement('div');
    toolbarStatus.className = 'lesson-draft-toolbar-status upload-message';
    toolbarStatus.setAttribute('aria-live', 'polite');
    toolbarLeft.appendChild(forReviewLabel);
    toolbarLeft.appendChild(btnSaveCompleteness);
    toolbarLeft.appendChild(toolbarStatus);
    toolbarMain.appendChild(toolbarLeft);
    toolbarMain.appendChild(btnFindRelevant);
    toolbar.appendChild(toolbarMain);
    card.appendChild(toolbar);

    function setToolbarStatus(msg, isError = false) {
        toolbarStatus.textContent = msg || '';
        toolbarStatus.classList.remove('upload-message--success', 'upload-message--error');
        if (!msg) return;
        toolbarStatus.classList.add(isError ? 'upload-message--error' : 'upload-message--success');
    }

    btnFindRelevant.addEventListener('click', async () => {
        try {
            btnFindRelevant.disabled = true;
            await showFindRelevantLessonText(supabase, row, {
                organizationId,
                projectId: pid,
                onRankedResults: (results) => {
                    renderRelevantLessonCards(relevantResultsHost, results, {
                        onOpen: (openedLessonId, openedProjectId) => {
                            void openRankedRelevantLesson(openedLessonId, openedProjectId);
                        },
                    });
                },
            });
        } catch (err) {
            console.error(err);
            setToolbarStatus(err.message || 'Could not build lesson text.', true);
        } finally {
            btnFindRelevant.disabled = false;
        }
    });

    const relevantResultsHost = document.createElement('div');
    relevantResultsHost.id = 'lessonRelevantResultsHost';
    relevantResultsHost.className = 'lesson-relevant-results';
    relevantResultsHost.hidden = true;

    async function openRankedRelevantLesson(openedLessonId, openedProjectIdFromCard = null) {
        try {
            const { data, error } = await supabase
                .from('lessons_learned')
                .select('id, title, high_level_title, category, review, created_by, project_id')
                .eq('id', openedLessonId)
                .eq('organization_id', organizationId)
                .maybeSingle();
            if (error) throw new Error(error.message || 'Failed to load lesson.');
            if (!data) throw new Error('Lesson not found.');

            const openedProjectId =
                data.project_id != null
                    ? data.project_id
                    : openedProjectIdFromCard != null
                      ? openedProjectIdFromCard
                      : null;
            if (openedProjectId == null) throw new Error('Lesson has no project.');

            const { data: openedProject, error: projErr } = await supabase
                .from('projects')
                .select('project_id, project_type_id, project_name')
                .eq('project_id', openedProjectId)
                .eq('organization_id', organizationId)
                .maybeSingle();
            if (projErr) throw new Error(projErr.message || 'Failed to load project.');

            const projectForOpen = openedProject || { project_id: openedProjectId };
            await mountLessonFullPage(mountEl, data, projectForOpen, {
                ...ctx,
                projectId: openedProjectId,
            });
        } catch (err) {
            console.error(err);
            setToolbarStatus(err.message || 'Could not open lesson.', true);
        }
    }

    card.appendChild(relevantResultsHost);
    card.appendChild(buildLessonPrimaryTitle(row));

    const detailHost = document.createElement('div');
    card.appendChild(detailHost);
    mountEl.appendChild(card);

    const {
        openLessonDraftTextModal,
        refreshLessonCompleteness,
        formatCompletenessStatusMessage,
    } = await import('./my-projects-lesson-draft-editor.js');

    btnSaveCompleteness.addEventListener('click', async () => {
        try {
            btnSaveCompleteness.disabled = true;
            await refreshLessonCompleteness(lessonId, organizationId, userId);
            setToolbarStatus(formatCompletenessStatusMessage(), false);
        } catch (err) {
            console.error(err);
            setToolbarStatus(err.message || 'Could not update completeness.', true);
        } finally {
            btnSaveCompleteness.disabled = false;
        }
    });

    function effectiveNoteOwnerId(noteRow) {
        const cb = noteRow && noteRow.created_by;
        if (cb != null && String(cb).trim() !== '') return String(cb);
        if (lessonCreatorId != null) return String(lessonCreatorId);
        return '';
    }

    function canMutateNote(noteRow) {
        return userId != null && effectiveNoteOwnerId(noteRow) === String(userId);
    }

    async function refreshBody() {
        detailHost.innerHTML = '';
        const loading = document.createElement('div');
        loading.className = 'org-lesson-detail-loading';
        loading.textContent = 'Loading lesson structure…';
        detailHost.appendChild(loading);

        try {
            const detail = await fetchLessonStructure(supabase, {
                organizationId,
                projectId: pid,
                lessonId,
            });
            const notes = Array.isArray(detail.notes) ? detail.notes.slice() : [];
            const detailSansNotes = { ...detail, notes: [] };
            detailHost.innerHTML = '';
            renderLessonStructureInto(
                detailHost,
                detailSansNotes,
                { supabase, organizationId, projectId: pid },
                { skipNotes: true }
            );

            const ns = document.createElement('section');
            ns.className = 'org-lesson-notes-section org-lesson-notes-section--interactive';
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
                        organization_id: organizationId,
                        project_id: pid,
                    });
                    if (error) throw error;
                    await refreshBody();
                } catch (err) {
                    console.error(err);
                    alert(err.message || 'Save failed.');
                }
            });

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
                    li.appendChild(span);
                    if (canMutateNote(n)) {
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
                                showDelete: true,
                            });
                            try {
                                let qDel = supabase
                                    .from('lessons_learned_notes')
                                    .delete()
                                    .eq('id', n.id)
                                    .eq('organization_id', organizationId);
                                let qUp = supabase
                                    .from('lessons_learned_notes')
                                    .update({ notes: r.value })
                                    .eq('id', n.id)
                                    .eq('organization_id', organizationId);
                                if (n.created_by != null) {
                                    qDel = qDel.eq('created_by', userId);
                                    qUp = qUp.eq('created_by', userId);
                                }
                                if (r.action === 'delete') {
                                    const { error } = await qDel;
                                    if (error) throw error;
                                } else if (r.action === 'save') {
                                    const { error } = await qUp;
                                    if (error) throw error;
                                } else return;
                                await refreshBody();
                            } catch (err) {
                                console.error(err);
                                alert(err.message || 'Update failed.');
                            }
                        });
                        li.appendChild(eb);
                    }
                    ul.appendChild(li);
                });
                ns.appendChild(ul);
            }

            detailHost.appendChild(ns);
        } catch (err) {
            detailHost.innerHTML = '';
            const errEl = document.createElement('div');
            errEl.className = 'upload-message upload-message--error';
            errEl.textContent = (err && err.message) || 'Failed to load lesson details.';
            detailHost.appendChild(errEl);
        }
    }

    await refreshBody();
}

const MY_PROJECTS_CATEGORY_LABEL_PREVIEW_COUNT = 4;

/**
 * @param {string[]} labels
 * @returns {string}
 */
function formatMyProjectsCategoryLabelsLine(labels) {
    return `Lessons Learned Categories: ${labels.join(', ')}`;
}

/**
 * @param {{ id?: unknown, category?: unknown, title?: unknown }} row
 * @param {{ project_id?: unknown }} project
 * @param {{
 *   onOpenLesson?: (args: { row: typeof row, project: typeof project }) => void,
 *   useYellowHighlight?: boolean,
 *   projectName?: string|null,
 *   categoryLabels?: string[],
 * }} deps
 */
export function createMyProjectsLessonWrap(row, project, deps) {
    const { onOpenLesson, useYellowHighlight, projectName, categoryLabels } = deps || {};
    const wrap = document.createElement('div');
    wrap.className = 'my-projects-lesson-wrap';

    const categoryRaw = row && row.category ? String(row.category).trim() : '';
    const categoryLower = categoryRaw.toLowerCase();
    const categoryDisplay = categoryRaw
        ? categoryRaw.charAt(0).toUpperCase() + categoryRaw.slice(1)
        : 'Issue';
    const title = row && row.title ? String(row.title).trim() : '(Untitled)';
    const projectNameText =
        projectName != null && String(projectName).trim()
            ? String(projectName).trim()
            : project && project.project_name != null && String(project.project_name).trim()
              ? String(project.project_name).trim()
              : '';

    const card = document.createElement('div');
    card.className = 'my-projects-lesson-card my-projects-lesson-card--clickable';
    if (useYellowHighlight) {
        card.classList.add('my-projects-lesson-card--for-review-creator');
    } else if (categoryLower === 'success') {
        card.classList.add('my-projects-lesson-card--success');
    } else {
        card.classList.add('my-projects-lesson-card--issue');
    }

    const titleLine = document.createElement('div');
    titleLine.className = 'my-projects-lesson-card-title';
    const label = document.createElement('strong');
    label.textContent = `${categoryDisplay}: `;
    titleLine.appendChild(label);
    titleLine.appendChild(document.createTextNode(title));
    card.appendChild(titleLine);

    if (projectNameText) {
        const projectLine = document.createElement('div');
        projectLine.className = 'my-projects-lesson-card-project';
        projectLine.textContent = projectNameText;
        card.appendChild(projectLine);
    }

    card.setAttribute('role', 'button');
    card.tabIndex = 0;
    card.setAttribute(
        'aria-label',
        projectNameText
            ? `Open full lesson, ${categoryDisplay}: ${title}, project ${projectNameText}`
            : `Open full lesson, ${categoryDisplay}: ${title}`
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

    const labels = Array.isArray(categoryLabels)
        ? categoryLabels.map((t) => String(t || '').trim()).filter(Boolean)
        : [];
    if (labels.length > 0) {
        const categoriesEl = document.createElement('div');
        categoriesEl.className = 'my-projects-lesson-categories';

        const textEl = document.createElement('span');
        textEl.className = 'my-projects-lesson-categories-text';
        categoriesEl.appendChild(textEl);

        if (labels.length <= MY_PROJECTS_CATEGORY_LABEL_PREVIEW_COUNT) {
            textEl.textContent = formatMyProjectsCategoryLabelsLine(labels);
        } else {
            const previewLabels = labels.slice(0, MY_PROJECTS_CATEGORY_LABEL_PREVIEW_COUNT);
            let expanded = false;
            const toggleBtn = document.createElement('button');
            toggleBtn.type = 'button';
            toggleBtn.className = 'my-projects-lesson-categories-toggle';
            toggleBtn.setAttribute('aria-expanded', 'false');
            toggleBtn.textContent = 'Show more';

            const renderCategories = () => {
                textEl.textContent = formatMyProjectsCategoryLabelsLine(
                    expanded ? labels : previewLabels
                );
                toggleBtn.textContent = expanded ? 'Show less' : 'Show more';
                toggleBtn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
            };
            renderCategories();

            toggleBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                expanded = !expanded;
                renderCategories();
            });

            categoriesEl.appendChild(document.createTextNode(' '));
            categoriesEl.appendChild(toggleBtn);
        }

        wrap.appendChild(categoriesEl);
    }

    return wrap;
}
