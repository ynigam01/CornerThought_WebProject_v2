/**
 * Project Analysis — UI for the My Projects project workspace “Project Analysis” flow.
 */

/**
 * @param {{
 *   mountEl: HTMLElement | null,
 *   project?: { id?: string | number, project_id?: string | number, project_name?: string, [key: string]: unknown } | null,
 *   ctUser?: { id?: string | number, usertype?: string, organizationid?: string | number, [key: string]: unknown } | null,
 *   organizationId?: string | number | null,
 *   loadLessonsCategoriesForSelect?: (selectEl: HTMLSelectElement) => Promise<void>,
 * }} ctx
 */
export function mountProjectAnalysisPortal(ctx) {
    const { mountEl, project, ctUser, loadLessonsCategoriesForSelect } = ctx || {};
    if (!mountEl) return;

    const organizationId =
        ctx && ctx.organizationId != null
            ? ctx.organizationId
            : ctUser && ctUser.organizationid != null
              ? ctUser.organizationid
              : null;
    const projectId =
        project && project.project_id != null
            ? project.project_id
            : project && project.id != null
              ? project.id
              : null;

    mountEl.innerHTML = '';

    const wrap = document.createElement('div');
    wrap.className = 'my-projects-project-analysis-placeholder';
    wrap.setAttribute('role', 'region');
    wrap.setAttribute('aria-label', 'Project analysis');

    const actions = document.createElement('div');
    actions.className = 'my-projects-project-analysis-actions';

    const projectBtn = document.createElement('button');
    projectBtn.type = 'button';
    projectBtn.className = 'side-button';
    projectBtn.id = 'projectAnalysisProjectLevelButton';
    projectBtn.textContent = 'Project Level Analysis';

    const categoryBtn = document.createElement('button');
    categoryBtn.type = 'button';
    categoryBtn.className = 'side-button';
    categoryBtn.id = 'projectAnalysisCategoryLevelButton';
    categoryBtn.textContent = 'Category Level Analysis';
    categoryBtn.setAttribute('aria-expanded', 'false');
    categoryBtn.setAttribute('aria-controls', 'myProjectsProjectAnalysisCategoriesPanel');

    const createEmbeddingsBtn = document.createElement('button');
    createEmbeddingsBtn.type = 'button';
    createEmbeddingsBtn.className = 'side-button';
    createEmbeddingsBtn.id = 'projectAnalysisCreateEmbeddingsButton';
    createEmbeddingsBtn.textContent = 'Create Embeddings';

    actions.appendChild(projectBtn);
    actions.appendChild(categoryBtn);
    actions.appendChild(createEmbeddingsBtn);

    const categoryPanel = document.createElement('div');
    categoryPanel.id = 'myProjectsProjectAnalysisCategoriesPanel';
    categoryPanel.className = 'my-projects-project-analysis-category-panel';
    categoryPanel.hidden = true;

    const formGroup = document.createElement('div');
    formGroup.className = 'form-group';
    formGroup.style.marginBottom = '0';

    const label = document.createElement('label');
    const categoriesSelect = document.createElement('select');
    categoriesSelect.id = 'myProjectsProjectAnalysisCategoriesSelect';
    categoriesSelect.setAttribute('aria-label', 'Lessons Learned Categories');
    label.setAttribute('for', categoriesSelect.id);
    label.textContent = 'Lessons Learned Categories';

    const placeholderOpt = document.createElement('option');
    placeholderOpt.value = '';
    placeholderOpt.textContent = 'Select a category';
    categoriesSelect.appendChild(placeholderOpt);

    formGroup.appendChild(label);
    formGroup.appendChild(categoriesSelect);
    categoryPanel.appendChild(formGroup);

    const embeddingsStatus = document.createElement('div');
    embeddingsStatus.id = 'projectAnalysisCreateEmbeddingsStatus';
    embeddingsStatus.className = 'upload-message';
    embeddingsStatus.setAttribute('aria-live', 'polite');
    embeddingsStatus.style.display = 'none';

    wrap.appendChild(actions);
    wrap.appendChild(categoryPanel);
    wrap.appendChild(embeddingsStatus);
    mountEl.appendChild(wrap);

    function setEmbeddingsStatus(message, kind) {
        embeddingsStatus.classList.remove('upload-message--success', 'upload-message--error');
        if (!message) {
            embeddingsStatus.style.display = 'none';
            embeddingsStatus.textContent = '';
            return;
        }
        embeddingsStatus.style.display = '';
        embeddingsStatus.textContent = message;
        if (kind === 'error') {
            embeddingsStatus.classList.add('upload-message--error');
        } else if (kind === 'success') {
            embeddingsStatus.classList.add('upload-message--success');
        }
    }

    projectBtn.addEventListener('click', () => {
        categoryPanel.hidden = true;
        categoryBtn.setAttribute('aria-expanded', 'false');
    });

    categoryBtn.addEventListener('click', async () => {
        const willOpen = categoryPanel.hidden;
        if (!willOpen) {
            categoryPanel.hidden = true;
            categoryBtn.setAttribute('aria-expanded', 'false');
            return;
        }
        categoryPanel.hidden = false;
        categoryBtn.setAttribute('aria-expanded', 'true');

        if (typeof loadLessonsCategoriesForSelect !== 'function') return;

        try {
            await loadLessonsCategoriesForSelect(categoriesSelect);
        } catch (err) {
            console.error('Project Analysis: failed to load lessons categories', err);
        }
    });

    createEmbeddingsBtn.addEventListener('click', async () => {
        if (organizationId == null || organizationId === '') {
            setEmbeddingsStatus(
                'Could not determine your organization. Please log out and log back in.',
                'error',
            );
            return;
        }
        if (projectId == null || projectId === '') {
            setEmbeddingsStatus('Could not determine the current project.', 'error');
            return;
        }

        createEmbeddingsBtn.disabled = true;
        createEmbeddingsBtn.textContent = 'Creating Embeddings...';
        setEmbeddingsStatus(
            'Checking completed lessons for missing embeddings…',
            null,
        );

        try {
            const response = await fetch('/api/draft-lessons/embed-completed-for-project', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ organizationId, projectId }),
            });

            if (!response.ok) {
                let data = null;
                try {
                    data = await response.json();
                } catch (_) {
                    data = null;
                }
                throw new Error(
                    (data && data.error) ||
                        `Failed to create embeddings (${response.status}).`,
                );
            }

            if (!response.body || typeof response.body.getReader !== 'function') {
                throw new Error('Streaming response is not supported in this browser.');
            }

            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            let buffer = '';
            let lessonTotal = 0;

            const handleEvent = (event) => {
                if (!event || !event.type) return;
                if (event.type === 'start') {
                    lessonTotal = Number(event.total) || 0;
                    if (lessonTotal === 0) {
                        setEmbeddingsStatus(
                            'No completed lessons need embeddings for this project.',
                            'success',
                        );
                    } else {
                        setEmbeddingsStatus(
                            `Embedding completed lessons (0/${lessonTotal})…`,
                            null,
                        );
                    }
                } else if (event.type === 'lesson') {
                    const index = Number(event.index) || 0;
                    const total = Number(event.total) || lessonTotal;
                    setEmbeddingsStatus(
                        `Embedding completed lessons (${index}/${total})…`,
                        null,
                    );
                } else if (event.type === 'done') {
                    const processed = Number(event.processed) || 0;
                    const skipped = Number(event.skipped) || 0;
                    const failed = Number(event.failed) || 0;
                    const hasFailures = failed > 0;
                    setEmbeddingsStatus(
                        `Done. Updated ${processed} lesson${processed === 1 ? '' : 's'}. ` +
                            `Skipped ${skipped} already complete` +
                            (failed > 0 ? `. Failed: ${failed}` : '') +
                            '.',
                        hasFailures ? 'error' : 'success',
                    );
                } else if (event.type === 'error') {
                    setEmbeddingsStatus(
                        event.message || 'Failed to create embeddings.',
                        'error',
                    );
                }
            };

            while (true) {
                const { value, done } = await reader.read();
                if (done) break;
                buffer += decoder.decode(value, { stream: true });

                let newlineIndex = buffer.indexOf('\n');
                while (newlineIndex !== -1) {
                    const line = buffer.slice(0, newlineIndex).trim();
                    buffer = buffer.slice(newlineIndex + 1);
                    newlineIndex = buffer.indexOf('\n');
                    if (!line) continue;
                    try {
                        handleEvent(JSON.parse(line));
                    } catch (_) {
                        // ignore incomplete/malformed lines
                    }
                }
            }

            const trailing = buffer.trim();
            if (trailing) {
                try {
                    handleEvent(JSON.parse(trailing));
                } catch (_) {
                    // ignore incomplete trailing buffer
                }
            }
        } catch (err) {
            console.error('Create Embeddings failed:', err);
            setEmbeddingsStatus(
                err && err.message ? err.message : 'Failed to create embeddings.',
                'error',
            );
        } finally {
            createEmbeddingsBtn.disabled = false;
            createEmbeddingsBtn.textContent = 'Create Embeddings';
        }
    });
}

/** @param {HTMLElement | null} mountEl */
export function clearProjectAnalysisPortal(mountEl) {
    if (mountEl) mountEl.innerHTML = '';
}
