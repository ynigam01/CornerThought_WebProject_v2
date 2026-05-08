/**
 * Project Analysis — UI for the My Projects project workspace “Project Analysis” flow.
 */

/**
 * @param {{
 *   mountEl: HTMLElement | null,
 *   project?: { id?: string | number, project_id?: string | number, project_name?: string, [key: string]: unknown } | null,
 *   ctUser?: { id?: string | number, usertype?: string, [key: string]: unknown } | null,
 *   loadLessonsCategoriesForSelect?: (selectEl: HTMLSelectElement) => Promise<void>,
 * }} ctx
 */
export function mountProjectAnalysisPortal(ctx) {
    const { mountEl, loadLessonsCategoriesForSelect } = ctx || {};
    if (!mountEl) return;

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

    actions.appendChild(projectBtn);
    actions.appendChild(categoryBtn);

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

    wrap.appendChild(actions);
    wrap.appendChild(categoryPanel);
    mountEl.appendChild(wrap);

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
}

/** @param {HTMLElement | null} mountEl */
export function clearProjectAnalysisPortal(mountEl) {
    if (mountEl) mountEl.innerHTML = '';
}
