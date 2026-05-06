/**
 * Project Analysis — UI for the My Projects project workspace “Project Analysis” flow.
 */

/**
 * @param {{
 *   mountEl: HTMLElement | null,
 *   project?: { id?: string | number, project_id?: string | number, project_name?: string, [key: string]: unknown } | null,
 *   ctUser?: { id?: string | number, usertype?: string, [key: string]: unknown } | null,
 * }} ctx
 */
export function mountProjectAnalysisPortal(ctx) {
    const { mountEl } = ctx || {};
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

    actions.appendChild(projectBtn);
    actions.appendChild(categoryBtn);

    wrap.appendChild(actions);
    mountEl.appendChild(wrap);
}

/** @param {HTMLElement | null} mountEl */
export function clearProjectAnalysisPortal(mountEl) {
    if (mountEl) mountEl.innerHTML = '';
}
