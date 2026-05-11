/**
 * Workshop — UI and behavior for the My Projects “Setup Workshop” flow.
 */

/**
 * @param {string} startTimeHHMM Start time from `<input type="time">` (`HH:MM` or `HH:MM:SS`).
 * @param {number} durationMinutes
 * @returns {string} End time as `HH:MM:SS` (same-day; wraps past midnight modulo 24h for MVP).
 */
export function computeWorkshopEndTime(startTimeHHMM, durationMinutes) {
    const raw = String(startTimeHHMM || '').trim();
    const parts = raw.split(':');
    const h = parseInt(parts[0], 10);
    const min = parseInt(parts[1] !== undefined ? parts[1] : '0', 10);
    const dm = Number(durationMinutes);
    if (!Number.isFinite(h) || !Number.isFinite(min) || !Number.isFinite(dm) || dm < 0) {
        throw new Error('Invalid start time or duration.');
    }
    let total = h * 60 + min + dm;
    total %= 24 * 60;
    if (total < 0) total += 24 * 60;
    const eh = Math.floor(total / 60);
    const em = total % 60;
    return `${String(eh).padStart(2, '0')}:${String(em).padStart(2, '0')}:00`;
}

/** @param {string} timeFromInput */
export function normalizeWorkshopTimeForDb(timeFromInput) {
    const t = String(timeFromInput || '').trim();
    if (!t) throw new Error('Start time is required.');
    const parts = t.split(':');
    const hh = String(parts[0] || '0').padStart(2, '0');
    const mm = String(parts[1] !== undefined ? parts[1] : '0').padStart(2, '0');
    const ss = parts[2] !== undefined ? String(parseInt(parts[2], 10) || 0).padStart(2, '0') : '00';
    return `${hh}:${mm}:${ss}`;
}

/**
 * @param {{
 *   title: string,
 *   description: string,
 *   dateStr: string,
 *   startTime: string,
 *   durationMinutes: number,
 * }} formValues
 * @param {{ projectId: string | number, organizationId: string | number, createdByUserId: string | number }} ctx
 * @returns {Record<string, unknown>}
 */
export function buildWorkshopsInsertRow(formValues, ctx) {
    const { title, description, dateStr, startTime, durationMinutes } = formValues;
    const { projectId, organizationId, createdByUserId } = ctx;

    const workshop_title = String(title || '').trim();
    if (!workshop_title) throw new Error('Workshop name is required.');

    const date = String(dateStr || '').trim();
    if (!date) throw new Error('Date is required.');

    const startNorm = normalizeWorkshopTimeForDb(startTime);
    const end_time = computeWorkshopEndTime(startTime, durationMinutes);

    const pid = projectId != null ? Number(projectId) : NaN;
    const oid = organizationId != null ? Number(organizationId) : NaN;
    const uid = createdByUserId != null ? Number(createdByUserId) : NaN;
    if (!Number.isFinite(pid) || !Number.isFinite(oid) || !Number.isFinite(uid)) {
        throw new Error('Missing project, organization, or user context.');
    }

    const row = {
        workshop_title,
        workshop_description: String(description || '').trim() || null,
        date,
        start_time: startNorm,
        end_time,
        project_id: pid,
        organization_id: oid,
        created_by: uid,
    };

    return row;
}

/**
 * Insert a row into `workshops`. Requires Supabase RLS to allow INSERT for the portal user.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {Record<string, unknown>} row
 */
export async function insertWorkshop(supabase, row) {
    return supabase.from('workshops').insert(row);
}

/**
 * @param {{
 *   mountEl: HTMLElement | null,
 *   project?: { project_id?: string | number, project_name?: string, [key: string]: unknown } | null,
 *   ctUser?: { id?: string | number, [key: string]: unknown } | null,
 *   loadLessonsCategoriesForSelect?: (selectEl: HTMLSelectElement) => Promise<void>,
 * }} ctx
 */
export function mountWorkshopModule(ctx) {
    const { mountEl, loadLessonsCategoriesForSelect } = ctx || {};
    if (!mountEl) return;

    mountEl.innerHTML = '';

    const wrap = document.createElement('div');
    wrap.className = 'my-projects-workshop-mount-inner';
    wrap.setAttribute('role', 'region');
    wrap.setAttribute('aria-label', 'Workshop');

    const categoriesGroup = document.createElement('div');
    categoriesGroup.className = 'form-group';
    categoriesGroup.style.marginBottom = '14px';

    const categoriesLabel = document.createElement('label');
    const categoriesSelect = document.createElement('select');
    categoriesSelect.id = 'myProjectsWorkshopCategoriesSelect';
    categoriesSelect.setAttribute('aria-label', 'Lessons Learned Categories');
    categoriesLabel.setAttribute('for', categoriesSelect.id);
    categoriesLabel.textContent = 'Lessons Learned Categories';

    const placeholderOpt = document.createElement('option');
    placeholderOpt.value = '';
    placeholderOpt.textContent = 'Select a category';
    categoriesSelect.appendChild(placeholderOpt);

    categoriesGroup.appendChild(categoriesLabel);
    categoriesGroup.appendChild(categoriesSelect);

    const msg = document.createElement('p');
    msg.className = 'subtitle';
    msg.style.margin = '0';
    msg.textContent = 'Workshop Module Coming Soon';

    wrap.appendChild(categoriesGroup);
    wrap.appendChild(msg);
    mountEl.appendChild(wrap);

    if (typeof loadLessonsCategoriesForSelect !== 'function') return;

    loadLessonsCategoriesForSelect(categoriesSelect).catch((err) => {
        console.error('Workshop: failed to load lessons categories', err);
    });
}

/** @param {HTMLElement | null} mountEl */
export function clearWorkshopModule(mountEl) {
    if (mountEl) mountEl.innerHTML = '';
}
