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
 * Update an existing workshop row by id.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} workshopId
 * @param {Record<string, unknown>} row
 */
export async function updateWorkshop(supabase, workshopId, row) {
    return supabase.from('workshops').update(row).eq('id', workshopId);
}

/**
 * Delete a workshop row by id.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} workshopId
 */
export async function deleteWorkshop(supabase, workshopId) {
    return supabase.from('workshops').delete().eq('id', workshopId);
}

/**
 * Compute the duration in minutes between two HH:MM(:SS) time strings.
 * Handles overnight wrap-around.
 * @param {string} startTime
 * @param {string} endTime
 * @returns {number}
 */
export function computeWorkshopDurationMinutes(startTime, endTime) {
    const toMins = (t) => {
        const parts = String(t || '').split(':');
        const h = parseInt(parts[0], 10);
        const m = parseInt(parts[1] || '0', 10);
        return Number.isFinite(h) && Number.isFinite(m) ? h * 60 + m : 0;
    };
    const diff = toMins(endTime) - toMins(startTime);
    return diff < 0 ? diff + 24 * 60 : diff;
}

/**
 * Fetch all workshop_lessons_learned rows for a workshop.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} workshopId
 */
export async function fetchWorkshopLessons(supabase, workshopId) {
    return supabase
        .from('workshop_lessons_learned')
        .select('id, lessons_learned_id, grouping_id')
        .eq('workshop_id', workshopId);
}

/**
 * Insert a row into workshop_lessons_learned.
 * Returns the new row so the caller can capture its id.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} workshopId
 * @param {string | number} lessonId
 * @param {string | number} addedBy
 */
export async function addLessonToWorkshop(supabase, workshopId, lessonId, addedBy) {
    return supabase
        .from('workshop_lessons_learned')
        .insert({ workshop_id: Number(workshopId), lessons_learned_id: Number(lessonId), added_by: Number(addedBy) })
        .select('id, lessons_learned_id')
        .single();
}

/**
 * Delete a workshop_lessons_learned row by its own id.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} linkRowId
 */
export async function removeLessonFromWorkshop(supabase, linkRowId) {
    return supabase
        .from('workshop_lessons_learned')
        .delete()
        .eq('id', linkRowId);
}

/**
 * Fetch all workshop_attendees rows for a workshop.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} workshopId
 */
export async function fetchWorkshopAttendees(supabase, workshopId) {
    return supabase
        .from('workshop_attendees')
        .select('id, user_id, name, email, notification_status, confirmation')
        .eq('workshop_id', workshopId);
}

/**
 * Insert a row into workshop_attendees.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {{ workshopId: string | number, userId: string | number, name: string, email: string }} params
 */
export async function addAttendeeToWorkshop(supabase, { workshopId, userId, name, email }) {
    return supabase
        .from('workshop_attendees')
        .insert({
            workshop_id: Number(workshopId),
            user_id: Number(userId),
            name: name || null,
            email: email || null,
            confirmation: false,
            attendance: false,
            internal_attendee: true,
        })
        .select('id, user_id, name, email')
        .single();
}

/**
 * Delete a workshop_attendees row by its own id.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} attendeeRowId
 */
export async function removeAttendeeFromWorkshop(supabase, attendeeRowId) {
    return supabase
        .from('workshop_attendees')
        .delete()
        .eq('id', attendeeRowId);
}

/**
 * Set notification_status to "sent" for attendees of a workshop.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} workshopId
 * @param {{ onlyNew?: boolean }} [options] - When onlyNew is true, only updates rows where notification_status IS NULL.
 */
export async function notifyWorkshopAttendees(supabase, workshopId, { onlyNew = false } = {}) {
    let query = supabase
        .from('workshop_attendees')
        .update({ notification_status: 'sent' })
        .eq('workshop_id', Number(workshopId));
    if (onlyNew) {
        query = query.is('notification_status', null);
    }
    return query;
}

/**
 * Accept or decline a workshop invite.
 * Sets confirmation to the given value and notification_status to 'confirmation_sent'.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} attendeeId  The workshop_attendees row id.
 * @param {boolean} accepted  true = Accept, false = Decline.
 */
export async function respondToWorkshopInvite(supabase, attendeeId, accepted) {
    return supabase
        .from('workshop_attendees')
        .update({ confirmation: accepted, notification_status: 'confirmation_sent' })
        .eq('id', Number(attendeeId));
}

/**
 * Fetch all workshops for a given project, ordered newest date first.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} projectId
 */
export async function fetchWorkshopsForProject(supabase, projectId) {
    return supabase
        .from('workshops')
        .select('id, workshop_title, workshop_description, date, start_time, end_time')
        .eq('project_id', projectId)
        .order('date', { ascending: false });
}

/** @param {string} timeHHMMSS */
export function formatWorkshopTime(timeHHMMSS) {
    const parts = String(timeHHMMSS || '').split(':');
    const h = parseInt(parts[0], 10);
    const m = parts[1] || '00';
    if (!Number.isFinite(h)) return timeHHMMSS;
    const period = h >= 12 ? 'PM' : 'AM';
    const h12 = h % 12 || 12;
    return `${h12}:${m} ${period}`;
}

/**
 * Insert a new row into workshop_lessons_groupings.
 * Returns the new row so the caller can capture its id.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {{ groupingDescription: string, createdBy: string | number, workshopId: string | number }} params
 */
export async function insertWorkshopLessonsGrouping(supabase, { groupingDescription, createdBy, workshopId }) {
    return supabase
        .from('workshop_lessons_groupings')
        .insert({
            grouping_description: String(groupingDescription || '').trim(),
            created_by: Number(createdBy),
            workshop_id: Number(workshopId),
        })
        .select('id')
        .single();
}

/**
 * Set grouping_id on a list of workshop_lessons_learned rows.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {Array<string | number>} wllIds  IDs of workshop_lessons_learned rows to update
 * @param {string | number} groupingId
 */
export async function setGroupingOnWorkshopLessons(supabase, wllIds, groupingId) {
    return supabase
        .from('workshop_lessons_learned')
        .update({ grouping_id: Number(groupingId) })
        .in('id', wllIds.map(Number));
}

/**
 * Clear grouping_id (set to null) on a single workshop_lessons_learned row.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} wllId
 */
export async function clearGroupingFromWorkshopLesson(supabase, wllId) {
    return supabase
        .from('workshop_lessons_learned')
        .update({ grouping_id: null })
        .eq('id', Number(wllId));
}

/**
 * Clear grouping_id (set to null) on all workshop_lessons_learned rows that belong to a grouping.
 * Call this before deleting the grouping to avoid orphaned references.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} groupingId
 */
export async function clearGroupingFromWorkshopLessons(supabase, groupingId) {
    return supabase
        .from('workshop_lessons_learned')
        .update({ grouping_id: null })
        .eq('grouping_id', Number(groupingId));
}

/**
 * Fetch all groupings for a workshop, ordered oldest first.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} workshopId
 */
export async function fetchWorkshopGroupings(supabase, workshopId) {
    return supabase
        .from('workshop_lessons_groupings')
        .select('id, grouping_description, created_by')
        .eq('workshop_id', workshopId)
        .order('id', { ascending: true });
}

/**
 * Delete a workshop_lessons_groupings row by id.
 * Clear all lesson grouping references first with clearGroupingFromWorkshopLessons.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @param {string | number} groupingId
 */
export async function deleteWorkshopGrouping(supabase, groupingId) {
    return supabase
        .from('workshop_lessons_groupings')
        .delete()
        .eq('id', Number(groupingId));
}

/**
 * Renders the categories / attendees dropdown + a list of workshops into mountEl.
 * @param {{
 *   mountEl: HTMLElement | null,
 *   workshops?: Array<Record<string, unknown>> | null,
 *   loading?: boolean,
 *   error?: { message?: string } | null,
 *   loadLessonsCategoriesForSelect?: (selectEl: HTMLSelectElement) => Promise<void>,
 *   onEdit?: (workshop: Record<string, unknown>) => void,
 *   onDelete?: (workshop: Record<string, unknown>) => void,
 *   onAddLessons?: (workshop: Record<string, unknown>) => void,
 *   onCancelLessons?: () => void,
 *   lessonsWorkshopId?: string | number | null,
 *   onMountLessons?: (containerEl: HTMLElement, categoryId: string) => void,
 *   onAddAttendees?: (workshop: Record<string, unknown>) => void,
 *   onCancelAttendees?: () => void,
 *   attendeesWorkshopId?: string | number | null,
 *   onMountAttendees?: (containerEl: HTMLElement, workshopId: string | number) => void,
 *   onAddAttendee?: (user: { userId: string, name: string, email: string }, containerEl: HTMLElement, selectEl: HTMLSelectElement) => void,
 *   loadAttendeesForSelect?: (selectEl: HTMLSelectElement, workshopId: string | number) => Promise<void>,
 *   onGroupLessons?: (workshop: Record<string, unknown>) => void,
 *   onCancelGroupLessons?: () => void,
 *   groupLessonsWorkshopId?: string | number | null,
 *   onMountGroupLessons?: (containerEl: HTMLElement, workshopId: string | number) => void,
 * }} ctx
 */
export function mountManageWorkshopsPanel({
    mountEl, workshops, loading, error,
    loadLessonsCategoriesForSelect,
    onEdit, onDelete, onAddLessons, onCancelLessons,
    lessonsWorkshopId = null,
    onMountLessons,
    onAddAttendees, onCancelAttendees,
    attendeesWorkshopId = null,
    onMountAttendees, onAddAttendee, loadAttendeesForSelect,
    onGroupLessons, onCancelGroupLessons,
    groupLessonsWorkshopId = null,
    onMountGroupLessons,
}) {
    if (!mountEl) return;
    mountEl.innerHTML = '';

    const inLessonsMode = lessonsWorkshopId != null;
    const inAttendeesMode = attendeesWorkshopId != null;
    const inGroupLessonsMode = groupLessonsWorkshopId != null;

    const wrap = document.createElement('div');
    wrap.className = 'my-projects-workshop-mount-inner';
    wrap.setAttribute('role', 'region');
    wrap.setAttribute('aria-label', 'Workshops list');

    // Categories dropdown — only visible in lessons mode
    const categoriesGroup = document.createElement('div');
    categoriesGroup.className = 'form-group';
    categoriesGroup.style.marginBottom = '14px';
    categoriesGroup.style.display = inLessonsMode ? '' : 'none';

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
    wrap.appendChild(categoriesGroup);

    // Attendees dropdown — only visible in attendees mode
    const attendeesGroup = document.createElement('div');
    attendeesGroup.className = 'form-group';
    attendeesGroup.style.marginBottom = '14px';
    attendeesGroup.style.display = inAttendeesMode ? '' : 'none';

    const attendeesLabel = document.createElement('label');
    const attendeesSelect = document.createElement('select');
    attendeesSelect.id = 'myProjectsWorkshopAttendeesSelect';
    attendeesSelect.setAttribute('aria-label', 'Add Attendee');
    attendeesLabel.setAttribute('for', attendeesSelect.id);
    attendeesLabel.textContent = 'Add Attendee';

    const attendeesPlaceholderOpt = document.createElement('option');
    attendeesPlaceholderOpt.value = '';
    attendeesPlaceholderOpt.textContent = 'Select a user to add…';
    attendeesSelect.appendChild(attendeesPlaceholderOpt);

    attendeesGroup.appendChild(attendeesLabel);
    attendeesGroup.appendChild(attendeesSelect);
    wrap.appendChild(attendeesGroup);

    // Track the lessons/attendees containers so handlers can reference them
    let lessonsContainer = null;
    let attendeesContainer = null;

    if (inLessonsMode) {
        if (typeof loadLessonsCategoriesForSelect === 'function') {
            loadLessonsCategoriesForSelect(categoriesSelect).catch((err) => {
                console.error('Manage Workshops: failed to load lessons categories', err);
            });
        }
        categoriesSelect.addEventListener('change', () => {
            if (lessonsContainer && typeof onMountLessons === 'function') {
                onMountLessons(lessonsContainer, categoriesSelect.value || '');
            }
        });
    }

    if (inAttendeesMode) {
        if (typeof loadAttendeesForSelect === 'function') {
            loadAttendeesForSelect(attendeesSelect, attendeesWorkshopId).catch((err) => {
                console.error('Manage Workshops: failed to load attendees for select', err);
            });
        }
        attendeesSelect.addEventListener('change', () => {
            if (!attendeesSelect.value) return;
            const selectedOption = attendeesSelect.options[attendeesSelect.selectedIndex];
            const userId = attendeesSelect.value;
            const name = selectedOption.dataset.name || '';
            const email = selectedOption.dataset.email || '';
            attendeesSelect.value = '';
            if (attendeesContainer && typeof onAddAttendee === 'function') {
                onAddAttendee({ userId, name, email }, attendeesContainer, attendeesSelect);
            }
        });
    }

    // Workshops content area
    if (loading) {
        const msg = document.createElement('p');
        msg.className = 'subtitle';
        msg.style.margin = '0';
        msg.textContent = 'Loading workshops…';
        wrap.appendChild(msg);
    } else if (error) {
        const msg = document.createElement('p');
        msg.className = 'upload-message upload-message--error';
        msg.textContent = error.message || 'Failed to load workshops.';
        wrap.appendChild(msg);
    } else if (!workshops || workshops.length === 0) {
        const msg = document.createElement('p');
        msg.className = 'subtitle';
        msg.style.margin = '0';
        msg.textContent = 'No workshops have been created for this project yet.';
        wrap.appendChild(msg);
    } else {
        // In lessons/attendees/group-lessons mode only show the selected card; otherwise show all
        const activeId = inLessonsMode ? lessonsWorkshopId
            : inAttendeesMode ? attendeesWorkshopId
            : inGroupLessonsMode ? groupLessonsWorkshopId
            : null;
        const displayWorkshops = activeId != null
            ? workshops.filter((w) => String(w.id) === String(activeId))
            : workshops;

        if (!inLessonsMode && !inAttendeesMode && !inGroupLessonsMode) {
            const heading = document.createElement('p');
            heading.style.cssText = 'font-weight: 600; margin: 0 0 12px; font-size: 14px; color: #444;';
            heading.textContent = `${workshops.length} workshop${workshops.length !== 1 ? 's' : ''}`;
            wrap.appendChild(heading);
        }

        const list = document.createElement('div');
        list.className = 'lessons-results';
        list.style.marginTop = '0';

        displayWorkshops.forEach((w) => {
            const card = document.createElement('div');
            card.className = 'lesson-card';

            const title = document.createElement('div');
            title.style.cssText = 'font-weight: 600; font-size: 15px; margin-bottom: 6px;';
            title.textContent = String(w.workshop_title || '');
            card.appendChild(title);

            const meta = document.createElement('div');
            meta.style.cssText = 'font-size: 13px; color: #555; display: flex; gap: 16px; flex-wrap: wrap;';

            if (w.date) {
                const dateSpan = document.createElement('span');
                dateSpan.textContent = `\uD83D\uDCC5 ${w.date}`;
                meta.appendChild(dateSpan);
            }
            if (w.start_time) {
                const timeSpan = document.createElement('span');
                const startFmt = formatWorkshopTime(w.start_time);
                const endFmt = w.end_time ? ` \u2013 ${formatWorkshopTime(w.end_time)}` : '';
                timeSpan.textContent = `\u23F0 ${startFmt}${endFmt}`;
                meta.appendChild(timeSpan);
            }
            card.appendChild(meta);

            if (w.workshop_description) {
                const desc = document.createElement('div');
                desc.style.cssText = 'margin-top: 8px; font-size: 14px; color: #444;';
                desc.textContent = String(w.workshop_description);
                card.appendChild(desc);
            }

            if (inLessonsMode) {
                const goBackBtn = document.createElement('button');
                goBackBtn.type = 'button';
                goBackBtn.className = 'secondary-button';
                goBackBtn.style.cssText = 'margin-top: 12px; font-size: 13px;';
                goBackBtn.textContent = 'Go Back';
                goBackBtn.addEventListener('click', () => onCancelLessons && onCancelLessons());
                card.appendChild(goBackBtn);

                list.appendChild(card);

                lessonsContainer = document.createElement('div');
                lessonsContainer.className = 'workshop-lessons-container';
                list.appendChild(lessonsContainer);

                wrap.appendChild(list);

                if (typeof onMountLessons === 'function') {
                    onMountLessons(lessonsContainer, '');
                }

                mountEl.appendChild(wrap);
                return;
            } else if (inAttendeesMode) {
                const goBackBtn = document.createElement('button');
                goBackBtn.type = 'button';
                goBackBtn.className = 'secondary-button';
                goBackBtn.style.cssText = 'margin-top: 12px; font-size: 13px;';
                goBackBtn.textContent = 'Go Back';
                goBackBtn.addEventListener('click', () => onCancelAttendees && onCancelAttendees());
                card.appendChild(goBackBtn);

                list.appendChild(card);

                attendeesContainer = document.createElement('div');
                attendeesContainer.className = 'workshop-attendees-container';
                list.appendChild(attendeesContainer);

                wrap.appendChild(list);

                if (typeof onMountAttendees === 'function') {
                    onMountAttendees(attendeesContainer, attendeesWorkshopId);
                }

                mountEl.appendChild(wrap);
                return;
            } else if (inGroupLessonsMode) {
                const goBackBtn = document.createElement('button');
                goBackBtn.type = 'button';
                goBackBtn.className = 'secondary-button';
                goBackBtn.style.cssText = 'margin-top: 12px; font-size: 13px;';
                goBackBtn.textContent = 'Go Back';
                goBackBtn.addEventListener('click', () => onCancelGroupLessons && onCancelGroupLessons());
                card.appendChild(goBackBtn);

                list.appendChild(card);

                const groupLessonsContainer = document.createElement('div');
                groupLessonsContainer.className = 'workshop-group-lessons-container';
                list.appendChild(groupLessonsContainer);

                wrap.appendChild(list);

                if (typeof onMountGroupLessons === 'function') {
                    onMountGroupLessons(groupLessonsContainer, groupLessonsWorkshopId);
                }

                mountEl.appendChild(wrap);
                return;
            } else {
                // Normal mode — icon action buttons
                const actions = document.createElement('div');
                actions.className = 'workshop-card-actions';

                const iconBtn = (faClass, label, danger = false) => {
                    const btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className = 'workshop-icon-btn' + (danger ? ' workshop-icon-btn--danger' : '');
                    btn.title = label;
                    btn.setAttribute('aria-label', label);
                    btn.innerHTML = `<i class="${faClass}"></i>`;
                    return btn;
                };

                const editBtn = iconBtn('fa-solid fa-pen', 'Edit Workshop Details');
                editBtn.addEventListener('click', () => onEdit && onEdit(w));

                const attendeesBtn = iconBtn('fa-solid fa-user-plus', 'Add Attendees');
                attendeesBtn.addEventListener('click', () => onAddAttendees && onAddAttendees(w));

                const lessonsBtn = iconBtn('fa-solid fa-lightbulb', 'Add Lessons Learned');
                lessonsBtn.innerHTML = `<span style="position:relative;display:inline-flex;align-items:center;justify-content:center;"><i class="fa-solid fa-lightbulb"></i><i class="fa-solid fa-plus" style="position:absolute;font-size:0.52em;bottom:-1px;right:-4px;"></i></span>`;
                lessonsBtn.addEventListener('click', () => onAddLessons && onAddLessons(w));

                const groupLessonsBtn = iconBtn('fa-solid fa-layer-group', 'Group Lessons Learned');
                groupLessonsBtn.addEventListener('click', () => onGroupLessons && onGroupLessons(w));

                const deleteBtn = iconBtn('fa-solid fa-trash', 'Delete Workshop', true);
                deleteBtn.addEventListener('click', () => onDelete && onDelete(w));

                actions.appendChild(editBtn);
                actions.appendChild(attendeesBtn);
                actions.appendChild(lessonsBtn);
                actions.appendChild(groupLessonsBtn);
                actions.appendChild(deleteBtn);
                card.appendChild(actions);
            }

            list.appendChild(card);
        });

        wrap.appendChild(list);
    }

    mountEl.appendChild(wrap);
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
