// js/admin-portal.js
import { supabase } from './supabase-client.js';

document.addEventListener("DOMContentLoaded", () => {
    const addOrgButton = document.querySelector(".add-organization-button");
    const modal = document.getElementById("addOrgModal");
    const closeButton = document.querySelector(".close-button");
    const form = document.getElementById("addOrgForm");
    const orgTableBody = document.querySelector(".organizations-table tbody");
    const editModal = document.getElementById('editOrgModal');
    const editCloseButton = document.querySelector('.edit-close-button');
    const editForm = document.getElementById('editOrgForm');
    const editNameInput = document.getElementById('editOrgName');
    const editTypeInput = document.getElementById('editOrgType');
    const editEmailInput = document.getElementById('editOrgEmail');
    const editUsersInput = document.getElementById('editOrgUsers');
    const editFirstNameInput = document.getElementById('editAdminFirstName');
    const editLastNameInput = document.getElementById('editAdminLastName');
    let currentEditingId = null;

    // Helper to render an organization row
    function renderOrganizationRow(org) {
        const newRow = document.createElement('tr');
        if (org?.id !== undefined && org?.id !== null) {
            newRow.dataset.orgId = String(org.id);
        }

        const idCell = document.createElement('td');
        idCell.textContent = org?.id ?? '—';

        const nameCell = document.createElement('td');
        nameCell.textContent = org?.name ?? '—';

        const typeCell = document.createElement('td');
        typeCell.textContent = org?.type ?? '—';

        const emailCell = document.createElement('td');
        emailCell.textContent = org?.admin_email ?? '—';

        const usersCell = document.createElement('td');
        usersCell.textContent = String(org?.users_allotted ?? '—');

        const actionsCell = document.createElement('td');
        actionsCell.innerHTML = '<button class="edit-button">Edit</button><button class="delete-button">Delete</button>';

        newRow.appendChild(idCell);
        newRow.appendChild(nameCell);
        newRow.appendChild(typeCell);
        newRow.appendChild(emailCell);
        newRow.appendChild(usersCell);
        newRow.appendChild(actionsCell);
        orgTableBody.appendChild(newRow);
    }

    // Load and render organizations on page load
    async function loadOrganizations() {
        // show loading state
        orgTableBody.innerHTML = '';
        const loadingRow = document.createElement('tr');
        const loadingCell = document.createElement('td');
        loadingCell.colSpan = 6;
        loadingCell.textContent = 'Loading organizations...';
        loadingRow.appendChild(loadingCell);
        orgTableBody.appendChild(loadingRow);

        const { data, error } = await supabase
            .from('organizations')
            .select('id, name, type, admin_email, users_allotted')
            .order('date_added', { ascending: false });

        orgTableBody.innerHTML = '';
        if (error) {
            const errRow = document.createElement('tr');
            const errCell = document.createElement('td');
            errCell.colSpan = 6;
            errCell.textContent = 'Failed to load organizations.';
            errRow.appendChild(errCell);
            orgTableBody.appendChild(errRow);
            console.error('Supabase select error:', error);
            return;
        }

        if (!data || data.length === 0) {
            const emptyRow = document.createElement('tr');
            const emptyCell = document.createElement('td');
            emptyCell.colSpan = 6;
            emptyCell.textContent = 'No organizations yet.';
            emptyRow.appendChild(emptyCell);
            orgTableBody.appendChild(emptyRow);
            return;
        }

        data.forEach(renderOrganizationRow);
    }

    // Handle table action buttons (edit, delete)
    orgTableBody.addEventListener('click', async (event) => {
        // Edit button
        const editBtn = event.target.closest('.edit-button');
        if (editBtn) {
            const row = editBtn.closest('tr');
            const idAttr = row?.dataset?.orgId;
            const idText = row?.querySelector('td')?.textContent?.trim();
            const idValue = idAttr ?? idText;
            const id = idValue ? Number(idValue) : NaN;
            if (!Number.isFinite(id)) {
                alert('Invalid organization id.');
                return;
            }
            currentEditingId = id;
            // Prefill from row cells
            const cells = row.querySelectorAll('td');
            editNameInput.value = cells[1]?.textContent?.trim() ?? '';
            editTypeInput.value = cells[2]?.textContent?.trim() ?? '';
            editEmailInput.value = cells[3]?.textContent?.trim() ?? '';
            const usersVal = cells[4]?.textContent?.trim();
            editUsersInput.value = usersVal ? Number(usersVal) : '';
            // Try to prefill first/last name from users table by email
            try {
                const adminEmail = editEmailInput.value.trim();
                if (adminEmail) {
                    const { data: userRows, error: selErr } = await supabase
                        .from('users')
                        .select('name')
                        .eq('email', adminEmail)
                        .limit(1);
                    if (!selErr && userRows && userRows.length > 0 && userRows[0]?.name) {
                        const name = String(userRows[0].name);
                        const parts = name.trim().split(/\\s+/);
                        const first = parts.shift() || '';
                        const last = parts.join(' ');
                        if (editFirstNameInput) editFirstNameInput.value = first;
                        if (editLastNameInput) editLastNameInput.value = last;
                    } else {
                        if (editFirstNameInput) editFirstNameInput.value = '';
                        if (editLastNameInput) editLastNameInput.value = '';
                    }
                }
            } catch (e) {
                console.warn('Prefill admin name failed:', e);
                if (editFirstNameInput) editFirstNameInput.value = '';
                if (editLastNameInput) editLastNameInput.value = '';
            }
            // Show edit modal
            editModal.classList.add('show');
            return;
        }

        // Delete button
        const deleteBtn = event.target.closest('.delete-button');
        if (!deleteBtn) return;

        const row = deleteBtn.closest('tr');
        const idAttr = row?.dataset?.orgId;
        const idText = row?.querySelector('td')?.textContent?.trim();
        const idValue = idAttr ?? idText;

        const id = idValue ? Number(idValue) : NaN;
        if (!Number.isFinite(id)) {
            alert('Invalid organization id.');
            return;
        }

        const confirmed = confirm('Are you sure you want to delete this organization? This cannot be undone.');
        if (!confirmed) return;

        deleteBtn.disabled = true;
        const originalText = deleteBtn.textContent;
        deleteBtn.textContent = 'Deleting...';

        const { error } = await supabase
            .from('organizations')
            .delete()
            .eq('id', id);

        if (error) {
            console.error('Supabase delete error:', error);
            alert('Failed to delete organization.');
            deleteBtn.disabled = false;
            deleteBtn.textContent = originalText;
            return;
        }

        row?.remove();
    });

    // Edit modal close behavior
    if (editCloseButton) {
        editCloseButton.addEventListener('click', () => {
            editModal.classList.remove('show');
            currentEditingId = null;
        });
    }
    window.addEventListener('click', (e) => {
        if (e.target === editModal) {
            editModal.classList.remove('show');
            currentEditingId = null;
        }
    });

    // Edit form submit -> persist update
    if (editForm) {
        editForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const id = currentEditingId;
            if (!Number.isFinite(id)) {
                alert('Invalid organization id.');
                return;
            }
            const payload = {
                name: editNameInput.value.trim(),
                type: editTypeInput.value.trim(),
                admin_email: editEmailInput.value.trim(),
                users_allotted: Number(editUsersInput.value)
            };
            if (!payload.name || !payload.type || !payload.admin_email || !payload.users_allotted) {
                alert('Please fill in all fields.');
                return;
            }
            const saveBtn = editForm.querySelector('.save-edit-button');
            if (saveBtn) {
                saveBtn.disabled = true;
                saveBtn.textContent = 'Saving...';
            }
            const { data, error } = await supabase
                .from('organizations')
                .update(payload)
                .eq('id', id)
                .select()
                .single();
            if (error) {
                console.error('Supabase update error:', error);
                alert('Failed to update organization. You may not have permission.');
                if (saveBtn) {
                    saveBtn.disabled = false;
                    saveBtn.textContent = 'Save Changes';
                }
                return;
            }
            // Upsert admin user for this organization (create or update)
            try {
                const first = (editFirstNameInput?.value || '').trim();
                const last = (editLastNameInput?.value || '').trim();
                const fullName = [first, last].filter(Boolean).join(' ') || null;
                const adminEmail = payload.admin_email;
                if (adminEmail) {
                    // Check if user exists by email
                    const { data: existing, error: existErr } = await supabase
                        .from('users')
                        .select('id')
                        .eq('email', adminEmail)
                        .limit(1);
                    if (existErr) {
                        console.warn('Could not check existing user:', existErr);
                    }
                    const hasUser = Array.isArray(existing) && existing.length > 0;
                    if (hasUser) {
                        const { error: updErr } = await supabase
                            .from('users')
                            .update({
                                name: fullName,
                                organization: data?.name ?? payload.name,
                                organizationid: data?.id ?? id,
                                usertype: 'Company Administrator'
                            })
                            .eq('email', adminEmail);
                        if (updErr) {
                            console.error('Supabase users update error:', updErr);
                        }
                    } else if (first || last) {
                        const { error: insErr } = await supabase
                            .from('users')
                            .insert({
                                name: fullName,
                                organization: data?.name ?? payload.name,
                                email: adminEmail,
                                organizationid: data?.id ?? id,
                                usertype: 'Company Administrator'
                            });
                        if (insErr) {
                            console.error('Supabase users insert error:', insErr);
                        }
                    }
                }
            } catch (e) {
                console.warn('Upsert admin user failed:', e);
            }
            // Update UI row
            const row = orgTableBody.querySelector(`tr[data-org-id=\"${id}\"]`) || Array.from(orgTableBody.querySelectorAll('tr')).find(tr => tr.querySelector('td')?.textContent?.trim() === String(id));
            if (row) {
                const cells = row.querySelectorAll('td');
                if (cells[1]) cells[1].textContent = data?.name ?? payload.name;
                if (cells[2]) cells[2].textContent = data?.type ?? payload.type;
                if (cells[3]) cells[3].textContent = data?.admin_email ?? payload.admin_email;
                if (cells[4]) cells[4].textContent = String(data?.users_allotted ?? payload.users_allotted);
            }
            editModal.classList.remove('show');
            if (saveBtn) {
                saveBtn.disabled = false;
                saveBtn.textContent = 'Save Changes';
            }
            alert('Organization updated successfully!');
        });
    }

    addOrgButton.addEventListener("click", () => {
        modal.classList.add("show");
    });

    closeButton.addEventListener("click", () => {
        modal.classList.remove("show");
    });

    window.addEventListener("click", (e) => {
        if (e.target === modal) modal.classList.remove("show");
    });

    document.getElementById("addOrgForm").addEventListener("submit", async (e) => {
        e.preventDefault();
        const nameInput = document.getElementById("orgName");
        const typeInput = document.getElementById("orgType");
        const emailInput = document.getElementById("orgEmail");
        const usersInput = document.getElementById("orgUsers");

        const orgName = nameInput.value.trim();
        const orgType = typeInput.value.trim();
        const adminEmail = emailInput.value.trim();
        const usersAllotted = Number(usersInput.value);

        if (!orgName || !orgType || !adminEmail || !usersAllotted) {
            alert('Please fill in all fields.');
            return;
        }

        const saveBtn = form.querySelector('.save-org-button');
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.textContent = 'Saving...';
        }

        try {
            const payload = {
                name: orgName,
                type: orgType,
                admin_email: adminEmail,
                users_allotted: usersAllotted
            };

            const { data, error } = await supabase.from('organizations').insert(payload).select().single();
            if (error) {
                console.error('Supabase insert error:', error);
                alert('Failed to add organization. Please check console for details.');
                return;
            }

            // Update table UI with the new organization row
            renderOrganizationRow(data);

            // Also create admin user in users table if first/last name provided
            const adminFirst = (document.getElementById('orgAdminFirstName')?.value || '').trim();
            const adminLast = (document.getElementById('orgAdminLastName')?.value || '').trim();
            if (adminFirst && adminLast) {
                const fullName = `${adminFirst} ${adminLast}`.trim();
                const { error: userInsertError } = await supabase
                    .from('users')
                    .insert({
                        name: fullName,
                        organization: orgName,
                        email: adminEmail,
                        organizationid: data?.id ?? null,
                        usertype: 'Company Administrator'
                    });
                if (userInsertError) {
                    console.error('Supabase users insert error:', userInsertError);
                    // Non-blocking: org created successfully; inform user but don’t roll back
                    alert('Organization saved, but failed to add admin user. Check console for details.');
                }
            }

            // Reset and close modal
            form.reset();
            modal.classList.remove('show');
            alert('Organization added successfully!');
        } catch (err) {
            console.error('Unexpected error:', err);
            alert('An unexpected error occurred.');
        } finally {
            if (saveBtn) {
                saveBtn.disabled = false;
                saveBtn.textContent = 'Save';
            }
        }
    });

    // Collapsible form functionality
    const addOrgToggle = document.getElementById("addOrgToggle");
    const addOrgToggleIcon = document.getElementById("addOrgToggleIcon");

    // Toggle Add Organization form
    addOrgToggle.addEventListener("click", () => {
        const modal = document.getElementById("addOrgModal");
        modal.classList.toggle("collapsed");
        
        if (modal.classList.contains("collapsed")) {
            addOrgToggleIcon.style.transform = "rotate(0deg)";
        } else {
            addOrgToggleIcon.style.transform = "rotate(180deg)";
        }
    });

    // Sidebar navigation
    const sidebarNav = document.querySelector('.sidebar .main-nav');
    const navItems = sidebarNav.querySelectorAll('.nav-item');
    const viewContents = {
        'organizations': document.getElementById('organizations-view'),
        'public-database': document.getElementById('public-database-view')
    };

    // Function to switch views
    function switchView(viewName) {
        // Update active state in sidebar
        navItems.forEach(item => {
            if (item.dataset.view === viewName) {
                item.classList.add('active');
            } else {
                item.classList.remove('active');
            }
        });

        // Show/hide view contents
        Object.keys(viewContents).forEach(key => {
            const view = viewContents[key];
            if (view) {
                if (key === viewName) {
                    view.style.display = '';
                } else {
                    view.style.display = 'none';
                }
            }
        });
    }

    // Handle sidebar navigation clicks
    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const viewName = item.dataset.view;
            if (viewName) {
                if (viewContents[viewName]) {
                    switchView(viewName);
                } else {
                    // View not implemented yet - update active state only
                    navItems.forEach(navItem => {
                        if (navItem === item) {
                            navItem.classList.add('active');
                        } else {
                            navItem.classList.remove('active');
                        }
                    });
                    // For now, hide all views if clicking on unimplemented view
                    Object.values(viewContents).forEach(view => {
                        if (view) view.style.display = 'none';
                    });
                }
            }
        });
    });

    // Kick off initial load
    loadOrganizations();
});
