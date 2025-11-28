// js/admin-portal.js
import { supabase } from './supabase-client.js';
import * as XLSX from 'https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs';

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
        'public-projects': document.getElementById('public-projects-view'),
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
                    // When entering Public Database view, refresh dropdowns
                    if (key === 'public-database') {
                        loadPublicProjectsForDbDropdown();
                        loadPublicProjectTypesForDbDropdown();
                    }
                } else {
                    view.style.display = 'none';
                }
            }
        });
    }

    // Public Database preview modal
    const publicDbPreviewModal = document.getElementById('publicDbPreviewModal');
    const publicDbPreviewBody = document.getElementById('publicDbPreviewBody');
    const closePublicDbPreview = document.getElementById('closePublicDbPreview');

    function openPublicDbPreview(entries) {
        if (!publicDbPreviewModal || !publicDbPreviewBody) return;

        publicDbPreviewBody.innerHTML = '';

        if (!entries || entries.length === 0) {
            const emptyMsg = document.createElement('p');
            emptyMsg.textContent = 'No matching Issues or Successes were found in this file.';
            publicDbPreviewBody.appendChild(emptyMsg);
        } else {
            entries.forEach(entry => {
                const wrapper = document.createElement('div');
                wrapper.className = 'public-db-entry';

                const title = document.createElement('div');
                title.className = 'public-db-entry-title';
                title.textContent = `${entry.dataId}: ${entry.type}: ${entry.text}`;
                wrapper.appendChild(title);

                if (entry.causes && entry.causes.length > 0) {
                    const ul = document.createElement('ul');
                    ul.className = 'public-db-entry-causes';
                    entry.causes.forEach(cause => {
                        const li = document.createElement('li');
                        li.textContent = cause;
                        ul.appendChild(li);
                    });
                    wrapper.appendChild(ul);
                }

                publicDbPreviewBody.appendChild(wrapper);
            });
        }

        publicDbPreviewModal.classList.add('show');
    }

    if (closePublicDbPreview && publicDbPreviewModal) {
        closePublicDbPreview.addEventListener('click', () => {
            publicDbPreviewModal.classList.remove('show');
        });

        publicDbPreviewModal.addEventListener('click', (e) => {
            if (e.target === publicDbPreviewModal) {
                publicDbPreviewModal.classList.remove('show');
            }
        });
    }

    // Public Projects inline create form (Admin Portal)
    const publicInlineForm = document.getElementById('publicInlineCreateProjectForm');
    const publicCreateSummary = document.getElementById('publicCreateSummary');
    const publicCreateRightCol = document.getElementById('publicCreateRightCol');
    const publicProjectIndustryInput = document.getElementById('publicProjectIndustry');
    const publicProjectTypeInput = document.getElementById('publicProjectType');
    const manageProjectTypesButton = document.getElementById('manageProjectTypesButton');
    const managePublicProjectsButton = document.getElementById('managePublicProjectsButton');
    const projectTypesPanel = document.getElementById('projectTypesPanel');
    const projectTypesForm = document.getElementById('projectTypesForm');
    const projectTypesBackButton = document.getElementById('projectTypesBackButton');
    const publicCreateColumns = document.getElementById('publicCreateColumns');
    const publicProjectsManagePanel = document.getElementById('publicProjectsManagePanel');
    const publicProjectsBackButton = document.getElementById('publicProjectsBackButton');
    const publicProjectSelect = document.getElementById('publicProjectSelect');
    const publicProjectsStatus = document.getElementById('publicProjectsStatus');
    const openPublicProjectEditFromManage = document.getElementById('openPublicProjectEditFromManage');
    const openPublicProjectDetailsFromManage = document.getElementById('openPublicProjectDetailsFromManage');
    const publicProjectDetailsUpload = document.getElementById('publicProjectDetailsUpload');
    const publicProjectDetailsFileInput = document.getElementById('publicProjectDetailsFileInput');
    const publicProjectDetailsValidateButton = document.getElementById('publicProjectDetailsValidateButton');
    const publicProjectDetailsStatus = document.getElementById('publicProjectDetailsStatus');

    let currentPublicProject = null;

    if (publicInlineForm && publicCreateSummary && publicCreateRightCol) {
        publicInlineForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const name = document.getElementById('publicProjectName').value.trim();
            const industry = (document.getElementById('publicProjectIndustry')?.value || '').trim();
            const type = document.getElementById('publicProjectType').value.trim();

            if (!name || !type || !industry) {
                alert('Please fill in Project Name, Industry, and Type.');
                return;
            }

            const assetRaw = (document.getElementById('publicAssetType').value || '').trim();
            const descRaw = (document.getElementById('publicProjectDescription').value || '').trim();
            const startRaw = (document.getElementById('publicStartDate').value || '').trim();
            const endRaw = (document.getElementById('publicEndDate').value || '').trim();

            const asset = assetRaw === 'N/A' || assetRaw === '' ? null : assetRaw;
            const desc = descRaw || null;
            const start = startRaw || null;
            const end = endRaw || null;

            const saveBtn = publicInlineForm.querySelector('.save-project-button');
            if (saveBtn) {
                saveBtn.disabled = true;
                saveBtn.textContent = 'Saving...';
            }

            try {
                // 1) Find or create matching public project_type
                let projectTypeId = null;
                const { data: existingTypes, error: selectErr } = await supabase
                    .from('project_type')
                    .select('id, project_type, industry, is_public')
                    .eq('is_public', true)
                    .eq('project_type', type)
                    .eq('industry', industry)
                    .limit(1);

                if (selectErr) {
                    console.error('Error looking up project_type:', selectErr);
                    alert('Failed to look up project type. Please try again.');
                    if (saveBtn) {
                        saveBtn.disabled = false;
                        saveBtn.textContent = 'Create Public Project';
                    }
                    return;
                }

                if (existingTypes && existingTypes.length > 0) {
                    projectTypeId = existingTypes[0].id;
                } else {
                    const { data: insertedTypes, error: insertTypeErr } = await supabase
                        .from('project_type')
                        .insert({
                            project_type: type,
                            industry,
                            organization_id: null,
                            is_public: true
                        })
                        .select('id')
                        .single();

                    if (insertTypeErr) {
                        console.error('Error inserting project_type:', insertTypeErr);
                        alert('Failed to create project type. Please try again.');
                        if (saveBtn) {
                            saveBtn.disabled = false;
                            saveBtn.textContent = 'Create Public Project';
                        }
                        return;
                    }

                    projectTypeId = insertedTypes.id;
                }

                // 2) Insert project into projects table
                const projectPayload = {
                    project_name: name,
                    project_type_id: projectTypeId,
                    project_description: desc,
                    asset_new_existing: asset,
                    start_date: start,
                    end_date: end,
                    organization_id: null
                };

                const { data: projectRow, error: projectErr } = await supabase
                    .from('projects')
                    .insert(projectPayload)
                    .select()
                    .single();

                if (projectErr) {
                    console.error('Error inserting project:', projectErr);
                    alert('Failed to create public project. Please try again.');
                    if (saveBtn) {
                        saveBtn.disabled = false;
                        saveBtn.textContent = 'Create Public Project';
                    }
                    return;
                }

                // 3) Update local state and summary for UI
                currentPublicProject = {
                    // Local id corresponds to projects.project_id in the database
                    id: projectRow?.project_id ?? null,
                    name,
                    industry,
                    type,
                    asset: asset ?? 'N/A',
                    desc: desc ?? 'N/A',
                    start: start ?? 'N/A',
                    end: end ?? 'N/A'
                };

                publicCreateSummary.innerHTML = `
                    <div><strong>Project Name:</strong> ${name}</div>
                    <div><strong>Industry:</strong> ${industry}</div>
                    <div><strong>Type:</strong> ${type}</div>
                    <div><strong>New Asset or Existing:</strong> ${currentPublicProject.asset}</div>
                    <div><strong>Description:</strong> ${currentPublicProject.desc}</div>
                    <div><strong>Start Date:</strong> ${currentPublicProject.start}</div>
                    <div><strong>End Date:</strong> ${currentPublicProject.end}</div>
                `;
                publicCreateSummary.style.display = '';

                // Insert only the Project Details button (no Project Team / Lessons Learned Metadata)
                if (!publicCreateRightCol.querySelector('#publicProjectDetailsBtn')) {
                    publicCreateRightCol.insertAdjacentHTML('beforeend', `
                        <button id="publicProjectDetailsBtn" class="side-button">Project Details</button>
                    `);
                }

                alert('Public project created successfully.');
            } catch (err) {
                console.error('Unexpected error creating public project:', err);
                alert('An unexpected error occurred while creating the public project.');
            } finally {
                if (saveBtn) {
                    saveBtn.disabled = false;
                    saveBtn.textContent = 'Create Public Project';
                }
            }
        });
    }

    // Autocomplete for Industry field in Create Public Project form
    if (publicProjectIndustryInput) {
        let industryDropdown = null;
        let lastQuery = '';
        let debounceTimer = null;

        function closeIndustryDropdown() {
            if (industryDropdown && industryDropdown.parentNode) {
                industryDropdown.parentNode.removeChild(industryDropdown);
            }
            industryDropdown = null;
        }

        function createIndustryDropdown() {
            if (industryDropdown) return industryDropdown;
            industryDropdown = document.createElement('div');
            industryDropdown.className = 'autocomplete-list';
            publicProjectIndustryInput.parentNode.appendChild(industryDropdown);
            return industryDropdown;
        }

        async function fetchIndustrySuggestions(query) {
            // Don't query for very short input
            if (!query || query.length < 2) {
                closeIndustryDropdown();
                return;
            }

            // Avoid duplicate queries for same text
            if (query === lastQuery) return;
            lastQuery = query;

            try {
                const { data, error } = await supabase
                    .from('project_type')
                    .select('industry, is_public')
                    .ilike('industry', `%${query}%`)
                    .eq('is_public', true)
                    .limit(10);

                if (error) {
                    console.error('Error fetching industry suggestions:', error);
                    return;
                }

                const suggestions = (data || [])
                    .map(row => row.industry)
                    .filter(Boolean);

                if (suggestions.length === 0) {
                    closeIndustryDropdown();
                    return;
                }

                const uniqueSuggestions = Array.from(new Set(suggestions));
                const dropdown = createIndustryDropdown();
                dropdown.innerHTML = '';

                uniqueSuggestions.forEach(value => {
                    const item = document.createElement('div');
                    item.className = 'autocomplete-item';
                    item.textContent = value;
                    item.addEventListener('mousedown', (e) => {
                        // Use mousedown so it fires before input blur
                        e.preventDefault();
                        publicProjectIndustryInput.value = value;
                        closeIndustryDropdown();
                    });
                    dropdown.appendChild(item);
                });
            } catch (err) {
                console.error('Unexpected error fetching industry suggestions:', err);
            }
        }

        publicProjectIndustryInput.addEventListener('input', () => {
            const query = publicProjectIndustryInput.value.trim();
            if (debounceTimer) {
                clearTimeout(debounceTimer);
            }
            debounceTimer = setTimeout(() => {
                fetchIndustrySuggestions(query);
            }, 250);
        });

        publicProjectIndustryInput.addEventListener('focus', () => {
            const query = publicProjectIndustryInput.value.trim();
            if (query.length >= 2) {
                fetchIndustrySuggestions(query);
            }
        });

        publicProjectIndustryInput.addEventListener('blur', () => {
            // Slight delay so click on suggestion can register
            setTimeout(() => {
                closeIndustryDropdown();
            }, 150);
        });
    }

    // Autocomplete for Type field in Create Public Project form
    if (publicProjectTypeInput) {
        let typeDropdown = null;
        let lastTypeQuery = '';
        let typeDebounceTimer = null;

        function closeTypeDropdown() {
            if (typeDropdown && typeDropdown.parentNode) {
                typeDropdown.parentNode.removeChild(typeDropdown);
            }
            typeDropdown = null;
        }

        function createTypeDropdown() {
            if (typeDropdown) return typeDropdown;
            typeDropdown = document.createElement('div');
            typeDropdown.className = 'autocomplete-list';
            publicProjectTypeInput.parentNode.appendChild(typeDropdown);
            return typeDropdown;
        }

        async function fetchTypeSuggestions(query) {
            if (!query || query.length < 2) {
                closeTypeDropdown();
                return;
            }

            if (query === lastTypeQuery) return;
            lastTypeQuery = query;

            try {
                const { data, error } = await supabase
                    .from('project_type')
                    .select('project_type, is_public')
                    .ilike('project_type', `%${query}%`)
                    .eq('is_public', true)
                    .limit(10);

                if (error) {
                    console.error('Error fetching type suggestions:', error);
                    return;
                }

                const suggestions = (data || [])
                    .map(row => row.project_type)
                    .filter(Boolean);

                if (suggestions.length === 0) {
                    closeTypeDropdown();
                    return;
                }

                const uniqueSuggestions = Array.from(new Set(suggestions));
                const dropdown = createTypeDropdown();
                dropdown.innerHTML = '';

                uniqueSuggestions.forEach(value => {
                    const item = document.createElement('div');
                    item.className = 'autocomplete-item';
                    item.textContent = value;
                    item.addEventListener('mousedown', (e) => {
                        e.preventDefault();
                        publicProjectTypeInput.value = value;
                        closeTypeDropdown();
                    });
                    dropdown.appendChild(item);
                });
            } catch (err) {
                console.error('Unexpected error fetching type suggestions:', err);
            }
        }

        publicProjectTypeInput.addEventListener('input', () => {
            const query = publicProjectTypeInput.value.trim();
            if (typeDebounceTimer) {
                clearTimeout(typeDebounceTimer);
            }
            typeDebounceTimer = setTimeout(() => {
                fetchTypeSuggestions(query);
            }, 250);
        });

        publicProjectTypeInput.addEventListener('focus', () => {
            const query = publicProjectTypeInput.value.trim();
            if (query.length >= 2) {
                fetchTypeSuggestions(query);
            }
        });

        publicProjectTypeInput.addEventListener('blur', () => {
            setTimeout(() => {
                closeTypeDropdown();
            }, 150);
        });
    }

    // Manage Project Types panel behavior
    if (manageProjectTypesButton && projectTypesPanel && publicCreateColumns) {
        manageProjectTypesButton.addEventListener('click', () => {
            projectTypesPanel.style.display = '';
            publicCreateColumns.style.display = 'none';
            projectTypesPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    }

    if (projectTypesBackButton && projectTypesPanel && publicCreateColumns) {
        projectTypesBackButton.addEventListener('click', () => {
            projectTypesPanel.style.display = 'none';
            publicCreateColumns.style.display = '';
            publicCreateColumns.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    }

    if (projectTypesForm) {
        projectTypesForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const type = document.getElementById('projectTypeName').value.trim();
            const industry = document.getElementById('projectTypeIndustry').value.trim();
            if (!type || !industry) {
                alert('Please fill in both Project Type and Industry.');
                return;
            }

            const submitBtn = projectTypesForm.querySelector('button[type="submit"]');
            if (submitBtn) {
                submitBtn.disabled = true;
                submitBtn.textContent = 'Saving...';
            }

            try {
                const payload = {
                    project_type: type,
                    industry,
                    organization_id: null,
                    is_public: true
                };

                const { error } = await supabase
                    .from('project_type')
                    .insert(payload);

                if (error) {
                    console.error('Supabase project_type insert error:', error);
                    alert('Failed to save project type. Please try again.');
                    return;
                }

                alert(`Project Type "${type}" for Industry "${industry}" saved successfully.`);
                projectTypesForm.reset();
            } catch (err) {
                console.error('Unexpected error saving project type:', err);
                alert('An unexpected error occurred while saving the project type.');
            } finally {
                if (submitBtn) {
                    submitBtn.disabled = false;
                    submitBtn.textContent = 'Save Project Type';
                }
            }
        });
    }

    // Load public projects into Manage Public Projects dropdown
    async function loadPublicProjectsForDropdown() {
        if (!publicProjectSelect || !publicProjectsStatus) return;

        publicProjectSelect.innerHTML = '<option value="">Loading...</option>';
        publicProjectSelect.disabled = true;
        publicProjectsStatus.textContent = '';

        try {
            // 1) Get all public project types
            const { data: typeRows, error: typeErr } = await supabase
                .from('project_type')
                .select('id')
                .eq('is_public', true);

            if (typeErr) {
                console.error('Error loading public project types:', typeErr);
                publicProjectsStatus.textContent = 'Failed to load public project types.';
                publicProjectSelect.innerHTML = '<option value=\"\">Error loading projects</option>';
                return;
            }

            if (!typeRows || typeRows.length === 0) {
                publicProjectsStatus.textContent = 'No public project types found.';
                publicProjectSelect.innerHTML = '<option value=\"\">No public projects available</option>';
                return;
            }

            const typeIds = typeRows
                .map(row => row && row.id)
                .filter(id => id !== null && id !== undefined);

            if (typeIds.length === 0) {
                publicProjectsStatus.textContent = 'No public project types found.';
                publicProjectSelect.innerHTML = '<option value=\"\">No public projects available</option>';
                return;
            }

            // 2) Load projects that use those public project_type ids
            //    and have organization_id = null (public projects only)
            const { data: projectRows, error: projectErr } = await supabase
                .from('projects')
                .select('project_id, project_name, project_type_id, organization_id')
                .in('project_type_id', typeIds)
                .is('organization_id', null);

            if (projectErr) {
                console.error('Error loading public projects:', projectErr);
                publicProjectsStatus.textContent = 'Failed to load public projects.';
                publicProjectSelect.innerHTML = '<option value=\"\">Error loading projects</option>';
                return;
            }

            if (!projectRows || projectRows.length === 0) {
                publicProjectsStatus.textContent = 'No public projects found.';
                publicProjectSelect.innerHTML = '<option value=\"\">No public projects available</option>';
                return;
            }

            publicProjectSelect.innerHTML = '<option value=\"\">Select Public Project</option>';
            projectRows.forEach(project => {
                if (!project || !project.project_id || !project.project_name) return;
                const option = document.createElement('option');
                option.value = project.project_id;
                option.textContent = project.project_name;
                publicProjectSelect.appendChild(option);
            });

            publicProjectSelect.disabled = false;
            publicProjectsStatus.textContent = '';
        } catch (err) {
            console.error('Unexpected error loading public projects:', err);
            publicProjectsStatus.textContent = 'An unexpected error occurred while loading public projects.';
            publicProjectSelect.innerHTML = '<option value=\"\">Error loading projects</option>';
            publicProjectSelect.disabled = true;
        }
    }

    // Load full details for a single public project (for editing)
    async function loadSinglePublicProject(projectId) {
        if (!projectId || !publicProjectsStatus) return;

        publicProjectsStatus.textContent = 'Loading project details...';

        try {
            // Load project row
            const { data: projectRow, error: projectErr } = await supabase
                .from('projects')
                .select('project_id, project_name, project_type_id, project_description, asset_new_existing, start_date, end_date, organization_id')
                .eq('project_id', projectId)
                .is('organization_id', null)
                .single();

            if (projectErr) {
                console.error('Error loading public project details:', projectErr);
                publicProjectsStatus.textContent = 'Failed to load project details.';
                return;
            }

            if (!projectRow) {
                publicProjectsStatus.textContent = 'Project not found.';
                return;
            }

            // Load project_type row for type + industry
            let industry = '';
            let type = '';
            if (projectRow.project_type_id != null) {
                const { data: typeRow, error: typeErr } = await supabase
                    .from('project_type')
                    .select('project_type, industry')
                    .eq('id', projectRow.project_type_id)
                    .single();

                if (typeErr) {
                    console.error('Error loading project type for public project:', typeErr);
                } else if (typeRow) {
                    type = typeRow.project_type || '';
                    industry = typeRow.industry || '';
                }
            }

            currentPublicProject = {
                id: projectRow.project_id,
                name: projectRow.project_name || '',
                industry: industry,
                type: type,
                asset: projectRow.asset_new_existing || 'N/A',
                desc: projectRow.project_description || 'N/A',
                start: projectRow.start_date || 'N/A',
                end: projectRow.end_date || 'N/A'
            };

            publicProjectsStatus.textContent = '';
        } catch (err) {
            console.error('Unexpected error loading public project details:', err);
            publicProjectsStatus.textContent = 'An unexpected error occurred while loading project details.';
        }
    }

    // Manage Public Projects sub-module behavior
    if (managePublicProjectsButton && publicProjectsManagePanel && publicCreateColumns) {
        managePublicProjectsButton.addEventListener('click', () => {
            // Hide create form and project types panel
            publicCreateColumns.style.display = 'none';
            if (projectTypesPanel) {
                projectTypesPanel.style.display = 'none';
            }

            // Show Manage Public Projects panel
            publicProjectsManagePanel.style.display = '';
            loadPublicProjectsForDropdown();
            publicProjectsManagePanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    }

    if (publicProjectsBackButton && publicProjectsManagePanel && publicCreateColumns) {
        publicProjectsBackButton.addEventListener('click', () => {
            publicProjectsManagePanel.style.display = 'none';
            publicCreateColumns.style.display = '';
            publicCreateColumns.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    }

    // When a project is selected in the dropdown, load its details and show Edit button
    if (publicProjectSelect && openPublicProjectEditFromManage && openPublicProjectDetailsFromManage) {
        publicProjectSelect.addEventListener('change', async () => {
            const value = publicProjectSelect.value;
            if (!value) {
                openPublicProjectEditFromManage.style.display = 'none';
                openPublicProjectDetailsFromManage.style.display = 'none';
                if (publicProjectDetailsUpload) {
                    publicProjectDetailsUpload.style.display = 'none';
                }
                currentPublicProject = null;
                return;
            }

            await loadSinglePublicProject(value);
            if (currentPublicProject && currentPublicProject.id) {
                openPublicProjectEditFromManage.style.display = '';
                openPublicProjectDetailsFromManage.style.display = '';
            } else {
                openPublicProjectEditFromManage.style.display = 'none';
                openPublicProjectDetailsFromManage.style.display = 'none';
            }
        });

        openPublicProjectEditFromManage.addEventListener('click', () => {
            if (!currentPublicProject || !currentPublicProject.id) {
                alert('Please select a valid public project first.');
                return;
            }
            openPublicProjectEditModal();
        });

        // Show Project Details upload section when clicked
        openPublicProjectDetailsFromManage.addEventListener('click', () => {
            if (!currentPublicProject || !currentPublicProject.id) {
                alert('Please select a valid public project first.');
                return;
            }
            if (publicProjectDetailsUpload) {
                publicProjectDetailsUpload.style.display = '';
            }
            if (publicProjectDetailsStatus) {
                publicProjectDetailsStatus.textContent = '';
            }
            if (publicProjectDetailsFileInput) {
                publicProjectDetailsFileInput.value = '';
            }
        });
    }

    // Read Excel for Project Details and write rows to Supabase project_details
    if (publicProjectDetailsValidateButton && publicProjectDetailsFileInput && publicProjectDetailsStatus) {
        publicProjectDetailsValidateButton.addEventListener('click', async () => {
            publicProjectDetailsStatus.classList.remove('upload-message--success', 'upload-message--error');

            if (!currentPublicProject || !currentPublicProject.id) {
                publicProjectDetailsStatus.textContent = 'Please select a public project before uploading details.';
                publicProjectDetailsStatus.classList.add('upload-message--error');
                return;
            }

            const file = publicProjectDetailsFileInput.files && publicProjectDetailsFileInput.files[0];
            if (!file) {
                publicProjectDetailsStatus.textContent = 'Please choose a file first.';
                publicProjectDetailsStatus.classList.add('upload-message--error');
                return;
            }

            const name = file.name || '';
            const isExcel = /\.xlsx$/i.test(name) || /\.xls$/i.test(name);
            if (!isExcel) {
                publicProjectDetailsStatus.textContent = 'Only Excel files (.xlsx, .xls) are supported.';
                publicProjectDetailsStatus.classList.add('upload-message--error');
                return;
            }

            try {
                publicProjectDetailsStatus.textContent = `Reading "${name}"...`;

                const arrayBuffer = await file.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });

                // Use the first sheet
                const firstSheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

                if (!rows || rows.length === 0) {
                    publicProjectDetailsStatus.textContent = 'The Excel file is empty.';
                    publicProjectDetailsStatus.classList.add('upload-message--error');
                    return;
                }

                const headerRow = rows[0];
                const normalize = (v) => String(v || '').trim().toLowerCase();

                let idxParamName = -1;
                let idxParamEntry = -1;
                headerRow.forEach((cell, idx) => {
                    const key = normalize(cell);
                    if (key === 'parameter name') idxParamName = idx;
                    if (key === 'parameter entry') idxParamEntry = idx;
                });

                if (idxParamName === -1 || idxParamEntry === -1) {
                    publicProjectDetailsStatus.textContent = 'Could not find required headers "Parameter Name" and "Parameter Entry".';
                    publicProjectDetailsStatus.classList.add('upload-message--error');
                    return;
                }

                const projectId = currentPublicProject.id;
                const inserts = [];

                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row) continue;
                    const rawName = row[idxParamName];
                    const rawEntry = row[idxParamEntry];
                    const paramName = rawName != null ? String(rawName).trim() : '';
                    const paramEntry = rawEntry != null ? String(rawEntry).trim() : '';
                    if (!paramName && !paramEntry) continue;

                    inserts.push({
                        project_id: projectId,
                        organization_id: null,
                        parameter_name: paramName,
                        parameter_entry: paramEntry
                    });
                }

                if (inserts.length === 0) {
                    publicProjectDetailsStatus.textContent = 'No parameter rows found under the required headers.';
                    publicProjectDetailsStatus.classList.add('upload-message--error');
                    return;
                }

                const { error } = await supabase
                    .from('project_details')
                    .insert(inserts);

                if (error) {
                    console.error('Error inserting project_details rows:', error);
                    publicProjectDetailsStatus.textContent = 'Failed to save project details to the database.';
                    publicProjectDetailsStatus.classList.add('upload-message--error');
                    return;
                }

                publicProjectDetailsStatus.textContent = `Saved ${inserts.length} project detail entries for this project.`;
                publicProjectDetailsStatus.classList.add('upload-message--success');
            } catch (err) {
                console.error('Unexpected error processing project details Excel:', err);
                publicProjectDetailsStatus.textContent = 'An unexpected error occurred while processing the Excel file.';
                publicProjectDetailsStatus.classList.add('upload-message--error');
            }
        });
    }

    // Edit Public Project modal wiring
    const publicProjectEditModal = document.getElementById('publicProjectEditModal');
    const closePublicProjectEdit = document.getElementById('closePublicProjectEdit');
    const publicProjectEditForm = document.getElementById('publicProjectEditForm');
    const deletePublicProjectBtn = document.getElementById('deletePublicProjectBtn');

    function openPublicProjectEditModal() {
        if (!publicProjectEditModal || !currentPublicProject) return;

        // Populate fields
        document.getElementById('editPublicProjectName').value = currentPublicProject.name || '';
        document.getElementById('editPublicProjectIndustry').value = currentPublicProject.industry || '';
        document.getElementById('editPublicProjectType').value = currentPublicProject.type || '';
        document.getElementById('editPublicAssetType').value = currentPublicProject.asset === 'N/A' ? '' : (currentPublicProject.asset || '');
        document.getElementById('editPublicProjectDescription').value = currentPublicProject.desc || '';
        document.getElementById('editPublicStartDate').value = currentPublicProject.start === 'N/A' ? '' : (currentPublicProject.start || '');
        document.getElementById('editPublicEndDate').value = currentPublicProject.end === 'N/A' ? '' : (currentPublicProject.end || '');

        publicProjectEditModal.classList.add('show');
    }

    // Note: inline Edit button has been removed from the create form.

    if (closePublicProjectEdit && publicProjectEditModal) {
        closePublicProjectEdit.addEventListener('click', () => {
            publicProjectEditModal.classList.remove('show');
        });

        publicProjectEditModal.addEventListener('click', (e) => {
            if (e.target === publicProjectEditModal) {
                publicProjectEditModal.classList.remove('show');
            }
        });
    }

    if (publicProjectEditForm) {
        publicProjectEditForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!currentPublicProject) {
                publicProjectEditModal.classList.remove('show');
                return;
            }

            const name = document.getElementById('editPublicProjectName').value.trim();
            const industry = (document.getElementById('editPublicProjectIndustry')?.value || '').trim();
            const type = document.getElementById('editPublicProjectType').value.trim();

            if (!name || !type || !industry) {
                alert('Please fill in Project Name, Industry, and Type.');
                return;
            }

            const assetRaw = (document.getElementById('editPublicAssetType').value || '').trim();
            const descRaw = (document.getElementById('editPublicProjectDescription').value || '').trim();
            const startRaw = (document.getElementById('editPublicStartDate').value || '').trim();
            const endRaw = (document.getElementById('editPublicEndDate').value || '').trim();

            const asset = assetRaw === 'N/A' || assetRaw === '' ? null : assetRaw;
            const desc = descRaw || null;
            const start = startRaw || null;
            const end = endRaw || null;

            const saveBtn = publicProjectEditForm.querySelector('button[type="submit"]');
            if (saveBtn) {
                saveBtn.disabled = true;
                saveBtn.textContent = 'Saving...';
            }

            try {
                // 1) Find or create matching public project_type
                let projectTypeId = null;
                const { data: existingTypes, error: selectErr } = await supabase
                    .from('project_type')
                    .select('id, project_type, industry, is_public')
                    .eq('is_public', true)
                    .eq('project_type', type)
                    .eq('industry', industry)
                    .limit(1);

                if (selectErr) {
                    console.error('Error looking up project_type (edit):', selectErr);
                    alert('Failed to look up project type. Please try again.');
                    if (saveBtn) {
                        saveBtn.disabled = false;
                        saveBtn.textContent = 'Save Edits';
                    }
                    return;
                }

                if (existingTypes && existingTypes.length > 0) {
                    projectTypeId = existingTypes[0].id;
                } else {
                    const { data: insertedTypes, error: insertTypeErr } = await supabase
                        .from('project_type')
                        .insert({
                            project_type: type,
                            industry,
                            organization_id: null,
                            is_public: true
                        })
                        .select('id')
                        .single();

                    if (insertTypeErr) {
                        console.error('Error inserting project_type (edit):', insertTypeErr);
                        alert('Failed to create project type. Please try again.');
                        if (saveBtn) {
                            saveBtn.disabled = false;
                            saveBtn.textContent = 'Save Edits';
                        }
                        return;
                    }

                    projectTypeId = insertedTypes.id;
                }

                // 2) Ensure we know which project row to edit.
                //    If currentPublicProject.id is missing (e.g., after a refresh),
                //    try to look it up safely based on name + project_type_id + organization_id null.
                const projectPayload = {
                    project_name: name,
                    project_type_id: projectTypeId,
                    project_description: desc,
                    asset_new_existing: asset,
                    start_date: start,
                    end_date: end,
                    organization_id: null
                };

                if (!currentPublicProject.id) {
                    const { data: matches, error: lookupErr } = await supabase
                        .from('projects')
                        .select('project_id')
                        .eq('project_name', name)
                        .eq('project_type_id', projectTypeId)
                        .is('organization_id', null)
                        .order('project_id', { ascending: false })
                        .limit(1);

                    if (lookupErr) {
                        console.error('Error looking up existing project to edit:', lookupErr);
                        alert('Could not find the original public project to edit. Please recreate it.');
                        if (saveBtn) {
                            saveBtn.disabled = false;
                            saveBtn.textContent = 'Save Edits';
                        }
                        return;
                    }

                    if (!matches || matches.length === 0) {
                        alert('Could not find the original public project to edit. Please recreate it.');
                        if (saveBtn) {
                            saveBtn.disabled = false;
                            saveBtn.textContent = 'Save Edits';
                        }
                        return;
                    }

                    currentPublicProject.id = matches[0].project_id;
                }

                const { data: projectRow, error: projectErr } = await supabase
                    .from('projects')
                    .update(projectPayload)
                    .eq('project_id', currentPublicProject.id)
                    .select()
                    .single();

                if (projectErr) {
                    console.error('Error updating project (edit):', projectErr);
                    alert('Failed to save edits to public project. Please try again.');
                    if (saveBtn) {
                        saveBtn.disabled = false;
                        saveBtn.textContent = 'Save Edits';
                    }
                    return;
                }

                // 3) Update local state and main form
                currentPublicProject = {
                    id: projectRow?.project_id ?? currentPublicProject.id ?? null,
                    name,
                    industry,
                    type,
                    asset: asset ?? 'N/A',
                    desc: desc ?? 'N/A',
                    start: start ?? 'N/A',
                    end: end ?? 'N/A'
                };

                // Update main form fields
                document.getElementById('publicProjectName').value = name;
                document.getElementById('publicProjectIndustry').value = industry;
                document.getElementById('publicProjectType').value = type;
                document.getElementById('publicAssetType').value = currentPublicProject.asset === 'N/A' ? '' : currentPublicProject.asset;
                document.getElementById('publicProjectDescription').value = currentPublicProject.desc === 'N/A' ? '' : currentPublicProject.desc;
                document.getElementById('publicStartDate').value = currentPublicProject.start === 'N/A' ? '' : currentPublicProject.start;
                document.getElementById('publicEndDate').value = currentPublicProject.end === 'N/A' ? '' : currentPublicProject.end;

                // Update summary
                publicCreateSummary.innerHTML = `
                    <div><strong>Project Name:</strong> ${currentPublicProject.name}</div>
                    <div><strong>Industry:</strong> ${currentPublicProject.industry}</div>
                    <div><strong>Type:</strong> ${currentPublicProject.type}</div>
                    <div><strong>New Asset or Existing:</strong> ${currentPublicProject.asset}</div>
                    <div><strong>Description:</strong> ${currentPublicProject.desc}</div>
                    <div><strong>Start Date:</strong> ${currentPublicProject.start}</div>
                    <div><strong>End Date:</strong> ${currentPublicProject.end}</div>
                `;
                publicCreateSummary.style.display = '';

                alert('Public project edits saved successfully.');
                publicProjectEditModal.classList.remove('show');
            } catch (err) {
                console.error('Unexpected error editing public project:', err);
                alert('An unexpected error occurred while saving edits to the public project.');
            } finally {
                if (saveBtn) {
                    saveBtn.disabled = false;
                    saveBtn.textContent = 'Save Edits';
                }
            }
        });
    }

    if (deletePublicProjectBtn) {
        deletePublicProjectBtn.addEventListener('click', () => {
            if (!currentPublicProject) {
                publicProjectEditModal.classList.remove('show');
                return;
            }
            const confirmed = confirm('Are you sure you want to delete this public project?');
            if (!confirmed) return;

            // Clear current project data
            currentPublicProject = null;

            // Clear form fields
            document.getElementById('publicProjectName').value = '';
            document.getElementById('publicProjectType').value = '';
            document.getElementById('publicAssetType').value = '';
            document.getElementById('publicProjectDescription').value = '';
            document.getElementById('publicStartDate').value = '';
            document.getElementById('publicEndDate').value = '';

            // Hide summary
            publicCreateSummary.innerHTML = '';
            publicCreateSummary.style.display = 'none';

            // Hide buttons
            if (publicEditProjectBtn) {
                publicEditProjectBtn.style.display = 'none';
            }
            const detailsBtn = publicCreateRightCol.querySelector('#publicProjectDetailsBtn');
            if (detailsBtn) {
                detailsBtn.remove();
            }

            publicProjectEditModal.classList.remove('show');
        });
    }

    // Public Database file dropzone setup
    const publicDbDropzone = document.getElementById('publicDbDropzone');
    const publicDbFileInput = document.getElementById('publicDbFileInput');
    const publicDbMessage = document.getElementById('publicDbMessage');
    const publicDbProjectSelect = document.getElementById('publicDbProjectSelect');
    const publicDbProjectTypeSelect = document.getElementById('publicDbProjectTypeSelect');

    function setPublicDbMessage(text, type) {
        if (!publicDbMessage) return;
        publicDbMessage.textContent = text;
        publicDbMessage.classList.remove('upload-message--success', 'upload-message--error');
        if (type === 'success') {
            publicDbMessage.classList.add('upload-message--success');
        } else if (type === 'error') {
            publicDbMessage.classList.add('upload-message--error');
        }
    }

    async function loadPublicProjectsForDbDropdown() {
        if (!publicDbProjectSelect) return;

        publicDbProjectSelect.innerHTML = '<option value=\"\">Loading...</option>';
        publicDbProjectSelect.disabled = true;

        try {
            // 1) Get all public project types
            const { data: typeRows, error: typeErr } = await supabase
                .from('project_type')
                .select('id')
                .eq('is_public', true);

            if (typeErr) {
                console.error('Error loading public project types for Public Database:', typeErr);
                setPublicDbMessage('Failed to load public project types.', 'error');
                publicDbProjectSelect.innerHTML = '<option value=\"\">Error loading projects</option>';
                return;
            }

            if (!typeRows || typeRows.length === 0) {
                setPublicDbMessage('No public project types found.', 'error');
                publicDbProjectSelect.innerHTML = '<option value=\"\">No public projects available</option>';
                return;
            }

            const typeIds = typeRows
                .map(row => row && row.id)
                .filter(id => id !== null && id !== undefined);

            if (typeIds.length === 0) {
                setPublicDbMessage('No public project types found.', 'error');
                publicDbProjectSelect.innerHTML = '<option value=\"\">No public projects available</option>';
                return;
            }

            // 2) Load projects that use those public project_type ids
            //    and have organization_id = null (public projects only)
            const { data: projectRows, error: projectErr } = await supabase
                .from('projects')
                .select('project_id, project_name, project_type_id, organization_id')
                .in('project_type_id', typeIds)
                .is('organization_id', null);

            if (projectErr) {
                console.error('Error loading public projects for Public Database:', projectErr);
                setPublicDbMessage('Failed to load public projects.', 'error');
                publicDbProjectSelect.innerHTML = '<option value=\"\">Error loading projects</option>';
                return;
            }

            if (!projectRows || projectRows.length === 0) {
                setPublicDbMessage('No public projects found.', 'error');
                publicDbProjectSelect.innerHTML = '<option value=\"\">No public projects available</option>';
                return;
            }

            publicDbProjectSelect.innerHTML = '<option value=\"\">Select Public Project</option>';
            projectRows.forEach(project => {
                if (!project || !project.project_id || !project.project_name) return;
                const option = document.createElement('option');
                option.value = project.project_id;
                option.textContent = project.project_name;
                publicDbProjectSelect.appendChild(option);
            });

            publicDbProjectSelect.disabled = false;
        } catch (err) {
            console.error('Unexpected error loading public projects for Public Database:', err);
            setPublicDbMessage('An unexpected error occurred while loading public projects.', 'error');
            publicDbProjectSelect.innerHTML = '<option value=\"\">Error loading projects</option>';
            publicDbProjectSelect.disabled = true;
        }
    }

    async function loadPublicProjectTypesForDbDropdown() {
        if (!publicDbProjectTypeSelect) return;

        publicDbProjectTypeSelect.innerHTML = '<option value=\"\">Loading...</option>';
        publicDbProjectTypeSelect.disabled = true;

        try {
            const { data, error } = await supabase
                .from('project_type')
                .select('id, project_type, is_public')
                .eq('is_public', true)
                .order('project_type', { ascending: true });

            if (error) {
                console.error('Error loading public project types for Public Database:', error);
                setPublicDbMessage('Failed to load public project types.', 'error');
                publicDbProjectTypeSelect.innerHTML = '<option value=\"\">Error loading project types</option>';
                return;
            }

            if (!data || data.length === 0) {
                setPublicDbMessage('No public project types found.', 'error');
                publicDbProjectTypeSelect.innerHTML = '<option value=\"\">No public project types available</option>';
                return;
            }

            publicDbProjectTypeSelect.innerHTML = '<option value=\"\">Select Public Project Type</option>';
            data.forEach(row => {
                if (!row || !row.id || !row.project_type) return;
                const option = document.createElement('option');
                option.value = row.id;
                option.textContent = row.project_type;
                publicDbProjectTypeSelect.appendChild(option);
            });

            publicDbProjectTypeSelect.disabled = false;
        } catch (err) {
            console.error('Unexpected error loading public project types for Public Database:', err);
            setPublicDbMessage('An unexpected error occurred while loading public project types.', 'error');
            publicDbProjectTypeSelect.innerHTML = '<option value=\"\">Error loading project types</option>';
            publicDbProjectTypeSelect.disabled = true;
        }
    }

    async function parsePublicDbExcel(file) {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        function normalizeHeader(value) {
            return String(value || '').trim().toLowerCase();
        }

        function findSheetWithHeaders(requiredHeaders) {
            for (const sheetName of workbook.SheetNames) {
                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
                if (!rows || rows.length === 0) continue;

                const headerRow = rows[0];
                const headerMap = {};
                headerRow.forEach((cell, idx) => {
                    const key = normalizeHeader(cell);
                    if (key) headerMap[key] = idx;
                });

                const hasAll = requiredHeaders.every(h => headerMap[normalizeHeader(h)] !== undefined);
                if (hasAll) {
                    return {
                        sheetName,
                        rows,
                        headerMap
                    };
                }
            }
            return null;
        }

        // Find Issues/Successes sheet
        const issuesSheetInfo = findSheetWithHeaders(['Data ID', 'Issue', 'Success']);
        if (!issuesSheetInfo) {
            setPublicDbMessage('Could not find a sheet with columns: Data ID, Issue, Success.', 'error');
            return;
        }

        const causesSheetInfo = findSheetWithHeaders(['Data ID', 'Cause']);
        // causesSheetInfo can be null; that just means no causes

        const entriesById = new Map();

        const dataIdKey = normalizeHeader('Data ID');
        const issueKey = normalizeHeader('Issue');
        const successKey = normalizeHeader('Success');

        const idxDataId = issuesSheetInfo.headerMap[dataIdKey];
        const idxIssue = issuesSheetInfo.headerMap[issueKey];
        const idxSuccess = issuesSheetInfo.headerMap[successKey];

        for (let i = 1; i < issuesSheetInfo.rows.length; i++) {
            const row = issuesSheetInfo.rows[i];
            if (!row) continue;
            const dataIdRaw = row[idxDataId];
            const issueRaw = row[idxIssue];
            const successRaw = row[idxSuccess];

            const dataId = dataIdRaw != null ? String(dataIdRaw).trim() : '';
            const issueText = issueRaw != null ? String(issueRaw).trim() : '';
            const successText = successRaw != null ? String(successRaw).trim() : '';

            if (!dataId) continue;
            if (!issueText && !successText) continue;

            const type = issueText ? 'Issue' : 'Success';
            const text = issueText || successText;

            if (!entriesById.has(dataId)) {
                entriesById.set(dataId, {
                    dataId,
                    type,
                    text,
                    causes: []
                });
            }
        }

        // Attach causes if a causes sheet exists
        if (causesSheetInfo) {
            const causeDataIdKey = normalizeHeader('Data ID');
            const causeKey = normalizeHeader('Cause');
            const idxCauseDataId = causesSheetInfo.headerMap[causeDataIdKey];
            const idxCause = causesSheetInfo.headerMap[causeKey];

            for (let i = 1; i < causesSheetInfo.rows.length; i++) {
                const row = causesSheetInfo.rows[i];
                if (!row) continue;
                const dataIdRaw = row[idxCauseDataId];
                const causeRaw = row[idxCause];
                const dataId = dataIdRaw != null ? String(dataIdRaw).trim() : '';
                const causeText = causeRaw != null ? String(causeRaw).trim() : '';
                if (!dataId || !causeText) continue;

                const entry = entriesById.get(dataId);
                if (entry) {
                    entry.causes.push(causeText);
                }
            }
        }

        const entries = Array.from(entriesById.values());
        if (entries.length === 0) {
            setPublicDbMessage('No Issues or Successes were found in this file.', 'error');
            return;
        }

        setPublicDbMessage(`Loaded ${entries.length} Issue/Success entries from the Excel file.`, 'success');
        openPublicDbPreview(entries);
    }

    function handlePublicDbFiles(fileList) {
        if (!fileList || fileList.length === 0) {
            return;
        }

        // --- Public Database selection rules before upload ---
        // 1) If BOTH a Public Project and Public Project Type are selected, block upload.
        // 2) If NEITHER is selected, ask if user wants to upload general data not
        //    attached to a project or project type. Only proceed if they confirm.
        // 3) If exactly ONE of them is selected, allow upload.
        if (publicDbProjectSelect && publicDbProjectTypeSelect) {
            const projectValue = publicDbProjectSelect.value || '';
            const projectTypeValue = publicDbProjectTypeSelect.value || '';
            const hasProject = projectValue !== '';
            const hasProjectType = projectTypeValue !== '';

            // Case: both selected -> block upload
            if (hasProject && hasProjectType) {
                setPublicDbMessage('Please select either a Public Project or a Public Project Type, not both.', 'error');
                alert('Please select either a Public Project or a Public Project Type, not both.');
                return;
            }

            // Case: neither selected -> ask if they want general (unattached) data
            if (!hasProject && !hasProjectType) {
                const proceedAsGeneral = confirm(
                    'You have not selected a Public Project or a Public Project Type.\n\n' +
                    'Do you want to upload general data that is NOT attached to any project or project type?'
                );
                if (!proceedAsGeneral) {
                    setPublicDbMessage(
                        'Upload cancelled. Please select a Public Project or a Public Project Type, or confirm general upload.',
                        'error'
                    );
                    return;
                }
                // If they confirm, continue with the upload as general data.
            }
        }

        const file = fileList[0];
        const name = file.name || '';
        const isExcel = /\.xlsx$/i.test(name) || /\.xls$/i.test(name);

        if (isExcel) {
            setPublicDbMessage(`Reading "${name}"...`, 'success');
            parsePublicDbExcel(file).catch((err) => {
                console.error('Error parsing Excel file:', err);
                setPublicDbMessage('An error occurred while reading the Excel file.', 'error');
            });
        } else {
            setPublicDbMessage('Only Excel files (.xlsx, .xls) are supported at this time.', 'error');
        }
    }

    if (publicDbDropzone && publicDbFileInput) {
        // Click opens file picker
        publicDbDropzone.addEventListener('click', () => {
            publicDbFileInput.click();
        });

        // File selected via picker
        publicDbFileInput.addEventListener('change', (e) => {
            const files = e.target.files;
            handlePublicDbFiles(files);
        });

        // Drag & drop support
        ['dragenter', 'dragover'].forEach(eventName => {
            publicDbDropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                publicDbDropzone.classList.add('drag-over');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            publicDbDropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                publicDbDropzone.classList.remove('drag-over');
            });
        });

        publicDbDropzone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt && dt.files;
            handlePublicDbFiles(files);
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
