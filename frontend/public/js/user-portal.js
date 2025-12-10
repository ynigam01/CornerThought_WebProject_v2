// js/user-portal.js
import { supabase } from './supabase-client.js';
import * as XLSX from 'https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs';

// Require login: redirect to user-login if no session is present
try {
    const ctUser = JSON.parse(sessionStorage.getItem('ct_user'));
    if (!ctUser) {
        window.location.href = 'user-login.html';
    }
} catch (_) {
    window.location.href = 'user-login.html';
}
document.addEventListener("DOMContentLoaded", () => {
    const container = document.getElementById("projectFormContainer");

    // Get user info and check permissions
    let ctUser = null;
    let canCreateProjects = false;
    let canManageOrgSettings = false;
    
    try {
        ctUser = JSON.parse(sessionStorage.getItem('ct_user'));
        if (ctUser && ctUser.usertype) {
            const userType = String(ctUser.usertype).trim();
            // Only Administrators and Company Administrators can create projects and manage org settings
            canCreateProjects = userType === 'Administrator' || userType === 'Company Administrator';
            canManageOrgSettings = userType === 'Administrator' || userType === 'Company Administrator';
        }
    } catch (_) {}

    // Organization id for the logged-in user (used for project-type lookups and org projects)
    const organizationId = ctUser && ctUser.organizationid ? ctUser.organizationid : null;

    // Cache of organization-specific project types for use in create project forms
    let orgProjectTypes = [];
    let orgProjectTypesLoaded = false;

    // State for organization project details sub-module (Create view)
    let orgProjectDetailsPanel = null;
    let orgProjectDetailsProjectSelect = null;
    let orgProjectDetailsFileInput = null;
    let orgProjectDetailsValidateButton = null;
    let orgProjectDetailsStatus = null;
    let orgProjectDetailsBackButton = null;
    let orgProjectDetailsChoices = null;
    let currentOrgProject = null;
    let orgProjectsById = new Map();

    async function loadOrgProjectTypes() {
        if (!organizationId || orgProjectTypesLoaded) return;
        try {
            const { data, error } = await supabase
                .from('project_type')
                .select('project_type')
                .eq('organization_id', organizationId)
                .order('project_type', { ascending: true });

            if (error) {
                console.error('Error loading organization project types for Create Project:', error);
                return;
            }

            const names = (data || [])
                .map(row => row && row.project_type)
                .filter(Boolean);

            orgProjectTypes = Array.from(new Set(names));
            orgProjectTypesLoaded = true;
        } catch (err) {
            console.error('Unexpected error loading organization project types for Create Project:', err);
        }
    }

    function isValidOrgProjectType(value) {
        if (!value) return false;
        const target = String(value).trim().toLowerCase();
        return orgProjectTypes.some(name => String(name).trim().toLowerCase() === target);
    }

    // Load organization projects for the Project Details dropdown in Create view
    async function loadOrgProjectsForDropdown() {
        if (!organizationId || !orgProjectDetailsProjectSelect || !orgProjectDetailsStatus) return;

        orgProjectDetailsStatus.classList.remove('upload-message--success', 'upload-message--error');
        orgProjectDetailsStatus.textContent = 'Loading projects...';

        try {
            const { data: projectRows, error: projectErr } = await supabase
                .from('projects')
                .select('project_id, project_name')
                .eq('organization_id', organizationId)
                .order('project_name', { ascending: true });

            if (projectErr) {
                console.error('Error loading organization projects for details dropdown:', projectErr);
                orgProjectDetailsStatus.textContent = 'Failed to load projects.';
                orgProjectDetailsStatus.classList.add('upload-message--error');
                return;
            }

            const rows = projectRows || [];
            if (rows.length === 0) {
                orgProjectDetailsStatus.textContent = 'No projects found for your organization.';
                orgProjectDetailsStatus.classList.add('upload-message--error');
                if (orgProjectDetailsChoices) {
                    orgProjectDetailsChoices.clearChoices();
                } else {
                    orgProjectDetailsProjectSelect.innerHTML = '<option value=\"\">No projects available</option>';
                }
                currentOrgProject = null;
                orgProjectsById = new Map();
                return;
            }

            orgProjectsById = new Map();
            const choicesData = rows.map(row => {
                const idStr = String(row.project_id);
                const name = row.project_name || `Project ${idStr}`;
                orgProjectsById.set(idStr, name);
                return {
                    value: idStr,
                    label: name
                };
            });

            function updateCurrentOrgProjectFromSelect() {
                const value = orgProjectDetailsProjectSelect.value;
                if (!value) {
                    currentOrgProject = null;
                    return;
                }
                const name = orgProjectsById.get(value) || '';
                const numericId = Number.isNaN(Number(value)) ? value : Number(value);
                currentOrgProject = { id: numericId, name };
            }

            // Initialize Choices once
            if (!orgProjectDetailsChoices) {
                if (typeof Choices === 'undefined') {
                    console.warn('Choices library not loaded; falling back to native select for org projects dropdown.');
                    orgProjectDetailsProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                    choicesData.forEach(choice => {
                        const option = document.createElement('option');
                        option.value = choice.value;
                        option.textContent = choice.label;
                        orgProjectDetailsProjectSelect.appendChild(option);
                    });
                    orgProjectDetailsProjectSelect.addEventListener('change', updateCurrentOrgProjectFromSelect);
                } else {
                    orgProjectDetailsChoices = new Choices(orgProjectDetailsProjectSelect, {
                        searchEnabled: true,
                        shouldSort: false,
                        placeholder: true,
                        placeholderValue: 'Select Project',
                        searchPlaceholderValue: 'Type to search...'
                    });
                    orgProjectDetailsProjectSelect.addEventListener('change', updateCurrentOrgProjectFromSelect);
                }
            }

            if (orgProjectDetailsChoices) {
                orgProjectDetailsChoices.setChoices(choicesData, 'value', 'label', true);
            } else {
                orgProjectDetailsProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                choicesData.forEach(choice => {
                    const option = document.createElement('option');
                    option.value = choice.value;
                    option.textContent = choice.label;
                    orgProjectDetailsProjectSelect.appendChild(option);
                });
            }

            orgProjectDetailsStatus.textContent = '';
            orgProjectDetailsStatus.classList.remove('upload-message--error');
        } catch (err) {
            console.error('Unexpected error loading organization projects for details dropdown:', err);
            orgProjectDetailsStatus.textContent = 'An unexpected error occurred while loading projects.';
            orgProjectDetailsStatus.classList.add('upload-message--error');
        }
    }

    // Insert a new organization project into Supabase `projects` table.
    // Expects `type` to already be validated as a valid organization project type.
    async function createOrganizationProject({ name, type, assetRaw, descRaw, startRaw, endRaw }) {
        if (!organizationId) {
            alert('Could not determine your organization. Please log out and log back in, or contact your administrator.');
            return { success: false };
        }

        // Normalize optional fields
        const asset = assetRaw === 'N/A' || assetRaw === '' ? null : (assetRaw || null);
        const desc = descRaw || null;
        const start = startRaw || null;
        const end = endRaw || null;

        try {
            // Look up the project_type id for this organization and type name
            const { data: typeRows, error: typeErr } = await supabase
                .from('project_type')
                .select('id')
                .eq('organization_id', organizationId)
                .eq('project_type', type)
                .limit(1);

            if (typeErr) {
                console.error('Error looking up organization project_type:', typeErr);
                alert('Failed to look up project type. Please try again.');
                return { success: false };
            }

            if (!typeRows || typeRows.length === 0) {
                alert('The selected Project Type could not be found for your organization. Please refresh and try again.');
                return { success: false };
            }

            const projectTypeId = typeRows[0].id;

            const projectPayload = {
                project_name: name,
                project_type_id: projectTypeId,
                project_description: desc,
                asset_new_existing: asset,
                start_date: start,
                end_date: end,
                organization_id: organizationId,
                search_embedding: null
            };

            const { data: projectRow, error: projectErr } = await supabase
                .from('projects')
                .insert(projectPayload)
                .select()
                .single();

            if (projectErr) {
                console.error('Error inserting organization project:', projectErr);
                alert('Failed to create project. Please try again.');
                return { success: false };
            }

            return { success: true, project: projectRow };
        } catch (err) {
            console.error('Unexpected error creating organization project:', err);
            alert('An unexpected error occurred while creating the project.');
            return { success: false };
        }
    }

    // Ensure the organization Project Details panel exists in the Create view
    function ensureOrgProjectDetailsPanel() {
        if (orgProjectDetailsPanel) return;

        const createView = document.querySelector('#createView');
        if (!createView) return;

        createView.insertAdjacentHTML('beforeend', `
            <section id="orgProjectDetailsPanel" class="project-types-panel" style="display: none; margin-top: 16px;">
                <div class="project-types-panel-header">
                    <div>
                        <h3>Project Details</h3>
                        <p class="subtitle">Select a project from your organization and upload an Excel file of project parameters.</p>
                    </div>
                    <button type="button" id="orgProjectDetailsBackButton" class="secondary-button">Back to Create Project</button>
                </div>
                <div class="form-group">
                    <label for="orgProjectDetailsProjectSelect">Select Project</label>
                    <select id="orgProjectDetailsProjectSelect">
                        <option value=\"\">Select Project</option>
                    </select>
                </div>
                <div class="form-group" id="orgProjectDetailsUpload">
                    <p>Upload an Excel file with project details. Only .xlsx and .xls formats are accepted.</p>
                    <input
                        type="file"
                        id="orgProjectDetailsFileInput"
                        accept=".xlsx,.xls"
                    >
                </div>
                <div class="form-buttons">
                    <button type="button" id="orgProjectDetailsValidateButton" class="secondary-button">
                        Upload Project Details
                    </button>
                </div>
                <div id="orgProjectDetailsStatus" class="upload-message" aria-live="polite"></div>
            </section>
        `);

        orgProjectDetailsPanel = document.getElementById('orgProjectDetailsPanel');
        orgProjectDetailsProjectSelect = document.getElementById('orgProjectDetailsProjectSelect');
        orgProjectDetailsFileInput = document.getElementById('orgProjectDetailsFileInput');
        orgProjectDetailsValidateButton = document.getElementById('orgProjectDetailsValidateButton');
        orgProjectDetailsStatus = document.getElementById('orgProjectDetailsStatus');
        orgProjectDetailsBackButton = document.getElementById('orgProjectDetailsBackButton');

        if (orgProjectDetailsBackButton) {
            orgProjectDetailsBackButton.addEventListener('click', () => {
                hideOrgProjectDetailsPanel();
            });
        }

        // Wire Excel upload behavior
        if (orgProjectDetailsValidateButton && orgProjectDetailsFileInput && orgProjectDetailsStatus) {
            orgProjectDetailsValidateButton.addEventListener('click', async () => {
                orgProjectDetailsStatus.classList.remove('upload-message--success', 'upload-message--error');

                if (!currentOrgProject || !currentOrgProject.id) {
                    orgProjectDetailsStatus.textContent = 'Please select a project before uploading details.';
                    orgProjectDetailsStatus.classList.add('upload-message--error');
                    return;
                }

                const file = orgProjectDetailsFileInput.files && orgProjectDetailsFileInput.files[0];
                if (!file) {
                    orgProjectDetailsStatus.textContent = 'Please choose a file first.';
                    orgProjectDetailsStatus.classList.add('upload-message--error');
                    return;
                }

                const name = file.name || '';
                const isExcel = /\.xlsx$/i.test(name) || /\.xls$/i.test(name);
                if (!isExcel) {
                    orgProjectDetailsStatus.textContent = 'Only Excel files (.xlsx, .xls) are supported.';
                    orgProjectDetailsStatus.classList.add('upload-message--error');
                    return;
                }

                try {
                    orgProjectDetailsStatus.textContent = `Reading \"${name}\"...`;

                    const arrayBuffer = await file.arrayBuffer();
                    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

                    const firstSheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[firstSheetName];
                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

                    if (!rows || rows.length === 0) {
                        orgProjectDetailsStatus.textContent = 'The Excel file is empty.';
                        orgProjectDetailsStatus.classList.add('upload-message--error');
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
                        orgProjectDetailsStatus.textContent = 'Could not find required headers \"Parameter Name\" and \"Parameter Entry\".';
                        orgProjectDetailsStatus.classList.add('upload-message--error');
                        return;
                    }

                    const projectId = currentOrgProject.id;
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
                            organization_id: organizationId,
                            parameter_name: paramName,
                            parameter_entry: paramEntry
                        });
                    }

                    if (inserts.length === 0) {
                        orgProjectDetailsStatus.textContent = 'No parameter rows found under the required headers.';
                        orgProjectDetailsStatus.classList.add('upload-message--error');
                        return;
                    }

                    const { error } = await supabase
                        .from('project_details')
                        .insert(inserts);

                    if (error) {
                        console.error('Error inserting organization project_details rows:', error);
                        orgProjectDetailsStatus.textContent = 'Failed to save project details to the database.';
                        orgProjectDetailsStatus.classList.add('upload-message--error');
                        return;
                    }

                    orgProjectDetailsStatus.textContent = `Saved ${inserts.length} project detail entries for this project.`;
                    orgProjectDetailsStatus.classList.add('upload-message--success');
                    if (orgProjectDetailsFileInput) {
                        orgProjectDetailsFileInput.value = '';
                    }
                } catch (err) {
                    console.error('Unexpected error processing organization project details Excel:', err);
                    orgProjectDetailsStatus.textContent = 'An unexpected error occurred while processing the Excel file.';
                    orgProjectDetailsStatus.classList.add('upload-message--error');
                }
            });
        }
    }

    function showOrgProjectDetailsPanel() {
        ensureOrgProjectDetailsPanel();
        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = 'none';
        }
        if (orgProjectDetailsPanel) {
            orgProjectDetailsPanel.style.display = '';
            orgProjectDetailsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        // Load projects into the dropdown when panel is shown
        loadOrgProjectsForDropdown();
    }

    function hideOrgProjectDetailsPanel() {
        if (orgProjectDetailsPanel) {
            orgProjectDetailsPanel.style.display = 'none';
        }
        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = '';
        }
    }

    // Attach a simple searchable dropdown of organization project types to a text input.
    // Users can only select from existing types; free-form values will be rejected on submit.
    function attachOrgProjectTypeDropdown(inputEl) {
        if (!inputEl) return;

        let dropdown = null;

        function closeDropdown() {
            if (dropdown && dropdown.parentNode) {
                dropdown.parentNode.removeChild(dropdown);
            }
            dropdown = null;
        }

        function createDropdown() {
            if (dropdown) return dropdown;
            if (!inputEl.parentNode) return null;
            dropdown = document.createElement('div');
            dropdown.className = 'autocomplete-list';
            inputEl.parentNode.appendChild(dropdown);
            return dropdown;
        }

        function renderOptions(options) {
            if (!options || options.length === 0) {
                closeDropdown();
                return;
            }
            const list = createDropdown();
            if (!list) return;
            list.innerHTML = '';

            options.forEach(name => {
                const item = document.createElement('div');
                item.className = 'autocomplete-item';
                item.textContent = name;
                item.addEventListener('mousedown', (e) => {
                    // Use mousedown so it fires before input blur
                    e.preventDefault();
                    inputEl.value = name;
                    closeDropdown();
                });
                list.appendChild(item);
            });
        }

        inputEl.addEventListener('focus', async () => {
            await loadOrgProjectTypes();
            renderOptions(orgProjectTypes);
        });

        inputEl.addEventListener('input', () => {
            const query = inputEl.value.trim().toLowerCase();
            if (!query) {
                renderOptions(orgProjectTypes);
                return;
            }
            const filtered = orgProjectTypes.filter(name =>
                String(name).toLowerCase().includes(query)
            );
            renderOptions(filtered);
        });

        inputEl.addEventListener('blur', () => {
            setTimeout(() => {
                closeDropdown();
            }, 150);
        });
    }

    // Hide Create New Project button if user doesn't have permission
    const createProjectButton = document.querySelector(".create-project-button");
    if (createProjectButton && !canCreateProjects) {
        createProjectButton.style.display = 'none';
    }

    // Hide Organization Settings menu item if user doesn't have permission
    const sidebarNav = document.querySelector('.sidebar .main-nav');
    if (sidebarNav && !canManageOrgSettings) {
        const navItems = sidebarNav.querySelectorAll('.nav-item');
        navItems.forEach(item => {
            const text = item.textContent.trim();
            if (text === 'Organization Settings' || text.includes('Organization Settings')) {
                item.style.display = 'none';
            }
        });
    }

    // Personalize welcome header with first name if session exists
    try {
        const stored = JSON.parse(sessionStorage.getItem('ct_user'));
        const fullName = stored && stored.name ? String(stored.name).trim() : '';
        if (fullName) {
            const firstName = fullName.split(/\s+/)[0] || 'User';
            const headerEl = document.querySelector('#homeView h1');
            if (headerEl) headerEl.textContent = `Welcome ${firstName}`;
        }
    } catch (_) {}

const projectFormHTML = `
    <div class="modal" id="createProjectModal">
        <div class="modal-content collapsible-form">
            <div class="form-toggle" id="createProjectToggle">
                <i class="fas fa-chevron-right" id="createProjectToggleIcon"></i>
            </div>
            <div class="form-content">
            <span class="close-button" id="closeCreateProject">&times;</span>
            <h3>Create New Project</h3>
            <form id="createProjectForm">
                <div class="input-group">
                    <label for="projectName">Project Name</label>
                    <input type="text" id="projectName" required>
                </div>
                <div class="input-group">
                    <label for="projectType">Type</label>
                    <input type="text" id="projectType" required>
                </div>
                <div class="input-group">
                    <label for="assetType">New Asset or Existing</label>
                    <select id="assetType" required>
                        <option value="">Select an option</option>
                        <option value="New">New</option>
                        <option value="Existing">Existing</option>
                        <option value="N/A">N/A</option>
                    </select>
                </div>
                <div class="input-group">
                    <label for="projectDescription">Description</label>
                    <textarea id="projectDescription" required></textarea>
                </div>
                <div class="input-group">
                    <label for="startDate">Start Date</label>
                    <input type="date" id="startDate">
                </div>
                <div class="input-group">
                    <label for="endDate">End Date</label>
                    <input type="date" id="endDate">
                </div>
                <button type="submit" class="save-project-button">Create Project</button>
            </form>
            </div>
        </div>
    </div>`;

    const addDataFormHTML = `
    <div class="modal" id="addDataModal">
        <div class="modal-content collapsible-form">
            <div class="form-toggle" id="addDataToggle">
                <i class="fas fa-chevron-right" id="addDataToggleIcon"></i>
            </div>
            <div class="form-content">
            <span class="close-button" id="closeAddData">&times;</span>
            <h3>Add Data</h3>
            <form id="addDataForm">
                <div class="input-group">
                    <label for="issueSuccess">Issue/Success</label>
                    <input type="text" id="issueSuccess" required>
                </div>
                    <div class="button-group">
                        <button type="button" class="issue-button">Issue</button>
                        <button type="button" class="success-button">Success</button>
                    </div>
                    <div class="input-group" id="causesSection" style="display: none;">
                        <label for="causes">Causes</label>
                        <input type="text" id="causes" placeholder="Enter a cause...">
                        <button type="button" class="add-cause-button">Add</button>
                    </div>
                    <div class="input-group" id="impactsSection" style="display: none;">
                        <label for="impacts">Impacts</label>
                        <input type="text" id="impacts" placeholder="Enter an impact...">
                        <button type="button" class="add-impact-button">Add</button>
                    </div>
                    <div class="input-group" id="actionsSection" style="display: none;">
                        <label for="actions">Actions Taken</label>
                        <input type="text" id="actions" placeholder="Enter an action...">
                        <button type="button" class="add-action-button">Add</button>
                    </div>
                    <div class="input-group" id="lessonsSection" style="display: none;">
                        <label for="lessons">Lessons Learned</label>
                        <input type="text" id="lessons" placeholder="Enter a lesson...">
                        <button type="button" class="add-lesson-button">Add</button>
                    </div>
            </form>
            </div>
        </div>
    </div>`;

    const addUserModalHTML = `
    <div class="modal" id="addUserModal">
        <div class="modal-content" style="width: 500px; max-width: 90%; padding: 30px;">
            <button class="close-button" id="closeAddUser">&times;</button>
            <h3>Add New User</h3>
            <form id="addUserForm">
                <div class="input-group">
                    <label for="userName">Name</label>
                    <input type="text" id="userName" required>
                </div>
                <div class="input-group">
                    <label for="userEmail">Email</label>
                    <input type="email" id="userEmail" required>
                </div>
                <div class="input-group">
                    <label for="userType">User Type</label>
                    <select id="userType" required>
                        <option value="">Select a user type</option>
                        <option value="Administrator">Administrator</option>
                        <option value="Leadership">Leadership</option>
                        <option value="Project Manager">Project Manager</option>
                        <option value="Subject Matter Expert">Subject Matter Expert</option>
                        <option value="Team Member">Team Member</option>
                    </select>
                </div>
                <div class="modal-actions">
                    <button type="button" id="cancelAddUser">Cancel</button>
                    <button type="submit">Add User</button>
                </div>
            </form>
        </div>
    </div>`;

    document.body.insertAdjacentHTML("beforeend", projectFormHTML);
    document.body.insertAdjacentHTML("beforeend", addDataFormHTML);
    document.body.insertAdjacentHTML("beforeend", addUserModalHTML);

    const addDataButton = document.querySelector(".add-data-button");
    const logoutButton = document.querySelector(".logout-button");
    const homeLink = document.getElementById("homeLink");

    const createProjectModal = document.getElementById("createProjectModal");
    const addDataModal = document.getElementById("addDataModal");
    const addUserModal = document.getElementById("addUserModal");

    document.getElementById("closeCreateProject").onclick = () => createProjectModal.classList.remove("show");
    document.getElementById("closeAddData").onclick = () => {
        addDataModal.classList.remove("show");
        resetAddDataForm();
    };
    document.getElementById("closeAddUser").onclick = () => addUserModal.classList.remove("show");
    document.getElementById("cancelAddUser").onclick = () => addUserModal.classList.remove("show");

    // Attach searchable org project-type dropdown to the modal Create Project form
    const modalProjectTypeInput = document.getElementById('projectType');
    attachOrgProjectTypeDropdown(modalProjectTypeInput);

    if (createProjectButton) {
        createProjectButton.onclick = () => {
            // Check permissions before allowing navigation
            if (!canCreateProjects) {
                alert('You do not have permission to create new projects. Please contact your administrator.');
                return;
            }
            if (confirmNavigation("Create New Project")) {
                location.hash = 'create';
            }
        };
    }
    if (homeLink) {
        homeLink.onclick = (e) => {
            e.preventDefault();
            if (!confirmNavigation('home')) return;
            const currentRoute = (location.hash || '#home').replace('#','');
            if (currentRoute !== 'home') {
                location.hash = 'home';
            } else {
                navigate('home');
            }
        };
    }
    logoutButton.onclick = () => {
        if (confirmNavigation("logout")) {
            try { sessionStorage.removeItem('ct_user'); } catch (_) {}
            window.location.href = "index.html";
        }
    };

    window.onclick = (e) => {
    if (e.target === createProjectModal) createProjectModal.classList.remove("show");
    if (e.target === addDataModal) {
        addDataModal.classList.remove("show");
        resetAddDataForm();
    }
    if (e.target === addUserModal) addUserModal.classList.remove("show");
    };

    // Handle Add User form submission -> persist to Supabase 'users'
    document.getElementById("addUserForm").onsubmit = async (e) => {
        e.preventDefault();
        const userName = document.getElementById("userName").value.trim();
        const userEmail = document.getElementById("userEmail").value.trim().toLowerCase();
        let userType = document.getElementById("userType").value;

        // Normalize: ensure Company Administrator is never used from this portal
        if (userType === 'Company Administrator') {
            userType = 'Administrator';
        }

        if (!userName || !userEmail || !userType) {
            alert('Please fill in Name, Email, and User Type.');
            return;
        }

        // Pull organizationid from session
        let organizationid = null;
        try {
            const stored = JSON.parse(sessionStorage.getItem('ct_user'));
            if (stored && stored.organizationid) {
                organizationid = stored.organizationid;
            }
        } catch (_) {}

        // Optionally fetch organization name by id
        let organization = null;
        try {
            if (organizationid != null) {
                const { data: orgRow, error: orgErr } = await supabase
                    .from('organizations')
                    .select('name')
                    .eq('id', organizationid)
                    .maybeSingle();
                if (!orgErr && orgRow && orgRow.name) {
                    organization = orgRow.name;
                }
            }
        } catch (_) {}

        // Insert into public.users with correct column names
        const { error } = await supabase
            .from('users')
            .insert({
                name: userName,
                email: userEmail,
                usertype: userType,
                organization,
                organizationid
            });

        if (error) {
            console.error('Supabase users insert error:', error);
            alert('Failed to add user. Please try again.');
            return;
        }

        alert(`User added successfully!`);
        // Close modal and reset form
        addUserModal.classList.remove("show");
        document.getElementById("addUserForm").reset();
    };

    document.getElementById("createProjectForm").onsubmit = async (e) => {
        e.preventDefault();
        const name = document.getElementById('projectName').value.trim();
        const type = document.getElementById('projectType').value.trim();
        if (!name || !type) {
            alert('Please fill in both Project Name and Type.');
            return;
        }
        if (!isValidOrgProjectType(type)) {
            alert('Please choose a valid Project Type from the list for your organization.');
            return;
        }

        const assetRaw = (document.getElementById('assetType').value || '').trim();
        const descRaw = (document.getElementById('projectDescription').value || '').trim();
        const startRaw = (document.getElementById('startDate').value || '').trim();
        const endRaw = (document.getElementById('endDate').value || '').trim();

        const saveBtn = e.target.querySelector('.save-project-button');
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.textContent = 'Saving...';
        }

        const result = await createOrganizationProject({
            name,
            type,
            assetRaw,
            descRaw,
            startRaw,
            endRaw
        });

        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.textContent = 'Create Project';
        }

        if (!result.success) return;

        alert("Project created successfully!");
        createProjectModal.classList.remove("show");
        e.target.reset();
    };

    // Handle Issue button click
    document.querySelector(".issue-button").onclick = () => {
        const text = document.getElementById("issueSuccess").value.trim();
        if (text) {
            addIssueSuccessEntry("Issue", text);
            // Grey out the field and disable it
            const inputField = document.getElementById("issueSuccess");
            inputField.style.backgroundColor = "#f5f5f5";
            inputField.style.color = "#999";
            inputField.disabled = true;
            // Show all sections
            document.getElementById("causesSection").style.display = "block";
            document.getElementById("impactsSection").style.display = "block";
            document.getElementById("actionsSection").style.display = "block";
            document.getElementById("lessonsSection").style.display = "block";
            // Keep the form open - don't close it
        }
    };

    // Handle Success button click
    document.querySelector(".success-button").onclick = () => {
        const text = document.getElementById("issueSuccess").value.trim();
        if (text) {
            addIssueSuccessEntry("Success", text);
            // Grey out the field and disable it
            const inputField = document.getElementById("issueSuccess");
            inputField.style.backgroundColor = "#f5f5f5";
            inputField.style.color = "#999";
            inputField.disabled = true;
            // Show all sections
            document.getElementById("causesSection").style.display = "block";
            document.getElementById("impactsSection").style.display = "block";
            document.getElementById("actionsSection").style.display = "block";
            document.getElementById("lessonsSection").style.display = "block";
            // Keep the form open - don't close it
        }
    };

    // Function to add Issue/Success entry to display area
    function addIssueSuccessEntry(type, text) {
        const displayArea = getActiveDisplayArea();
        const entry = document.createElement("div");
        entry.className = "issue-success-entry";
        
        // Create the main content with edit/delete buttons
        const mainContent = document.createElement("div");
        mainContent.className = "entry-main-content";
        mainContent.innerHTML = `
            <div class="entry-text">
                <strong>${type}:</strong> ${text}
            </div>
            <div class="entry-actions">
                <button class="edit-entry-button" title="Edit ${type}">
                    <i class="fas fa-edit"></i>
                </button>
                <button class="delete-entry-button" title="Delete ${type}">
                    <i class="fas fa-trash"></i>
                </button>
            </div>
        `;
        
        entry.appendChild(mainContent);
        
        // Add containers for all sub-items
        const causesContainer = document.createElement("div");
        causesContainer.className = "sub-item-container";
        causesContainer.style.display = "none";
        causesContainer.innerHTML = '<div class="sub-item-header"><strong>Causes:</strong></div><ul class="sub-item-list"></ul>';
        
        const impactsContainer = document.createElement("div");
        impactsContainer.className = "sub-item-container";
        impactsContainer.style.display = "none";
        impactsContainer.innerHTML = '<div class="sub-item-header"><strong>Impacts:</strong></div><ul class="sub-item-list"></ul>';
        
        const actionsContainer = document.createElement("div");
        actionsContainer.className = "sub-item-container";
        actionsContainer.style.display = "none";
        actionsContainer.innerHTML = '<div class="sub-item-header"><strong>Actions Taken:</strong></div><ul class="sub-item-list"></ul>';
        
        const lessonsContainer = document.createElement("div");
        lessonsContainer.className = "sub-item-container";
        lessonsContainer.style.display = "none";
        lessonsContainer.innerHTML = '<div class="sub-item-header"><strong>Lessons Learned:</strong></div><ul class="sub-item-list"></ul>';
        
        entry.appendChild(causesContainer);
        entry.appendChild(impactsContainer);
        entry.appendChild(actionsContainer);
        entry.appendChild(lessonsContainer);
        displayArea.appendChild(entry);
        
        // Store references to the lists for this entry
        entry.causesList = causesContainer.querySelector('.sub-item-list');
        entry.impactsList = impactsContainer.querySelector('.sub-item-list');
        entry.actionsList = actionsContainer.querySelector('.sub-item-list');
        entry.lessonsList = lessonsContainer.querySelector('.sub-item-list');
        
        // Add event listener for the delete button
        const deleteButton = mainContent.querySelector('.delete-entry-button');
        deleteButton.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent triggering the entry click
            // Show confirmation dialog
            const confirmed = confirm(`Are you sure you want to delete this ${type.toLowerCase()} entry and all its sub-items?`);
            if (confirmed) {
                // Add fade out animation before removing
                entry.style.transition = 'opacity 0.3s ease, transform 0.3s ease';
                entry.style.opacity = '0';
                entry.style.transform = 'translateX(-20px)';
                
                // Remove the entry after animation completes
                setTimeout(() => {
                    entry.remove();
                }, 300);
            }
        });

        // Add event listener for the edit button
        const editButton = mainContent.querySelector('.edit-entry-button');
        editButton.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent triggering the entry click
            openEditModal(type, text, (newText) => {
                // Update the main text
                const textElement = mainContent.querySelector('.entry-text');
                textElement.innerHTML = `<strong>${type}:</strong> ${newText}`;
            });
        });

        // Add event listener for clicking on the entry text to edit main content
        const textElement = mainContent.querySelector('.entry-text');
        textElement.addEventListener('click', (e) => {
            e.stopPropagation();
            openEditModal(type, text, (newText) => {
                // Update the main text
                textElement.innerHTML = `<strong>${type}:</strong> ${newText}`;
            });
        });
    }

    // Handle Add Cause button click
    document.querySelector(".add-cause-button").onclick = () => {
        const causeText = document.getElementById("causes").value.trim();
        if (causeText) {
            addSubItemEntry('causes', causeText);
            // Reset the causes field
            document.getElementById("causes").value = "";
        }
    };

    // Handle Add Impact button click
    document.querySelector(".add-impact-button").onclick = () => {
        const impactText = document.getElementById("impacts").value.trim();
        if (impactText) {
            addSubItemEntry('impacts', impactText);
            // Reset the impacts field
            document.getElementById("impacts").value = "";
        }
    };

    // Handle Add Action button click
    document.querySelector(".add-action-button").onclick = () => {
        const actionText = document.getElementById("actions").value.trim();
        if (actionText) {
            addSubItemEntry('actions', actionText);
            // Reset the actions field
            document.getElementById("actions").value = "";
        }
    };

    // Handle Add Lesson button click
    document.querySelector(".add-lesson-button").onclick = () => {
        const lessonText = document.getElementById("lessons").value.trim();
        if (lessonText) {
            addSubItemEntry('lessons', lessonText);
            // Reset the lessons field
            document.getElementById("lessons").value = "";
        }
    };

    // Function to add sub-item entry to the most recent Issue/Success
    function addSubItemEntry(listType, text) {
        const displayArea = getActiveDisplayArea();
        const entries = displayArea.querySelectorAll('.issue-success-entry');
        
        if (entries.length > 0) {
            const lastEntry = entries[entries.length - 1];
            const container = lastEntry.querySelector(`.sub-item-container:nth-child(${listType === 'causes' ? '2' : listType === 'impacts' ? '3' : listType === 'actions' ? '4' : '5'})`);
            const list = lastEntry[`${listType}List`];
            
            // Show the container if it's hidden
            container.style.display = "block";
            
            // Add the item as a clickable bullet point
            const item = document.createElement("li");
            item.textContent = text;
            item.style.cursor = "pointer";
            item.style.padding = "4px 8px";
            item.style.borderRadius = "4px";
            item.style.transition = "background-color 0.2s ease";
            item.style.marginBottom = "2px";
            
            // Add hover effect
            item.addEventListener('mouseenter', () => {
                item.style.backgroundColor = "#f0f0f0";
            });
            
            item.addEventListener('mouseleave', () => {
                item.style.backgroundColor = "transparent";
            });
            
            // Add click event to edit the sub-item
            item.addEventListener('click', () => {
                const capitalizedType = listType.charAt(0).toUpperCase() + listType.slice(1);
                openEditModal(capitalizedType, text, (newText) => {
                    item.textContent = newText;
                });
            });
            
            list.appendChild(item);
        }
    }

    // Function to reset the Add Data form
    function resetAddDataForm() {
        const inputField = document.getElementById("issueSuccess");
        inputField.value = "";
        inputField.disabled = false;
        inputField.style.backgroundColor = "";
        inputField.style.color = "";
        // Hide all sections
        document.getElementById("causesSection").style.display = "none";
        document.getElementById("impactsSection").style.display = "none";
        document.getElementById("actionsSection").style.display = "none";
        document.getElementById("lessonsSection").style.display = "none";
        // Reset all fields
        document.getElementById("causes").value = "";
        document.getElementById("impacts").value = "";
        document.getElementById("actions").value = "";
        document.getElementById("lessons").value = "";
    }

    // Collapsible form functionality
    const addDataToggle = document.getElementById("addDataToggle");
    const addDataToggleIcon = document.getElementById("addDataToggleIcon");
    const createProjectToggle = document.getElementById("createProjectToggle");
    const createProjectToggleIcon = document.getElementById("createProjectToggleIcon");

    // Toggle Add Data form
    if (addDataToggle) {
        addDataToggle.addEventListener("click", (e) => {
            e.stopPropagation();
            const modal = document.getElementById("addDataModal");
            modal.classList.toggle("collapsed");
            
            if (modal.classList.contains("collapsed")) {
                addDataToggleIcon.style.transform = "rotate(0deg)";
            } else {
                addDataToggleIcon.style.transform = "rotate(180deg)";
            }
        });
    }

    // Toggle Create Project form
    if (createProjectToggle) {
        createProjectToggle.addEventListener("click", (e) => {
            e.stopPropagation();
            const modal = document.getElementById("createProjectModal");
            modal.classList.toggle("collapsed");
            
            if (modal.classList.contains("collapsed")) {
                createProjectToggleIcon.style.transform = "rotate(0deg)";
            } else {
                createProjectToggleIcon.style.transform = "rotate(180deg)";
            }
        });
    }

    // Show display area when Add Data is clicked
    addDataButton.onclick = () => {
        // If there's already data, ask for confirmation
        if (hasUnsavedData()) {
            const confirmed = confirm("You have existing data in the main area. Do you want to add more data to the current entries?");
            if (!confirmed) {
                return;
            }
        }
        // Navigate to add data then open modal and reveal display area
        location.hash = 'adddata';
        setTimeout(() => {
            addDataModal.classList.add('show');
            const display = getActiveDisplayArea();
            if (display) display.style.display = 'block';
        }, 0);
    };

    function getActiveDisplayArea() {
        // Prefer Add Data view's area if present; fallback to homeView's area
        return document.getElementById('addDataDisplay') || document.getElementById('issueSuccessDisplay');
    }

    // Function to check if there's unsaved data in the main area
    function hasUnsavedData() {
        const displayArea = getActiveDisplayArea();
        if (!displayArea) return false;
        const entries = displayArea.querySelectorAll('.issue-success-entry');
        return entries.length > 0;
    }

    // Function to reset the main area to blank state
    function resetMainArea() {
        const displayArea = getActiveDisplayArea();
        if (displayArea) {
            displayArea.innerHTML = '';
            displayArea.style.display = "none";
        }
        
        // Reset the Add Data form
        resetAddDataForm();
    }

    // Function to show confirmation dialog for unsaved data
    function confirmNavigation(action) {
        if (hasUnsavedData()) {
            const confirmed = confirm("You have unsaved data in the main area. Do you want to proceed without saving?");
            if (confirmed) {
                resetMainArea();
                return true;
            }
            return false;
        }
        return true;
    }

    // Edit Modal functionality
    const editModal = document.getElementById("editModal");
    const editTextarea = document.getElementById("editTextarea");
    const editModalTitle = document.getElementById("editModalTitle");
    const closeEditModal = document.getElementById("closeEditModal");
    const cancelEdit = document.getElementById("cancelEdit");
    const saveEdit = document.getElementById("saveEdit");

    let currentEditCallback = null;

    // Function to open edit modal
    function openEditModal(title, currentText, callback) {
        editModalTitle.textContent = `Edit ${title}`;
        editTextarea.value = currentText;
        editModal.classList.add("show");
        currentEditCallback = callback;
        
        // Focus on textarea and select all text
        setTimeout(() => {
            editTextarea.focus();
            editTextarea.select();
        }, 100);
    }

    // Function to close edit modal
    function closeEditModalFunc() {
        editModal.classList.remove("show");
        currentEditCallback = null;
        editTextarea.value = "";
    }

    // Event listeners for edit modal
    closeEditModal.onclick = closeEditModalFunc;
    cancelEdit.onclick = closeEditModalFunc;
    
    saveEdit.onclick = () => {
        const newText = editTextarea.value.trim();
        if (newText && currentEditCallback) {
            currentEditCallback(newText);
        }
        closeEditModalFunc();
    };

    // Close modal when clicking outside
    editModal.onclick = (e) => {
        if (e.target === editModal) {
            closeEditModalFunc();
        }
    };

    // Close modal with Escape key
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && editModal.classList.contains('show')) {
            closeEditModalFunc();
        }
    });

    // Sidebar navigation (delegation)
    const sidebar = document.querySelector('.sidebar');
    sidebar.addEventListener('click', (e) => {
        const link = e.target.closest('.nav-item');
        if (!link) return;
        e.preventDefault();
        if (!confirmNavigation('navigate')) return;
        const text = link.textContent.trim();
        const route = text === 'Organization Settings' ? 'org'
                    : text === 'My Projects' ? 'projects'
                    : text === 'General Search' || text === 'Search Projects' ? 'search'
                    : 'home';
        location.hash = route;
    });

    // Simple router
    function setActiveSidebar(route) {
        const links = document.querySelectorAll('.sidebar .nav-item');
        links.forEach(l => {
            const text = l.textContent.trim();
            const r = text === 'Organization Settings' ? 'org'
                    : text === 'My Projects' ? 'projects'
                    : text === 'General Search' || text === 'Search Projects' ? 'search'
                    : 'home';
            if (r === route) l.classList.add('active'); else l.classList.remove('active');
            l.setAttribute('aria-current', r === route ? 'page' : 'false');
        });
    }

    function show(elSelector) {
        const el = document.querySelector(elSelector);
        if (el) el.hidden = false;
    }
    function hide(elSelector) {
        const el = document.querySelector(elSelector);
        if (el) el.hidden = true;
    }

    function showHomeView() {
        // Remove org view if present, show home, hide search
        exitOrganizationSettings();
        show('#homeView');
        hide('#searchView');
        hide('#createView');
        hide('#addDataView');
        hide('#projectsView');
        // Ensure header visible and display area hidden by default
        const header = document.querySelector('#homeView h1');
        if (header) header.style.display = '';
    }

    function showSearchView() {
        // Remove org view, hide home, show search
        exitOrganizationSettings();
        hide('#homeView');
        show('#searchView');
        hide('#createView');
        hide('#addDataView');
        hide('#projectsView');
    }

    function showProjectsView() {
        exitOrganizationSettings();
        hide('#homeView');
        hide('#searchView');
        hide('#createView');
        hide('#addDataView');
        show('#projectsView');
    }

    function showCreateView() {
        exitOrganizationSettings();
        hide('#homeView');
        hide('#searchView');
        show('#createView');
        hide('#addDataView');
        hide('#projectsView');
        // When entering Create view, default to showing the Create Project form (hide Project Details panel)
        hideOrgProjectDetailsPanel();
        // Ensure two-column container exists
        const createView = document.querySelector('#createView');
        if (createView && !createView.querySelector('#createColumns')) {
            createView.insertAdjacentHTML('beforeend', `
                <div id="createColumns">
                    <div class="left-col" id="createLeftCol"></div>
                    <div class="right-col" id="createRightCol"></div>
                </div>
            `);
        }
        const leftCol = document.getElementById('createLeftCol');
        const rightCol = document.getElementById('createRightCol');

        // Always show the side buttons in the Create view, whether or not a project has been created yet
        if (rightCol && !rightCol.querySelector('#projectDetailsBtn')) {
            rightCol.insertAdjacentHTML('beforeend', `
                <button id="projectDetailsBtn" class="side-button">Project Details</button>
                <button id="projectTeamBtn" class="side-button">Project Team</button>
                <button id="lessonsMetadataBtn" class="side-button">Lessons Learned Metadata</button>
            `);

            const projectDetailsBtn = document.getElementById('projectDetailsBtn');
            if (projectDetailsBtn && !projectDetailsBtn.dataset.wired) {
                projectDetailsBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    showOrgProjectDetailsPanel();
                });
                projectDetailsBtn.dataset.wired = 'true';
            }
        } else if (rightCol) {
            // Ensure click handler is wired even if buttons already exist
            const projectDetailsBtn = document.getElementById('projectDetailsBtn');
            if (projectDetailsBtn && !projectDetailsBtn.dataset.wired) {
                projectDetailsBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    showOrgProjectDetailsPanel();
                });
                projectDetailsBtn.dataset.wired = 'true';
            }
        }

        if (leftCol && !document.getElementById('inlineCreateProjectForm')) {
            leftCol.insertAdjacentHTML('beforeend', `
                <form id="inlineCreateProjectForm" class="project-form half-width">
                    <div class="form-group">
                        <label for="inlineProjectName">Project Name</label>
                        <input type="text" id="inlineProjectName" required>
                    </div>
                    <div class="form-group">
                        <label for="inlineProjectType">Type</label>
                        <input type="text" id="inlineProjectType" required>
                    </div>
                    <div class="form-group">
                        <label for="inlineAssetType">New Asset or Existing</label>
                        <select id="inlineAssetType">
                            <option value="">Select an option</option>
                            <option value="New">New</option>
                            <option value="Existing">Existing</option>
                            <option value="N/A">N/A</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label for="inlineProjectDescription">Description</label>
                        <textarea id="inlineProjectDescription"></textarea>
                    </div>
                    <div class="form-group">
                        <label for="inlineStartDate">Start Date</label>
                        <input type="date" id="inlineStartDate">
                    </div>
                    <div class="form-group">
                        <label for="inlineEndDate">End Date</label>
                        <input type="date" id="inlineEndDate">
                    </div>
                    <button type="submit" class="save-project-button">Create Project</button>
                </form>
                <div id="createSummary" class="create-summary" style="display:none;"></div>
            `);

            const inlineForm = document.getElementById('inlineCreateProjectForm');

            // Attach searchable org project-type dropdown to the inline Create Project form
            const inlineProjectTypeInput = document.getElementById('inlineProjectType');
            attachOrgProjectTypeDropdown(inlineProjectTypeInput);

            inlineForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                const name = document.getElementById('inlineProjectName').value.trim();
                const type = document.getElementById('inlineProjectType').value.trim();
                if (!name || !type) {
                    alert('Please fill in both Project Name and Type.');
                    return;
                }
                if (!isValidOrgProjectType(type)) {
                    alert('Please choose a valid Project Type from the list for your organization.');
                    return;
                }
                const assetRaw = (document.getElementById('inlineAssetType').value || '').trim();
                const descRaw = (document.getElementById('inlineProjectDescription').value || '').trim();
                const startRaw = (document.getElementById('inlineStartDate').value || '').trim();
                const endRaw = (document.getElementById('inlineEndDate').value || '').trim();

                const saveBtn = inlineForm.querySelector('.save-project-button');
                if (saveBtn) {
                    saveBtn.disabled = true;
                    saveBtn.textContent = 'Saving...';
                }

                const result = await createOrganizationProject({
                    name,
                    type,
                    assetRaw,
                    descRaw,
                    startRaw,
                    endRaw
                });

                if (saveBtn) {
                    saveBtn.disabled = false;
                    saveBtn.textContent = 'Create Project';
                }

                if (!result.success) return;

                const asset = assetRaw || 'N/A';
                const desc = descRaw || 'N/A';
                const start = startRaw || 'N/A';
                const end = endRaw || 'N/A';

                const summary = document.getElementById('createSummary');
                summary.innerHTML = `
                    <div><strong>Project Name:</strong> ${name}</div>
                    <div><strong>Type:</strong> ${type}</div>
                    <div><strong>New Asset or Existing:</strong> ${asset}</div>
                    <div><strong>Description:</strong> ${desc}</div>
                    <div><strong>Start Date:</strong> ${start}</div>
                    <div><strong>End Date:</strong> ${end}</div>
                `;
                summary.style.display = '';
            });
        }
    }

    function showAddDataView() {
        exitOrganizationSettings();
        hide('#homeView');
        hide('#searchView');
        hide('#createView');
        show('#addDataView');
        hide('#projectsView');
    }

    // Function to display Organization Settings
    function displayOrganizationSettings() {
        // Hide other views
        hide('#homeView');
        hide('#searchView');
        hide('#createView');
        hide('#addDataView');
        hide('#projectsView');
        // If already present, don't duplicate
        if (!document.getElementById('orgView')) {
            const tpl = document.getElementById('orgSettingsTemplate');
            if (tpl) {
                document.querySelector('.main-content').appendChild(tpl.content.cloneNode(true));

                // Wire up "Add New User" button to open the Add User modal
                const addBtn = document.getElementById('addNewUserButton');
                if (addBtn) addBtn.onclick = () => addUserModal.classList.add('show');

                // Manage Projects module within Organization Settings
                const manageProjectsBtn = document.getElementById('manageProjectsButton');
                const orgMainSummary = document.getElementById('orgMainSummary');
                const orgProjectsPanel = document.getElementById('orgProjectsPanel');
                const orgProjectsBackButton = document.getElementById('orgProjectsBackButton');
                const orgProjectTypesForm = document.getElementById('orgProjectTypesForm');
                const orgProjectTypeInput = document.getElementById('orgProjectTypeName');
                const orgProjectsSubtitle = document.getElementById('orgProjectsSubtitle');
                const orgManageProjectsSubmodule = document.getElementById('orgManageProjectsSubmodule');
                const orgProjectsSelect = document.getElementById('orgProjectsSelect');
                const orgProjectDetailsSelect = document.getElementById('orgProjectDetailsSelect');

                // Look up organization id once for this view
                let organizationId = null;
                try {
                    const stored = JSON.parse(sessionStorage.getItem('ct_user'));
                    if (stored && stored.organizationid) {
                        organizationId = stored.organizationid;
                    }
                } catch (_) {}

                // Simple in-place autocomplete dropdown for org project types
                let orgTypeDropdown = null;
                let orgTypeOptions = [];
                let orgProjectsLoaded = false;

                function closeOrgTypeDropdown() {
                    if (orgTypeDropdown && orgTypeDropdown.parentNode) {
                        orgTypeDropdown.parentNode.removeChild(orgTypeDropdown);
                    }
                    orgTypeDropdown = null;
                }

                function createOrgTypeDropdown() {
                    if (orgTypeDropdown) return orgTypeDropdown;
                    if (!orgProjectTypeInput || !orgProjectTypeInput.parentNode) return null;
                    orgTypeDropdown = document.createElement('div');
                    orgTypeDropdown.className = 'autocomplete-list';
                    orgProjectTypeInput.parentNode.appendChild(orgTypeDropdown);
                    return orgTypeDropdown;
                }

                function renderOrgTypeDropdown(options) {
                    if (!options || options.length === 0) {
                        closeOrgTypeDropdown();
                        return;
                    }
                    const dropdown = createOrgTypeDropdown();
                    if (!dropdown) return;
                    dropdown.innerHTML = '';

                    options.forEach(value => {
                        const item = document.createElement('div');
                        item.className = 'autocomplete-item';
                        item.textContent = value;
                        item.addEventListener('mousedown', (e) => {
                            // Use mousedown so it fires before input blur
                            e.preventDefault();
                            if (orgProjectTypeInput) {
                                orgProjectTypeInput.value = value;
                            }
                            closeOrgTypeDropdown();
                        });
                        dropdown.appendChild(item);
                    });
                }

                async function loadOrgProjectTypes() {
                    // Only load once and only if we know the organization id
                    if (!organizationId || orgTypeOptions.length > 0) return;

                    try {
                        const { data, error } = await supabase
                            .from('project_type')
                            .select('project_type')
                            .eq('organization_id', organizationId)
                            .order('project_type', { ascending: true });

                        if (error) {
                            console.error('Error loading organization project types:', error);
                            return;
                        }

                        const names = (data || [])
                            .map(row => row && row.project_type)
                            .filter(Boolean);

                        // Remove duplicates and store locally
                        orgTypeOptions = Array.from(new Set(names));
                    } catch (err) {
                        console.error('Unexpected error loading organization project types:', err);
                    }
                }

                // Load organization projects into the Manage Projects sub-module dropdown
                async function loadOrgProjectsForManagePanel() {
                    if (!organizationId || !orgProjectsSelect) return;

                    // Basic guard against reloading on every click; we still handle empty state.
                    if (orgProjectsLoaded && orgProjectsSelect.options.length > 1) return;

                    orgProjectsSelect.innerHTML = '<option value=\"\">Select Project</option>';

                    try {
                        const { data, error } = await supabase
                            .from('projects')
                            .select('project_id, project_name')
                            .eq('organization_id', organizationId)
                            .order('project_name', { ascending: true });

                        if (error) {
                            console.error('Error loading organization projects for Manage Projects dropdown:', error);
                            const opt = document.createElement('option');
                            opt.value = '';
                            opt.textContent = 'Failed to load projects';
                            opt.disabled = true;
                            orgProjectsSelect.appendChild(opt);
                            return;
                        }

                        const rows = data || [];
                        if (rows.length === 0) {
                            const opt = document.createElement('option');
                            opt.value = '';
                            opt.textContent = 'No projects found for your organization';
                            opt.disabled = true;
                            orgProjectsSelect.appendChild(opt);
                            return;
                        }

                        rows.forEach(row => {
                            const idStr = String(row.project_id);
                            const name = row.project_name || `Project ${idStr}`;
                            const opt = document.createElement('option');
                            opt.value = idStr;
                            opt.textContent = name;
                            orgProjectsSelect.appendChild(opt);
                        });

                        orgProjectsLoaded = true;
                    } catch (err) {
                        console.error('Unexpected error loading organization projects for Manage Projects dropdown:', err);
                        const opt = document.createElement('option');
                        opt.value = '';
                        opt.textContent = 'Unexpected error loading projects';
                        opt.disabled = true;
                        orgProjectsSelect.appendChild(opt);
                    }
                }

                // Load project details for selected project into the Project Details dropdown
                async function loadOrgProjectDetailsForManagePanel(projectId) {
                    if (!organizationId || !orgProjectDetailsSelect) return;

                    orgProjectDetailsSelect.innerHTML = '<option value=\"\">Loading project details...</option>';
                    orgProjectDetailsSelect.disabled = true;

                    if (!projectId) {
                        orgProjectDetailsSelect.innerHTML = '<option value=\"\">Select a project to see its details</option>';
                        return;
                    }

                    try {
                        const { data, error } = await supabase
                            .from('project_details')
                            .select('id, parameter_name, parameter_entry')
                            .eq('organization_id', organizationId)
                            .eq('project_id', projectId)
                            .order('parameter_name', { ascending: true });

                        if (error) {
                            console.error('Error loading project_details for Manage Projects dropdown:', error);
                            orgProjectDetailsSelect.innerHTML = '<option value=\"\">Failed to load project details</option>';
                            return;
                        }

                        const rows = data || [];
                        if (rows.length === 0) {
                            orgProjectDetailsSelect.innerHTML = '<option value=\"\">No project details found</option>';
                            return;
                        }

                        orgProjectDetailsSelect.innerHTML = '<option value=\"\">Select Project Detail</option>';
                        rows.forEach(row => {
                            const name = row.parameter_name != null ? String(row.parameter_name).trim() : '';
                            const entry = row.parameter_entry != null ? String(row.parameter_entry).trim() : '';
                            const parts = [];
                            if (name) parts.push(name);
                            if (entry) parts.push(entry);
                            const label = parts.length ? parts.join(': ') : '(no name / entry)';

                            const opt = document.createElement('option');
                            opt.value = String(row.id);
                            opt.textContent = label;
                            orgProjectDetailsSelect.appendChild(opt);
                        });

                        orgProjectDetailsSelect.disabled = false;
                    } catch (err) {
                        console.error('Unexpected error loading project_details for Manage Projects dropdown:', err);
                        orgProjectDetailsSelect.innerHTML = '<option value=\"\">Unexpected error loading project details</option>';
                    }
                }

                // Attach autocomplete behaviour to the Project Type input
                if (orgProjectTypeInput) {
                    orgProjectTypeInput.addEventListener('focus', async () => {
                        await loadOrgProjectTypes();
                        renderOrgTypeDropdown(orgTypeOptions);
                    });

                    orgProjectTypeInput.addEventListener('input', () => {
                        const query = orgProjectTypeInput.value.trim().toLowerCase();
                        if (!query) {
                            renderOrgTypeDropdown(orgTypeOptions);
                            return;
                        }
                        const filtered = orgTypeOptions.filter(name =>
                            String(name).toLowerCase().includes(query)
                        );
                        renderOrgTypeDropdown(filtered);
                    });

                    orgProjectTypeInput.addEventListener('blur', () => {
                        // Slight delay so mousedown on dropdown items still works
                        setTimeout(() => {
                            closeOrgTypeDropdown();
                        }, 150);
                    });
                }

                // Show the Manage Projects & Project Types module
                if (manageProjectsBtn && orgMainSummary && orgProjectsPanel) {
                    manageProjectsBtn.onclick = () => {
                        orgMainSummary.style.display = 'none';
                        orgProjectsPanel.style.display = '';
                    };
                }

                // "Manage Projects" button inside the Manage Projects panel
                const orgManageProjectsButton = document.getElementById('orgManageProjectsButton');
                if (orgManageProjectsButton && orgProjectTypesForm && orgManageProjectsSubmodule) {
                    orgManageProjectsButton.addEventListener('click', async (e) => {
                        e.preventDefault();

                        // Hide project-type specific UI
                        orgProjectTypesForm.style.display = 'none';
                        if (orgProjectsSubtitle) {
                            orgProjectsSubtitle.style.display = 'none';
                        }

                        // Show Manage Projects sub-module
                        orgManageProjectsSubmodule.style.display = '';

                        // Reset dropdowns to a clean state
                        if (orgProjectsSelect) {
                            orgProjectsSelect.innerHTML = '<option value=\"\">Select Project</option>';
                        }
                        if (orgProjectDetailsSelect) {
                            orgProjectDetailsSelect.innerHTML = '<option value=\"\">Select a project to see its details</option>';
                            orgProjectDetailsSelect.disabled = true;
                        }

                        // Load projects for this organization
                        await loadOrgProjectsForManagePanel();
                    });
                }

                // When a project is chosen, load its project details
                if (orgProjectsSelect && orgProjectDetailsSelect) {
                    orgProjectsSelect.addEventListener('change', (e) => {
                        const projectIdRaw = e.target.value;
                        if (!projectIdRaw) {
                            orgProjectDetailsSelect.innerHTML = '<option value=\"\">Select a project to see its details</option>';
                            orgProjectDetailsSelect.disabled = true;
                            return;
                        }
                        const numericId = Number.isNaN(Number(projectIdRaw)) ? projectIdRaw : Number(projectIdRaw);
                        loadOrgProjectDetailsForManagePanel(numericId);
                    });
                }

                // Go Back button returns to the main Organization Settings summary
                if (orgProjectsBackButton && orgMainSummary && orgProjectsPanel) {
                    orgProjectsBackButton.onclick = () => {
                        orgProjectsPanel.style.display = 'none';
                        orgMainSummary.style.display = '';

                        // Restore original Manage Projects & Project Types view state
                        if (orgProjectTypesForm) {
                            orgProjectTypesForm.style.display = '';
                        }
                        if (orgManageProjectsSubmodule) {
                            orgManageProjectsSubmodule.style.display = 'none';
                        }
                        if (orgProjectsSubtitle) {
                            orgProjectsSubtitle.style.display = '';
                        }
                        if (orgProjectsSelect) {
                            orgProjectsSelect.selectedIndex = 0;
                        }
                        if (orgProjectDetailsSelect) {
                            orgProjectDetailsSelect.innerHTML = '<option value=\"\">Select a project to see its details</option>';
                            orgProjectDetailsSelect.disabled = true;
                        }
                    };
                }

                // Handle "Create Project Type" form submission for this organization
                if (orgProjectTypesForm) {
                    orgProjectTypesForm.addEventListener('submit', async (e) => {
                        e.preventDefault();

                        const typeInput = document.getElementById('orgProjectTypeName');
                        const type = typeInput ? typeInput.value.trim() : '';

                        if (!type) {
                            alert('Please enter a Project Type.');
                            return;
                        }

                        if (!organizationId) {
                            alert('Could not determine your organization. Please log out and log back in, or contact your administrator.');
                            return;
                        }

                        const submitBtn = orgProjectTypesForm.querySelector('button[type="submit"]');
                        if (submitBtn) {
                            submitBtn.disabled = true;
                            submitBtn.textContent = 'Saving...';
                        }

                        try {
                            const payload = {
                                project_type: type,
                                industry: null,          // No Industry field in this module
                                organization_id: organizationId,
                                is_public: false         // Organization-specific project type
                            };

                            const { error } = await supabase
                                .from('project_type')
                                .insert(payload);

                            if (error) {
                                console.error('Supabase project_type insert error (org portal):', error);
                                alert('Failed to save project type. Please try again.');
                                return;
                            }

                            alert(`Project Type "${type}" saved successfully for your organization.`);
                            orgProjectTypesForm.reset();
                        } catch (err) {
                            console.error('Unexpected error saving project type (org portal):', err);
                            alert('An unexpected error occurred while saving the project type.');
                        } finally {
                            if (submitBtn) {
                                submitBtn.disabled = false;
                                submitBtn.textContent = 'Save Project Type';
                            }
                        }
                    });
                }
            }
        }
    }

    // Function to exit Organization Settings and restore default main
    function exitOrganizationSettings() {
        const org = document.getElementById('orgView');
        if (org) org.remove();
        // Restore home view visibility if needed
        show('#homeView');
    }

    function navigate(route) {
        setActiveSidebar(route);
        if (route === 'org') {
            // Check permissions before allowing access to Organization Settings
            if (!canManageOrgSettings) {
                alert('You do not have permission to access Organization Settings. Please contact your administrator.');
                location.hash = 'home';
                showHomeView();
                return;
            }
            resetMainArea();
            displayOrganizationSettings();
        } else if (route === 'create') {
            // Check permissions before allowing access to Create Project
            if (!canCreateProjects) {
                alert('You do not have permission to create new projects. Please contact your administrator.');
                location.hash = 'home';
                showHomeView();
                return;
            }
            resetMainArea();
            showCreateView();
        } else if (route === 'search') {
            resetMainArea();
            showSearchView();
        } else if (route === 'projects') {
            resetMainArea();
            showProjectsView();
        } else if (route === 'adddata') {
            resetMainArea();
            showAddDataView();
        } else {
            resetMainArea();
            showHomeView();
        }
    }

    window.addEventListener('hashchange', () => {
        const route = (location.hash || '#home').replace('#','');
        navigate(route);
    });

    // initial route
    const initialRoute = (location.hash || '#home').replace('#','');
    navigate(initialRoute);
});
