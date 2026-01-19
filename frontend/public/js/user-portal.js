// js/user-portal.js
import { supabase } from './supabase-client.js';
import * as XLSX from 'https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs';
import { parseMsProjectTxt, importMsProjectTxtToSupabase } from './ms-project-txt-parser.js';
import { importProjectTeamListExcelToSupabase } from './excel-project-team-importer.js';
import { importProjectMiscListExcelToSupabase } from './excel-project-misc-importer.js';
import { importAssetGeneralExcelToSupabase } from './excel-asset-general-importer.js';

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
    let orgProjectDetailsManageButton = null;
    let orgProjectDetailsChoices = null;
    let currentOrgProject = null;
    let orgProjectsById = new Map();

    // State for Project Team sub-module (Create view)
    let projectTeamPanel = null;
    let projectTeamProjectSelect = null;
    let projectTeamUserSelect = null;
    let projectTeamBackButton = null;
    let projectTeamAddButton = null;
    let projectTeamAssignmentsButton = null;
    let projectTeamStatus = null;
    let projectTeamChoices = null;
    let projectTeamUsersById = new Map(); // user_id -> { name, email, usertype }
    let projectTeamProjectChoices = null;
    let projectTeamProjectsById = new Map(); // project_id -> { name }

    // State for Project Team Assignments panel (Create view)
    let projectTeamAssignmentsPanel = null;
    let projectTeamAssignmentsBackButton = null;
    let projectTeamAssignmentsRefreshButton = null;
    let projectTeamAssignmentsSaveButton = null;
    let projectTeamAssignmentsTypeSelect = null;
    let projectTeamAssignmentsSearchInput = null;
    let projectTeamAssignmentsStatus = null;
    let projectTeamAssignmentsTableWrap = null;
    let projectTeamAssignmentsRowsCache = [];
    let projectTeamAssignmentsSelectedIds = new Set();
    let projectTeamAssignmentsSavedIds = new Set();
    let projectTeamAssignmentsUserId = null;
    let projectTeamAssignmentsProjectId = null;

    // State for Manage Project Details panel (Create view)
    let orgProjectDetailsManagePanel = null;
    let orgProjectDetailsManageProjectSelect = null;
    let orgProjectDetailsManageBackButton = null;
    let orgProjectDetailsManageRefreshButton = null;
    let orgProjectDetailsManageSearchNameInput = null;
    let orgProjectDetailsManageSearchEntryInput = null;
    let orgProjectDetailsManageStatus = null;
    let orgProjectDetailsManageTableWrap = null;
    let orgProjectDetailsManageChoices = null;
    let orgProjectDetailsManageRowsCache = [];

    // State for Lessons Learned Metadata sub-module (Create view)
    let lessonsMetadataPanel = null;
    let lessonsMetadataProjectSelect = null;
    let lessonsMetadataStatus = null;
    let lessonsMetadataBackButton = null;
    let lessonsMetadataManageButton = null;
    let lessonsMetadataChoices = null;
    let lessonsMetadataFileTypeGroup = null;
    let lessonsMetadataFileTypeSelect = null;
    let lessonsMetadataUploadGroup = null;
    let lessonsMetadataFileInput = null;
    let lessonsMetadataUploadButton = null;
    let lessonsMetadataUpdateActionsGroup = null;
    let lessonsMetadataDeleteExistingXmlButton = null;
    let lessonsMetadataDeleteExistingTeamListButton = null;
    let lessonsProjectsById = new Map(); // project_id -> { name, project_type_id }

    // State for Manage Lessons Learned Metadata panel (Create view)
    let lessonsMetadataManagePanel = null;
    let lessonsMetadataManageProjectSelect = null;
    let lessonsMetadataManageBackButton = null;
    let lessonsMetadataManageRefreshButton = null;
    let lessonsMetadataManageTypeSelect = null;
    let lessonsMetadataManageSearchInput = null;
    let lessonsMetadataManageStatus = null;
    let lessonsMetadataManageTableWrap = null;
    let lessonsMetadataManageChoices = null;
    let lessonsMetadataManageRowsCache = [];

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
                .select('project_id, project_name, project_type_id')
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

    // Load organization projects for the Lessons Learned Metadata dropdown in Create view
    async function loadOrgProjectsForLessonsDropdown() {
        if (!organizationId || !lessonsMetadataProjectSelect || !lessonsMetadataStatus) return;

        lessonsMetadataStatus.classList.remove('upload-message--success', 'upload-message--error');
        lessonsMetadataStatus.textContent = 'Loading projects...';

        try {
            const { data: projectRows, error: projectErr } = await supabase
                .from('projects')
                .select('project_id, project_name')
                .eq('organization_id', organizationId)
                .order('project_name', { ascending: true });

            if (projectErr) {
                console.error('Error loading organization projects for lessons metadata dropdown:', projectErr);
                lessonsMetadataStatus.textContent = 'Failed to load projects.';
                lessonsMetadataStatus.classList.add('upload-message--error');
                return;
            }

            const rows = projectRows || [];
            if (rows.length === 0) {
                lessonsMetadataStatus.textContent = 'No projects found for your organization.';
                lessonsMetadataStatus.classList.add('upload-message--error');
                if (lessonsMetadataChoices) {
                    lessonsMetadataChoices.clearChoices();
                } else {
                    lessonsMetadataProjectSelect.innerHTML = '<option value=\"\">No projects available</option>';
                }
                lessonsProjectsById = new Map();
                return;
            }

            lessonsProjectsById = new Map();
            const choicesData = rows.map(row => {
                const idStr = String(row.project_id);
                const name = row.project_name || `Project ${idStr}`;
                const typeId = row.project_type_id != null ? String(row.project_type_id) : null;
                lessonsProjectsById.set(idStr, { name, project_type_id: typeId });
                return {
                    value: idStr,
                    label: name
                };
            });

            // Initialize Choices once
            if (!lessonsMetadataChoices) {
                if (typeof Choices === 'undefined') {
                    console.warn('Choices library not loaded; falling back to native select for lessons metadata dropdown.');
                    lessonsMetadataProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                    choicesData.forEach(choice => {
                        const option = document.createElement('option');
                        option.value = choice.value;
                        option.textContent = choice.label;
                        lessonsMetadataProjectSelect.appendChild(option);
                    });
                } else {
                    lessonsMetadataChoices = new Choices(lessonsMetadataProjectSelect, {
                        searchEnabled: true,
                        shouldSort: false,
                        placeholder: true,
                        placeholderValue: 'Select Project',
                        searchPlaceholderValue: 'Type to search...'
                    });
                }
            }

            if (lessonsMetadataChoices) {
                lessonsMetadataChoices.setChoices(choicesData, 'value', 'label', true);
            } else {
                lessonsMetadataProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                choicesData.forEach(choice => {
                    const option = document.createElement('option');
                    option.value = choice.value;
                    option.textContent = choice.label;
                    lessonsMetadataProjectSelect.appendChild(option);
                });
            }

            lessonsMetadataStatus.textContent = '';
            lessonsMetadataStatus.classList.remove('upload-message--error');
        } catch (err) {
            console.error('Unexpected error loading organization projects for lessons metadata dropdown:', err);
            lessonsMetadataStatus.textContent = 'An unexpected error occurred while loading projects.';
            lessonsMetadataStatus.classList.add('upload-message--error');
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
                    <div class="project-types-panel-header-actions">
                        <button type="button" id="orgProjectDetailsBackButton" class="secondary-button">Back to Create Project</button>
                        <button type="button" id="orgProjectDetailsManageButton" class="secondary-button">Manage Project Details</button>
                    </div>
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
        orgProjectDetailsManageButton = document.getElementById('orgProjectDetailsManageButton');

        if (orgProjectDetailsBackButton) {
            orgProjectDetailsBackButton.addEventListener('click', () => {
                hideOrgProjectDetailsPanel();
            });
        }

        if (orgProjectDetailsManageButton) {
            orgProjectDetailsManageButton.addEventListener('click', (e) => {
                e.preventDefault();
                showOrgProjectDetailsManagePanel();
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

    // Ensure the Lessons Learned Metadata panel exists in the Create view
    function ensureLessonsMetadataPanel() {
        if (lessonsMetadataPanel) return;

        const createView = document.querySelector('#createView');
        if (!createView) return;

        createView.insertAdjacentHTML('beforeend', `
            <section id="lessonsMetadataPanel" class="project-types-panel" style="display: none; margin-top: 16px;">
                <div class="project-types-panel-header">
                    <div>
                        <h3>Lessons Learned Metadata</h3>
                        <p class="subtitle">Select a project from your organization to work with its lessons learned metadata.</p>
                    </div>
                    <div class="project-types-panel-header-actions">
                        <button type="button" id="lessonsMetadataBackButton" class="secondary-button">Back to Create Project</button>
                        <button type="button" id="lessonsMetadataManageButton" class="secondary-button">Manage Lessons Learned Metadata</button>
                    </div>
                </div>
                <div class="form-group">
                    <label for="lessonsMetadataProjectSelect">Project</label>
                    <select id="lessonsMetadataProjectSelect">
                        <option value=\"\">Select Project</option>
                    </select>
                </div>
                <div class="form-group" id="lessonsMetadataFileTypeGroup" style="display: none;">
                    <label for="lessonsMetadataFileTypeSelect">Select Input Data File Type</label>
                    <select id="lessonsMetadataFileTypeSelect">
                        <option value="">Select File Type</option>
                        <option value="ms_project_xml">MS Project XML - New</option>
                        <option value="ms_project_xml_update">MS Project XML - Update</option>
                        <option value="excel_project_team_list_new">Excel Project Team List - New</option>
                        <option value="excel_project_team_list_update">Excel Project Team List - Update</option>
                        <option value="excel_project_misc_list_new">Excel Project Miscellaneous List - New</option>
                    </select>
                </div>
                <div class="form-group" id="lessonsMetadataUploadGroup" style="display: none;">
                    <label for="lessonsMetadataFileInput">Upload TXT File</label>
                    <input type="file" id="lessonsMetadataFileInput" accept=".txt">
                    <div class="form-buttons">
                        <button type="button" id="lessonsMetadataUploadButton" class="secondary-button">
                            Upload File
                        </button>
                    </div>
                </div>
                <div class="form-group" id="lessonsMetadataUpdateActionsGroup" style="display: none;">
                    <div class="form-buttons">
                        <button type="button" id="lessonsMetadataDeleteExistingXmlButton" class="secondary-button">
                            Delete Existing XML
                        </button>
                        <button type="button" id="lessonsMetadataDeleteExistingTeamListButton" class="secondary-button" style="display: none;">
                            Delete Existing Team List
                        </button>
                    </div>
                </div>
                <div id="lessonsMetadataStatus" class="upload-message" aria-live="polite"></div>
            </section>
        `);

        lessonsMetadataPanel = document.getElementById('lessonsMetadataPanel');
        lessonsMetadataProjectSelect = document.getElementById('lessonsMetadataProjectSelect');
        lessonsMetadataStatus = document.getElementById('lessonsMetadataStatus');
        lessonsMetadataBackButton = document.getElementById('lessonsMetadataBackButton');
        lessonsMetadataManageButton = document.getElementById('lessonsMetadataManageButton');
        lessonsMetadataFileTypeGroup = document.getElementById('lessonsMetadataFileTypeGroup');
        lessonsMetadataFileTypeSelect = document.getElementById('lessonsMetadataFileTypeSelect');
        lessonsMetadataUploadGroup = document.getElementById('lessonsMetadataUploadGroup');
        lessonsMetadataFileInput = document.getElementById('lessonsMetadataFileInput');
        lessonsMetadataUploadButton = document.getElementById('lessonsMetadataUploadButton');
        lessonsMetadataUpdateActionsGroup = document.getElementById('lessonsMetadataUpdateActionsGroup');
        lessonsMetadataDeleteExistingXmlButton = document.getElementById('lessonsMetadataDeleteExistingXmlButton');
        lessonsMetadataDeleteExistingTeamListButton = document.getElementById('lessonsMetadataDeleteExistingTeamListButton');

        if (lessonsMetadataBackButton) {
            lessonsMetadataBackButton.addEventListener('click', () => {
                hideLessonsMetadataPanel();
            });
        }

        if (lessonsMetadataManageButton) {
            lessonsMetadataManageButton.addEventListener('click', (e) => {
                e.preventDefault();
                showLessonsMetadataManagePanel();
            });
        }

        // Show/hide file type dropdown based on project selection
        if (lessonsMetadataProjectSelect && lessonsMetadataFileTypeGroup) {
            lessonsMetadataProjectSelect.addEventListener('change', () => {
                const value = lessonsMetadataProjectSelect.value;
                if (value) {
                    lessonsMetadataFileTypeGroup.style.display = '';
                    // Reset any previous file-type or upload state when a new project is chosen
                    if (lessonsMetadataFileTypeSelect) {
                        lessonsMetadataFileTypeSelect.value = '';
                    }
                    if (lessonsMetadataUploadGroup) {
                        lessonsMetadataUploadGroup.style.display = 'none';
                    }
                    if (lessonsMetadataFileInput) {
                        lessonsMetadataFileInput.value = '';
                    }
                    if (lessonsMetadataStatus) {
                        lessonsMetadataStatus.textContent = '';
                        lessonsMetadataStatus.classList.remove('upload-message--success', 'upload-message--error');
                    }
                } else {
                    lessonsMetadataFileTypeGroup.style.display = 'none';
                    if (lessonsMetadataUploadGroup) {
                        lessonsMetadataUploadGroup.style.display = 'none';
                    }
                    if (lessonsMetadataFileTypeSelect) {
                        lessonsMetadataFileTypeSelect.value = '';
                    }
                    if (lessonsMetadataFileInput) {
                        lessonsMetadataFileInput.value = '';
                    }
                    if (lessonsMetadataStatus) {
                        lessonsMetadataStatus.textContent = '';
                        lessonsMetadataStatus.classList.remove('upload-message--success', 'upload-message--error');
                    }
                }
            });
        }

        // Show/hide upload controls based on selected file type
        if (lessonsMetadataFileTypeSelect && lessonsMetadataUploadGroup) {
            lessonsMetadataFileTypeSelect.addEventListener('change', () => {
                const value = lessonsMetadataFileTypeSelect.value;
                const fileLabel = lessonsMetadataUploadGroup
                    ? lessonsMetadataUploadGroup.querySelector('label[for="lessonsMetadataFileInput"]')
                    : null;

                const resetStatus = () => {
                    if (!lessonsMetadataStatus) return;
                    lessonsMetadataStatus.textContent = '';
                    lessonsMetadataStatus.classList.remove('upload-message--error', 'upload-message--success');
                };

                const setHint = (msg) => {
                    if (!lessonsMetadataStatus) return;
                    lessonsMetadataStatus.textContent = msg;
                    lessonsMetadataStatus.classList.remove('upload-message--error', 'upload-message--success');
                };

                if (value === 'ms_project_xml') {
                    lessonsMetadataUploadGroup.style.display = '';
                    if (lessonsMetadataUpdateActionsGroup) lessonsMetadataUpdateActionsGroup.style.display = 'none';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.accept = '.txt';
                    if (fileLabel) fileLabel.textContent = 'Upload TXT File';
                    if (lessonsMetadataDeleteExistingXmlButton) lessonsMetadataDeleteExistingXmlButton.style.display = '';
                    if (lessonsMetadataDeleteExistingTeamListButton) lessonsMetadataDeleteExistingTeamListButton.style.display = 'none';
                    setHint('Please upload a .txt file exported from MS Project.');
                } else if (value === 'excel_project_team_list_new') {
                    lessonsMetadataUploadGroup.style.display = '';
                    if (lessonsMetadataUpdateActionsGroup) lessonsMetadataUpdateActionsGroup.style.display = 'none';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.accept = '.xlsx,.xls';
                    if (fileLabel) fileLabel.textContent = 'Upload Excel File';
                    if (lessonsMetadataDeleteExistingXmlButton) lessonsMetadataDeleteExistingXmlButton.style.display = '';
                    if (lessonsMetadataDeleteExistingTeamListButton) lessonsMetadataDeleteExistingTeamListButton.style.display = 'none';
                    setHint('Please upload an Excel file with headers: Team, Description, Team ID, Parent Team ID.');
                } else if (value === 'excel_project_misc_list_new') {
                    lessonsMetadataUploadGroup.style.display = '';
                    if (lessonsMetadataUpdateActionsGroup) lessonsMetadataUpdateActionsGroup.style.display = 'none';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.accept = '.xlsx,.xls';
                    if (fileLabel) fileLabel.textContent = 'Upload Excel File';
                    if (lessonsMetadataDeleteExistingXmlButton) lessonsMetadataDeleteExistingXmlButton.style.display = '';
                    if (lessonsMetadataDeleteExistingTeamListButton) lessonsMetadataDeleteExistingTeamListButton.style.display = 'none';
                    setHint('Please upload an Excel file with headers: Item ID, Item, Description, Parent Item ID.');
                } else if (value === 'ms_project_xml_update') {
                    lessonsMetadataUploadGroup.style.display = 'none';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    if (lessonsMetadataUpdateActionsGroup) lessonsMetadataUpdateActionsGroup.style.display = '';
                    if (lessonsMetadataDeleteExistingXmlButton) lessonsMetadataDeleteExistingXmlButton.style.display = '';
                    if (lessonsMetadataDeleteExistingTeamListButton) lessonsMetadataDeleteExistingTeamListButton.style.display = 'none';
                    resetStatus();
                } else if (value === 'excel_project_team_list_update') {
                    lessonsMetadataUploadGroup.style.display = 'none';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    if (lessonsMetadataUpdateActionsGroup) lessonsMetadataUpdateActionsGroup.style.display = '';
                    if (lessonsMetadataDeleteExistingXmlButton) lessonsMetadataDeleteExistingXmlButton.style.display = 'none';
                    if (lessonsMetadataDeleteExistingTeamListButton) lessonsMetadataDeleteExistingTeamListButton.style.display = '';
                    resetStatus();
                } else {
                    lessonsMetadataUploadGroup.style.display = 'none';
                    if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    if (lessonsMetadataUpdateActionsGroup) lessonsMetadataUpdateActionsGroup.style.display = 'none';
                    if (lessonsMetadataDeleteExistingXmlButton) lessonsMetadataDeleteExistingXmlButton.style.display = '';
                    if (lessonsMetadataDeleteExistingTeamListButton) lessonsMetadataDeleteExistingTeamListButton.style.display = 'none';
                    resetStatus();
                }
            });
        }

        // Placeholder: Delete Existing XML (no-op for now)
        if (lessonsMetadataDeleteExistingXmlButton) {
            lessonsMetadataDeleteExistingXmlButton.addEventListener('click', async (e) => {
                e.preventDefault();

                if (!lessonsMetadataProjectSelect || !lessonsMetadataProjectSelect.value) {
                    if (lessonsMetadataStatus) {
                        lessonsMetadataStatus.textContent = 'Please select a project first.';
                        lessonsMetadataStatus.classList.add('upload-message--error');
                    }
                    return;
                }

                const confirmed = confirm(
                    'Are you sure you want to delete the existing MS Project data for this project (tasks, resources, and assignments)? This cannot be undone.'
                );
                if (!confirmed) return;

                const projectIdStr = lessonsMetadataProjectSelect.value;
                const projectId = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);

                const chunkArray = (arr, size) => {
                    const n = Math.max(1, Number(size) || 1);
                    const out = [];
                    for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
                    return out;
                };

                const chunkSize = 250;
                const setStatus = (msg, type) => {
                    if (!lessonsMetadataStatus) return;
                    lessonsMetadataStatus.classList.remove('upload-message--success', 'upload-message--error');
                    lessonsMetadataStatus.textContent = msg;
                    if (type === 'success') lessonsMetadataStatus.classList.add('upload-message--success');
                    if (type === 'error') lessonsMetadataStatus.classList.add('upload-message--error');
                };

                try {
                    lessonsMetadataDeleteExistingXmlButton.disabled = true;
                    lessonsMetadataDeleteExistingXmlButton.textContent = 'Deleting...';

                    // 0) Delete assignment link rows (no metadata rows for assignments)
                    setStatus('Deleting assignments...', null);
                    const { data: assignmentRows, error: assignmentErr } = await supabase
                        .from('msproject_assignments')
                        .delete()
                        .eq('organization_id', organizationId)
                        .eq('project_id', projectId)
                        .select('id');

                    if (assignmentErr) {
                        throw new Error(assignmentErr.message || 'Failed deleting assignments.');
                    }

                    let deletedAssignments = (assignmentRows || []).length;

                    setStatus('Finding existing MS Project task entries...', null);

                    const { data: metaRows, error: metaErr } = await supabase
                        .from('lessons_learned_metadata_list')
                        .select('id')
                        .eq('organization_id', organizationId)
                        .eq('project_id', projectId)
                        .eq('metadata_source', 'ms project')
                        .eq('metadata_type', 'task');

                    if (metaErr) {
                        throw new Error(metaErr.message || 'Failed to look up existing MS Project entries.');
                    }

                    const taskIds = (metaRows || []).map(r => r && r.id).filter(Boolean);

                    setStatus('Finding existing MS Project resource entries...', null);
                    const { data: resourceMetaRows, error: resourceMetaErr } = await supabase
                        .from('lessons_learned_metadata_list')
                        .select('id')
                        .eq('organization_id', organizationId)
                        .eq('project_id', projectId)
                        .eq('metadata_source', 'ms project')
                        .eq('metadata_type', 'resource');

                    if (resourceMetaErr) {
                        throw new Error(resourceMetaErr.message || 'Failed to look up existing MS Project resource entries.');
                    }

                    const resourceIds = (resourceMetaRows || []).map(r => r && r.id).filter(Boolean);

                    if (taskIds.length === 0 && resourceIds.length === 0 && deletedAssignments === 0) {
                        setStatus('No existing MS Project entries were found for this project.', 'success');
                        return;
                    }

                    const taskIdChunks = chunkArray(taskIds, chunkSize);
                    const resourceIdChunks = chunkArray(resourceIds, chunkSize);
                    let deletedPred = 0;
                    let deletedDetails = 0;
                    let deletedTaskMeta = 0;
                    let deletedResourceDetails = 0;
                    let deletedResourceMeta = 0;

                    // 1) Delete predecessors
                    if (taskIdChunks.length) {
                        for (let i = 0; i < taskIdChunks.length; i++) {
                            setStatus(`Deleting predecessor links (${i + 1}/${taskIdChunks.length})...`, null);
                            const { data, error } = await supabase
                                .from('msproject_task_predecessors')
                                .delete()
                                .in('lessons_learned_metadata_list_id', taskIdChunks[i])
                                .select('id');
                            if (error) throw new Error(error.message || 'Failed deleting predecessor links.');
                            deletedPred += (data || []).length;
                        }
                    }

                    // 2) Delete task details
                    if (taskIdChunks.length) {
                        for (let i = 0; i < taskIdChunks.length; i++) {
                            setStatus(`Deleting task details (${i + 1}/${taskIdChunks.length})...`, null);
                            const { data, error } = await supabase
                                .from('msproject_task_details')
                                .delete()
                                .in('lessons_learned_metadata_list_id', taskIdChunks[i])
                                .select('id');
                            if (error) throw new Error(error.message || 'Failed deleting task details.');
                            deletedDetails += (data || []).length;
                        }
                    }

                    // 3) Delete task metadata rows
                    if (taskIdChunks.length) {
                        for (let i = 0; i < taskIdChunks.length; i++) {
                            setStatus(`Deleting task metadata (${i + 1}/${taskIdChunks.length})...`, null);
                            const { data, error } = await supabase
                                .from('lessons_learned_metadata_list')
                                .delete()
                                .in('id', taskIdChunks[i])
                                .select('id');
                            if (error) throw new Error(error.message || 'Failed deleting task metadata.');
                            deletedTaskMeta += (data || []).length;
                        }
                    }

                    // 4) Delete resource details rows
                    if (resourceIdChunks.length) {
                        for (let i = 0; i < resourceIdChunks.length; i++) {
                            setStatus(`Deleting resource details (${i + 1}/${resourceIdChunks.length})...`, null);
                            const { data, error } = await supabase
                                .from('msproject_resource_details')
                                .delete()
                                .in('lessons_learned_metadata_list_id', resourceIdChunks[i])
                                .select('id');
                            if (error) throw new Error(error.message || 'Failed deleting resource details.');
                            deletedResourceDetails += (data || []).length;
                        }
                    }

                    // 5) Delete resource metadata rows
                    if (resourceIdChunks.length) {
                        for (let i = 0; i < resourceIdChunks.length; i++) {
                            setStatus(`Deleting resource metadata (${i + 1}/${resourceIdChunks.length})...`, null);
                            const { data, error } = await supabase
                                .from('lessons_learned_metadata_list')
                                .delete()
                                .in('id', resourceIdChunks[i])
                                .select('id');
                            if (error) throw new Error(error.message || 'Failed deleting resource metadata.');
                            deletedResourceMeta += (data || []).length;
                        }
                    }

                    setStatus(
                        `Deleted MS Project data for this project: ${deletedAssignments} assignments, ${deletedTaskMeta} task metadata rows, ${deletedDetails} task detail rows, ${deletedPred} predecessor rows, ${deletedResourceMeta} resource metadata rows, ${deletedResourceDetails} resource detail rows.`,
                        'success'
                    );
                } catch (err) {
                    console.error('Delete Existing XML failed:', err);
                    setStatus(err && err.message ? err.message : 'Failed to delete existing MS Project data.', 'error');
                } finally {
                    lessonsMetadataDeleteExistingXmlButton.disabled = false;
                    lessonsMetadataDeleteExistingXmlButton.textContent = 'Delete Existing XML';
                }
            });
        }

        // Delete Existing Team List (Excel Project Team List - Update)
        if (lessonsMetadataDeleteExistingTeamListButton) {
            lessonsMetadataDeleteExistingTeamListButton.addEventListener('click', async (e) => {
                e.preventDefault();

                if (!lessonsMetadataProjectSelect || !lessonsMetadataProjectSelect.value) {
                    if (lessonsMetadataStatus) {
                        lessonsMetadataStatus.textContent = 'Please select a project first.';
                        lessonsMetadataStatus.classList.add('upload-message--error');
                    }
                    return;
                }

                const confirmed = confirm(
                    'Are you sure you want to delete the existing Excel Project Team List data for this project? This cannot be undone.'
                );
                if (!confirmed) return;

                const setStatus = (msg, type) => {
                    if (!lessonsMetadataStatus) return;
                    lessonsMetadataStatus.classList.remove('upload-message--success', 'upload-message--error');
                    lessonsMetadataStatus.textContent = msg;
                    if (type === 'success') lessonsMetadataStatus.classList.add('upload-message--success');
                    if (type === 'error') lessonsMetadataStatus.classList.add('upload-message--error');
                };

                const projectIdStr = lessonsMetadataProjectSelect.value;
                const projectId = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);

                try {
                    lessonsMetadataDeleteExistingTeamListButton.disabled = true;
                    lessonsMetadataDeleteExistingTeamListButton.textContent = 'Deleting...';

                    // 1) Delete team list rows first (avoids FK issues if present)
                    setStatus('Deleting existing project team list rows...', null);
                    const { data: deletedTeamRows, error: teamErr } = await supabase
                        .from('project_teams_excel')
                        .delete()
                        .eq('organization_id', organizationId)
                        .eq('project_id', projectId)
                        .select('id');

                    if (teamErr) {
                        throw new Error(teamErr.message || 'Failed deleting project team list rows.');
                    }

                    const deletedTeams = (deletedTeamRows || []).length;

                    // 2) Delete related metadata list rows (source: excel project team list)
                    setStatus('Deleting related lessons learned metadata rows...', null);
                    const { data: deletedMetaRows, error: metaErr } = await supabase
                        .from('lessons_learned_metadata_list')
                        .delete()
                        .eq('organization_id', organizationId)
                        .eq('project_id', projectId)
                        .eq('metadata_source', 'excel project team list')
                        .select('id');

                    if (metaErr) {
                        throw new Error(metaErr.message || 'Failed deleting lessons learned metadata rows.');
                    }

                    const deletedMeta = (deletedMetaRows || []).length;

                    setStatus(
                        `Deleted Excel Project Team List data for this project: ${deletedTeams} team rows and ${deletedMeta} metadata rows.`,
                        'success'
                    );
                } catch (err) {
                    console.error('Delete Existing Team List failed:', err);
                    setStatus(err && err.message ? err.message : 'Failed to delete existing team list data.', 'error');
                } finally {
                    lessonsMetadataDeleteExistingTeamListButton.disabled = false;
                    lessonsMetadataDeleteExistingTeamListButton.textContent = 'Delete Existing Team List';
                }
            });
        }

        // Upload handler for Lessons Learned Metadata importers (TXT + Excel)
        if (lessonsMetadataUploadButton && lessonsMetadataFileInput && lessonsMetadataStatus) {
            lessonsMetadataUploadButton.addEventListener('click', async () => {
                lessonsMetadataStatus.classList.remove('upload-message--success', 'upload-message--error');

                // Ensure a project is selected
                if (!lessonsMetadataProjectSelect || !lessonsMetadataProjectSelect.value) {
                    lessonsMetadataStatus.textContent = 'Please select a project first.';
                    lessonsMetadataStatus.classList.add('upload-message--error');
                    return;
                }

                // Ensure an uploadable file type is selected
                const selectedType = lessonsMetadataFileTypeSelect ? lessonsMetadataFileTypeSelect.value : '';
                const isMsProjectNew = selectedType === 'ms_project_xml';
                const isExcelTeamListNew = selectedType === 'excel_project_team_list_new';
                const isExcelMiscListNew = selectedType === 'excel_project_misc_list_new';

                if (!isMsProjectNew && !isExcelTeamListNew && !isExcelMiscListNew) {
                    lessonsMetadataStatus.textContent = 'Please select an input data file type that supports uploads.';
                    lessonsMetadataStatus.classList.add('upload-message--error');
                    return;
                }

                const file = lessonsMetadataFileInput.files && lessonsMetadataFileInput.files[0];
                if (!file) {
                    lessonsMetadataStatus.textContent = 'Please choose a file to upload.';
                    lessonsMetadataStatus.classList.add('upload-message--error');
                    return;
                }

                const name = file.name || '';
                const isTxt = /\.txt$/i.test(name);
                const isExcel = /\.xlsx$/i.test(name) || /\.xls$/i.test(name);

                if (isMsProjectNew) {
                    if (!isTxt) {
                        lessonsMetadataStatus.textContent = 'Invalid file type. Please upload a file with a .txt extension.';
                        lessonsMetadataStatus.classList.add('upload-message--error');
                        return;
                    }

                    // Basic MIME type check (not all browsers set this reliably, so it is secondary)
                    if (file.type && file.type !== 'text/plain') {
                        console.warn('File MIME type is not text/plain:', file.type);
                    }
                }

                if (isExcelTeamListNew) {
                    if (!isExcel) {
                        lessonsMetadataStatus.textContent = 'Invalid file type. Please upload an Excel file (.xlsx, .xls).';
                        lessonsMetadataStatus.classList.add('upload-message--error');
                        return;
                    }
                }

                if (isExcelMiscListNew) {
                    if (!isExcel) {
                        lessonsMetadataStatus.textContent = 'Invalid file type. Please upload an Excel file (.xlsx, .xls).';
                        lessonsMetadataStatus.classList.add('upload-message--error');
                        return;
                    }
                }

                try {
                    lessonsMetadataUploadButton.disabled = true;
                    lessonsMetadataUploadButton.textContent = 'Uploading...';

                    const projectIdStr = lessonsMetadataProjectSelect.value;
                    const createdBy = ctUser && ctUser.id != null ? Number(ctUser.id) : null;
                    if (!createdBy || Number.isNaN(createdBy)) {
                        throw new Error('Could not determine the uploader user id. Please log out and log back in.');
                    }

                    const projectId = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);

                    // Always fetch project_type_id live to avoid stale cache/nulls
                    const { data: projectRow, error: projectErr } = await supabase
                        .from('projects')
                        .select('project_type_id')
                        .eq('project_id', projectId)
                        .maybeSingle();

                    if (projectErr || !projectRow || projectRow.project_type_id == null) {
                        throw new Error('Could not determine the selected project type. Please refresh and try again.');
                    }

                    const projectTypeId = projectRow.project_type_id;

                    const context = {
                        organization_id: organizationId,
                        created_by: createdBy,
                        project_id: projectId,
                        project_type_id: projectTypeId
                    };

                    const onProgress = (msg) => {
                        lessonsMetadataStatus.textContent = msg;
                        lessonsMetadataStatus.classList.remove('upload-message--success', 'upload-message--error');
                    };

                    if (isMsProjectNew) {
                        // Delegate parse + persistence to the TXT importer.
                        const result = await importMsProjectTxtToSupabase({
                            supabase,
                            file,
                            context,
                            chunkSize: 250,
                            onProgress
                        });

                        // Build a user-friendly summary of which sections are present.
                        const count = result && typeof result.presentCount === 'number'
                            ? result.presentCount
                            : (result && Array.isArray(result.presentSections) ? result.presentSections.length : 0);
                        const sectionsList = result && Array.isArray(result.presentSections)
                            ? result.presentSections.join(', ')
                            : '';

                        let summaryMessage = '';
                        if (count === 0) {
                            summaryMessage = 'This TXT file does not contain Resources, Tasks, or Assignments sections.';
                        } else if (count === 1) {
                            summaryMessage = `This TXT file contains 1 of 3 expected sections: ${sectionsList}.`;
                        } else {
                            summaryMessage = `This TXT file contains ${count} of 3 expected sections: ${sectionsList}.`;
                        }

                        const inserted = result && result.inserted ? result.inserted : {};
                        const insertedTasks = inserted.msproject_task_details || 0;
                        const insertedPred = inserted.msproject_task_predecessors || 0;

                        lessonsMetadataStatus.textContent =
                            `Imported "${name}". ${summaryMessage} Inserted ${insertedTasks} tasks and ${insertedPred} predecessor links.`;
                        lessonsMetadataStatus.classList.add('upload-message--success');
                        if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    } else if (isExcelTeamListNew) {
                        // Delegate parse + persistence to the Excel importer.
                        const result = await importProjectTeamListExcelToSupabase({
                            supabase,
                            file,
                            context,
                            chunkSize: 250,
                            onProgress
                        });

                        const inserted = result && result.inserted ? result.inserted : {};
                        const insertedMeta = inserted.lessons_learned_metadata_list || 0;
                        const insertedTeams = inserted.project_teams_excel || 0;

                        lessonsMetadataStatus.textContent =
                            `Imported "${name}". Inserted ${insertedTeams} project team rows and ${insertedMeta} metadata rows.`;
                        lessonsMetadataStatus.classList.add('upload-message--success');
                        if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    } else if (isExcelMiscListNew) {
                        // Delegate parse + persistence to the Misc Excel importer.
                        const result = await importProjectMiscListExcelToSupabase({
                            supabase,
                            file,
                            context,
                            chunkSize: 250,
                            onProgress
                        });

                        const inserted = result && result.inserted ? result.inserted : {};
                        const insertedMeta = inserted.lessons_learned_metadata_list || 0;
                        const insertedMisc = inserted.project_miscellaneous_excel || 0;

                        lessonsMetadataStatus.textContent =
                            `Imported "${name}". Inserted ${insertedMisc} miscellaneous rows and ${insertedMeta} metadata rows.`;
                        lessonsMetadataStatus.classList.add('upload-message--success');
                        if (lessonsMetadataFileInput) lessonsMetadataFileInput.value = '';
                    } else {
                        throw new Error('Unsupported file type selection.');
                    }
                } catch (err) {
                    console.error('Error processing Lessons Learned Metadata upload:', err);
                    lessonsMetadataStatus.textContent = err && err.message
                        ? err.message
                        : 'Failed to process upload.';
                    lessonsMetadataStatus.classList.add('upload-message--error');
                } finally {
                    if (lessonsMetadataUploadButton) {
                        lessonsMetadataUploadButton.disabled = false;
                        lessonsMetadataUploadButton.textContent = 'Upload File';
                    }
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
        if (projectTeamPanel) {
            projectTeamPanel.style.display = 'none';
        }
        if (projectTeamAssignmentsPanel) {
            projectTeamAssignmentsPanel.style.display = 'none';
        }
        if (orgProjectDetailsPanel) {
            orgProjectDetailsPanel.style.display = '';
            orgProjectDetailsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        if (orgProjectDetailsManagePanel) {
            orgProjectDetailsManagePanel.style.display = 'none';
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

    function ensureProjectTeamPanel() {
        if (projectTeamPanel) return;

        const createView = document.querySelector('#createView');
        if (!createView) return;

        createView.insertAdjacentHTML('beforeend', `
            <section id="projectTeamPanel" class="project-types-panel" style="display: none; margin-top: 16px;">
                <div class="project-types-panel-header">
                    <div>
                        <h3>Project Team</h3>
                        <p class="subtitle">Select a user from your organization.</p>
                    </div>
                    <button type="button" id="projectTeamBackButton" class="secondary-button">Back to Create Project</button>
                </div>
                <div class="form-group">
                    <label for="projectTeamProjectSelect">Project</label>
                    <select id="projectTeamProjectSelect">
                        <option value=\"\">Select Project</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="projectTeamUserSelect">User</label>
                    <select id="projectTeamUserSelect">
                        <option value=\"\">Select User</option>
                    </select>
                </div>
                <div class="form-buttons">
                    <button type="button" id="projectTeamAddButton" class="secondary-button">Add Team Member</button>
                    <button type="button" id="projectTeamAssignmentsButton" class="secondary-button" style="display:none;">Team Member Assignments</button>
                </div>
                <div id="projectTeamStatus" class="upload-message" aria-live="polite"></div>
            </section>
        `);

        projectTeamPanel = document.getElementById('projectTeamPanel');
        projectTeamProjectSelect = document.getElementById('projectTeamProjectSelect');
        projectTeamUserSelect = document.getElementById('projectTeamUserSelect');
        projectTeamBackButton = document.getElementById('projectTeamBackButton');
        projectTeamAddButton = document.getElementById('projectTeamAddButton');
        projectTeamAssignmentsButton = document.getElementById('projectTeamAssignmentsButton');
        projectTeamStatus = document.getElementById('projectTeamStatus');

        if (projectTeamBackButton) {
            projectTeamBackButton.addEventListener('click', (e) => {
                e.preventDefault();
                hideProjectTeamPanel();
            });
        }

        if (projectTeamProjectSelect) {
            projectTeamProjectSelect.addEventListener('change', () => {
                updateProjectTeamAssignmentsButtonVisibility();
            });
        }

        if (projectTeamUserSelect) {
            projectTeamUserSelect.addEventListener('change', () => {
                updateProjectTeamAssignmentsButtonVisibility();
            });
        }

        if (projectTeamAddButton) {
            projectTeamAddButton.addEventListener('click', (e) => {
                e.preventDefault();
                addSelectedProjectTeamMember();
            });
        }

        if (projectTeamAssignmentsButton) {
            projectTeamAssignmentsButton.addEventListener('click', (e) => {
                e.preventDefault();
                showProjectTeamAssignmentsPanel();
            });
        }
    }

    function showProjectTeamPanel() {
        ensureProjectTeamPanel();
        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = 'none';
        }
        if (orgProjectDetailsPanel) orgProjectDetailsPanel.style.display = 'none';
        if (orgProjectDetailsManagePanel) orgProjectDetailsManagePanel.style.display = 'none';
        if (lessonsMetadataPanel) lessonsMetadataPanel.style.display = 'none';
        if (lessonsMetadataManagePanel) lessonsMetadataManagePanel.style.display = 'none';
        if (projectTeamAssignmentsPanel) projectTeamAssignmentsPanel.style.display = 'none';

        if (projectTeamPanel) {
            projectTeamPanel.style.display = '';
            projectTeamPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }

        loadOrgProjectsForProjectTeamDropdown();
        loadOrgUsersForProjectTeamDropdown();
    }

    function hideProjectTeamPanel() {
        if (projectTeamPanel) {
            projectTeamPanel.style.display = 'none';
        }
        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = '';
        }
    }

    function ensureProjectTeamAssignmentsPanel() {
        if (projectTeamAssignmentsPanel) return;

        const createView = document.querySelector('#createView');
        if (!createView) return;

        createView.insertAdjacentHTML('beforeend', `
            <section id="projectTeamAssignmentsPanel" class="project-types-panel" style="display: none; margin-top: 16px;">
                <div class="project-types-panel-header">
                    <div>
                        <h3 id="projectTeamAssignmentsTitle">Team Member Assignments</h3>
                        <p class="subtitle">View lessons learned metadata for the selected project.</p>
                    </div>
                    <button type="button" id="projectTeamAssignmentsBackButton" class="secondary-button">Back</button>
                </div>
                <div class="form-group">
                    <label for="projectTeamAssignmentsTypeSelect">Type (filter)</label>
                    <select id="projectTeamAssignmentsTypeSelect">
                        <option value=\"\">All Types</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="projectTeamAssignmentsSearchInput">Search (metadata)</label>
                    <input
                        type="text"
                        id="projectTeamAssignmentsSearchInput"
                        placeholder="Type to filter metadata..."
                        autocomplete="off"
                    >
                </div>
                <div class="form-buttons">
                    <button type="button" id="projectTeamAssignmentsRefreshButton" class="secondary-button">Refresh List</button>
                    <button type="button" id="projectTeamAssignmentsSaveButton" class="secondary-button">Save Selected Assignments</button>
                </div>
                <div id="projectTeamAssignmentsStatus" class="upload-message" aria-live="polite"></div>
                <div id="projectTeamAssignmentsTableWrap" style="margin-top: 12px;"></div>
            </section>
        `);

        projectTeamAssignmentsPanel = document.getElementById('projectTeamAssignmentsPanel');
        projectTeamAssignmentsBackButton = document.getElementById('projectTeamAssignmentsBackButton');
        projectTeamAssignmentsRefreshButton = document.getElementById('projectTeamAssignmentsRefreshButton');
        projectTeamAssignmentsSaveButton = document.getElementById('projectTeamAssignmentsSaveButton');
        projectTeamAssignmentsTypeSelect = document.getElementById('projectTeamAssignmentsTypeSelect');
        projectTeamAssignmentsSearchInput = document.getElementById('projectTeamAssignmentsSearchInput');
        projectTeamAssignmentsStatus = document.getElementById('projectTeamAssignmentsStatus');
        projectTeamAssignmentsTableWrap = document.getElementById('projectTeamAssignmentsTableWrap');

        if (projectTeamAssignmentsBackButton) {
            projectTeamAssignmentsBackButton.addEventListener('click', (e) => {
                e.preventDefault();
                hideProjectTeamAssignmentsPanel();
            });
        }

        if (projectTeamAssignmentsRefreshButton) {
            projectTeamAssignmentsRefreshButton.addEventListener('click', (e) => {
                e.preventDefault();
                refreshProjectTeamAssignmentsList();
            });
        }

        if (projectTeamAssignmentsSaveButton) {
            projectTeamAssignmentsSaveButton.addEventListener('click', (e) => {
                e.preventDefault();
                saveProjectTeamAssignments();
            });
        }

        if (projectTeamAssignmentsTypeSelect) {
            projectTeamAssignmentsTypeSelect.addEventListener('change', () => {
                renderProjectTeamAssignmentsTable();
            });
        }

        if (projectTeamAssignmentsSearchInput) {
            projectTeamAssignmentsSearchInput.addEventListener('input', () => {
                renderProjectTeamAssignmentsTable();
            });
        }
    }

    async function showProjectTeamAssignmentsPanel() {
        ensureProjectTeamAssignmentsPanel();

        // Ensure we have a selected project (scope is fixed to Project Team selection)
        const projectIdStr = projectTeamProjectSelect ? projectTeamProjectSelect.value : '';
        const userIdStr = projectTeamUserSelect ? projectTeamUserSelect.value : '';
        if (!projectIdStr) {
            if (projectTeamStatus) {
                projectTeamStatus.classList.remove('upload-message--success');
                projectTeamStatus.classList.add('upload-message--error');
                projectTeamStatus.textContent = 'Please select a project first.';
            }
            return;
        }
        if (!userIdStr) {
            if (projectTeamStatus) {
                projectTeamStatus.classList.remove('upload-message--success');
                projectTeamStatus.classList.add('upload-message--error');
                projectTeamStatus.textContent = 'Please select a user first.';
            }
            return;
        }

        // Update title: Team Member Assignments for <Name>
        const titleEl = projectTeamAssignmentsPanel
            ? projectTeamAssignmentsPanel.querySelector('#projectTeamAssignmentsTitle')
            : null;
        if (titleEl) {
            const userInfo = projectTeamUsersById ? projectTeamUsersById.get(String(userIdStr)) : null;
            const name = userInfo && userInfo.name ? userInfo.name : 'Selected User';
            titleEl.textContent = `Team Member Assignments for ${name}`;
        }

        // Establish scope for bulk save + preload
        projectTeamAssignmentsProjectId = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);
        projectTeamAssignmentsUserId = Number.isNaN(Number(userIdStr)) ? userIdStr : Number(userIdStr);
        projectTeamAssignmentsSavedIds = new Set();
        projectTeamAssignmentsSelectedIds = new Set();

        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = 'none';
        }

        if (projectTeamPanel) {
            projectTeamPanel.style.display = 'none';
        }
        if (orgProjectDetailsPanel) orgProjectDetailsPanel.style.display = 'none';
        if (orgProjectDetailsManagePanel) orgProjectDetailsManagePanel.style.display = 'none';
        if (lessonsMetadataPanel) lessonsMetadataPanel.style.display = 'none';
        if (lessonsMetadataManagePanel) lessonsMetadataManagePanel.style.display = 'none';

        if (projectTeamAssignmentsPanel) {
            projectTeamAssignmentsPanel.style.display = '';
            projectTeamAssignmentsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }

        await preloadProjectTeamAssignmentsSavedIds();
        await refreshProjectTeamAssignmentsList();
    }

    function hideProjectTeamAssignmentsPanel() {
        if (projectTeamAssignmentsPanel) {
            projectTeamAssignmentsPanel.style.display = 'none';
        }
        // Return to Project Team submodule
        showProjectTeamPanel();
    }

    function updateProjectTeamAssignmentsTypeOptions(rows) {
        if (!projectTeamAssignmentsTypeSelect) return;

        const previous = projectTeamAssignmentsTypeSelect.value || '';
        const types = Array.from(
            new Set((rows || []).map(r => r && r.metadata_type).filter(Boolean))
        ).sort((a, b) => String(a).localeCompare(String(b)));

        projectTeamAssignmentsTypeSelect.innerHTML = '';

        const allOpt = document.createElement('option');
        allOpt.value = '';
        allOpt.textContent = 'All Types';
        projectTeamAssignmentsTypeSelect.appendChild(allOpt);

        types.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t;
            opt.textContent = t;
            projectTeamAssignmentsTypeSelect.appendChild(opt);
        });

        projectTeamAssignmentsTypeSelect.value = types.includes(previous) ? previous : '';
    }

    async function preloadProjectTeamAssignmentsSavedIds() {
        if (!organizationId || projectTeamAssignmentsProjectId == null || projectTeamAssignmentsUserId == null) return;

        if (projectTeamAssignmentsStatus) {
            projectTeamAssignmentsStatus.classList.remove('upload-message--success', 'upload-message--error');
            projectTeamAssignmentsStatus.textContent = 'Loading saved assignments...';
        }

        try {
            const { data: rows, error } = await supabase
                .from('project_team_member_assignments')
                .select('lessons_learned_metadata_list_id')
                .eq('organization_id', organizationId)
                .eq('project_id', projectTeamAssignmentsProjectId)
                .eq('user_id', projectTeamAssignmentsUserId)
                .limit(5000);

            if (error) {
                console.error('Error preloading project team member assignments:', error);
                if (projectTeamAssignmentsStatus) {
                    projectTeamAssignmentsStatus.textContent = 'Failed to load saved assignments.';
                    projectTeamAssignmentsStatus.classList.add('upload-message--error');
                }
                projectTeamAssignmentsSavedIds = new Set();
                projectTeamAssignmentsSelectedIds = new Set();
                return;
            }

            const ids = (rows || [])
                .map(r => r && r.lessons_learned_metadata_list_id)
                .filter(v => v != null)
                .map(v => String(v));

            projectTeamAssignmentsSavedIds = new Set(ids);
            // Start selection equal to saved state (pre-highlight)
            projectTeamAssignmentsSelectedIds = new Set(ids);
        } catch (err) {
            console.error('Unexpected error preloading project team member assignments:', err);
            if (projectTeamAssignmentsStatus) {
                projectTeamAssignmentsStatus.textContent = 'An unexpected error occurred while loading saved assignments.';
                projectTeamAssignmentsStatus.classList.add('upload-message--error');
            }
            projectTeamAssignmentsSavedIds = new Set();
            projectTeamAssignmentsSelectedIds = new Set();
        }
    }

    function renderProjectTeamAssignmentsTable() {
        if (!projectTeamAssignmentsTableWrap) return;

        const selectedType = projectTeamAssignmentsTypeSelect ? projectTeamAssignmentsTypeSelect.value : '';
        const rows = Array.isArray(projectTeamAssignmentsRowsCache) ? projectTeamAssignmentsRowsCache : [];

        const searchRaw = projectTeamAssignmentsSearchInput ? projectTeamAssignmentsSearchInput.value : '';
        const search = String(searchRaw || '').trim().toLowerCase();

        let filtered = rows;
        if (selectedType) {
            filtered = filtered.filter(r => r && r.metadata_type === selectedType);
        }
        if (search) {
            filtered = filtered.filter(r => String((r && r.metadata) || '').toLowerCase().includes(search));
        }

        projectTeamAssignmentsTableWrap.innerHTML = '';

        if (filtered.length === 0) {
            const empty = document.createElement('div');
            const searchDisplay = String(searchRaw || '').trim();
            empty.textContent =
                selectedType && searchDisplay
                    ? `No metadata entries match type "${selectedType}" and "${searchDisplay}".`
                    : searchDisplay
                        ? `No metadata entries match "${searchDisplay}".`
                        : selectedType
                            ? `No metadata entries found for type "${selectedType}".`
                            : 'No metadata has been imported for this project yet.';
            projectTeamAssignmentsTableWrap.appendChild(empty);

            if (projectTeamAssignmentsStatus) {
                projectTeamAssignmentsStatus.classList.remove('upload-message--success', 'upload-message--error');
                if (selectedType || search) {
                    projectTeamAssignmentsStatus.textContent = selectedType
                        ? `Showing 0 ${search ? 'matching ' : ''}"${selectedType}" entries (of ${rows.length} total).`
                        : `Showing 0 matching entries (of ${rows.length} total).`;
                } else {
                    projectTeamAssignmentsStatus.textContent = `Loaded ${rows.length} metadata entr${rows.length === 1 ? 'y' : 'ies'}.`;
                }
                if (rows.length > 0) projectTeamAssignmentsStatus.classList.add('upload-message--success');
            }
            return;
        }

        const table = document.createElement('table');
        table.className = 'organizations-table';

        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        ['Type', 'Metadata'].forEach(label => {
            const th = document.createElement('th');
            th.textContent = label;
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        filtered.forEach(row => {
            const tr = document.createElement('tr');
            const lessonsIdStr = String(row.id);
            tr.style.cursor = 'pointer';
            if (projectTeamAssignmentsSelectedIds && projectTeamAssignmentsSelectedIds.has(lessonsIdStr)) {
                tr.classList.add('row-selected');
            }
            tr.addEventListener('click', () => {
                if (!projectTeamAssignmentsSelectedIds) {
                    projectTeamAssignmentsSelectedIds = new Set();
                }
                if (projectTeamAssignmentsSelectedIds.has(lessonsIdStr)) {
                    projectTeamAssignmentsSelectedIds.delete(lessonsIdStr);
                    tr.classList.remove('row-selected');
                } else {
                    projectTeamAssignmentsSelectedIds.add(lessonsIdStr);
                    tr.classList.add('row-selected');
                }
            });

            const tdType = document.createElement('td');
            tdType.textContent = row.metadata_type || '';
            tr.appendChild(tdType);

            const tdMetadata = document.createElement('td');
            tdMetadata.textContent = row.metadata || '';
            tr.appendChild(tdMetadata);

            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        projectTeamAssignmentsTableWrap.appendChild(table);

        if (projectTeamAssignmentsStatus) {
            projectTeamAssignmentsStatus.classList.remove('upload-message--success', 'upload-message--error');
            if (selectedType || search) {
                const typePart = selectedType ? `"${selectedType}" ` : '';
                const matchPart = search ? 'matching ' : '';
                projectTeamAssignmentsStatus.textContent =
                    `Showing ${filtered.length} ${matchPart}${typePart}entr${filtered.length === 1 ? 'y' : 'ies'} (of ${rows.length} total).`;
            } else {
                projectTeamAssignmentsStatus.textContent = `Loaded ${rows.length} metadata entr${rows.length === 1 ? 'y' : 'ies'}.`;
            }
            projectTeamAssignmentsStatus.classList.add('upload-message--success');
        }
    }

    async function refreshProjectTeamAssignmentsList() {
        if (!organizationId || !projectTeamAssignmentsStatus || !projectTeamAssignmentsTableWrap) return;
        if (!projectTeamProjectSelect) return;

        const projectIdStr = projectTeamProjectSelect.value;
        if (!projectIdStr) {
            projectTeamAssignmentsTableWrap.innerHTML = '';
            projectTeamAssignmentsStatus.textContent = 'Please select a project first.';
            projectTeamAssignmentsStatus.classList.remove('upload-message--success');
            projectTeamAssignmentsStatus.classList.add('upload-message--error');
            projectTeamAssignmentsRowsCache = [];
            if (projectTeamAssignmentsTypeSelect) projectTeamAssignmentsTypeSelect.value = '';
            if (projectTeamAssignmentsSearchInput) projectTeamAssignmentsSearchInput.value = '';
            return;
        }

        const project_id = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);
        projectTeamAssignmentsStatus.classList.remove('upload-message--success', 'upload-message--error');
        projectTeamAssignmentsStatus.textContent = 'Loading metadata...';
        projectTeamAssignmentsTableWrap.innerHTML = '';

        try {
            const { data: rows, error } = await supabase
                .from('lessons_learned_metadata_list')
                .select('id, metadata_type, metadata')
                .eq('organization_id', organizationId)
                .eq('project_id', project_id)
                .order('id', { ascending: true })
                .limit(5000);

            if (error) {
                console.error('Error loading lessons learned metadata for project team assignments:', error);
                projectTeamAssignmentsStatus.textContent = 'Failed to load metadata.';
                projectTeamAssignmentsStatus.classList.add('upload-message--error');
                projectTeamAssignmentsRowsCache = [];
                return;
            }

            const data = rows || [];
            projectTeamAssignmentsRowsCache = data;
            updateProjectTeamAssignmentsTypeOptions(data);
            renderProjectTeamAssignmentsTable();
        } catch (err) {
            console.error('Unexpected error loading lessons learned metadata for project team assignments:', err);
            projectTeamAssignmentsStatus.textContent = 'An unexpected error occurred while loading metadata.';
            projectTeamAssignmentsStatus.classList.add('upload-message--error');
            projectTeamAssignmentsRowsCache = [];
        }
    }

    async function saveProjectTeamAssignments() {
        if (!organizationId || !projectTeamAssignmentsStatus) return;

        if (projectTeamAssignmentsProjectId == null || projectTeamAssignmentsUserId == null) {
            projectTeamAssignmentsStatus.classList.remove('upload-message--success');
            projectTeamAssignmentsStatus.classList.add('upload-message--error');
            projectTeamAssignmentsStatus.textContent = 'Missing project or user selection.';
            return;
        }

        const selected = projectTeamAssignmentsSelectedIds instanceof Set ? projectTeamAssignmentsSelectedIds : new Set();
        const saved = projectTeamAssignmentsSavedIds instanceof Set ? projectTeamAssignmentsSavedIds : new Set();

        const selectedIds = Array.from(selected);
        const savedIds = Array.from(saved);

        const selectedSet = new Set(selectedIds);
        const savedSet = new Set(savedIds);

        const toInsert = selectedIds.filter(id => !savedSet.has(id));
        const toDelete = savedIds.filter(id => !selectedSet.has(id));

        const coerceId = (idStr) => {
            const n = Number(idStr);
            return Number.isNaN(n) ? idStr : n;
        };

        const chunkArray = (arr, size) => {
            const n = Math.max(1, Number(size) || 1);
            const out = [];
            for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
            return out;
        };

        // Build lookup from lessons id -> row for assignment fields
        const rowById = new Map();
        (projectTeamAssignmentsRowsCache || []).forEach(r => {
            if (r && r.id != null) rowById.set(String(r.id), r);
        });

        // Resolve user name
        const userInfo = projectTeamUsersById ? projectTeamUsersById.get(String(projectTeamAssignmentsUserId)) : null;
        const name = userInfo && userInfo.name ? userInfo.name : null;

        projectTeamAssignmentsStatus.classList.remove('upload-message--success', 'upload-message--error');
        projectTeamAssignmentsStatus.textContent = 'Saving assignments...';

        if (projectTeamAssignmentsSaveButton) {
            projectTeamAssignmentsSaveButton.disabled = true;
            projectTeamAssignmentsSaveButton.textContent = 'Saving...';
        }

        try {
            // 1) Insert new selections
            let insertedCount = 0;
            if (toInsert.length > 0) {
                const payload = [];
                toInsert.forEach(idStr => {
                    const row = rowById.get(String(idStr));
                    if (!row) return;
                    payload.push({
                        organization_id: organizationId,
                        project_id: projectTeamAssignmentsProjectId,
                        user_id: projectTeamAssignmentsUserId,
                        name,
                        lessons_learned_metadata_list_id: coerceId(idStr),
                        assignment: row.metadata || null,
                        assignment_type: row.metadata_type || null
                    });
                });

                const chunks = chunkArray(payload, 250);
                for (let i = 0; i < chunks.length; i++) {
                    const chunk = chunks[i];
                    if (chunk.length === 0) continue;
                    const { error } = await supabase
                        .from('project_team_member_assignments')
                        .insert(chunk);
                    if (error) throw error;
                    insertedCount += chunk.length;
                }
            }

            // 2) Delete removed selections
            let deletedCount = 0;
            if (toDelete.length > 0) {
                const idsToDelete = toDelete.map(coerceId);
                const chunks = chunkArray(idsToDelete, 250);
                for (let i = 0; i < chunks.length; i++) {
                    const chunk = chunks[i];
                    if (chunk.length === 0) continue;
                    const { data, error } = await supabase
                        .from('project_team_member_assignments')
                        .delete()
                        .eq('organization_id', organizationId)
                        .eq('project_id', projectTeamAssignmentsProjectId)
                        .eq('user_id', projectTeamAssignmentsUserId)
                        .in('lessons_learned_metadata_list_id', chunk)
                        .select('id');
                    if (error) throw error;
                    deletedCount += (data || []).length;
                }
            }

            // Update saved set to match selected set
            projectTeamAssignmentsSavedIds = new Set(Array.from(selectedSet));

            projectTeamAssignmentsStatus.classList.remove('upload-message--error');
            projectTeamAssignmentsStatus.classList.add('upload-message--success');
            projectTeamAssignmentsStatus.textContent =
                `Saved assignments. Added ${toInsert.length} and removed ${toDelete.length}.`;
        } catch (err) {
            console.error('Failed saving project team member assignments:', err);
            projectTeamAssignmentsStatus.classList.remove('upload-message--success');
            projectTeamAssignmentsStatus.classList.add('upload-message--error');
            projectTeamAssignmentsStatus.textContent = 'Failed to save selected assignments.';
        } finally {
            if (projectTeamAssignmentsSaveButton) {
                projectTeamAssignmentsSaveButton.disabled = false;
                projectTeamAssignmentsSaveButton.textContent = 'Save Selected Assignments';
            }
        }
    }

    async function loadOrgUsersForProjectTeamDropdown() {
        if (!organizationId || !projectTeamUserSelect || !projectTeamStatus) return;

        projectTeamStatus.classList.remove('upload-message--success', 'upload-message--error');
        projectTeamStatus.textContent = 'Loading users...';

        try {
            const { data: userRows, error: userErr } = await supabase
                .from('users')
                .select('id, name, email, usertype, organizationid')
                .eq('organizationid', organizationId)
                .order('name', { ascending: true })
                .limit(5000);

            if (userErr) {
                console.error('Error loading organization users for project team dropdown:', userErr);
                projectTeamStatus.textContent = 'Failed to load users.';
                projectTeamStatus.classList.add('upload-message--error');
                return;
            }

            const rows = userRows || [];
            if (rows.length === 0) {
                projectTeamStatus.textContent = 'No users found for your organization.';
                projectTeamStatus.classList.add('upload-message--error');
                if (projectTeamChoices) {
                    projectTeamChoices.clearChoices();
                } else {
                    projectTeamUserSelect.innerHTML = '<option value=\"\">No users available</option>';
                }
                projectTeamUsersById = new Map();
                return;
            }

            projectTeamUsersById = new Map();
            const choicesData = rows.map(row => {
                const idStr = String(row.id);
                const name = row.name || `User ${idStr}`;
                const email = row.email || '';
                const usertype = row.usertype || '';
                projectTeamUsersById.set(idStr, { name, email, usertype });
                const label = email ? `${name} (${email})` : name;
                return { value: idStr, label };
            });

            if (!projectTeamChoices) {
                if (typeof Choices === 'undefined') {
                    console.warn('Choices library not loaded; falling back to native select for project team users dropdown.');
                    projectTeamUserSelect.innerHTML = '<option value=\"\">Select User</option>';
                    choicesData.forEach(choice => {
                        const option = document.createElement('option');
                        option.value = choice.value;
                        option.textContent = choice.label;
                        projectTeamUserSelect.appendChild(option);
                    });
                } else {
                    projectTeamChoices = new Choices(projectTeamUserSelect, {
                        searchEnabled: true,
                        shouldSort: false,
                        placeholder: true,
                        placeholderValue: 'Select User',
                        searchPlaceholderValue: 'Type to search...'
                    });
                }
            }

            if (projectTeamChoices) {
                projectTeamChoices.setChoices(choicesData, 'value', 'label', true);
            } else {
                projectTeamUserSelect.innerHTML = '<option value=\"\">Select User</option>';
                choicesData.forEach(choice => {
                    const option = document.createElement('option');
                    option.value = choice.value;
                    option.textContent = choice.label;
                    projectTeamUserSelect.appendChild(option);
                });
            }

            projectTeamStatus.textContent = '';
            projectTeamStatus.classList.remove('upload-message--error');
        } catch (err) {
            console.error('Unexpected error loading organization users for project team dropdown:', err);
            projectTeamStatus.textContent = 'An unexpected error occurred while loading users.';
            projectTeamStatus.classList.add('upload-message--error');
        }
    }

    async function loadOrgProjectsForProjectTeamDropdown() {
        if (!organizationId || !projectTeamProjectSelect || !projectTeamStatus) return;

        projectTeamStatus.classList.remove('upload-message--success', 'upload-message--error');
        projectTeamStatus.textContent = 'Loading projects...';

        try {
            const { data: projectRows, error: projectErr } = await supabase
                .from('projects')
                .select('project_id, project_name')
                .eq('organization_id', organizationId)
                .order('project_name', { ascending: true });

            if (projectErr) {
                console.error('Error loading organization projects for project team dropdown:', projectErr);
                projectTeamStatus.textContent = 'Failed to load projects.';
                projectTeamStatus.classList.add('upload-message--error');
                return;
            }

            const rows = projectRows || [];
            if (rows.length === 0) {
                projectTeamStatus.textContent = 'No projects found for your organization.';
                projectTeamStatus.classList.add('upload-message--error');
                if (projectTeamProjectChoices) {
                    projectTeamProjectChoices.clearChoices();
                } else {
                    projectTeamProjectSelect.innerHTML = '<option value=\"\">No projects available</option>';
                }
                projectTeamProjectsById = new Map();
                return;
            }

            projectTeamProjectsById = new Map();
            const choicesData = rows.map(row => {
                const idStr = String(row.project_id);
                const name = row.project_name || `Project ${idStr}`;
                projectTeamProjectsById.set(idStr, { name });
                return { value: idStr, label: name };
            });

            if (!projectTeamProjectChoices) {
                if (typeof Choices === 'undefined') {
                    console.warn('Choices library not loaded; falling back to native select for project team projects dropdown.');
                    projectTeamProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                    choicesData.forEach(choice => {
                        const option = document.createElement('option');
                        option.value = choice.value;
                        option.textContent = choice.label;
                        projectTeamProjectSelect.appendChild(option);
                    });
                } else {
                    projectTeamProjectChoices = new Choices(projectTeamProjectSelect, {
                        searchEnabled: true,
                        shouldSort: false,
                        placeholder: true,
                        placeholderValue: 'Select Project',
                        searchPlaceholderValue: 'Type to search...'
                    });
                }
            }

            if (projectTeamProjectChoices) {
                projectTeamProjectChoices.setChoices(choicesData, 'value', 'label', true);
            } else {
                projectTeamProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                choicesData.forEach(choice => {
                    const option = document.createElement('option');
                    option.value = choice.value;
                    option.textContent = choice.label;
                    projectTeamProjectSelect.appendChild(option);
                });
            }

            // Don't clear status entirely if users are still loading; but keep it non-error
            if (projectTeamStatus.classList.contains('upload-message--error')) {
                projectTeamStatus.classList.remove('upload-message--error');
            }
            if (projectTeamStatus.textContent === 'Loading projects...') {
                projectTeamStatus.textContent = '';
            }
        } catch (err) {
            console.error('Unexpected error loading organization projects for project team dropdown:', err);
            projectTeamStatus.textContent = 'An unexpected error occurred while loading projects.';
            projectTeamStatus.classList.add('upload-message--error');
        }
    }

    async function updateProjectTeamAssignmentsButtonVisibility() {
        if (!projectTeamAssignmentsButton) return;
        if (!organizationId || !projectTeamProjectSelect || !projectTeamUserSelect) {
            projectTeamAssignmentsButton.style.display = 'none';
            return;
        }

        const projectIdStr = projectTeamProjectSelect.value;
        const userIdStr = projectTeamUserSelect.value;
        if (!projectIdStr || !userIdStr) {
            projectTeamAssignmentsButton.style.display = 'none';
            return;
        }

        const project_id = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);
        const user_id = Number.isNaN(Number(userIdStr)) ? userIdStr : Number(userIdStr);

        try {
            const { data: existing, error } = await supabase
                .from('project_team_members')
                .select('id')
                .eq('organization_id', organizationId)
                .eq('project_id', project_id)
                .eq('user_id', user_id)
                .limit(1);

            if (error) {
                console.warn('Failed checking project team membership for assignments button visibility:', error);
                projectTeamAssignmentsButton.style.display = 'none';
                return;
            }

            projectTeamAssignmentsButton.style.display = (existing && existing.length > 0) ? '' : 'none';
        } catch (err) {
            console.warn('Unexpected error checking project team membership for assignments button visibility:', err);
            projectTeamAssignmentsButton.style.display = 'none';
        }
    }

    async function addSelectedProjectTeamMember() {
        if (!organizationId || !projectTeamStatus) return;
        if (!projectTeamProjectSelect || !projectTeamUserSelect) return;

        projectTeamStatus.classList.remove('upload-message--success', 'upload-message--error');

        const projectIdStr = projectTeamProjectSelect.value;
        const userIdStr = projectTeamUserSelect.value;

        if (!projectIdStr) {
            projectTeamStatus.textContent = 'Please select a project.';
            projectTeamStatus.classList.add('upload-message--error');
            return;
        }
        if (!userIdStr) {
            projectTeamStatus.textContent = 'Please select a user.';
            projectTeamStatus.classList.add('upload-message--error');
            return;
        }

        const project_id = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);
        const user_id = Number.isNaN(Number(userIdStr)) ? userIdStr : Number(userIdStr);

        const userInfo = projectTeamUsersById.get(String(userIdStr));
        const name = userInfo && userInfo.name ? userInfo.name : null;

        try {
            if (projectTeamAddButton) {
                projectTeamAddButton.disabled = true;
                projectTeamAddButton.textContent = 'Adding...';
            }

            // Prevent duplicates (same org + project + user)
            const { data: existing, error: existingErr } = await supabase
                .from('project_team_members')
                .select('id')
                .eq('organization_id', organizationId)
                .eq('project_id', project_id)
                .eq('user_id', user_id)
                .limit(1);

            if (existingErr) {
                console.warn('Duplicate check failed; proceeding with insert:', existingErr);
            } else if (existing && existing.length > 0) {
                projectTeamStatus.textContent = 'That user is already a team member for the selected project.';
                projectTeamStatus.classList.add('upload-message--error');
                await updateProjectTeamAssignmentsButtonVisibility();
                return;
            }

            const { error } = await supabase
                .from('project_team_members')
                .insert({
                    organization_id: organizationId,
                    project_id,
                    user_id,
                    name,
                    assignments: false
                });

            if (error) {
                console.error('Error inserting project team member:', error);
                projectTeamStatus.textContent = 'Failed to add team member.';
                projectTeamStatus.classList.add('upload-message--error');
                return;
            }

            projectTeamStatus.textContent = 'Team member added.';
            projectTeamStatus.classList.add('upload-message--success');
            await updateProjectTeamAssignmentsButtonVisibility();
        } catch (err) {
            console.error('Unexpected error adding team member:', err);
            projectTeamStatus.textContent = 'An unexpected error occurred while adding the team member.';
            projectTeamStatus.classList.add('upload-message--error');
        } finally {
            if (projectTeamAddButton) {
                projectTeamAddButton.disabled = false;
                projectTeamAddButton.textContent = 'Add Team Member';
            }
        }
    }

    function ensureOrgProjectDetailsManagePanel() {
        if (orgProjectDetailsManagePanel) return;

        const createView = document.querySelector('#createView');
        if (!createView) return;

        createView.insertAdjacentHTML('beforeend', `
            <section id="orgProjectDetailsManagePanel" class="project-types-panel" style="display: none; margin-top: 16px;">
                <div class="project-types-panel-header">
                    <div>
                        <h3>Manage Project Details</h3>
                        <p class="subtitle">View project detail parameters that have been uploaded for a project.</p>
                    </div>
                    <button type="button" id="orgProjectDetailsManageBackButton" class="secondary-button">Back</button>
                </div>
                <div class="form-group">
                    <label for="orgProjectDetailsManageProjectSelect">Project</label>
                    <select id="orgProjectDetailsManageProjectSelect">
                        <option value=\"\">Select Project</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="orgProjectDetailsManageSearchNameInput">Search (parameter name)</label>
                    <input
                        type="text"
                        id="orgProjectDetailsManageSearchNameInput"
                        placeholder="Type to filter parameter names..."
                        autocomplete="off"
                    >
                </div>
                <div class="form-group">
                    <label for="orgProjectDetailsManageSearchEntryInput">Search (parameter entry)</label>
                    <input
                        type="text"
                        id="orgProjectDetailsManageSearchEntryInput"
                        placeholder="Type to filter parameter entries..."
                        autocomplete="off"
                    >
                </div>
                <div class="form-buttons">
                    <button type="button" id="orgProjectDetailsManageRefreshButton" class="secondary-button">Refresh List</button>
                </div>
                <div id="orgProjectDetailsManageStatus" class="upload-message" aria-live="polite"></div>
                <div id="orgProjectDetailsManageTableWrap" style="margin-top: 12px;"></div>
            </section>
        `);

        orgProjectDetailsManagePanel = document.getElementById('orgProjectDetailsManagePanel');
        orgProjectDetailsManageProjectSelect = document.getElementById('orgProjectDetailsManageProjectSelect');
        orgProjectDetailsManageBackButton = document.getElementById('orgProjectDetailsManageBackButton');
        orgProjectDetailsManageRefreshButton = document.getElementById('orgProjectDetailsManageRefreshButton');
        orgProjectDetailsManageSearchNameInput = document.getElementById('orgProjectDetailsManageSearchNameInput');
        orgProjectDetailsManageSearchEntryInput = document.getElementById('orgProjectDetailsManageSearchEntryInput');
        orgProjectDetailsManageStatus = document.getElementById('orgProjectDetailsManageStatus');
        orgProjectDetailsManageTableWrap = document.getElementById('orgProjectDetailsManageTableWrap');

        if (orgProjectDetailsManageBackButton) {
            orgProjectDetailsManageBackButton.addEventListener('click', (e) => {
                e.preventDefault();
                hideOrgProjectDetailsManagePanel();
            });
        }

        if (orgProjectDetailsManageProjectSelect) {
            orgProjectDetailsManageProjectSelect.addEventListener('change', () => {
                refreshOrgProjectDetailsManageList();
            });
        }

        if (orgProjectDetailsManageRefreshButton) {
            orgProjectDetailsManageRefreshButton.addEventListener('click', (e) => {
                e.preventDefault();
                refreshOrgProjectDetailsManageList();
            });
        }

        if (orgProjectDetailsManageSearchNameInput) {
            orgProjectDetailsManageSearchNameInput.addEventListener('input', () => {
                renderOrgProjectDetailsManageTable();
            });
        }

        if (orgProjectDetailsManageSearchEntryInput) {
            orgProjectDetailsManageSearchEntryInput.addEventListener('input', () => {
                renderOrgProjectDetailsManageTable();
            });
        }
    }

    function showOrgProjectDetailsManagePanel() {
        ensureOrgProjectDetailsManagePanel();
        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = 'none';
        }
        if (projectTeamPanel) {
            projectTeamPanel.style.display = 'none';
        }
        if (projectTeamAssignmentsPanel) {
            projectTeamAssignmentsPanel.style.display = 'none';
        }
        if (orgProjectDetailsPanel) {
            orgProjectDetailsPanel.style.display = 'none';
        }
        if (orgProjectDetailsManagePanel) {
            orgProjectDetailsManagePanel.style.display = '';
            orgProjectDetailsManagePanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        loadOrgProjectsForProjectDetailsManageDropdown();
    }

    function hideOrgProjectDetailsManagePanel() {
        if (orgProjectDetailsManagePanel) {
            orgProjectDetailsManagePanel.style.display = 'none';
        }
        // Return to the Project Details submodule (still within Create view)
        showOrgProjectDetailsPanel();
    }

    function showLessonsMetadataPanel() {
        ensureLessonsMetadataPanel();
        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = 'none';
        }
        if (projectTeamPanel) {
            projectTeamPanel.style.display = 'none';
        }
        if (projectTeamAssignmentsPanel) {
            projectTeamAssignmentsPanel.style.display = 'none';
        }
        if (lessonsMetadataPanel) {
            lessonsMetadataPanel.style.display = '';
            lessonsMetadataPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        if (lessonsMetadataManagePanel) {
            lessonsMetadataManagePanel.style.display = 'none';
        }
        // Load projects into the dropdown when panel is shown
        loadOrgProjectsForLessonsDropdown();
    }

    function hideLessonsMetadataPanel() {
        if (lessonsMetadataPanel) {
            lessonsMetadataPanel.style.display = 'none';
        }
        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = '';
        }
    }

    function ensureLessonsMetadataManagePanel() {
        if (lessonsMetadataManagePanel) return;

        const createView = document.querySelector('#createView');
        if (!createView) return;

        createView.insertAdjacentHTML('beforeend', `
            <section id="lessonsMetadataManagePanel" class="project-types-panel" style="display: none; margin-top: 16px;">
                <div class="project-types-panel-header">
                    <div>
                        <h3>Manage Lessons Learned Metadata</h3>
                        <p class="subtitle">View the metadata entries that have been imported for a project.</p>
                    </div>
                    <button type="button" id="lessonsMetadataManageBackButton" class="secondary-button">Back</button>
                </div>
                <div class="form-group">
                    <label for="lessonsMetadataManageProjectSelect">Project</label>
                    <select id="lessonsMetadataManageProjectSelect">
                        <option value=\"\">Select Project</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="lessonsMetadataManageTypeSelect">Type (filter)</label>
                    <select id="lessonsMetadataManageTypeSelect">
                        <option value=\"\">All Types</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="lessonsMetadataManageSearchInput">Search (metadata)</label>
                    <input
                        type="text"
                        id="lessonsMetadataManageSearchInput"
                        placeholder="Type to filter metadata..."
                        autocomplete="off"
                    >
                </div>
                <div class="form-buttons">
                    <button type="button" id="lessonsMetadataManageRefreshButton" class="secondary-button">Refresh List</button>
                </div>
                <div id="lessonsMetadataManageStatus" class="upload-message" aria-live="polite"></div>
                <div id="lessonsMetadataManageTableWrap" style="margin-top: 12px;"></div>
            </section>
        `);

        lessonsMetadataManagePanel = document.getElementById('lessonsMetadataManagePanel');
        lessonsMetadataManageProjectSelect = document.getElementById('lessonsMetadataManageProjectSelect');
        lessonsMetadataManageBackButton = document.getElementById('lessonsMetadataManageBackButton');
        lessonsMetadataManageRefreshButton = document.getElementById('lessonsMetadataManageRefreshButton');
        lessonsMetadataManageTypeSelect = document.getElementById('lessonsMetadataManageTypeSelect');
        lessonsMetadataManageSearchInput = document.getElementById('lessonsMetadataManageSearchInput');
        lessonsMetadataManageStatus = document.getElementById('lessonsMetadataManageStatus');
        lessonsMetadataManageTableWrap = document.getElementById('lessonsMetadataManageTableWrap');

        if (lessonsMetadataManageBackButton) {
            lessonsMetadataManageBackButton.addEventListener('click', (e) => {
                e.preventDefault();
                hideLessonsMetadataManagePanel();
            });
        }

        if (lessonsMetadataManageProjectSelect) {
            lessonsMetadataManageProjectSelect.addEventListener('change', () => {
                refreshLessonsMetadataManageList();
            });
        }

        if (lessonsMetadataManageRefreshButton) {
            lessonsMetadataManageRefreshButton.addEventListener('click', (e) => {
                e.preventDefault();
                refreshLessonsMetadataManageList();
            });
        }

        if (lessonsMetadataManageTypeSelect) {
            lessonsMetadataManageTypeSelect.addEventListener('change', () => {
                renderLessonsMetadataManageTable();
            });
        }

        if (lessonsMetadataManageSearchInput) {
            lessonsMetadataManageSearchInput.addEventListener('input', () => {
                renderLessonsMetadataManageTable();
            });
        }
    }

    function showLessonsMetadataManagePanel() {
        ensureLessonsMetadataManagePanel();
        const createColumns = document.getElementById('createColumns');
        if (createColumns) {
            createColumns.style.display = 'none';
        }
        if (projectTeamPanel) {
            projectTeamPanel.style.display = 'none';
        }
        if (projectTeamAssignmentsPanel) {
            projectTeamAssignmentsPanel.style.display = 'none';
        }
        if (lessonsMetadataPanel) {
            lessonsMetadataPanel.style.display = 'none';
        }
        if (lessonsMetadataManagePanel) {
            lessonsMetadataManagePanel.style.display = '';
            lessonsMetadataManagePanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        loadOrgProjectsForLessonsManageDropdown();
    }

    function hideLessonsMetadataManagePanel() {
        if (lessonsMetadataManagePanel) {
            lessonsMetadataManagePanel.style.display = 'none';
        }
        // Return to the Lessons Learned Metadata panel (still within Create view)
        showLessonsMetadataPanel();
    }

    async function loadOrgProjectsForLessonsManageDropdown() {
        if (!organizationId || !lessonsMetadataManageProjectSelect || !lessonsMetadataManageStatus) return;

        lessonsMetadataManageStatus.classList.remove('upload-message--success', 'upload-message--error');
        lessonsMetadataManageStatus.textContent = 'Loading projects...';

        try {
            const { data: projectRows, error: projectErr } = await supabase
                .from('projects')
                .select('project_id, project_name, project_type_id')
                .eq('organization_id', organizationId)
                .order('project_name', { ascending: true });

            if (projectErr) {
                console.error('Error loading organization projects for manage lessons metadata dropdown:', projectErr);
                lessonsMetadataManageStatus.textContent = 'Failed to load projects.';
                lessonsMetadataManageStatus.classList.add('upload-message--error');
                return;
            }

            const rows = projectRows || [];
            if (rows.length === 0) {
                lessonsMetadataManageStatus.textContent = 'No projects found for your organization.';
                lessonsMetadataManageStatus.classList.add('upload-message--error');
                if (lessonsMetadataManageChoices) {
                    lessonsMetadataManageChoices.clearChoices();
                } else {
                    lessonsMetadataManageProjectSelect.innerHTML = '<option value=\"\">No projects available</option>';
                }
                return;
            }

            const choicesData = rows.map(row => {
                const idStr = String(row.project_id);
                const name = row.project_name || `Project ${idStr}`;
                const typeId = row.project_type_id != null ? String(row.project_type_id) : null;
                lessonsProjectsById.set(idStr, { name, project_type_id: typeId });
                return { value: idStr, label: name };
            });

            if (!lessonsMetadataManageChoices) {
                if (typeof Choices === 'undefined') {
                    console.warn('Choices library not loaded; falling back to native select for manage lessons metadata dropdown.');
                    lessonsMetadataManageProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                    choicesData.forEach(choice => {
                        const option = document.createElement('option');
                        option.value = choice.value;
                        option.textContent = choice.label;
                        lessonsMetadataManageProjectSelect.appendChild(option);
                    });
                } else {
                    lessonsMetadataManageChoices = new Choices(lessonsMetadataManageProjectSelect, {
                        searchEnabled: true,
                        shouldSort: false,
                        placeholder: true,
                        placeholderValue: 'Select Project',
                        searchPlaceholderValue: 'Type to search...'
                    });
                }
            }

            if (lessonsMetadataManageChoices) {
                lessonsMetadataManageChoices.setChoices(choicesData, 'value', 'label', true);
            } else {
                lessonsMetadataManageProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                choicesData.forEach(choice => {
                    const option = document.createElement('option');
                    option.value = choice.value;
                    option.textContent = choice.label;
                    lessonsMetadataManageProjectSelect.appendChild(option);
                });
            }

            lessonsMetadataManageStatus.textContent = '';
            lessonsMetadataManageStatus.classList.remove('upload-message--error');
        } catch (err) {
            console.error('Unexpected error loading organization projects for manage lessons metadata dropdown:', err);
            lessonsMetadataManageStatus.textContent = 'An unexpected error occurred while loading projects.';
            lessonsMetadataManageStatus.classList.add('upload-message--error');
        }
    }

    function updateLessonsMetadataManageTypeOptions(rows) {
        if (!lessonsMetadataManageTypeSelect) return;

        const previous = lessonsMetadataManageTypeSelect.value || '';
        const types = Array.from(
            new Set((rows || []).map(r => r && r.metadata_type).filter(Boolean))
        ).sort((a, b) => String(a).localeCompare(String(b)));

        lessonsMetadataManageTypeSelect.innerHTML = '';

        const allOpt = document.createElement('option');
        allOpt.value = '';
        allOpt.textContent = 'All Types';
        lessonsMetadataManageTypeSelect.appendChild(allOpt);

        types.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t;
            opt.textContent = t;
            lessonsMetadataManageTypeSelect.appendChild(opt);
        });

        // Restore previous value if it still exists; otherwise default to All
        lessonsMetadataManageTypeSelect.value = types.includes(previous) ? previous : '';
    }

    function renderLessonsMetadataManageTable() {
        if (!lessonsMetadataManageTableWrap) return;

        const selectedType = lessonsMetadataManageTypeSelect ? lessonsMetadataManageTypeSelect.value : '';
        const rows = Array.isArray(lessonsMetadataManageRowsCache) ? lessonsMetadataManageRowsCache : [];

        const searchRaw = lessonsMetadataManageSearchInput ? lessonsMetadataManageSearchInput.value : '';
        const search = String(searchRaw || '').trim().toLowerCase();

        let filtered = rows;

        if (selectedType) {
            filtered = filtered.filter(r => r && r.metadata_type === selectedType);
        }

        if (search) {
            filtered = filtered.filter(r => String((r && r.metadata) || '').toLowerCase().includes(search));
        }

        lessonsMetadataManageTableWrap.innerHTML = '';

        if (filtered.length === 0) {
            const empty = document.createElement('div');
            const searchDisplay = String(searchRaw || '').trim();
            empty.textContent =
                selectedType && searchDisplay
                    ? `No metadata entries match type "${selectedType}" and "${searchDisplay}".`
                    : searchDisplay
                        ? `No metadata entries match "${searchDisplay}".`
                        : selectedType
                            ? `No metadata entries found for type "${selectedType}".`
                            : 'No metadata has been imported for this project yet.';
            lessonsMetadataManageTableWrap.appendChild(empty);

            if (lessonsMetadataManageStatus) {
                lessonsMetadataManageStatus.classList.remove('upload-message--success', 'upload-message--error');
                if (selectedType || search) {
                    lessonsMetadataManageStatus.textContent = selectedType
                        ? `Showing 0 ${search ? 'matching ' : ''}"${selectedType}" entries (of ${rows.length} total).`
                        : `Showing 0 matching entries (of ${rows.length} total).`;
                } else {
                    lessonsMetadataManageStatus.textContent = `Loaded ${rows.length} metadata entr${rows.length === 1 ? 'y' : 'ies'}.`;
                }
                if (rows.length > 0) lessonsMetadataManageStatus.classList.add('upload-message--success');
            }

            return;
        }

        const table = document.createElement('table');
        table.className = 'organizations-table';

        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        ['Type', 'Metadata'].forEach(label => {
            const th = document.createElement('th');
            th.textContent = label;
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        filtered.forEach(row => {
            const tr = document.createElement('tr');

            const tdType = document.createElement('td');
            tdType.textContent = row.metadata_type || '';
            tr.appendChild(tdType);

            const tdMetadata = document.createElement('td');
            tdMetadata.textContent = row.metadata || '';
            tr.appendChild(tdMetadata);

            tbody.appendChild(tr);
        });
        table.appendChild(tbody);

        lessonsMetadataManageTableWrap.appendChild(table);

        if (lessonsMetadataManageStatus) {
            lessonsMetadataManageStatus.classList.remove('upload-message--success', 'upload-message--error');
            if (selectedType || search) {
                const typePart = selectedType ? `"${selectedType}" ` : '';
                const matchPart = search ? 'matching ' : '';
                lessonsMetadataManageStatus.textContent =
                    `Showing ${filtered.length} ${matchPart}${typePart}entr${filtered.length === 1 ? 'y' : 'ies'} (of ${rows.length} total).`;
            } else {
                lessonsMetadataManageStatus.textContent = `Loaded ${rows.length} metadata entr${rows.length === 1 ? 'y' : 'ies'}.`;
            }
            lessonsMetadataManageStatus.classList.add('upload-message--success');
        }
    }

    async function refreshLessonsMetadataManageList() {
        if (!organizationId || !lessonsMetadataManageProjectSelect || !lessonsMetadataManageStatus || !lessonsMetadataManageTableWrap) return;

        const projectIdStr = lessonsMetadataManageProjectSelect.value;
        if (!projectIdStr) {
            lessonsMetadataManageTableWrap.innerHTML = '';
            lessonsMetadataManageStatus.textContent = 'Please select a project.';
            lessonsMetadataManageStatus.classList.remove('upload-message--success');
            lessonsMetadataManageStatus.classList.add('upload-message--error');
            lessonsMetadataManageRowsCache = [];
            if (lessonsMetadataManageTypeSelect) lessonsMetadataManageTypeSelect.value = '';
            if (lessonsMetadataManageSearchInput) lessonsMetadataManageSearchInput.value = '';
            return;
        }

        const projectId = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);
        lessonsMetadataManageStatus.classList.remove('upload-message--success', 'upload-message--error');
        lessonsMetadataManageStatus.textContent = 'Loading metadata...';
        lessonsMetadataManageTableWrap.innerHTML = '';

        try {
            const { data: rows, error } = await supabase
                .from('lessons_learned_metadata_list')
                .select('id, metadata_type, metadata')
                .eq('organization_id', organizationId)
                .eq('project_id', projectId)
                .order('id', { ascending: true })
                .limit(5000);

            if (error) {
                console.error('Error loading lessons learned metadata list:', error);
                lessonsMetadataManageStatus.textContent = 'Failed to load metadata.';
                lessonsMetadataManageStatus.classList.add('upload-message--error');
                lessonsMetadataManageRowsCache = [];
                return;
            }

            const data = rows || [];
            lessonsMetadataManageRowsCache = data;
            updateLessonsMetadataManageTypeOptions(data);
            renderLessonsMetadataManageTable();
        } catch (err) {
            console.error('Unexpected error loading lessons learned metadata list:', err);
            lessonsMetadataManageStatus.textContent = 'An unexpected error occurred while loading metadata.';
            lessonsMetadataManageStatus.classList.add('upload-message--error');
            lessonsMetadataManageRowsCache = [];
        }
    }

    async function loadOrgProjectsForProjectDetailsManageDropdown() {
        if (!organizationId || !orgProjectDetailsManageProjectSelect || !orgProjectDetailsManageStatus) return;

        orgProjectDetailsManageStatus.classList.remove('upload-message--success', 'upload-message--error');
        orgProjectDetailsManageStatus.textContent = 'Loading projects...';

        try {
            const { data: projectRows, error: projectErr } = await supabase
                .from('projects')
                .select('project_id, project_name, project_type_id')
                .eq('organization_id', organizationId)
                .order('project_name', { ascending: true });

            if (projectErr) {
                console.error('Error loading organization projects for manage project details dropdown:', projectErr);
                orgProjectDetailsManageStatus.textContent = 'Failed to load projects.';
                orgProjectDetailsManageStatus.classList.add('upload-message--error');
                return;
            }

            const rows = projectRows || [];
            if (rows.length === 0) {
                orgProjectDetailsManageStatus.textContent = 'No projects found for your organization.';
                orgProjectDetailsManageStatus.classList.add('upload-message--error');
                if (orgProjectDetailsManageChoices) {
                    orgProjectDetailsManageChoices.clearChoices();
                } else {
                    orgProjectDetailsManageProjectSelect.innerHTML = '<option value=\"\">No projects available</option>';
                }
                return;
            }

            const choicesData = rows.map(row => {
                const idStr = String(row.project_id);
                const name = row.project_name || `Project ${idStr}`;
                const typeId = row.project_type_id != null ? String(row.project_type_id) : null;
                orgProjectsById.set(idStr, { name, project_type_id: typeId });
                return { value: idStr, label: name };
            });

            if (!orgProjectDetailsManageChoices) {
                if (typeof Choices === 'undefined') {
                    console.warn('Choices library not loaded; falling back to native select for manage project details dropdown.');
                    orgProjectDetailsManageProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                    choicesData.forEach(choice => {
                        const option = document.createElement('option');
                        option.value = choice.value;
                        option.textContent = choice.label;
                        orgProjectDetailsManageProjectSelect.appendChild(option);
                    });
                } else {
                    orgProjectDetailsManageChoices = new Choices(orgProjectDetailsManageProjectSelect, {
                        searchEnabled: true,
                        shouldSort: false,
                        placeholder: true,
                        placeholderValue: 'Select Project',
                        searchPlaceholderValue: 'Type to search...'
                    });
                }
            }

            if (orgProjectDetailsManageChoices) {
                orgProjectDetailsManageChoices.setChoices(choicesData, 'value', 'label', true);
            } else {
                orgProjectDetailsManageProjectSelect.innerHTML = '<option value=\"\">Select Project</option>';
                choicesData.forEach(choice => {
                    const option = document.createElement('option');
                    option.value = choice.value;
                    option.textContent = choice.label;
                    orgProjectDetailsManageProjectSelect.appendChild(option);
                });
            }

            orgProjectDetailsManageStatus.textContent = '';
            orgProjectDetailsManageStatus.classList.remove('upload-message--error');
        } catch (err) {
            console.error('Unexpected error loading organization projects for manage project details dropdown:', err);
            orgProjectDetailsManageStatus.textContent = 'An unexpected error occurred while loading projects.';
            orgProjectDetailsManageStatus.classList.add('upload-message--error');
        }
    }

    function renderOrgProjectDetailsManageTable() {
        if (!orgProjectDetailsManageTableWrap) return;

        const rows = Array.isArray(orgProjectDetailsManageRowsCache) ? orgProjectDetailsManageRowsCache : [];

        const nameRaw = orgProjectDetailsManageSearchNameInput ? orgProjectDetailsManageSearchNameInput.value : '';
        const entryRaw = orgProjectDetailsManageSearchEntryInput ? orgProjectDetailsManageSearchEntryInput.value : '';
        const nameFilter = String(nameRaw || '').trim().toLowerCase();
        const entryFilter = String(entryRaw || '').trim().toLowerCase();

        let filtered = rows;

        if (nameFilter) {
            filtered = filtered.filter(r => String((r && r.parameter_name) || '').toLowerCase().includes(nameFilter));
        }
        if (entryFilter) {
            filtered = filtered.filter(r => String((r && r.parameter_entry) || '').toLowerCase().includes(entryFilter));
        }

        orgProjectDetailsManageTableWrap.innerHTML = '';

        if (filtered.length === 0) {
            const empty = document.createElement('div');
            if (rows.length === 0) {
                empty.textContent = 'No project details have been uploaded for this project yet.';
            } else {
                empty.textContent = 'No rows match your filters.';
            }
            orgProjectDetailsManageTableWrap.appendChild(empty);

            if (orgProjectDetailsManageStatus) {
                orgProjectDetailsManageStatus.classList.remove('upload-message--success', 'upload-message--error');
                if (nameFilter || entryFilter) {
                    orgProjectDetailsManageStatus.textContent = `Showing 0 matching rows (of ${rows.length} total).`;
                } else {
                    orgProjectDetailsManageStatus.textContent = `Loaded ${rows.length} project detail entr${rows.length === 1 ? 'y' : 'ies'}.`;
                }
                if (rows.length > 0) orgProjectDetailsManageStatus.classList.add('upload-message--success');
            }
            return;
        }

        const table = document.createElement('table');
        table.className = 'organizations-table';

        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        ['Parameter Name', 'Parameter Entry'].forEach(label => {
            const th = document.createElement('th');
            th.textContent = label;
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        filtered.forEach(row => {
            const tr = document.createElement('tr');

            const tdName = document.createElement('td');
            tdName.textContent = row.parameter_name || '';
            tr.appendChild(tdName);

            const tdEntry = document.createElement('td');
            tdEntry.textContent = row.parameter_entry || '';
            tr.appendChild(tdEntry);

            tbody.appendChild(tr);
        });
        table.appendChild(tbody);

        orgProjectDetailsManageTableWrap.appendChild(table);

        if (orgProjectDetailsManageStatus) {
            orgProjectDetailsManageStatus.classList.remove('upload-message--success', 'upload-message--error');
            if (nameFilter || entryFilter) {
                orgProjectDetailsManageStatus.textContent =
                    `Showing ${filtered.length} matching row${filtered.length === 1 ? '' : 's'} (of ${rows.length} total).`;
            } else {
                orgProjectDetailsManageStatus.textContent = `Loaded ${rows.length} project detail entr${rows.length === 1 ? 'y' : 'ies'}.`;
            }
            orgProjectDetailsManageStatus.classList.add('upload-message--success');
        }
    }

    async function refreshOrgProjectDetailsManageList() {
        if (!organizationId || !orgProjectDetailsManageProjectSelect || !orgProjectDetailsManageStatus || !orgProjectDetailsManageTableWrap) return;

        const projectIdStr = orgProjectDetailsManageProjectSelect.value;
        if (!projectIdStr) {
            orgProjectDetailsManageTableWrap.innerHTML = '';
            orgProjectDetailsManageStatus.textContent = 'Please select a project.';
            orgProjectDetailsManageStatus.classList.remove('upload-message--success');
            orgProjectDetailsManageStatus.classList.add('upload-message--error');
            orgProjectDetailsManageRowsCache = [];
            if (orgProjectDetailsManageSearchNameInput) orgProjectDetailsManageSearchNameInput.value = '';
            if (orgProjectDetailsManageSearchEntryInput) orgProjectDetailsManageSearchEntryInput.value = '';
            return;
        }

        const projectId = Number.isNaN(Number(projectIdStr)) ? projectIdStr : Number(projectIdStr);
        orgProjectDetailsManageStatus.classList.remove('upload-message--success', 'upload-message--error');
        orgProjectDetailsManageStatus.textContent = 'Loading project details...';
        orgProjectDetailsManageTableWrap.innerHTML = '';

        try {
            const { data: rows, error } = await supabase
                .from('project_details')
                .select('id, parameter_name, parameter_entry')
                .eq('organization_id', organizationId)
                .eq('project_id', projectId)
                .order('parameter_name', { ascending: true })
                .limit(5000);

            if (error) {
                console.error('Error loading project details list:', error);
                orgProjectDetailsManageStatus.textContent = 'Failed to load project details.';
                orgProjectDetailsManageStatus.classList.add('upload-message--error');
                orgProjectDetailsManageRowsCache = [];
                return;
            }

            const data = rows || [];
            orgProjectDetailsManageRowsCache = data;
            renderOrgProjectDetailsManageTable();
        } catch (err) {
            console.error('Unexpected error loading project details list:', err);
            orgProjectDetailsManageStatus.textContent = 'An unexpected error occurred while loading project details.';
            orgProjectDetailsManageStatus.classList.add('upload-message--error');
            orgProjectDetailsManageRowsCache = [];
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

    function wireCreateProjectButtons(root = document) {
        const buttons = root.querySelectorAll('.create-project-button');
        buttons.forEach((btn) => {
            if (!btn) return;
            if (!canCreateProjects) {
                btn.style.display = 'none';
                return;
            }
            btn.style.display = '';
            if (btn.dataset && btn.dataset.wired === 'true') return;
            btn.addEventListener('click', () => {
                // Check permissions before allowing navigation
                if (!canCreateProjects) {
                    alert('You do not have permission to create new projects. Please contact your administrator.');
                    return;
                }
                if (confirmNavigation("Create New Project")) {
                    location.hash = 'create';
                }
            });
            if (btn.dataset) btn.dataset.wired = 'true';
        });
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

    wireCreateProjectButtons(document);
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
                    : text === 'Search Projects' ? 'search-projects'
                    : text === 'General Search' ? 'general-search'
                    : 'home';
        location.hash = route;
    });

    // Simple router
    function setActiveSidebar(route) {
        // Backward compatibility for older hashes like "#search"
        const normalizedRoute = route === 'search' ? 'search-projects' : route;
        const links = document.querySelectorAll('.sidebar .nav-item');
        links.forEach(l => {
            const text = l.textContent.trim();
            const r = text === 'Organization Settings' ? 'org'
                    : text === 'My Projects' ? 'projects'
                    : text === 'Search Projects' ? 'search-projects'
                    : text === 'General Search' ? 'general-search'
                    : 'home';
            if (r === normalizedRoute) l.classList.add('active'); else l.classList.remove('active');
            l.setAttribute('aria-current', r === normalizedRoute ? 'page' : 'false');
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

    function showSearchProjectsView() {
        showSearchView();
        renderSearchProjectsModule();
    }

    function showGeneralSearchView() {
        showSearchView();
        renderGeneralSearchModule();
    }

    function renderSearchProjectsModule() {
        const searchView = document.getElementById('searchView');
        if (!searchView) return;

        // Re-render from scratch each time we enter this route.
        searchView.innerHTML = `
            <h1>Search Projects</h1>
            <div id="searchProjectsMain">
                <div id="searchProjectsPanel" class="project-types-panel">
                    <p class="subtitle" style="margin: 0 0 10px;">Search by project name</p>
                    <div class="form-group" style="margin-bottom: 4px;">
                        <input
                            id="searchProjectsInput"
                            type="text"
                            placeholder="Search by project name"
                            autocomplete="off"
                            aria-label="Search projects by project name"
                        >
                    </div>
                </div>

                <div id="searchProjectsListStatus" class="search-status" aria-live="polite"></div>
                <div id="searchProjectsListWrap">
                    <ul id="searchProjectsList" class="search-projects-list"></ul>
                </div>

                <div id="searchProjectsStatus" class="search-status" aria-live="polite"></div>
                <div id="searchProjectsResults" class="lessons-results"></div>
            </div>

            <div id="searchProjectsDetailsScreen" style="display: none;">
                <div class="search-projects-details-layout">
                    <div id="searchProjectsDetailsPanel" class="project-types-panel">
                        <div class="project-types-panel-header">
                            <div>
                                <h3>Project Info</h3>
                                <p class="subtitle">Project summary information.</p>
                            </div>
                            <button type="button" id="searchProjectsDetailsBack" class="secondary-button">Go Back</button>
                        </div>
                        <div class="form-group">
                            <label>Project Name</label>
                            <div id="searchProjectsDetailsName"></div>
                        </div>
                        <div class="form-group">
                            <label>Project Description</label>
                            <div id="searchProjectsDetailsDescription"></div>
                        </div>
                        <div class="form-group">
                            <label>Project Type</label>
                            <div id="searchProjectsDetailsType"></div>
                        </div>
                    </div>
                    <div class="search-projects-details-actions">
                        <button type="button" id="searchProjectsDetailsAction" class="side-button">Project Details</button>
                        <button type="button" id="searchProjectsTeamAction" class="side-button">Project Team</button>
                        <button type="button" id="searchProjectsLessonsAction" class="side-button">Lessons Learned Metadata</button>
                    </div>
                </div>
            </div>

            <div id="searchProjectsMetadataPanel" class="project-types-panel" style="display: none;">
                <div class="project-types-panel-header">
                    <div>
                        <h3 id="searchProjectsMetadataTitle">Lessons Learned Metadata</h3>
                        <p class="subtitle">View the metadata entries that have been imported for this project.</p>
                    </div>
                    <button type="button" id="searchProjectsMetadataBackButton" class="secondary-button">Back</button>
                </div>
                <div class="form-group">
                    <label for="searchProjectsMetadataTypeSelect">Type (filter)</label>
                    <select id="searchProjectsMetadataTypeSelect">
                        <option value="">All Types</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="searchProjectsMetadataSearchInput">Search (metadata)</label>
                    <input
                        type="text"
                        id="searchProjectsMetadataSearchInput"
                        placeholder="Type to filter metadata..."
                        autocomplete="off"
                    >
                </div>
                <div class="form-buttons">
                    <button type="button" id="searchProjectsMetadataRefreshButton" class="secondary-button">Refresh List</button>
                </div>
                <div id="searchProjectsMetadataStatus" class="upload-message" aria-live="polite"></div>
                <div id="searchProjectsMetadataTableWrap" style="margin-top: 12px;"></div>
            </div>

            <div id="searchProjectsTaskDetailsPanel" class="project-types-panel" style="display: none;">
                <div class="project-types-panel-header">
                    <div>
                        <h3 id="searchProjectsTaskDetailsTitle">Task Details</h3>
                    </div>
                    <button type="button" id="searchProjectsTaskDetailsBackButton" class="secondary-button">Back</button>
                </div>
                <div id="searchProjectsTaskDetailsStatus" class="upload-message" aria-live="polite"></div>
                <div id="searchProjectsTaskDetailsBody"></div>
            </div>
        `;

        const input = document.getElementById('searchProjectsInput');
        const MIN_TERM_LENGTH = 2;
        const DEBOUNCE_MS = 350;
        let debounceTimer = null;
        let activeSearchController = null;

        const clearResults = () => {
            setSearchProjectsStatus('');
            renderSearchProjectsResults([], '');
        };

        const runSearch = async (rawTerm) => {
            const term = (rawTerm || '').trim();
            if (!term) {
                clearResults();
                return;
            }
            if (term.length < MIN_TERM_LENGTH) {
                clearResults();
                setSearchProjectsStatus(`Type at least ${MIN_TERM_LENGTH} characters to search.`);
                return;
            }

            // Cancel any in-flight search
            if (activeSearchController) {
                try { activeSearchController.abort(); } catch (_) {}
            }
            activeSearchController = new AbortController();

            setSearchProjectsStatus('Searching projects…');
            renderSearchProjectsResults([], '');

            try {
                const results = await searchProjectsApi(term, activeSearchController.signal);
                renderSearchProjectsResults(results, term);
            } catch (err) {
                // Ignore aborted requests (user kept typing)
                if (err && (err.name === 'AbortError' || String(err.message || '').includes('aborted'))) {
                    return;
                }
                console.error('Project search error:', err);
                setSearchProjectsStatus('Something went wrong while searching projects. Please try again.');
                renderSearchProjectsResults([], '');
            }
        };

        if (input) {
            input.addEventListener('input', () => {
                const term = input.value || '';
                updateSearchProjectsNameList(term);
                if (debounceTimer) clearTimeout(debounceTimer);
                debounceTimer = setTimeout(() => runSearch(term), DEBOUNCE_MS);
            });

            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    if (debounceTimer) clearTimeout(debounceTimer);
                    runSearch(input.value || '');
                } else if (e.key === 'Escape') {
                    e.preventDefault();
                    input.value = '';
                    updateSearchProjectsNameList('');
                    clearResults();
                }
            });
        }

        // Autofocus when entering the module
        if (input) {
            setTimeout(() => input.focus(), 0);
        }

        // Load organization projects list beneath the search bar.
        loadOrganizationProjectsForSearch();

        const backBtn = document.getElementById('searchProjectsDetailsBack');
        if (backBtn) {
            backBtn.addEventListener('click', () => {
                hideSearchProjectsDetails();
            });
        }

        const lessonsBtn = document.getElementById('searchProjectsLessonsAction');
        if (lessonsBtn) {
            lessonsBtn.addEventListener('click', () => {
                showSearchProjectsMetadataPanel();
            });
        }

        const detailsBtn = document.getElementById('searchProjectsDetailsAction');
        if (detailsBtn) {
            detailsBtn.addEventListener('click', () => {
                hideSearchProjectsMetadataPanel();
            });
        }

        const metadataBackBtn = document.getElementById('searchProjectsMetadataBackButton');
        if (metadataBackBtn) {
            metadataBackBtn.addEventListener('click', () => {
                hideSearchProjectsMetadataPanel();
            });
        }

        const metadataTypeSelect = document.getElementById('searchProjectsMetadataTypeSelect');
        if (metadataTypeSelect) {
            metadataTypeSelect.addEventListener('change', () => {
                renderSearchProjectsMetadataTable();
            });
        }

        const metadataSearchInput = document.getElementById('searchProjectsMetadataSearchInput');
        if (metadataSearchInput) {
            metadataSearchInput.addEventListener('input', () => {
                renderSearchProjectsMetadataTable();
            });
        }

        const metadataRefreshBtn = document.getElementById('searchProjectsMetadataRefreshButton');
        if (metadataRefreshBtn) {
            metadataRefreshBtn.addEventListener('click', (e) => {
                e.preventDefault();
                refreshSearchProjectsMetadataList();
            });
        }

        const taskDetailsBackBtn = document.getElementById('searchProjectsTaskDetailsBackButton');
        if (taskDetailsBackBtn) {
            taskDetailsBackBtn.addEventListener('click', () => {
                hideSearchProjectsTaskDetails();
            });
        }
    }

    let searchProjectsCache = [];
    let searchProjectsCacheLoaded = false;
    let searchProjectsCacheLoading = false;
    let searchProjectsSelectedProject = null;
    let searchProjectsMetadataRowsCache = [];

    async function loadOrganizationProjectsForSearch() {
        if (!organizationId) {
            setSearchProjectsListStatus('No organization found for this user.');
            return;
        }
        if (searchProjectsCacheLoaded || searchProjectsCacheLoading) {
            updateSearchProjectsNameList(document.getElementById('searchProjectsInput')?.value || '');
            return;
        }

        searchProjectsCacheLoading = true;
        setSearchProjectsListStatus('Loading projects…');

        try {
            const { data, error } = await supabase
                .from('projects')
                .select('project_id, project_name, project_description, project_type_id, project_type:project_type_id (project_type)')
                .eq('organization_id', organizationId)
                .order('project_name', { ascending: true });

            if (error) {
                console.error('Error loading organization projects for search list:', error);
                setSearchProjectsListStatus('Failed to load projects.');
                searchProjectsCache = [];
                searchProjectsCacheLoaded = false;
                searchProjectsCacheLoading = false;
                updateSearchProjectsNameList('');
                return;
            }

            searchProjectsCache = Array.isArray(data) ? data : [];
            searchProjectsCacheLoaded = true;
            searchProjectsCacheLoading = false;
            updateSearchProjectsNameList('');
        } catch (err) {
            console.error('Unexpected error loading organization projects for search list:', err);
            setSearchProjectsListStatus('An unexpected error occurred while loading projects.');
            searchProjectsCache = [];
            searchProjectsCacheLoaded = false;
            searchProjectsCacheLoading = false;
            updateSearchProjectsNameList('');
        }
    }

    function setSearchProjectsListStatus(message) {
        const el = document.getElementById('searchProjectsListStatus');
        if (el) el.textContent = message || '';
    }

    function updateSearchProjectsNameList(term) {
        const listEl = document.getElementById('searchProjectsList');
        if (!listEl) return;

        const normalized = (term || '').trim().toLowerCase();
        const projects = Array.isArray(searchProjectsCache) ? searchProjectsCache : [];

        let filtered = projects;
        if (normalized) {
            filtered = projects.filter((p) => {
                const name = p && p.project_name ? String(p.project_name).toLowerCase() : '';
                return name.includes(normalized);
            });
        }

        listEl.innerHTML = '';

        if (!projects.length) {
            setSearchProjectsListStatus('No projects found for your organization.');
            return;
        }

        if (normalized && filtered.length === 0) {
            setSearchProjectsListStatus('No matching projects.');
            return;
        }

        setSearchProjectsListStatus('');

        filtered.forEach((p) => {
            const li = document.createElement('li');
            li.className = 'search-projects-list-item';
            const nameSpan = document.createElement('span');
            nameSpan.className = 'search-projects-list-name';
            nameSpan.textContent = p && p.project_name ? p.project_name : '(Unnamed Project)';

            const detailsBtn = document.createElement('button');
            detailsBtn.type = 'button';
            detailsBtn.className = 'secondary-button search-projects-details-button';
            detailsBtn.textContent = 'Project Info';
            detailsBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                showSearchProjectsDetails(p);
            });

            li.appendChild(nameSpan);
            li.appendChild(detailsBtn);
            listEl.appendChild(li);
        });
    }

    function updateSearchProjectsMetadataTypeOptions(rows) {
        const typeSelect = document.getElementById('searchProjectsMetadataTypeSelect');
        if (!typeSelect) return;

        const previous = typeSelect.value || '';
        const types = Array.from(
            new Set((rows || []).map(r => r && r.metadata_type).filter(Boolean))
        ).sort((a, b) => String(a).localeCompare(String(b)));

        typeSelect.innerHTML = '';

        const allOpt = document.createElement('option');
        allOpt.value = '';
        allOpt.textContent = 'All Types';
        typeSelect.appendChild(allOpt);

        types.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t;
            opt.textContent = t;
            typeSelect.appendChild(opt);
        });

        typeSelect.value = types.includes(previous) ? previous : '';
    }

    function renderSearchProjectsMetadataTable() {
        const tableWrap = document.getElementById('searchProjectsMetadataTableWrap');
        const typeSelect = document.getElementById('searchProjectsMetadataTypeSelect');
        const searchInput = document.getElementById('searchProjectsMetadataSearchInput');
        const statusEl = document.getElementById('searchProjectsMetadataStatus');
        if (!tableWrap) return;

        const selectedType = typeSelect ? typeSelect.value : '';
        const rows = Array.isArray(searchProjectsMetadataRowsCache) ? searchProjectsMetadataRowsCache : [];
        const searchRaw = searchInput ? searchInput.value : '';
        const search = String(searchRaw || '').trim().toLowerCase();

        let filtered = rows;
        if (selectedType) {
            filtered = filtered.filter(r => r && r.metadata_type === selectedType);
        }
        if (search) {
            filtered = filtered.filter(r => String((r && r.metadata) || '').toLowerCase().includes(search));
        }

        tableWrap.innerHTML = '';

        if (filtered.length === 0) {
            const empty = document.createElement('div');
            const searchDisplay = String(searchRaw || '').trim();
            empty.textContent =
                selectedType && searchDisplay
                    ? `No metadata entries match type "${selectedType}" and "${searchDisplay}".`
                    : searchDisplay
                        ? `No metadata entries match "${searchDisplay}".`
                        : selectedType
                            ? `No metadata entries found for type "${selectedType}".`
                            : 'No metadata has been imported for this project yet.';
            tableWrap.appendChild(empty);

            if (statusEl) {
                statusEl.classList.remove('upload-message--success', 'upload-message--error');
                if (selectedType || search) {
                    statusEl.textContent = selectedType
                        ? `Showing 0 ${search ? 'matching ' : ''}"${selectedType}" entries (of ${rows.length} total).`
                        : `Showing 0 matching entries (of ${rows.length} total).`;
                } else {
                    statusEl.textContent = `Loaded ${rows.length} metadata entr${rows.length === 1 ? 'y' : 'ies'}.`;
                }
                if (rows.length > 0) statusEl.classList.add('upload-message--success');
            }
            return;
        }

        const table = document.createElement('table');
        table.className = 'organizations-table';

        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        ['Type', 'Metadata'].forEach(label => {
            const th = document.createElement('th');
            th.textContent = label;
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        filtered.forEach(row => {
            const tr = document.createElement('tr');
            const metadataType = row && row.metadata_type ? String(row.metadata_type).toLowerCase() : '';
            const metadataSource = row && row.metadata_source ? String(row.metadata_source).toLowerCase() : '';
            const isTask = metadataType === 'task' && metadataSource === 'ms project';
            if (isTask) {
                tr.classList.add('metadata-row-clickable');
                tr.addEventListener('click', () => {
                    showSearchProjectsTaskDetails(row);
                });
            }

            const tdType = document.createElement('td');
            tdType.textContent = row.metadata_type || '';
            tr.appendChild(tdType);

            const tdMetadata = document.createElement('td');
            tdMetadata.textContent = row.metadata || '';
            tr.appendChild(tdMetadata);

            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        tableWrap.appendChild(table);

        if (statusEl) {
            statusEl.classList.remove('upload-message--success', 'upload-message--error');
            if (selectedType || search) {
                const typePart = selectedType ? `"${selectedType}" ` : '';
                const matchPart = search ? 'matching ' : '';
                statusEl.textContent =
                    `Showing ${filtered.length} ${matchPart}${typePart}entr${filtered.length === 1 ? 'y' : 'ies'} (of ${rows.length} total).`;
            } else {
                statusEl.textContent = `Loaded ${rows.length} metadata entr${rows.length === 1 ? 'y' : 'ies'}.`;
            }
            statusEl.classList.add('upload-message--success');
        }
    }

    async function refreshSearchProjectsMetadataList() {
        const statusEl = document.getElementById('searchProjectsMetadataStatus');
        const tableWrap = document.getElementById('searchProjectsMetadataTableWrap');
        const typeSelect = document.getElementById('searchProjectsMetadataTypeSelect');
        const searchInput = document.getElementById('searchProjectsMetadataSearchInput');

        if (!organizationId || !statusEl || !tableWrap) return;

        const projectId = searchProjectsSelectedProject ? searchProjectsSelectedProject.project_id : null;
        if (!projectId) {
            tableWrap.innerHTML = '';
            statusEl.textContent = 'No project selected.';
            statusEl.classList.remove('upload-message--success');
            statusEl.classList.add('upload-message--error');
            searchProjectsMetadataRowsCache = [];
            if (typeSelect) typeSelect.value = '';
            if (searchInput) searchInput.value = '';
            return;
        }

        statusEl.classList.remove('upload-message--success', 'upload-message--error');
        statusEl.textContent = 'Loading metadata...';
        tableWrap.innerHTML = '';

        try {
            const { data: rows, error } = await supabase
                .from('lessons_learned_metadata_list')
                .select('id, metadata_type, metadata, metadata_source')
                .eq('organization_id', organizationId)
                .eq('project_id', projectId)
                .order('id', { ascending: true })
                .limit(5000);

            if (error) {
                console.error('Error loading lessons learned metadata list (search projects):', error);
                statusEl.textContent = 'Failed to load metadata.';
                statusEl.classList.add('upload-message--error');
                searchProjectsMetadataRowsCache = [];
                return;
            }

            const data = rows || [];
            searchProjectsMetadataRowsCache = data;
            updateSearchProjectsMetadataTypeOptions(data);
            renderSearchProjectsMetadataTable();
        } catch (err) {
            console.error('Unexpected error loading lessons learned metadata list (search projects):', err);
            statusEl.textContent = 'An unexpected error occurred while loading metadata.';
            statusEl.classList.add('upload-message--error');
            searchProjectsMetadataRowsCache = [];
        }
    }

    async function loadSearchProjectsTaskDetails(metadataRow) {
        const statusEl = document.getElementById('searchProjectsTaskDetailsStatus');
        const bodyEl = document.getElementById('searchProjectsTaskDetailsBody');
        if (!statusEl || !bodyEl) return;

        if (!organizationId) {
            statusEl.textContent = 'No organization found for this user.';
            statusEl.classList.add('upload-message--error');
            return;
        }

        const metadataId = metadataRow && metadataRow.id ? metadataRow.id : null;
        if (!metadataId) {
            statusEl.textContent = 'Missing metadata record.';
            statusEl.classList.add('upload-message--error');
            return;
        }

        try {
            const { data: taskRows, error } = await supabase
                .from('msproject_task_details')
                .select('id, uid, task_name, created_at, updated_at, start, finish, duration, percent_complete, baseline_start, baseline_finish, actual_start, actual_finish, predecessors, wbs_parent, fixed_cost, notes, project_id, organization_id')
                .eq('lessons_learned_metadata_list_id', metadataId)
                .limit(1);

            if (error) {
                console.error('Error loading MS Project task details:', error);
                statusEl.textContent = 'Failed to load task details.';
                statusEl.classList.add('upload-message--error');
                return;
            }

            const task = Array.isArray(taskRows) && taskRows.length > 0 ? taskRows[0] : null;
            if (!task) {
                statusEl.textContent = 'No task details found for this metadata entry.';
                statusEl.classList.add('upload-message--error');
                return;
            }

            const predecessorNames = await loadSearchProjectsTaskPredecessorNames(task);
            const wbsParentName = await loadSearchProjectsWbsParentName(task);
            renderSearchProjectsTaskDetails(task, predecessorNames, wbsParentName);
            statusEl.textContent = '';
            statusEl.classList.remove('upload-message--error');
        } catch (err) {
            console.error('Unexpected error loading MS Project task details:', err);
            statusEl.textContent = 'An unexpected error occurred while loading task details.';
            statusEl.classList.add('upload-message--error');
        }
    }

    async function loadSearchProjectsTaskPredecessorNames(task) {
        const taskId = task && task.id ? task.id : null;
        if (!taskId) return [];

        try {
            const { data: rows, error } = await supabase
                .from('msproject_task_predecessors')
                .select('predecessor_uid, project_id, organization_id')
                .eq('msproject_task_details_id', taskId);

            if (error) {
                console.error('Error loading task predecessors:', error);
                return [];
            }

            const uids = Array.from(
                new Set((rows || []).map(r => r && r.predecessor_uid).filter(Boolean))
            );
            if (uids.length === 0) return [];

            const projectId = task.project_id;
            const orgId = task.organization_id;
            const { data: nameRows, error: nameErr } = await supabase
                .from('msproject_task_details')
                .select('uid, task_name')
                .eq('project_id', projectId)
                .eq('organization_id', orgId)
                .in('uid', uids);

            if (nameErr) {
                console.error('Error loading predecessor task names:', nameErr);
                return [];
            }

            const nameByUid = new Map(
                (nameRows || []).map(r => [String(r.uid), r.task_name || `UID ${r.uid}`])
            );
            return uids.map(uid => nameByUid.get(String(uid)) || `UID ${uid}`);
        } catch (err) {
            console.error('Unexpected error loading predecessor names:', err);
            return [];
        }
    }

    function renderSearchProjectsTaskDetails(task, predecessorNames, wbsParentName) {
        const bodyEl = document.getElementById('searchProjectsTaskDetailsBody');
        if (!bodyEl) return;

        const items = [];
        const addItem = (label, value) => {
            if (value == null || value === '') return;
            items.push({ label, value });
        };

        addItem('UID', task.uid);
        addItem('Task Name', task.task_name);
        addItem('Created At', task.created_at);
        addItem('Updated At', task.updated_at);
        addItem('Start', formatDateOnly(task.start));
        addItem('Finish', formatDateOnly(task.finish));
        addItem('Duration', task.duration);
        addItem('Percent Complete', task.percent_complete);
        addItem('Baseline Start', formatDateOnly(task.baseline_start));
        addItem('Baseline Finish', formatDateOnly(task.baseline_finish));
        addItem('Actual Start', formatDateOnly(task.actual_start));
        addItem('Actual Finish', formatDateOnly(task.actual_finish));
        if (task.predecessors != null) {
            addItem('Predecessors', task.predecessors ? 'Yes' : 'No');
        }
        if (wbsParentName) {
            addItem('WBS Parent', wbsParentName);
        }
        if (Array.isArray(predecessorNames) && predecessorNames.length > 0) {
            addItem('Predecessors', predecessorNames.join(', '));
        }
        addItem('Fixed Cost', task.fixed_cost);
        addItem('Notes', task.notes);

        if (items.length === 0) {
            bodyEl.textContent = 'No task details available.';
            return;
        }

        const list = document.createElement('div');
        list.className = 'task-details-list';

        items.forEach(({ label, value }) => {
            const row = document.createElement('div');
            row.className = 'task-details-row';

            const labelEl = document.createElement('div');
            labelEl.className = 'task-details-label';
            labelEl.textContent = label;

            const valueEl = document.createElement('div');
            valueEl.className = 'task-details-value';
            valueEl.textContent = value;

            row.appendChild(labelEl);
            row.appendChild(valueEl);
            list.appendChild(row);
        });

        bodyEl.innerHTML = '';
        bodyEl.appendChild(list);
    }

    function formatDateOnly(value) {
        if (!value) return '';
        const date = value instanceof Date ? value : new Date(value);
        if (Number.isNaN(date.getTime())) return String(value);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    async function loadSearchProjectsWbsParentName(task) {
        const wbsParent = task && task.wbs_parent ? String(task.wbs_parent) : '';
        if (!wbsParent) return '';
        const projectId = task.project_id;
        const orgId = task.organization_id;
        if (projectId == null || orgId == null) return '';

        try {
            const { data, error } = await supabase
                .from('msproject_task_details')
                .select('task_name, wbs')
                .eq('project_id', projectId)
                .eq('organization_id', orgId)
                .eq('wbs', wbsParent)
                .limit(1);

            if (error) {
                console.error('Error loading WBS parent task name:', error);
                return '';
            }

            const row = Array.isArray(data) && data.length ? data[0] : null;
            return row && row.task_name ? row.task_name : '';
        } catch (err) {
            console.error('Unexpected error loading WBS parent task name:', err);
            return '';
        }
    }

    function showSearchProjectsDetails(project) {
        const panel = document.getElementById('searchProjectsDetailsPanel');
        const detailsScreen = document.getElementById('searchProjectsDetailsScreen');
        if (!panel || !detailsScreen) return;
        const main = document.getElementById('searchProjectsMain');
        const metadataPanel = document.getElementById('searchProjectsMetadataPanel');

        const nameEl = document.getElementById('searchProjectsDetailsName');
        const descEl = document.getElementById('searchProjectsDetailsDescription');
        const typeEl = document.getElementById('searchProjectsDetailsType');

        const name = project && project.project_name ? project.project_name : '(Unnamed Project)';
        const desc = project && project.project_description ? project.project_description : 'No description available.';
        const type =
            project && project.project_type && project.project_type.project_type
                ? project.project_type.project_type
                : 'Not set';

        if (nameEl) nameEl.textContent = name;
        if (descEl) descEl.textContent = desc;
        if (typeEl) typeEl.textContent = type;

        searchProjectsSelectedProject = project || null;

        if (main) main.style.display = 'none';
        if (metadataPanel) metadataPanel.style.display = 'none';
        detailsScreen.style.display = 'block';
        panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function hideSearchProjectsDetails() {
        const detailsScreen = document.getElementById('searchProjectsDetailsScreen');
        const metadataPanel = document.getElementById('searchProjectsMetadataPanel');
        const main = document.getElementById('searchProjectsMain');
        if (detailsScreen) detailsScreen.style.display = 'none';
        if (metadataPanel) metadataPanel.style.display = 'none';
        if (main) main.style.display = '';
    }

    function showSearchProjectsMetadataPanel() {
        const detailsScreen = document.getElementById('searchProjectsDetailsScreen');
        const panel = document.getElementById('searchProjectsMetadataPanel');
        const taskPanel = document.getElementById('searchProjectsTaskDetailsPanel');
        const title = document.getElementById('searchProjectsMetadataTitle');
        if (!panel) return;

        const projectName =
            searchProjectsSelectedProject && searchProjectsSelectedProject.project_name
                ? searchProjectsSelectedProject.project_name
                : 'Project';
        if (title) {
            title.textContent = `Lessons Learned Metadata for ${projectName}`;
        }

        if (detailsScreen) detailsScreen.style.display = 'none';
        if (taskPanel) taskPanel.style.display = 'none';
        panel.style.display = 'block';
        panel.scrollIntoView({ behavior: 'smooth', block: 'start' });

        refreshSearchProjectsMetadataList();
    }

    function hideSearchProjectsMetadataPanel() {
        const detailsScreen = document.getElementById('searchProjectsDetailsScreen');
        const panel = document.getElementById('searchProjectsMetadataPanel');
        if (panel) panel.style.display = 'none';
        if (detailsScreen) detailsScreen.style.display = 'block';
    }

    function showSearchProjectsTaskDetails(metadataRow) {
        const metadataPanel = document.getElementById('searchProjectsMetadataPanel');
        const taskPanel = document.getElementById('searchProjectsTaskDetailsPanel');
        const statusEl = document.getElementById('searchProjectsTaskDetailsStatus');
        const bodyEl = document.getElementById('searchProjectsTaskDetailsBody');
        if (!taskPanel || !statusEl || !bodyEl) return;

        if (metadataPanel) metadataPanel.style.display = 'none';
        taskPanel.style.display = 'block';
        taskPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });

        statusEl.textContent = 'Loading task details...';
        statusEl.classList.remove('upload-message--success', 'upload-message--error');
        bodyEl.innerHTML = '';

        loadSearchProjectsTaskDetails(metadataRow)
            .catch((err) => {
                console.error('Error loading task details:', err);
                statusEl.textContent = 'Failed to load task details.';
                statusEl.classList.add('upload-message--error');
            });
    }

    function hideSearchProjectsTaskDetails() {
        const metadataPanel = document.getElementById('searchProjectsMetadataPanel');
        const taskPanel = document.getElementById('searchProjectsTaskDetailsPanel');
        if (taskPanel) taskPanel.style.display = 'none';
        if (metadataPanel) metadataPanel.style.display = 'block';
    }

    function renderGeneralSearchModule() {
        const searchView = document.getElementById('searchView');
        if (!searchView) return;

        searchView.innerHTML = `
            <h1>General Search</h1>
            <div class="project-types-panel">
                <div class="project-types-panel-header">
                    <div>
                        <h3>General Search</h3>
                        <p class="subtitle">This section is coming soon.</p>
                    </div>
                </div>
            </div>
        `;
    }

    async function searchProjectsApi(term, signal) {
        const response = await fetch('/api/search-projects', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ queryText: term }),
            signal,
        });

        if (!response.ok) {
            const text = await response.text().catch(() => '');
            throw new Error(`Project search failed with ${response.status}: ${text}`);
        }

        const payload = await response.json().catch(() => ({}));
        return Array.isArray(payload.results) ? payload.results : [];
    }

    function setSearchProjectsStatus(message) {
        const el = document.getElementById('searchProjectsStatus');
        if (el) el.textContent = message || '';
    }

    function renderSearchProjectsResults(results, term) {
        const container = document.getElementById('searchProjectsResults');
        if (!container) return;

        container.innerHTML = '';

        if (!results || results.length === 0) {
            if (term) {
                setSearchProjectsStatus(`No projects found for “${term}”.`);
            }
            return;
        }

        setSearchProjectsStatus(
            `Showing ${results.length} project result${results.length === 1 ? '' : 's'} for “${term}”.`
        );

        results.forEach((item) => {
            const card = document.createElement('article');
            card.className = 'lesson-card';

            const nameLine = document.createElement('div');
            nameLine.className = 'lesson-card-fpc';
            const name = item && item.project_name ? String(item.project_name) : '(Unnamed Project)';
            nameLine.textContent = name;

            const descBox = document.createElement('div');
            descBox.className = 'lesson-card-inner lesson-card-generic-box';
            const desc = item && item.project_description ? String(item.project_description) : '';
            descBox.textContent = desc || 'No description available.';

            const meta = document.createElement('div');
            meta.className = 'lesson-card-meta';

            const idSpan = document.createElement('span');
            idSpan.textContent = `Project ID: ${item && item.project_id != null ? item.project_id : 'N/A'}`;

            const scoreSpan = document.createElement('span');
            const score = item && typeof item.score === 'number' ? item.score : null;
            scoreSpan.textContent = `Relevance: ${score == null ? 'N/A' : score.toFixed(4)}`;

            meta.appendChild(idSpan);
            meta.appendChild(scoreSpan);

            card.appendChild(nameLine);
            card.appendChild(descBox);
            card.appendChild(meta);

            container.appendChild(card);
        });
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
        // When entering Create view, default to showing the Create Project form (hide sub-modules)
        hideOrgProjectDetailsPanel();
        hideLessonsMetadataPanel();
        hideProjectTeamPanel();
        if (projectTeamAssignmentsPanel) {
            projectTeamAssignmentsPanel.style.display = 'none';
        }
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

            const projectTeamBtn = document.getElementById('projectTeamBtn');
            if (projectTeamBtn && !projectTeamBtn.dataset.wired) {
                projectTeamBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    showProjectTeamPanel();
                });
                projectTeamBtn.dataset.wired = 'true';
            }

            const lessonsMetadataBtn = document.getElementById('lessonsMetadataBtn');
            if (lessonsMetadataBtn && !lessonsMetadataBtn.dataset.wired) {
                lessonsMetadataBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    showLessonsMetadataPanel();
                });
                lessonsMetadataBtn.dataset.wired = 'true';
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

            const projectTeamBtn = document.getElementById('projectTeamBtn');
            if (projectTeamBtn && !projectTeamBtn.dataset.wired) {
                projectTeamBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    showProjectTeamPanel();
                });
                projectTeamBtn.dataset.wired = 'true';
            }

            const lessonsMetadataBtn = document.getElementById('lessonsMetadataBtn');
            if (lessonsMetadataBtn && !lessonsMetadataBtn.dataset.wired) {
                lessonsMetadataBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    showLessonsMetadataPanel();
                });
                lessonsMetadataBtn.dataset.wired = 'true';
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

                // Wire Create New Project button now that Org Settings template is mounted
                const orgView = document.getElementById('orgView');
                if (orgView) wireCreateProjectButtons(orgView);

                // Manage Projects module within Organization Settings
                const manageProjectsBtn = document.getElementById('manageProjectsButton');
                const orgMainSummary = document.getElementById('orgMainSummary');
                const manageOrgTeamsButton = document.getElementById('manageOrgTeamsButton');
                const orgTeamsPanel = document.getElementById('orgTeamsPanel');
                const orgTeamsBackButton = document.getElementById('orgTeamsBackButton');
                const orgTeamsUploadFile = document.getElementById('orgTeamsUploadFile');
                const orgTeamsUploadButton = document.getElementById('orgTeamsUploadButton');
                const orgTeamsUploadStatus = document.getElementById('orgTeamsUploadStatus');
                const orgTeamsUserSelect = document.getElementById('orgTeamsUserSelect');
                const orgTeamsUserStatus = document.getElementById('orgTeamsUserStatus');
                const orgTeamsAssignButton = document.getElementById('orgTeamsAssignButton');
                const orgTeamsAssignSubmodule = document.getElementById('orgTeamsAssignSubmodule');
                const orgTeamsAssignBackButton = document.getElementById('orgTeamsAssignBackButton');
                const orgTeamsAssignRefreshButton = document.getElementById('orgTeamsAssignRefreshButton');
                const orgTeamsAssignSaveButton = document.getElementById('orgTeamsAssignSaveButton');
                const orgTeamsAssignStatus = document.getElementById('orgTeamsAssignStatus');
                const orgTeamsAssignTableWrap = document.getElementById('orgTeamsAssignTableWrap');
                const orgProjectsPanel = document.getElementById('orgProjectsPanel');
                const orgProjectsBackButton = document.getElementById('orgProjectsBackButton');
                const orgProjectTypesForm = document.getElementById('orgProjectTypesForm');
                const orgProjectTypeInput = document.getElementById('orgProjectTypeName');
                const orgProjectsSubtitle = document.getElementById('orgProjectsSubtitle');
                const orgManageProjectsSubmodule = document.getElementById('orgManageProjectsSubmodule');
                const orgProjectsSelect = document.getElementById('orgProjectsSelect');
                const orgProjectDetailsSelect = document.getElementById('orgProjectDetailsSelect');

                // Create New Asset module within Organization Settings
                const createNewAssetButton = document.getElementById('createNewAssetButton');
                const orgAssetsPanel = document.getElementById('orgAssetsPanel');
                const orgAssetsCreateHeader = document.getElementById('orgAssetsCreateHeader');
                const orgAssetsBackButton = document.getElementById('orgAssetsBackButton');
                const orgAssetsForm = document.getElementById('orgAssetsForm');
                const orgAssetsStatus = document.getElementById('orgAssetsStatus');
                const assetDetailsButton = document.getElementById('assetDetailsButton');
                const orgAssetDetailsPanel = document.getElementById('orgAssetDetailsPanel');
                const orgAssetDetailsBackButton = document.getElementById('orgAssetDetailsBackButton');
                const orgAssetDetailsOrgBackButton = document.getElementById('orgAssetDetailsOrgBackButton');
                const orgAssetDetailsAssetSelect = document.getElementById('orgAssetDetailsAssetSelect');
                const orgAssetDetailsProjectSelect = document.getElementById('orgAssetDetailsProjectSelect');
                const orgAssetDetailsFileTypeSelect = document.getElementById('orgAssetDetailsFileTypeSelect');
                const orgAssetDetailsUploadWrap = document.getElementById('orgAssetDetailsUploadWrap');
                const orgAssetDetailsFileInput = document.getElementById('orgAssetDetailsFileInput');
                const orgAssetDetailsUploadButton = document.getElementById('orgAssetDetailsUploadButton');
                const orgAssetDetailsStatus = document.getElementById('orgAssetDetailsStatus');

                // Asset Details state
                let assetDetailsSelectedAssetId = null;
                let assetDetailsSelectedAssetName = null;
                let assetDetailsSelectedProjectId = null;
                let assetDetailsSelectedProjectName = null;
                let assetDetailsAssetsById = new Map();   // asset_id -> asset_name
                let assetDetailsProjectsById = new Map(); // project_id -> project_name

                let orgAssetDetailsAssetChoices = null;
                let orgAssetDetailsProjectChoices = null;

                // Manage Organizational Teams state
                let orgTeamsUserChoices = null;
                let orgTeamsUsersLoaded = false;
                let orgTeamsUsersById = new Map(); // user_id -> { name, email }
                let orgTeamsTeamsCache = []; // org_teams rows: [{ id, uid, department }]
                let orgTeamsAssignedTeamIds = new Set(); // org_team_id (uuid) assigned in DB
                let orgTeamsSelectedTeamIds = new Set(); // org_team_id (uuid) selected in UI

                function setOrgTeamsUploadStatus(message, kind = null) {
                    if (!orgTeamsUploadStatus) return;
                    orgTeamsUploadStatus.classList.remove('upload-message--success', 'upload-message--error');
                    orgTeamsUploadStatus.textContent = message || '';
                    if (kind === 'success') orgTeamsUploadStatus.classList.add('upload-message--success');
                    if (kind === 'error') orgTeamsUploadStatus.classList.add('upload-message--error');
                    orgTeamsUploadStatus.style.display = message ? '' : 'none';
                }

                function setOrgTeamsUserStatus(message, kind = 'success') {
                    if (!orgTeamsUserStatus) return;
                    orgTeamsUserStatus.textContent = message;
                    orgTeamsUserStatus.className = `upload-message${kind === 'error' ? ' upload-message--error' : kind === 'success' ? ' upload-message--success' : ''}`;
                    orgTeamsUserStatus.style.display = message ? '' : 'none';
                }

                function setOrgTeamsAssignStatus(message, kind = null) {
                    if (!orgTeamsAssignStatus) return;
                    orgTeamsAssignStatus.classList.remove('upload-message--success', 'upload-message--error');
                    orgTeamsAssignStatus.textContent = message || '';
                    if (kind === 'success') orgTeamsAssignStatus.classList.add('upload-message--success');
                    if (kind === 'error') orgTeamsAssignStatus.classList.add('upload-message--error');
                    orgTeamsAssignStatus.style.display = message ? '' : 'none';
                }

                function normalizeOrgTeamsHeaderKey(value) {
                    return String(value || '').trim().toLowerCase().replace(/\s+/g, '');
                }

                async function readFileAsArrayBuffer(file) {
                    return await new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onerror = () => reject(new Error('Failed reading file.'));
                        reader.onload = () => resolve(reader.result);
                        reader.readAsArrayBuffer(file);
                    });
                }

                async function readFileAsText(file) {
                    return await new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onerror = () => reject(new Error('Failed reading file.'));
                        reader.onload = () => resolve(String(reader.result || ''));
                        reader.readAsText(file);
                    });
                }

                async function parseOrgTeamsFile(file) {
                    const name = String(file && file.name ? file.name : '');
                    const ext = name.includes('.') ? name.split('.').pop().toLowerCase() : '';
                    const isCsv = ext === 'csv';
                    const isExcel = ext === 'xlsx' || ext === 'xls';

                    if (!isCsv && !isExcel) {
                        throw new Error('Invalid file type. Please upload .xlsx, .xls, or .csv.');
                    }

                    let workbook = null;
                    if (isCsv) {
                        const text = await readFileAsText(file);
                        workbook = XLSX.read(text, { type: 'string' });
                    } else {
                        const buf = await readFileAsArrayBuffer(file);
                        workbook = XLSX.read(buf, { type: 'array' });
                    }

                    const firstSheetName = workbook && Array.isArray(workbook.SheetNames) ? workbook.SheetNames[0] : null;
                    if (!firstSheetName) {
                        throw new Error('No worksheet found in the uploaded file.');
                    }
                    const ws = workbook.Sheets[firstSheetName];
                    if (!ws) {
                        throw new Error('Could not read the first worksheet.');
                    }

                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) || [];
                    if (rows.length < 2) {
                        throw new Error('No data rows found. Ensure the file includes a header row and at least one team row.');
                    }

                    const header = rows[0] || [];
                    const headerMap = new Map();
                    header.forEach((h, idx) => {
                        const key = normalizeOrgTeamsHeaderKey(h);
                        if (key) headerMap.set(key, idx);
                    });

                    const uidIdx = headerMap.get('uid');
                    const deptIdx = headerMap.get('department');
                    const descIdx = headerMap.get('description');

                    const missing = [];
                    if (uidIdx == null) missing.push('UID');
                    if (deptIdx == null) missing.push('Department');
                    if (descIdx == null) missing.push('Description');
                    if (missing.length > 0) {
                        throw new Error(`Missing required column(s): ${missing.join(', ')}. Expected columns: UID, Department, Description.`);
                    }

                    // De-duplicate by UID (last one wins)
                    const byUid = new Map();
                    let skippedEmptyUid = 0;
                    let duplicateUidCount = 0;

                    for (let i = 1; i < rows.length; i++) {
                        const row = rows[i] || [];
                        const uid = String(row[uidIdx] || '').trim();
                        if (!uid) {
                            skippedEmptyUid++;
                            continue;
                        }
                        if (byUid.has(uid)) duplicateUidCount++;
                        byUid.set(uid, {
                            uid,
                            department: String(row[deptIdx] || '').trim(),
                            description: String(row[descIdx] || '').trim()
                        });
                    }

                    const out = Array.from(byUid.values());
                    if (out.length === 0) {
                        throw new Error('No valid rows found (UID was empty for all rows).');
                    }

                    return { rows: out, skippedEmptyUid, duplicateUidCount };
                }

                async function fetchExistingOrgTeamUids(orgId) {
                    const existing = new Set();
                    const pageSize = 1000;
                    let offset = 0;
                    while (true) {
                        const { data, error } = await supabase
                            .from('org_teams')
                            .select('uid')
                            .eq('organization_id', orgId)
                            .order('uid', { ascending: true })
                            .range(offset, offset + pageSize - 1);

                        if (error) throw error;
                        const rows = data || [];
                        rows.forEach(r => {
                            if (r && r.uid != null) existing.add(String(r.uid));
                        });
                        if (rows.length < pageSize) break;
                        offset += pageSize;
                    }
                    return existing;
                }

                async function uploadOrgTeamsFromFile(file) {
                    const createdBy = ctUser && ctUser.id != null ? Number(ctUser.id) : null;
                    const orgId = organizationId != null ? Number(organizationId) : null;
                    if (!createdBy || Number.isNaN(createdBy)) {
                        throw new Error('Could not determine the uploader user id. Please log out and log back in.');
                    }
                    if (!orgId || Number.isNaN(orgId)) {
                        throw new Error('Could not determine your organization. Please log out and log back in.');
                    }

                    setOrgTeamsUploadStatus('Parsing file...', null);
                    const parsed = await parseOrgTeamsFile(file);

                    setOrgTeamsUploadStatus('Checking existing team UIDs...', null);
                    const existingUids = await fetchExistingOrgTeamUids(orgId);

                    const toInsert = [];
                    const toUpdate = [];
                    parsed.rows.forEach(r => {
                        if (existingUids.has(String(r.uid))) toUpdate.push(r);
                        else toInsert.push(r);
                    });

                    const chunkArray = (arr, size) => {
                        const n = Math.max(1, Number(size) || 1);
                        const out = [];
                        for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
                        return out;
                    };

                    let inserted = 0;
                    let updated = 0;

                    // Insert new rows in chunks
                    if (toInsert.length > 0) {
                        const payload = toInsert.map(r => ({
                            organization_id: orgId,
                            created_by: createdBy,
                            uid: r.uid,
                            department: r.department || null,
                            description: r.description || null
                        }));

                        const chunks = chunkArray(payload, 250);
                        for (let i = 0; i < chunks.length; i++) {
                            setOrgTeamsUploadStatus(`Uploading new teams... (${inserted}/${payload.length})`, null);
                            const chunk = chunks[i];
                            const { error } = await supabase.from('org_teams').insert(chunk);
                            if (error) throw error;
                            inserted += chunk.length;
                        }
                    }

                    // Update existing rows (per UID) with limited concurrency
                    if (toUpdate.length > 0) {
                        const updateChunks = chunkArray(toUpdate, 50);
                        for (let i = 0; i < updateChunks.length; i++) {
                            const chunk = updateChunks[i];
                            setOrgTeamsUploadStatus(`Updating existing teams... (${updated}/${toUpdate.length})`, null);
                            const results = await Promise.all(
                                chunk.map(r => supabase
                                    .from('org_teams')
                                    .update({
                                        department: r.department || null,
                                        description: r.description || null,
                                        created_by: createdBy
                                    })
                                    .eq('organization_id', orgId)
                                    .eq('uid', r.uid)
                                )
                            );
                            results.forEach(res => {
                                if (res && res.error) throw res.error;
                            });
                            updated += chunk.length;
                        }
                    }

                    const skipped = parsed.skippedEmptyUid || 0;
                    const dupes = parsed.duplicateUidCount || 0;
                    const note = [
                        skipped ? `${skipped} skipped (missing UID)` : null,
                        dupes ? `${dupes} duplicates in file (last value used)` : null
                    ].filter(Boolean).join(', ');

                    const summary = `Upload complete. Inserted ${inserted}, updated ${updated}.${note ? ' ' + note + '.' : ''}`;
                    setOrgTeamsUploadStatus(summary, 'success');
                }

                async function loadOrgTeamsForAssignList() {
                    if (!organizationId) throw new Error('Could not determine your organization. Please log out and log back in.');
                    const orgId = Number(organizationId);
                    if (!orgId || Number.isNaN(orgId)) throw new Error('Could not determine your organization. Please log out and log back in.');

                    const { data, error } = await supabase
                        .from('org_teams')
                        .select('id, uid, department')
                        .eq('organization_id', orgId)
                        .order('department', { ascending: true })
                        .order('uid', { ascending: true })
                        .limit(5000);

                    if (error) throw error;
                    orgTeamsTeamsCache = (data || []).map(r => ({
                        id: r.id,
                        uid: r.uid != null ? String(r.uid) : '',
                        department: r.department != null ? String(r.department) : ''
                    }));
                }

                async function loadOrgTeamsAssignmentsForUser(selectedUserId) {
                    if (!organizationId) throw new Error('Could not determine your organization. Please log out and log back in.');
                    const orgId = Number(organizationId);
                    const userId = Number(selectedUserId);
                    if (!orgId || Number.isNaN(orgId)) throw new Error('Could not determine your organization. Please log out and log back in.');
                    if (!userId || Number.isNaN(userId)) throw new Error('Please select a user first.');

                    const { data, error } = await supabase
                        .from('org_team_members')
                        .select('org_team_id')
                        .eq('organization_id', orgId)
                        .eq('team_member', userId)
                        .limit(5000);

                    if (error) throw error;
                    orgTeamsAssignedTeamIds = new Set((data || []).map(r => String(r.org_team_id)));
                    orgTeamsSelectedTeamIds = new Set(Array.from(orgTeamsAssignedTeamIds));
                }

                function renderOrgTeamsAssignTable() {
                    if (!orgTeamsAssignTableWrap) return;
                    orgTeamsAssignTableWrap.innerHTML = '';

                    if (!orgTeamsTeamsCache || orgTeamsTeamsCache.length === 0) {
                        const p = document.createElement('p');
                        p.textContent = 'No teams found for your organization.';
                        p.style.marginTop = '10px';
                        orgTeamsAssignTableWrap.appendChild(p);
                        return;
                    }

                    const table = document.createElement('table');
                    table.className = 'organizations-table';

                    const thead = document.createElement('thead');
                    const headRow = document.createElement('tr');
                    ['UID', 'Department'].forEach(label => {
                        const th = document.createElement('th');
                        th.textContent = label;
                        headRow.appendChild(th);
                    });
                    thead.appendChild(headRow);
                    table.appendChild(thead);

                    const tbody = document.createElement('tbody');
                    orgTeamsTeamsCache.forEach(team => {
                        const teamId = String(team.id);
                        const tr = document.createElement('tr');
                        tr.style.cursor = 'pointer';
                        if (orgTeamsSelectedTeamIds && orgTeamsSelectedTeamIds.has(teamId)) {
                            tr.classList.add('row-selected');
                        }
                        tr.addEventListener('click', () => {
                            if (!orgTeamsSelectedTeamIds) orgTeamsSelectedTeamIds = new Set();
                            if (orgTeamsSelectedTeamIds.has(teamId)) {
                                orgTeamsSelectedTeamIds.delete(teamId);
                                tr.classList.remove('row-selected');
                            } else {
                                orgTeamsSelectedTeamIds.add(teamId);
                                tr.classList.add('row-selected');
                            }
                        });

                        const tdUid = document.createElement('td');
                        tdUid.textContent = team.uid || '';
                        tr.appendChild(tdUid);

                        const tdDept = document.createElement('td');
                        tdDept.textContent = team.department || '';
                        tr.appendChild(tdDept);

                        tbody.appendChild(tr);
                    });
                    table.appendChild(tbody);
                    orgTeamsAssignTableWrap.appendChild(table);
                }

                async function showOrgTeamsAssignSubmodule() {
                    if (!orgTeamsAssignSubmodule) return;
                    const userIdStr = orgTeamsUserSelect ? String(orgTeamsUserSelect.value || '') : '';
                    if (!userIdStr) {
                        setOrgTeamsUserStatus('Please select a user first.', 'error');
                        return;
                    }

                    orgTeamsAssignSubmodule.style.display = '';
                    setOrgTeamsAssignStatus('Loading teams...', null);

                    await loadOrgTeamsForAssignList();
                    setOrgTeamsAssignStatus('Loading existing assignments...', null);
                    await loadOrgTeamsAssignmentsForUser(userIdStr);

                    renderOrgTeamsAssignTable();
                    setOrgTeamsAssignStatus('', null);
                }

                function hideOrgTeamsAssignSubmodule() {
                    if (!orgTeamsAssignSubmodule) return;
                    orgTeamsAssignSubmodule.style.display = 'none';
                    setOrgTeamsAssignStatus('', null);
                    if (orgTeamsAssignTableWrap) orgTeamsAssignTableWrap.innerHTML = '';
                    orgTeamsTeamsCache = [];
                    orgTeamsAssignedTeamIds = new Set();
                    orgTeamsSelectedTeamIds = new Set();
                }

                async function saveOrgTeamsAssignments() {
                    const userIdStr = orgTeamsUserSelect ? String(orgTeamsUserSelect.value || '') : '';
                    const userId = Number(userIdStr);
                    const orgId = organizationId != null ? Number(organizationId) : null;
                    const createdBy = ctUser && ctUser.id != null ? Number(ctUser.id) : null;

                    if (!userIdStr || !userId || Number.isNaN(userId)) {
                        throw new Error('Please select a user first.');
                    }
                    if (!orgId || Number.isNaN(orgId)) {
                        throw new Error('Could not determine your organization. Please log out and log back in.');
                    }
                    if (!createdBy || Number.isNaN(createdBy)) {
                        throw new Error('Could not determine the uploader user id. Please log out and log back in.');
                    }

                    const selectedSet = orgTeamsSelectedTeamIds || new Set();
                    const assignedSet = orgTeamsAssignedTeamIds || new Set();

                    const toInsertIds = Array.from(selectedSet).filter(id => !assignedSet.has(id));
                    const toDeleteIds = Array.from(assignedSet).filter(id => !selectedSet.has(id));

                    const chunkArray = (arr, size) => {
                        const n = Math.max(1, Number(size) || 1);
                        const out = [];
                        for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
                        return out;
                    };

                    const userInfo = orgTeamsUsersById && orgTeamsUsersById.get(userIdStr) ? orgTeamsUsersById.get(userIdStr) : null;
                    const userName = userInfo && userInfo.name ? String(userInfo.name) : `User ${userIdStr}`;

                    const teamsById = new Map((orgTeamsTeamsCache || []).map(t => [String(t.id), t]));

                    let inserted = 0;
                    let deleted = 0;

                    // Insert new assignments
                    if (toInsertIds.length > 0) {
                        const payload = toInsertIds.map(teamId => {
                            const team = teamsById.get(String(teamId));
                            return {
                                organization_id: orgId,
                                created_by: createdBy,
                                name: userName,
                                org_team_id: String(teamId),
                                team_member: userId,
                                department: team && team.department ? String(team.department) : null
                            };
                        });

                        const chunks = chunkArray(payload, 250);
                        for (let i = 0; i < chunks.length; i++) {
                            setOrgTeamsAssignStatus(`Saving... (adding ${inserted}/${payload.length})`, null);
                            const chunk = chunks[i];
                            const { error } = await supabase.from('org_team_members').insert(chunk);
                            if (error) throw error;
                            inserted += chunk.length;
                        }
                    }

                    // Delete removed assignments
                    if (toDeleteIds.length > 0) {
                        const idChunks = chunkArray(toDeleteIds.map(String), 250);
                        for (let i = 0; i < idChunks.length; i++) {
                            setOrgTeamsAssignStatus(`Saving... (removing ${deleted}/${toDeleteIds.length})`, null);
                            const chunk = idChunks[i];
                            const { data, error } = await supabase
                                .from('org_team_members')
                                .delete()
                                .eq('organization_id', orgId)
                                .eq('team_member', userId)
                                .in('org_team_id', chunk)
                                .select('id');
                            if (error) throw error;
                            deleted += (data || []).length;
                        }
                    }

                    // Update assigned set to match selected set
                    orgTeamsAssignedTeamIds = new Set(Array.from(selectedSet));

                    setOrgTeamsAssignStatus(`Saved. Added ${inserted} and removed ${deleted}.`, 'success');
                }

                async function loadOrgUsersForOrgTeamsDropdown({ force = false } = {}) {
                    if (!orgTeamsUserSelect) return;
                    if (!organizationId) {
                        setOrgTeamsUserStatus('Could not determine your organization. Please log out and log back in.', 'error');
                        return;
                    }

                    if (orgTeamsUsersLoaded && !force) {
                        setOrgTeamsUserStatus('', 'success');
                        return;
                    }

                    setOrgTeamsUserStatus('Loading users...', 'success');
                    try {
                        const { data: userRows, error: userErr } = await supabase
                            .from('users')
                            .select('id, name, email, organizationid')
                            .eq('organizationid', organizationId)
                            .order('name', { ascending: true })
                            .limit(5000);

                        if (userErr) {
                            console.error('Error loading organization users for org teams dropdown:', userErr);
                            setOrgTeamsUserStatus('Failed to load users.', 'error');
                            return;
                        }

                        const rows = userRows || [];
                        if (rows.length === 0) {
                            orgTeamsUsersLoaded = true;
                            setOrgTeamsUserStatus('No users found for your organization.', 'error');
                            if (orgTeamsUserChoices) {
                                orgTeamsUserChoices.clearChoices();
                            } else {
                                orgTeamsUserSelect.innerHTML = '<option value="">No users available</option>';
                            }
                            return;
                        }

                        orgTeamsUsersById = new Map();
                        const choicesData = rows.map(row => {
                            const idStr = String(row.id);
                            const name = row.name || `User ${idStr}`;
                            const email = row.email || '';
                            orgTeamsUsersById.set(idStr, { name, email });
                            const label = email ? `${name} (${email})` : name;
                            return { value: idStr, label };
                        });

                        if (!orgTeamsUserChoices) {
                            if (typeof Choices === 'undefined') {
                                console.warn('Choices library not loaded; falling back to native select for org teams users dropdown.');
                                orgTeamsUserSelect.innerHTML = '<option value="">Select User</option>';
                                choicesData.forEach(choice => {
                                    const option = document.createElement('option');
                                    option.value = choice.value;
                                    option.textContent = choice.label;
                                    orgTeamsUserSelect.appendChild(option);
                                });
                            } else {
                                orgTeamsUserChoices = new Choices(orgTeamsUserSelect, {
                                    searchEnabled: true,
                                    shouldSort: false,
                                    placeholder: true,
                                    placeholderValue: 'Select User',
                                    searchPlaceholderValue: 'Type to search...'
                                });
                            }
                        }

                        if (orgTeamsUserChoices) {
                            orgTeamsUserChoices.setChoices(choicesData, 'value', 'label', true);
                        } else {
                            orgTeamsUserSelect.innerHTML = '<option value="">Select User</option>';
                            choicesData.forEach(choice => {
                                const option = document.createElement('option');
                                option.value = choice.value;
                                option.textContent = choice.label;
                                orgTeamsUserSelect.appendChild(option);
                            });
                        }

                        orgTeamsUsersLoaded = true;
                        setOrgTeamsUserStatus('', 'success');
                    } catch (err) {
                        console.error('Unexpected error loading organization users for org teams dropdown:', err);
                        setOrgTeamsUserStatus('An unexpected error occurred while loading users.', 'error');
                    }
                }

                function setOrgAssetDetailsStatus(message, kind = 'success') {
                    if (!orgAssetDetailsStatus) return;
                    orgAssetDetailsStatus.textContent = message;
                    orgAssetDetailsStatus.className = `upload-message${kind === 'error' ? ' upload-message--error' : ' upload-message--success'}`;
                    orgAssetDetailsStatus.style.display = '';
                }

                async function loadOrgAssetsForAssetDetails() {
                    if (!orgAssetDetailsAssetSelect) return;

                    const orgId = ctUser && ctUser.organizationid != null ? Number(ctUser.organizationid) : null;
                    if (!orgId || Number.isNaN(orgId)) {
                        setOrgAssetsStatus('Could not determine your organization. Please log out and log back in.', 'error');
                        return;
                    }

                    orgAssetDetailsAssetSelect.innerHTML = '<option value=\"\">Select Asset</option>';
                    assetDetailsAssetsById = new Map();

                    try {
                        const { data, error } = await supabase
                            .from('assets')
                            .select('id, name')
                            .eq('organization_id', orgId)
                            .order('name', { ascending: true })
                            .limit(5000);

                        if (error) {
                            console.error('Error loading org assets for asset details dropdown:', error);
                            const opt = document.createElement('option');
                            opt.value = '';
                            opt.textContent = 'Failed to load assets';
                            opt.disabled = true;
                            orgAssetDetailsAssetSelect.appendChild(opt);
                            return;
                        }

                        const rows = data || [];
                        if (rows.length === 0) {
                            const opt = document.createElement('option');
                            opt.value = '';
                            opt.textContent = 'No assets found for your organization';
                            opt.disabled = true;
                            orgAssetDetailsAssetSelect.appendChild(opt);
                            return;
                        }

                        const choicesData = rows.map(row => {
                            const idStr = String(row.id);
                            const name = row.name || `Asset ${idStr}`;
                            assetDetailsAssetsById.set(idStr, name);
                            return { value: idStr, label: name };
                        });

                        if (typeof Choices === 'undefined') {
                            choicesData.forEach(choice => {
                                const option = document.createElement('option');
                                option.value = choice.value;
                                option.textContent = choice.label;
                                orgAssetDetailsAssetSelect.appendChild(option);
                            });
                            return;
                        }

                        if (!orgAssetDetailsAssetChoices) {
                            orgAssetDetailsAssetChoices = new Choices(orgAssetDetailsAssetSelect, {
                                searchEnabled: true,
                                shouldSort: false,
                                placeholder: true,
                                placeholderValue: 'Select Asset',
                                searchPlaceholderValue: 'Type to search...'
                            });
                        }
                        orgAssetDetailsAssetChoices.setChoices(choicesData, 'value', 'label', true);
                    } catch (err) {
                        console.error('Unexpected error loading org assets for asset details dropdown:', err);
                        const opt = document.createElement('option');
                        opt.value = '';
                        opt.textContent = 'Unexpected error loading assets';
                        opt.disabled = true;
                        orgAssetDetailsAssetSelect.appendChild(opt);
                    }
                }

                function setOrgAssetsStatus(message, kind = 'success') {
                    if (!orgAssetsStatus) return;
                    orgAssetsStatus.textContent = message;
                    orgAssetsStatus.className = `upload-message${kind === 'error' ? ' upload-message--error' : ' upload-message--success'}`;
                    orgAssetsStatus.style.display = '';
                }

                // Show Create New Asset panel
                if (createNewAssetButton && orgMainSummary && orgAssetsPanel) {
                    createNewAssetButton.onclick = (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        orgMainSummary.style.display = 'none';
                        if (orgProjectsPanel) orgProjectsPanel.style.display = 'none';
                        if (orgTeamsPanel) orgTeamsPanel.style.display = 'none';
                        orgAssetsPanel.style.display = '';
                        if (orgAssetsStatus) orgAssetsStatus.style.display = 'none';
                        if (orgAssetsCreateHeader) orgAssetsCreateHeader.style.display = '';
                        if (orgAssetDetailsPanel) orgAssetDetailsPanel.style.display = 'none';
                        if (orgAssetsForm) orgAssetsForm.style.display = '';
                    };
                }

                // Go Back from Create New Asset panel
                if (orgAssetsBackButton && orgMainSummary && orgAssetsPanel) {
                    orgAssetsBackButton.onclick = (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        orgAssetsPanel.style.display = 'none';
                        orgMainSummary.style.display = '';
                        if (orgAssetsStatus) orgAssetsStatus.style.display = 'none';
                        if (orgAssetsForm && typeof orgAssetsForm.reset === 'function') orgAssetsForm.reset();
                        if (orgAssetDetailsPanel) orgAssetDetailsPanel.style.display = 'none';
                        if (orgAssetsForm) orgAssetsForm.style.display = '';
                        if (orgAssetsCreateHeader) orgAssetsCreateHeader.style.display = '';
                    };
                }

                async function loadOrgProjectsForAssetDetails() {
                    if (!orgAssetDetailsProjectSelect) return;

                    const orgId = ctUser && ctUser.organizationid != null ? Number(ctUser.organizationid) : null;
                    if (!orgId || Number.isNaN(orgId)) {
                        // Keep UX simple: user will see empty list and can go back; also message in main status.
                        setOrgAssetsStatus('Could not determine your organization. Please log out and log back in.', 'error');
                        return;
                    }

                    // Reset options
                    orgAssetDetailsProjectSelect.innerHTML = '<option value=\"\">Select Project</option><option value=\"na\">N/A</option>';
                    assetDetailsProjectsById = new Map();
                    assetDetailsProjectsById.set('na', 'N/A');

                    try {
                        const { data, error } = await supabase
                            .from('projects')
                            .select('project_id, project_name')
                            .eq('organization_id', orgId)
                            .order('project_name', { ascending: true });

                        if (error) {
                            console.error('Error loading org projects for asset details dropdown:', error);
                            const opt = document.createElement('option');
                            opt.value = '';
                            opt.textContent = 'Failed to load projects';
                            opt.disabled = true;
                            orgAssetDetailsProjectSelect.appendChild(opt);
                            return;
                        }

                        const rows = data || [];
                        if (rows.length === 0) {
                            const opt = document.createElement('option');
                            opt.value = '';
                            opt.textContent = 'No projects found for your organization';
                            opt.disabled = true;
                            orgAssetDetailsProjectSelect.appendChild(opt);
                            return;
                        }

                        const choicesData = [
                            { value: 'na', label: 'N/A' },
                            ...rows.map(row => {
                                const idStr = String(row.project_id);
                                const name = row.project_name || `Project ${idStr}`;
                                assetDetailsProjectsById.set(idStr, name);
                                return { value: idStr, label: name };
                            })
                        ];

                        if (typeof Choices === 'undefined') {
                            // Fallback: native select
                            choicesData.forEach(choice => {
                                const option = document.createElement('option');
                                option.value = choice.value;
                                option.textContent = choice.label;
                                orgAssetDetailsProjectSelect.appendChild(option);
                            });
                            return;
                        }

                        if (!orgAssetDetailsProjectChoices) {
                            orgAssetDetailsProjectChoices = new Choices(orgAssetDetailsProjectSelect, {
                                searchEnabled: true,
                                shouldSort: false,
                                placeholder: true,
                                placeholderValue: 'Select Project',
                                searchPlaceholderValue: 'Type to search...'
                            });
                        }
                        orgAssetDetailsProjectChoices.setChoices(choicesData, 'value', 'label', true);
                    } catch (err) {
                        console.error('Unexpected error loading org projects for asset details dropdown:', err);
                        const opt = document.createElement('option');
                        opt.value = '';
                        opt.textContent = 'Unexpected error loading projects';
                        opt.disabled = true;
                        orgAssetDetailsProjectSelect.appendChild(opt);
                    }
                }

                function getAssetComponentsProjectFields() {
                    // N/A or blank -> nulls
                    if (!assetDetailsSelectedProjectId || !assetDetailsSelectedProjectName) {
                        return { project_id: null, project: null };
                    }
                    return {
                        project_id: assetDetailsSelectedProjectId,
                        project: assetDetailsSelectedProjectName
                    };
                }

                // Asset Details button -> show Asset Details panel
                if (assetDetailsButton && orgAssetsForm && orgAssetDetailsPanel) {
                    assetDetailsButton.addEventListener('click', async (e) => {
                        e.preventDefault();
                        orgAssetsForm.style.display = 'none';
                        if (orgAssetsStatus) orgAssetsStatus.style.display = 'none';
                        orgAssetDetailsPanel.style.display = '';
                        if (orgAssetsCreateHeader) orgAssetsCreateHeader.style.display = 'none';
                        if (orgAssetDetailsFileTypeSelect) {
                            orgAssetDetailsFileTypeSelect.selectedIndex = 0;
                        }
                        if (orgAssetDetailsUploadWrap) orgAssetDetailsUploadWrap.style.display = 'none';
                        if (orgAssetDetailsStatus) orgAssetDetailsStatus.style.display = 'none';
                        if (orgAssetDetailsFileInput) orgAssetDetailsFileInput.value = '';
                        if (orgAssetDetailsAssetSelect) {
                            orgAssetDetailsAssetSelect.selectedIndex = 0;
                        }
                        assetDetailsSelectedAssetId = null;
                        assetDetailsSelectedAssetName = null;
                        assetDetailsSelectedProjectId = null;
                        assetDetailsSelectedProjectName = null;
                        await loadOrgAssetsForAssetDetails();
                        await loadOrgProjectsForAssetDetails();
                    });
                }

                // Back to Create Asset (from Asset Details)
                if (orgAssetDetailsBackButton && orgAssetsForm && orgAssetDetailsPanel) {
                    orgAssetDetailsBackButton.addEventListener('click', (e) => {
                        e.preventDefault();
                        orgAssetDetailsPanel.style.display = 'none';
                        orgAssetsForm.style.display = '';
                        if (orgAssetsCreateHeader) orgAssetsCreateHeader.style.display = '';
                    });
                }

                // Go Back to Org Settings summary (from Asset Details)
                if (orgAssetDetailsOrgBackButton && orgMainSummary && orgAssetsPanel) {
                    orgAssetDetailsOrgBackButton.addEventListener('click', (e) => {
                        e.preventDefault();
                        orgAssetsPanel.style.display = 'none';
                        orgMainSummary.style.display = '';
                        if (orgAssetsStatus) orgAssetsStatus.style.display = 'none';
                        if (orgAssetDetailsPanel) orgAssetDetailsPanel.style.display = 'none';
                        if (orgAssetsForm) orgAssetsForm.style.display = '';
                        if (orgAssetsCreateHeader) orgAssetsCreateHeader.style.display = '';
                    });
                }

                // Track asset selection
                if (orgAssetDetailsAssetSelect) {
                    orgAssetDetailsAssetSelect.addEventListener('change', () => {
                        const val = String(orgAssetDetailsAssetSelect.value || '').trim();
                        if (!val) {
                            assetDetailsSelectedAssetId = null;
                            assetDetailsSelectedAssetName = null;
                            return;
                        }
                        assetDetailsSelectedAssetId = Number(val);
                        assetDetailsSelectedAssetName = assetDetailsAssetsById.get(val) || null;
                    });
                }

                // Track project selection (N/A -> null)
                if (orgAssetDetailsProjectSelect) {
                    orgAssetDetailsProjectSelect.addEventListener('change', () => {
                        const val = String(orgAssetDetailsProjectSelect.value || '').trim();
                        if (!val || val === 'na') {
                            assetDetailsSelectedProjectId = null;
                            assetDetailsSelectedProjectName = null;
                            return;
                        }
                        assetDetailsSelectedProjectId = Number(val);
                        assetDetailsSelectedProjectName = assetDetailsProjectsById.get(val) || null;
                    });
                }

                // File type toggle for Asset Details
                if (orgAssetDetailsFileTypeSelect) {
                    orgAssetDetailsFileTypeSelect.addEventListener('change', () => {
                        const val = String(orgAssetDetailsFileTypeSelect.value || '').trim();
                        const showUpload = val === 'excel_general';
                        if (orgAssetDetailsUploadWrap) orgAssetDetailsUploadWrap.style.display = showUpload ? '' : 'none';
                        if (!showUpload) {
                            if (orgAssetDetailsStatus) orgAssetDetailsStatus.style.display = 'none';
                            if (orgAssetDetailsFileInput) orgAssetDetailsFileInput.value = '';
                        }
                    });
                }

                // Upload handler (Excel - General)
                if (orgAssetDetailsUploadButton) {
                    orgAssetDetailsUploadButton.addEventListener('click', async (e) => {
                        e.preventDefault();

                        const fileType = orgAssetDetailsFileTypeSelect ? String(orgAssetDetailsFileTypeSelect.value || '').trim() : '';
                        if (fileType !== 'excel_general') {
                            setOrgAssetDetailsStatus('Please select the input data file type "Excel - General" first.', 'error');
                            return;
                        }

                        const orgId = ctUser && ctUser.organizationid != null ? Number(ctUser.organizationid) : null;
                        const createdBy = ctUser && ctUser.id != null ? Number(ctUser.id) : null;
                        if (!orgId || Number.isNaN(orgId)) {
                            setOrgAssetDetailsStatus('Could not determine your organization. Please log out and log back in.', 'error');
                            return;
                        }
                        if (!createdBy || Number.isNaN(createdBy)) {
                            setOrgAssetDetailsStatus('Could not determine your user id. Please log out and log back in.', 'error');
                            return;
                        }

                        if (!assetDetailsSelectedAssetId || !assetDetailsSelectedAssetName) {
                            setOrgAssetDetailsStatus('Please select an Asset first.', 'error');
                            return;
                        }

                        const file = orgAssetDetailsFileInput && orgAssetDetailsFileInput.files
                            ? orgAssetDetailsFileInput.files[0]
                            : null;
                        if (!file) {
                            setOrgAssetDetailsStatus('Please choose an Excel file (.xlsx or .xls) to upload.', 'error');
                            return;
                        }
                        const name = String(file.name || '').toLowerCase();
                        const isExcel = name.endsWith('.xlsx') || name.endsWith('.xls');
                        if (!isExcel) {
                            setOrgAssetDetailsStatus('Invalid file type. Please upload an Excel file (.xlsx, .xls).', 'error');
                            return;
                        }

                        const submitBtn = orgAssetDetailsUploadButton;
                        const originalText = submitBtn.textContent;
                        submitBtn.disabled = true;
                        submitBtn.textContent = 'Uploading...';
                        if (orgAssetDetailsStatus) orgAssetDetailsStatus.style.display = 'none';

                        try {
                            const projectFields = getAssetComponentsProjectFields();
                            const context = {
                                organization_id: orgId,
                                created_by: createdBy,
                                asset_id: assetDetailsSelectedAssetId,
                                asset_name: assetDetailsSelectedAssetName,
                                project_id: projectFields.project_id,
                                project: projectFields.project
                            };

                            const result = await importAssetGeneralExcelToSupabase({
                                supabase,
                                file,
                                context,
                                chunkSize: 250,
                                onProgress: (msg) => {
                                    setOrgAssetDetailsStatus(msg, 'success');
                                }
                            });

                            const compCount = result && result.inserted ? result.inserted.asset_components : 0;
                            const depCount = result && result.inserted ? result.inserted.asset_component_dependencies : 0;
                            setOrgAssetDetailsStatus(
                                `Upload complete. Inserted ${compCount} asset components and ${depCount} dependencies.`,
                                'success'
                            );
                        } catch (err) {
                            console.error('Excel - General upload failed:', err);
                            setOrgAssetDetailsStatus(err && err.message ? err.message : 'Upload failed.', 'error');
                        } finally {
                            submitBtn.disabled = false;
                            submitBtn.textContent = originalText || 'Upload Spreadsheet';
                        }
                    });
                }

                // Create New Asset form submit -> persist to Supabase 'assets'
                if (orgAssetsForm) {
                    orgAssetsForm.addEventListener('submit', async (e) => {
                        e.preventDefault();
                        const nameEl = document.getElementById('assetName');
                        const descEl = document.getElementById('assetDescription');
                        const startEl = document.getElementById('assetBuildStartDate');
                        const endEl = document.getElementById('assetCompletionDate');

                        const name = nameEl ? String(nameEl.value || '').trim() : '';
                        const desc = descEl ? String(descEl.value || '').trim() : '';
                        const start = startEl ? String(startEl.value || '').trim() : '';
                        const end = endEl ? String(endEl.value || '').trim() : '';

                        if (!name || !desc) {
                            setOrgAssetsStatus('Please fill in Asset Name and Asset Description.', 'error');
                            return;
                        }

                        // Pull required context from session (avoid referencing organizationId here due to TDZ in this function)
                        const orgId = ctUser && ctUser.organizationid != null ? Number(ctUser.organizationid) : null;
                        const createdBy = ctUser && ctUser.id != null ? Number(ctUser.id) : null;

                        if (!orgId || Number.isNaN(orgId)) {
                            setOrgAssetsStatus('Could not determine your organization. Please log out and log back in.', 'error');
                            return;
                        }
                        if (!createdBy || Number.isNaN(createdBy)) {
                            setOrgAssetsStatus('Could not determine your user id. Please log out and log back in.', 'error');
                            return;
                        }

                        const submitBtn = orgAssetsForm.querySelector('button[type="submit"]');
                        const originalBtnText = submitBtn ? submitBtn.textContent : '';
                        if (submitBtn) {
                            submitBtn.disabled = true;
                            submitBtn.textContent = 'Creating...';
                        }

                        try {
                            const payload = {
                                organization_id: orgId,
                                created_by: createdBy,
                                name,
                                description: desc,
                                build_start_date: start || null,
                                completion_date: end || null,
                                asset_details: false
                            };

                            const { data: assetRow, error } = await supabase
                                .from('assets')
                                .insert(payload)
                                .select()
                                .single();

                            if (error) {
                                console.error('Supabase assets insert error:', error);
                                setOrgAssetsStatus(error.message || 'Failed to create asset. Please try again.', 'error');
                                return;
                            }

                            const createdName = assetRow && assetRow.name ? String(assetRow.name) : name;
                            setOrgAssetsStatus(`Asset "${createdName}" created successfully.`, 'success');
                            if (typeof orgAssetsForm.reset === 'function') orgAssetsForm.reset();
                        } catch (err) {
                            console.error('Unexpected error creating asset:', err);
                            setOrgAssetsStatus('An unexpected error occurred while creating the asset.', 'error');
                        } finally {
                            if (submitBtn) {
                                submitBtn.disabled = false;
                                submitBtn.textContent = originalBtnText || 'Create Asset';
                            }
                        }
                    });
                }

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
                        if (orgTeamsPanel) orgTeamsPanel.style.display = 'none';
                        if (orgAssetsPanel) orgAssetsPanel.style.display = 'none';
                        orgProjectsPanel.style.display = '';
                    };
                }

                // Show the Manage Organizational Teams module (in main area)
                if (manageOrgTeamsButton && orgMainSummary && orgTeamsPanel) {
                    manageOrgTeamsButton.onclick = async (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        orgMainSummary.style.display = 'none';
                        if (orgProjectsPanel) orgProjectsPanel.style.display = 'none';
                        if (orgAssetsPanel) orgAssetsPanel.style.display = 'none';
                        orgTeamsPanel.style.display = '';
                        hideOrgTeamsAssignSubmodule();
                        if (orgTeamsUploadFile) orgTeamsUploadFile.value = '';
                        setOrgTeamsUploadStatus('', null);
                        setOrgTeamsUserStatus('', 'success');
                        await loadOrgUsersForOrgTeamsDropdown({ force: true });
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

                // Go Back from Teams panel returns to the main Organization Settings summary
                if (orgTeamsBackButton && orgMainSummary && orgTeamsPanel) {
                    orgTeamsBackButton.onclick = (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        orgTeamsPanel.style.display = 'none';
                        orgMainSummary.style.display = '';
                        setOrgTeamsUserStatus('', 'success');
                        setOrgTeamsUploadStatus('', null);
                        hideOrgTeamsAssignSubmodule();
                    };
                }

                // Upload Organizational Teams handler
                if (orgTeamsUploadButton && orgTeamsUploadFile) {
                    orgTeamsUploadButton.onclick = async (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        const file = orgTeamsUploadFile.files && orgTeamsUploadFile.files[0] ? orgTeamsUploadFile.files[0] : null;
                        if (!file) {
                            setOrgTeamsUploadStatus('Please choose a file first.', 'error');
                            return;
                        }

                        const name = String(file.name || '');
                        const ext = name.includes('.') ? name.split('.').pop().toLowerCase() : '';
                        const isAllowed = ext === 'xlsx' || ext === 'xls' || ext === 'csv';
                        if (!isAllowed) {
                            setOrgTeamsUploadStatus('Invalid file type. Please upload .xlsx, .xls, or .csv.', 'error');
                            return;
                        }

                        try {
                            orgTeamsUploadButton.disabled = true;
                            orgTeamsUploadButton.textContent = 'Uploading...';
                            await uploadOrgTeamsFromFile(file);
                        } catch (err) {
                            console.error('Org teams upload failed:', err);
                            setOrgTeamsUploadStatus(err && err.message ? err.message : 'Upload failed.', 'error');
                        } finally {
                            orgTeamsUploadButton.disabled = false;
                            orgTeamsUploadButton.textContent = 'Upload File';
                        }
                    };
                }

                // Assign Teams wiring
                if (orgTeamsAssignButton) {
                    orgTeamsAssignButton.onclick = async (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        try {
                            await showOrgTeamsAssignSubmodule();
                        } catch (err) {
                            console.error('Failed opening Assign Teams submodule:', err);
                            setOrgTeamsAssignStatus(err && err.message ? err.message : 'Failed to load teams.', 'error');
                        }
                    };
                }

                if (orgTeamsAssignBackButton) {
                    orgTeamsAssignBackButton.onclick = (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        hideOrgTeamsAssignSubmodule();
                    };
                }

                if (orgTeamsAssignRefreshButton) {
                    orgTeamsAssignRefreshButton.onclick = async (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        try {
                            await showOrgTeamsAssignSubmodule();
                        } catch (err) {
                            console.error('Failed refreshing Assign Teams submodule:', err);
                            setOrgTeamsAssignStatus(err && err.message ? err.message : 'Failed to refresh teams.', 'error');
                        }
                    };
                }

                if (orgTeamsAssignSaveButton) {
                    orgTeamsAssignSaveButton.onclick = async (e) => {
                        if (e && e.preventDefault) e.preventDefault();
                        try {
                            orgTeamsAssignSaveButton.disabled = true;
                            orgTeamsAssignSaveButton.textContent = 'Saving...';
                            await saveOrgTeamsAssignments();
                        } catch (err) {
                            console.error('Failed saving org team member assignments:', err);
                            setOrgTeamsAssignStatus(err && err.message ? err.message : 'Failed to save assignments.', 'error');
                        } finally {
                            orgTeamsAssignSaveButton.disabled = false;
                            orgTeamsAssignSaveButton.textContent = 'Save Assignments';
                        }
                    };
                }

                // If user changes, close the assign submodule (prevents mixing selections between users)
                if (orgTeamsUserSelect) {
                    orgTeamsUserSelect.addEventListener('change', () => {
                        setOrgTeamsUserStatus('', 'success');
                        hideOrgTeamsAssignSubmodule();
                    });
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
            // Backward compatibility: if a link still uses #search, treat it as Search Projects.
            showSearchProjectsView();
        } else if (route === 'search-projects') {
            resetMainArea();
            showSearchProjectsView();
        } else if (route === 'general-search') {
            resetMainArea();
            showGeneralSearchView();
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
