// js/user-portal.js
document.addEventListener("DOMContentLoaded", () => {
    const container = document.getElementById("projectFormContainer");

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
                    <input type="text" id="projectType">
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
                        <option value="Company Administrator">Company Administrator</option>
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

    const createProjectButton = document.querySelector(".create-project-button");
    const addDataButton = document.querySelector(".add-data-button");
    const logoutButton = document.querySelector(".logout-button");

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

    createProjectButton.onclick = () => {
        if (confirmNavigation("Create New Project")) {
            location.hash = 'create';
        }
    };
    logoutButton.onclick = () => {
        if (confirmNavigation("logout")) {
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

    // Handle Add User form submission
    document.getElementById("addUserForm").onsubmit = (e) => {
        e.preventDefault();
        const userName = document.getElementById("userName").value;
        const userEmail = document.getElementById("userEmail").value;
        const userType = document.getElementById("userType").value;
        
        // Here you would typically send this data to a backend API
        alert(`User added successfully!\nName: ${userName}\nEmail: ${userEmail}\nUser Type: ${userType}`);
        
        // Close modal and reset form
        addUserModal.classList.remove("show");
        document.getElementById("addUserForm").reset();
    };

    document.getElementById("createProjectForm").onsubmit = (e) => {
        e.preventDefault();
        alert("Project created successfully!");
        createProjectModal.style.display = "none";
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
    }

    function showCreateView() {
        exitOrganizationSettings();
        hide('#homeView');
        hide('#searchView');
        show('#createView');
        hide('#addDataView');
        // Inject form once
        const container = document.querySelector('#createView');
        if (container && !container.querySelector('#inlineCreateProjectForm')) {
            container.insertAdjacentHTML('beforeend', `
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
            inlineForm.addEventListener('submit', (e) => {
                e.preventDefault();
                const name = document.getElementById('inlineProjectName').value.trim();
                const type = document.getElementById('inlineProjectType').value.trim();
                if (!name || !type) {
                    alert('Please fill in both Project Name and Type.');
                    return;
                }
                const asset = document.getElementById('inlineAssetType').value.trim() || 'N/A';
                const desc = document.getElementById('inlineProjectDescription').value.trim() || 'N/A';
                const start = document.getElementById('inlineStartDate').value.trim() || 'N/A';
                const end = document.getElementById('inlineEndDate').value.trim() || 'N/A';

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
    }

    // Function to display Organization Settings
    function displayOrganizationSettings() {
        // Hide other views
        hide('#homeView');
        hide('#searchView');
        hide('#createView');
        hide('#addDataView');
        // If already present, don't duplicate
        if (!document.getElementById('orgView')) {
            const tpl = document.getElementById('orgSettingsTemplate');
            if (tpl) {
                document.querySelector('.main-content').appendChild(tpl.content.cloneNode(true));
                const addBtn = document.getElementById('addNewUserButton');
                if (addBtn) addBtn.onclick = () => addUserModal.classList.add('show');
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
            resetMainArea();
            displayOrganizationSettings();
        } else if (route === 'search') {
            resetMainArea();
            showSearchView();
        } else if (route === 'create') {
            resetMainArea();
            showCreateView();
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
