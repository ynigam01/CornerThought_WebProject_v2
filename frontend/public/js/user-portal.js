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
                    <button type="submit" class="save-project-button">Save Project</button>
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
                </form>
            </div>
        </div>
    </div>`;

    container.insertAdjacentHTML("beforeend", projectFormHTML);
    container.insertAdjacentHTML("beforeend", addDataFormHTML);

    const createProjectButton = document.querySelector(".create-project-button");
    const addDataButton = document.querySelector(".add-data-button");
    const logoutButton = document.querySelector(".logout-button");

    const createProjectModal = document.getElementById("createProjectModal");
    const addDataModal = document.getElementById("addDataModal");

    document.getElementById("closeCreateProject").onclick = () => createProjectModal.classList.remove("show");
    document.getElementById("closeAddData").onclick = () => {
        addDataModal.classList.remove("show");
        resetAddDataForm();
    };

    createProjectButton.onclick = () => createProjectModal.classList.add("show");
    addDataButton.onclick = () => {
        addDataModal.classList.add("show");
        document.getElementById("issueSuccessDisplay").style.display = "block";
    };
    logoutButton.onclick = () => window.location.href = "index.html";

    window.onclick = (e) => {
    if (e.target === createProjectModal) createProjectModal.classList.remove("show");
    if (e.target === addDataModal) {
        addDataModal.classList.remove("show");
        resetAddDataForm();
    }
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
            // Keep the form open - don't close it
        }
    };

    // Function to add Issue/Success entry to display area
    function addIssueSuccessEntry(type, text) {
        const displayArea = document.getElementById("issueSuccessDisplay");
        const entry = document.createElement("div");
        entry.className = "issue-success-entry";
        entry.innerHTML = `<strong>${type}:</strong> ${text}`;
        displayArea.appendChild(entry);
    }

    // Function to reset the Add Data form
    function resetAddDataForm() {
        const inputField = document.getElementById("issueSuccess");
        inputField.value = "";
        inputField.disabled = false;
        inputField.style.backgroundColor = "";
        inputField.style.color = "";
    }

    // Collapsible form functionality
    const addDataToggle = document.getElementById("addDataToggle");
    const addDataToggleIcon = document.getElementById("addDataToggleIcon");
    const createProjectToggle = document.getElementById("createProjectToggle");
    const createProjectToggleIcon = document.getElementById("createProjectToggleIcon");

    // Toggle Add Data form
    addDataToggle.addEventListener("click", () => {
        const modal = document.getElementById("addDataModal");
        modal.classList.toggle("collapsed");
        
        if (modal.classList.contains("collapsed")) {
            addDataToggleIcon.style.transform = "rotate(0deg)";
        } else {
            addDataToggleIcon.style.transform = "rotate(180deg)";
        }
    });

    // Toggle Create Project form
    createProjectToggle.addEventListener("click", () => {
        const modal = document.getElementById("createProjectModal");
        modal.classList.toggle("collapsed");
        
        if (modal.classList.contains("collapsed")) {
            createProjectToggleIcon.style.transform = "rotate(0deg)";
        } else {
            createProjectToggleIcon.style.transform = "rotate(180deg)";
        }
    });

    // Show display area when Add Data is clicked
    addDataButton.onclick = () => {
        addDataModal.classList.add("show");
        document.getElementById("issueSuccessDisplay").style.display = "block";
    };

    createProjectButton.onclick = () => {
        createProjectModal.classList.add("show");
    };
});
