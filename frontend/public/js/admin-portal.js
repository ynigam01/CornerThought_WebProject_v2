// js/admin-portal.js
import { supabase } from './supabase-client.js';

document.addEventListener("DOMContentLoaded", () => {
    const addOrgButton = document.querySelector(".add-organization-button");
    const modal = document.getElementById("addOrgModal");
    const closeButton = document.querySelector(".close-button");
    const form = document.getElementById("addOrgForm");
    const orgTableBody = document.querySelector(".organizations-table tbody");

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

    // Handle table action buttons (delete)
    orgTableBody.addEventListener('click', async (event) => {
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

    // Kick off initial load
    loadOrganizations();
});
