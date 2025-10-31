// js/admin-portal.js
import { supabase } from './supabase-client.js';

document.addEventListener("DOMContentLoaded", () => {
    const addOrgButton = document.querySelector(".add-organization-button");
    const modal = document.getElementById("addOrgModal");
    const closeButton = document.querySelector(".close-button");
    const form = document.getElementById("addOrgForm");
    const orgTableBody = document.querySelector(".organizations-table tbody");

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
            const newRow = document.createElement('tr');
            const idCell = document.createElement('td');
            idCell.textContent = data?.id ?? '—';
            const nameCell = document.createElement('td');
            nameCell.textContent = data?.name ?? orgName;
            const typeCell = document.createElement('td');
            typeCell.textContent = data?.type ?? orgType;
            const emailCell = document.createElement('td');
            emailCell.textContent = data?.admin_email ?? adminEmail;
            const usersCell = document.createElement('td');
            usersCell.textContent = String(data?.users_allotted ?? usersAllotted);
            const actionsCell = document.createElement('td');
            actionsCell.innerHTML = '<button class="edit-button">Edit</button><button class="delete-button">Delete</button>';

            newRow.appendChild(idCell);
            newRow.appendChild(nameCell);
            newRow.appendChild(typeCell);
            newRow.appendChild(emailCell);
            newRow.appendChild(usersCell);
            newRow.appendChild(actionsCell);
            orgTableBody.appendChild(newRow);

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
});
