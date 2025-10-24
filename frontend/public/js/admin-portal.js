// js/admin-portal.js
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

    document.getElementById("addOrgForm").addEventListener("submit", (e) => {
        e.preventDefault();
        alert("Organization added successfully!");
        modal.classList.add("hidden");
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
