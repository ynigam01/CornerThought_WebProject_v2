// js/admin-login.js
document.addEventListener("DOMContentLoaded", () => {
    const loginForm = document.getElementById("adminLoginForm");

    loginForm.addEventListener("submit", (e) => {
        e.preventDefault();
        const email = document.getElementById("email").value;
        const password = document.getElementById("password").value;

        if (email && password) {
            alert("Login successful! Redirecting to admin portal...");
            window.location.href = "admin-portal.html";
        } else {
            alert("Please enter valid credentials.");
        }
    });
});
