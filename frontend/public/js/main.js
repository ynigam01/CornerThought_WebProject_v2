// js/main.js
document.addEventListener("DOMContentLoaded", () => {
    const dropdownButton = document.querySelector(".dropdown-button");
    const dropdownContent = document.querySelector(".dropdown-content");

    dropdownButton.addEventListener("click", () => {
        dropdownContent.classList.toggle("show");
    });

    window.addEventListener("click", (e) => {
        if (!e.target.matches(".dropdown-button")) {
            dropdownContent.classList.remove("show");
        }
    });
});
