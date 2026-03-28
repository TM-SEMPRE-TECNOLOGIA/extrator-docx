// Theme Toggle
function toggleTheme() {
    document.documentElement.classList.toggle('dark');
    const isDark = document.documentElement.classList.contains('dark');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
}

// Load saved theme
function loadTheme() {
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'dark') {
        document.documentElement.classList.add('dark');
    }
}

// Set active navigation link
function setActiveNav() {
    const currentPage = window.location.pathname.split('/').pop() || 'index.html';
    const navLinks = document.querySelectorAll('.nav-link');
    
    navLinks.forEach(link => {
        const linkPage = link.getAttribute('href');
        if (linkPage === currentPage || (currentPage === '' && linkPage === 'index.html')) {
            link.classList.add('active');
        } else {
            link.classList.remove('active');
        }
    });
}

// Tab functionality
function initTabs() {
    const tabs = document.querySelectorAll('.tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            // Remove active from all tabs
            tabs.forEach(t => t.classList.remove('active'));
            // Add active to clicked tab
            tab.classList.add('active');
            
            // Trigger any custom event if needed
            const event = new CustomEvent('tabChange', { detail: { tab: tab.dataset.tab } });
            document.dispatchEvent(event);
        });
    });
}

// Density toggle
function toggleDensity() {
    const container = document.querySelector('.container');
    if (container.classList.contains('density-compact')) {
        container.classList.remove('density-compact');
        container.classList.add('density-comfortable');
    } else {
        container.classList.remove('density-comfortable');
        container.classList.add('density-compact');
    }
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    loadTheme();
    setActiveNav();
    initTabs();
});
