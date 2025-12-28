// Streamlit Enhancement Script
// Additional JavaScript for app enhancements

let streamlitUIHidden = true;

window.addEventListener('load', function() {
    // Hide Streamlit UI elements by default
    hideStreamlitUI();
    
    // Sync banner tabs with Streamlit tabs
    syncBannerTabs();
});

window.switchTab = function(tabIndex) {
    // Update banner tab appearance
    const bannerTabs = document.querySelectorAll('.banner-tab');
    bannerTabs.forEach((tab, index) => {
        if (index === tabIndex) {
            tab.classList.add('active');
        } else {
            tab.classList.remove('active');
        }
    });
    
    // Trigger Streamlit tab switch
    const streamlitTabs = document.querySelectorAll('[data-baseweb="tab"]');
    if (streamlitTabs[tabIndex]) {
        streamlitTabs[tabIndex].click();
    }
};

window.toggleStreamlitUI = function() {
    streamlitUIHidden = !streamlitUIHidden;
    
    if (streamlitUIHidden) {
        hideStreamlitUI();
    } else {
        showStreamlitUI();
    }
};

function hideStreamlitUI() {
    // Hide Streamlit header elements
    const header = document.querySelector('header[data-testid="stHeader"]');
    if (header) header.style.display = 'none';
    
    const deployButton = document.querySelector('.stDeployButton');
    if (deployButton) deployButton.style.display = 'none';
    
    const mainMenu = document.querySelector('#MainMenu');
    if (mainMenu) mainMenu.style.display = 'none';
    
    // Hide hamburger menu button
    const menuButton = document.querySelector('[data-testid="baseButton-header"]');
    if (menuButton) menuButton.style.display = 'none';
    
    // Hide any other Streamlit UI elements in header
    const headerElements = document.querySelectorAll('header *');
    headerElements.forEach(el => {
        if (el.classList.contains('stDeployButton') || 
            el.getAttribute('data-testid') === 'stHeader') {
            el.style.display = 'none';
        }
    });
}

function showStreamlitUI() {
    // Show Streamlit header elements
    const header = document.querySelector('header[data-testid="stHeader"]');
    if (header) header.style.display = '';
    
    const deployButton = document.querySelector('.stDeployButton');
    if (deployButton) deployButton.style.display = '';
    
    const mainMenu = document.querySelector('#MainMenu');
    if (mainMenu) mainMenu.style.display = '';
    
    const menuButton = document.querySelector('[data-testid="baseButton-header"]');
    if (menuButton) menuButton.style.display = '';
}

function syncBannerTabs() {
    // Watch for Streamlit tab changes and update banner
    const observer = new MutationObserver(function(mutations) {
        const streamlitTabs = document.querySelectorAll('[data-baseweb="tab"]');
        streamlitTabs.forEach((tab, index) => {
            if (tab.getAttribute('aria-selected') === 'true') {
                const bannerTabs = document.querySelectorAll('.banner-tab');
                bannerTabs.forEach((bannerTab, bannerIndex) => {
                    if (bannerIndex === index) {
                        bannerTab.classList.add('active');
                    } else {
                        bannerTab.classList.remove('active');
                    }
                });
            }
        });
    });
    
    const tabsContainer = document.querySelector('[data-baseweb="tabs"]');
    if (tabsContainer) {
        observer.observe(tabsContainer, { attributes: true, childList: true, subtree: true });
    }
}

