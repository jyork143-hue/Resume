document.addEventListener('DOMContentLoaded', () => {
    // 1. Mobile Navigation Toggle & Keyboard Accessibility
    const mobileMenuBtn = document.getElementById('mobile-menu-btn');
    const primaryNav = document.getElementById('primary-navigation');

    if (mobileMenuBtn && primaryNav) {
        const closeMobileMenu = () => {
            mobileMenuBtn.setAttribute('aria-expanded', 'false');
            primaryNav.classList.remove('is-open');
        };

        mobileMenuBtn.addEventListener('click', () => {
            const isExpanded = mobileMenuBtn.getAttribute('aria-expanded') === 'true';
            mobileMenuBtn.setAttribute('aria-expanded', String(!isExpanded));
            primaryNav.classList.toggle('is-open', !isExpanded);
        });

        primaryNav.querySelectorAll('a').forEach((link) => {
            link.addEventListener('click', closeMobileMenu);
        });

        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && primaryNav.classList.contains('is-open')) {
                closeMobileMenu();
                mobileMenuBtn.focus();
            }
        });
    }

    // 2. State-Aware Sticky Header Fallback (for browsers without @container scroll-state)
    const stickyContainer = document.querySelector('.sticky-nav-container');
    if (stickyContainer && 'IntersectionObserver' in window) {
        const topOffset = parseFloat(getComputedStyle(stickyContainer).top) || 0;
        const stickyObserver = new IntersectionObserver(
            ([entry]) => {
                stickyContainer.classList.toggle('is-stuck', entry.intersectionRatio < 1);
            },
            {
                threshold: [1],
                rootMargin: `-${topOffset + 1}px 0px 0px 0px`
            }
        );
        stickyObserver.observe(stickyContainer);
    }

    // 3. Experience Timeline: Expand All / Collapse All Toggle
    const toggleAllBtn = document.getElementById('toggle-all-experience');
    const expDetailsList = Array.from(document.querySelectorAll('.exp-details'));

    const syncToggleBtnState = () => {
        if (!toggleAllBtn || expDetailsList.length === 0) return;
        const allOpen = expDetailsList.every((details) => details.open);
        toggleAllBtn.setAttribute('aria-expanded', String(allOpen));
        toggleAllBtn.textContent = allOpen ? 'Collapse All Role Details' : 'Expand All Role Details';
    };

    if (toggleAllBtn && expDetailsList.length > 0) {
        toggleAllBtn.addEventListener('click', () => {
            const shouldOpen = !expDetailsList.every((details) => details.open);
            expDetailsList.forEach((details) => {
                details.open = shouldOpen;
            });
            syncToggleBtnState();
        });

        expDetailsList.forEach((details) => {
            details.addEventListener('toggle', syncToggleBtnState);
        });
    }

    // 4. Print-Friendly Résumé Mode (Auto-expand all <details> when printing)
    let savedDetailsState = [];

    const expandForPrint = () => {
        savedDetailsState = expDetailsList.map((d) => d.open);
        expDetailsList.forEach((d) => {
            d.open = true;
        });
    };

    const restoreAfterPrint = () => {
        if (savedDetailsState.length === expDetailsList.length) {
            expDetailsList.forEach((d, idx) => {
                d.open = savedDetailsState[idx];
            });
            syncToggleBtnState();
        }
    };

    window.addEventListener('beforeprint', expandForPrint);
    window.addEventListener('afterprint', restoreAfterPrint);

    const triggerPrintResume = () => {
        expandForPrint();
        window.print();
    };

    const heroPrintBtn = document.getElementById('hero-print-btn');
    const contactPrintBtn = document.getElementById('contact-print-btn');
    if (heroPrintBtn) heroPrintBtn.addEventListener('click', triggerPrintResume);
    if (contactPrintBtn) contactPrintBtn.addEventListener('click', triggerPrintResume);

    // 5. Active Section Navigation Indicator
    const sections = document.querySelectorAll('main > section[id]');
    const navLinks = document.querySelectorAll('.nav-link');

    if (sections.length > 0 && navLinks.length > 0 && 'IntersectionObserver' in window) {
        const sectionObserver = new IntersectionObserver(
            (entries) => {
                entries.forEach((entry) => {
                    if (entry.isIntersecting) {
                        const activeId = `#${entry.target.id}`;
                        navLinks.forEach((link) => {
                            if (link.getAttribute('href') === activeId) {
                                link.setAttribute('aria-current', 'true');
                            } else {
                                link.removeAttribute('aria-current');
                            }
                        });
                    }
                });
            },
            {
                rootMargin: '-20% 0px -65% 0px',
                threshold: 0
            }
        );

        sections.forEach((section) => sectionObserver.observe(section));
    }
});
