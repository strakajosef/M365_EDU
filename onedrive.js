/**
 * ============================================================================
 * INTERAKTIVNÍ MODUL MICROSOFT ONEDRIVE + PLOCHA WINDOWS 11 (onedrive.js)
 * Plně propojený s existujícím teams.js!
 * ============================================================================
 */

window.SandboxModules = window.SandboxModules || {};

window.SandboxModules.onedrive = {
    state: null,
    container: null,
    callbacks: null,

    init(container, mission, callbacks) {
        this.initStateOnly(container, callbacks);
        this.render();
    },

    initStateOnly(container, callbacks) {
        this.container = container;
        this.callbacks = callbacks;

        this.state = {
            environment: "desktop", // 'desktop' | 'teams' | 'outlook' | 'chrome'
            outlookOneDriveLoaded: false,
            outlookAppsFlyoutOpen: false,

            activeSection: "home",
            activeLibraryName: null,
            typeFilter: "all",
            searchQuery: "",
            nameFilterQuery: "",
            selectedFileId: "f_legacy",

            createMenuOpen: false,
            uploadModalOpen: false,
            shareModalOpen: false,
            sharePermission: "edit",
            versionModalOpen: false,
            contextMenu: { open: false, x: 0, y: 0, fileId: null },

            forYouCards: [
                { id: "c1", title: "Audit projektu - vzorek", icon: "📗", author: "Petr Vzorový edited this", time: "před 5 h" },
                { id: "c2", title: "Aplikace Výstřižky", icon: "🟣", author: "You edited", time: "před 27 min" },
                { id: "c3", title: "Zadani_scenare_priklad", icon: "📗", author: "Tomáš Ukázka shared this with you", time: "Včera" },
                { id: "c4", title: "Portal_spolecnosti_priklad", icon: "📕", author: "You recently opened", time: "před 3 h" }
            ],

            libraries: [
                { id: "lib_digi", name: "Projekt Alfa", badge: "PA", color: "#d83b01" },
                { id: "lib_czsk", name: "Tým Beta - Marketing", badge: "TB", color: "#a4262c" },
                { id: "lib_hrs", name: "Tým Gama (interní)", badge: "TG", color: "#323130" },
                { id: "lib_nzp", name: "Analýza - Příklad", badge: "AP", color: "#605e5c" },
                { id: "lib_ino", name: "Inovační komunita 🚀", badge: "💡", color: "#0078d4" },
                { id: "lib_zp", name: "Tým Delta - Provoz", badge: "TD", color: "#0078d4" },
                { id: "lib_comp", name: "Zaměstnanci - Benefity", badge: "ZB", color: "#c50f1f" }
            ],

            files: [
                { id: "f_local_plan", name: "Plan_obce_2026.docx", type: "word", icon: "📘", modified: "Dnes 9:12", owner: "Jana Testovací", size: "96 kB", sharing: "Soukromé", activity: "Uloženo pouze lokálně: C:\\Dokumenty\\Plan_obce_2026.docx (není zálohováno)", section: "my_files", lib: "Moje soubory", local: true },
                { id: "f_teams_chat", name: "Soubory z chatu aplikace Microsoft Teams", type: "folder", icon: "📁", modified: "20.10.2020", owner: "Jana Testovací", size: "87 položek", sharing: "Soukromé", activity: "", section: "my_files", lib: "Moje soubory" },
                { id: "f_m365", name: "Ukázkový sešit.xlsx", type: "excel", icon: "📗", modified: "před 4 h", owner: "Jana Testovací", size: "745 kB", sharing: "Sdíleno", activity: "Upravili jste tento dokument · před 4 h", section: "both", lib: "Moje soubory", versions: [
                    { ver: "Verze 3.0 (Aktuální)", date: "Dnes 16:40", author: "Jana Testovací", size: "745 kB" },
                    { ver: "Verze 2.0", date: "Včera 14:15", author: "Petr Vzorový", size: "712 kB" },
                    { ver: "Verze 1.0", date: "15.10.2026 09:10", author: "Jana Testovací", size: "680 kB" }
                ]},
                { id: "f_legacy", name: "legacy.csv", type: "excel", icon: "📗", modified: "před 2 h", owner: "Jana Testovací", size: "13,1 kB", sharing: "Soukromé", activity: "", section: "my_files", lib: "Moje soubory", versions: [
                    { ver: "Verze 2.0 (Aktuální)", date: "Dnes 11:20", author: "Jana Testovací", size: "13,1 kB" },
                    { ver: "Verze 1.0 (Původní export)", date: "Včera 17:05", author: "Jana Testovací", size: "12,4 kB" }
                ]},
                { id: "f_dana", name: "priklad_dokument.docx", type: "word", icon: "📘", modified: "Včera", owner: "Jana Testovací", size: "563 kB", sharing: "Soukromé", activity: "", section: "my_files", lib: "Moje soubory" },
                { id: "f_souhrn", name: "Souhrn vystupu tymu 2026.docx", type: "word", icon: "📘", modified: "23. 4.", owner: "Jana Testovací", size: "13,0 kB", sharing: "Sdíleno", activity: "Vy nasdílel(a) tento soubor. · 23. 4.", section: "my_files", lib: "Moje soubory" },
                { id: "f_vystrizky", name: "Aplikace Výstřižky", type: "ppt", icon: "🟣", modified: "před 27 min", owner: "Jana Testovací", size: "4,2 MB", sharing: "Soukromé", activity: "Upravili jste tento dokument. · před 30 min", section: "both", lib: "Moje soubory" },
                { id: "f_pyramida", name: "pyramida_dovednosti_priklad.pdf", type: "pdf", icon: "📕", modified: "před 2 min", owner: "Jana Testovací", size: "1,4 MB", sharing: "Sdíleno", activity: "", section: "both", lib: "Projekt Alfa" },
                { id: "f_citrix_pdf", name: "Portal_spolecnosti_priklad.pdf", type: "pdf", icon: "📕", modified: "před 3 h", owner: "Eva Ukázková", size: "2,1 MB", sharing: "Sdíleno", activity: "", section: "both", lib: "Tým Beta - Marketing" },
                { id: "f_citrix_xls", name: "Předběžný harmonogram migrace.xlsx", type: "excel", icon: "📗", modified: "před 3 h", owner: "Eva Ukázková", size: "890 kB", sharing: "Sdíleno", activity: "", section: "both", lib: "Tým Beta - Marketing" },
                { id: "f_koord", name: "Tabulka koordinátorů.xlsx", type: "excel", icon: "📗", modified: "před 3 h", owner: "Eva Ukázková", size: "450 kB", sharing: "Sdíleno", activity: "", section: "both", lib: "Tým Beta - Marketing" },
                { id: "f_audit_sprava", name: "Audit projektu - vzorek.xlsx", type: "excel", icon: "📗", modified: "před 4 h", owner: "Petr Vzorový", size: "1,1 MB", sharing: "Sdíleno", activity: "Eva Ukázková upravila tento dokument. · před 5 h", section: "both", lib: "Projekt Alfa" },
                { id: "f_audit_m365", name: "AUDIT_PRIKLAD_M365.xlsx", type: "excel", icon: "📗", modified: "před 4 h", owner: "Petr Vzorový", size: "920 kB", sharing: "Sdíleno", activity: "", section: "both", lib: "Projekt Alfa" }
            ]
        };
    },

    render() {
        const s = this.state;

        // Pokud si uživatel na ploše Windows 11 vybral TEAMS, použijeme tvůj hotový modul teams.js!
        if (s.environment === "teams" && window.SandboxModules.teams) {
            if (!window.SandboxModules.teams.state) {
                window.SandboxModules.teams.init(this.container, null, this.callbacks);
            }
            window.SandboxModules.teams.container = this.container;
            window.SandboxModules.teams.callbacks = this.callbacks;
            window.SandboxModules.teams.state.launchedFromDesktop = true;
            window.SandboxModules.teams.render();
            return;
        }

        this.container.innerHTML = `
            <div class="od-win11-desktop" id="od-root">
                <div class="od-win11-bloom"></div>
                ${s.environment === "desktop" ? this.renderWin11DesktopContent() : this.renderActiveEnvironmentWindow()}
                ${s.contextMenu.open ? this.renderContextMenu() : ""}
                ${s.uploadModalOpen ? this.renderUploadModal() : ""}
                ${s.shareModalOpen ? this.renderShareModal() : ""}
                ${s.versionModalOpen ? this.renderVersionModal() : ""}
                ${this.renderWin11Taskbar()}
            </div>
        `;

        this.bindEvents();
    },

    // Voláno přímo z tvého teams.js, když je v Teams otevřená záložka OneDrive
    renderEmbeddedOneDrive(hostType) {
        const s = this.state;
        return `
            <div class="od-main-layout" id="od-embedded-root" style="flex:1; display:flex; position:relative; overflow:hidden;">
                ${this.renderOneDriveInnerLayout(hostType)}
                ${s.contextMenu.open ? this.renderContextMenu() : ""}
                ${s.uploadModalOpen ? this.renderUploadModal() : ""}
                ${s.shareModalOpen ? this.renderShareModal() : ""}
                ${s.versionModalOpen ? this.renderVersionModal() : ""}
            </div>
        `;
    },

    renderWin11DesktopContent() {
        return `
            <div class="od-desktop-icons">
                <div class="od-desk-icon" data-launch-env="teams">
                    <div class="od-desk-icon-badge" style="background:#464eb8;">👥</div>
                    <span>Microsoft Teams</span>
                </div>
                <div class="od-desk-icon" data-launch-env="outlook">
                    <div class="od-desk-icon-badge" style="background:#0078d4;">✉️</div>
                    <span>Microsoft Outlook</span>
                </div>
                <div class="od-desk-icon" data-launch-env="chrome">
                    <div class="od-desk-icon-badge" style="background:linear-gradient(135deg,#ea4335,#fbbc05,#34a853,#4285f4);">🌐</div>
                    <span>Google Chrome</span>
                </div>
            </div>

            <div class="od-launcher-card">
                <div style="font-size:12px; text-transform:uppercase; letter-spacing:1px; color:#60cdff; font-weight:700;">Plocha Windows 11 • Rozcestník OneDrive</div>
                <h2 style="font-size:20px; margin:6px 0 6px;">Kudy chceš OneDrive otevřít?</h2>
                <p style="font-size:13px; color:#cbd5e1;">
                    Vyber si, zda chceš OneDrive spustit přes <strong>Microsoft Teams</strong>, <strong>Microsoft Outlook</strong> nebo <strong>Google Chrome</strong>:
                </p>

                <div class="od-launcher-grid">
                    <div class="od-launcher-tile" data-launch-env="teams">
                        <div class="od-desk-icon-badge" style="background:#464eb8;">👥</div>
                        <strong>Microsoft Teams</strong>
                        <small style="color:#cbd5e1; font-size:11px;">Otevře tvůj Teams -> vlevo klikni na <b>☁️ OneDrive</b> nebo na <b>•••</b></small>
                    </div>
                    <div class="od-launcher-tile" data-launch-env="outlook">
                        <div class="od-desk-icon-badge" style="background:#0078d4;">✉️</div>
                        <strong>Microsoft Outlook</strong>
                        <small style="color:#cbd5e1; font-size:11px;">Otevře Outlook -> vlevo klikni na <b>⊞ Aplikace -> OneDrive</b></small>
                    </div>
                    <div class="od-launcher-tile" data-launch-env="chrome">
                        <div class="od-desk-icon-badge" style="background:linear-gradient(135deg,#ea4335,#fbbc05,#34a853,#4285f4);">🌐</div>
                        <strong>Google Chrome</strong>
                        <small style="color:#cbd5e1; font-size:11px;">Otevře webový prohlížeč s adresou OneDrive</small>
                    </div>
                </div>
            </div>
        `;
    },

    renderWin11Taskbar() {
        const s = this.state;
        return `
            <footer class="od-win11-taskbar">
                <div style="display:flex; align-items:center; gap:8px; font-size:12px; cursor:pointer;" data-launch-env="desktop">
                    <span>🌤️ 16 °C Jasno</span>
                </div>
                <div class="od-taskbar-center">
                    <div class="od-tb-icon ${s.environment === "desktop" ? "active" : ""}" data-launch-env="desktop" title="Plocha Windows 11">🪟</div>
                    <div class="od-tb-icon ${s.environment === "teams" ? "active" : ""}" data-launch-env="teams" title="Microsoft Teams">👥</div>
                    <div class="od-tb-icon ${s.environment === "outlook" ? "active" : ""}" data-launch-env="outlook" title="Microsoft Outlook">✉️</div>
                    <div class="od-tb-icon ${s.environment === "chrome" ? "active" : ""}" data-launch-env="chrome" title="Google Chrome">🌐</div>
                </div>
                <div class="od-taskbar-tray">
                    <span>☁️ 🔊 🛜</span>
                    <span>CES</span>
                    <div style="text-align:right; line-height:1.15;">
                        <div>14:32</div>
                        <div style="font-size:10px; color:#94a3b8;">21.10.2026</div>
                    </div>
                </div>
            </footer>
        `;
    },

    renderActiveEnvironmentWindow() {
        const s = this.state;

        if (s.environment === "chrome") {
            return `
                <div class="od-app-window">
                    <div class="od-chrome-tabs">
                        <div class="od-chrome-tab">
                            <span>☁️ Domů - OneDrive</span>
                            <span style="margin-left:auto; color:#5f6368; cursor:pointer;" data-launch-env="desktop">✕</span>
                        </div>
                        <span style="padding:6px; color:#5f6368;">+</span>
                        <button style="margin-left:auto; margin-bottom:4px; border:none; background:transparent; cursor:pointer; font-size:12px;" data-launch-env="desktop">🪟 Plocha Windows 11</button>
                    </div>
                    <div class="od-chrome-navbar">
                        <span style="color:#5f6368; cursor:pointer;">⬅ ➡ ↻</span>
                        <div class="od-chrome-url">
                            <span>🔒</span>
                            <span>firma-priklad-my.sharepoint.com/?login_hint=jana%40example.com&source=waffle</span>
                        </div>
                        <span style="color:#5f6368;">⭐ 🧩 👤 ⋮</span>
                    </div>
                    <div class="od-chrome-bookmarks">
                        <span>⊞ Aplikace</span><span>🔶 Office</span><span>📁 Digitální archiv</span><span>📁 Projekty</span><span>📁 Šablony</span><span>📁 Nástroje</span>
                    </div>
                    <div class="od-web-header">
                        <div style="display:flex; align-items:center; gap:12px; font-weight:600; font-size:15px;">
                            <span>⠿</span><span>OneDrive</span>
                        </div>
                        ${this.renderSearchBar("Všechny soubory")}
                        <div style="display:flex; align-items:center; gap:14px; font-size:14px;">
                            <span>🖥️ ⚙️ ❓</span>
                            <div style="width:28px; height:28px; border-radius:50%; background:#0078d4; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:700;">JT</div>
                        </div>
                    </div>
                    <div class="od-main-layout">
                        ${this.renderOneDriveInnerLayout("chrome")}
                    </div>
                </div>
            `;
        }

        // PROSTŘEDÍ OUTLOOK (Screen 01)
        return `
            <div class="od-app-window">
                <div class="od-outlook-header">
                    <div style="display:flex; align-items:center; gap:12px; font-weight:600; font-size:13px;">
                        <span>📬 🔁 ↩️</span><span>Outlook • OneDrive</span>
                    </div>
                    ${this.renderSearchBar("Hledat")}
                    <div style="display:flex; align-items:center; gap:12px; font-size:12px;">
                        <span>🔔</span>
                        <div style="width:26px; height:26px; border-radius:50%; background:#004578; display:flex; align-items:center; justify-content:center; font-size:11px;">JT</div>
                        <button style="background:transparent; border:1px solid rgba(255,255,255,0.4); color:#fff; border-radius:4px; padding:2px 8px; cursor:pointer;" data-launch-env="desktop">🪟 Plocha</button>
                    </div>
                </div>
                <div class="od-outlook-subbar">
                    <div style="display:flex; gap:16px;"><span>Soubor</span><span>Nápověda</span></div>
                    <div style="display:flex; align-items:center; gap:10px;">
                        <span>Vyzkoušejte nový Outlook 🔘</span>
                        <span style="border:1px solid #d2d0ce; padding:2px 8px; border-radius:4px;">✨ Copilot ▾</span>
                    </div>
                </div>
                <div class="od-main-layout">
                    <aside class="od-host-rail">
                        <div class="od-host-rail-btn" title="Pošta">✉️</div>
                        <div class="od-host-rail-btn" title="Kalendář">📅</div>
                        <div class="od-host-rail-btn" title="Lidé">👥</div>
                        <div class="od-host-rail-btn" title="To Do">✔️</div>
                        <div class="od-host-rail-btn ${s.outlookOneDriveLoaded ? "active" : ""}" id="od-ol-rail-cloud" title="OneDrive">☁️</div>
                        <div class="od-host-rail-btn ${s.outlookAppsFlyoutOpen ? "active" : ""}" id="od-ol-rail-apps" title="Aplikace (OneDrive)">⊞</div>
                    </aside>

                    ${s.outlookAppsFlyoutOpen ? `
                        <div class="od-apps-popover" onclick="event.stopPropagation()">
                            <div style="font-size:12px; font-weight:600; color:#605e5c; margin-bottom:6px;">Aplikace v Outlooku</div>
                            <div class="od-apps-grid">
                                <div class="od-app-tile"><span style="font-size:22px;">👥</span><span>Groups</span></div>
                                <div class="od-app-tile"><span style="font-size:22px;">🟣</span><span>Loop</span></div>
                                <div class="od-app-tile"><span style="font-size:22px;">🌐</span><span>Viva Engage</span></div>
                                <div class="od-app-tile"><span style="font-size:22px;">📅</span><span>Bookings</span></div>
                                <div class="od-app-tile highlight" id="od-ol-pick-onedrive"><span style="font-size:22px;">☁️</span><span>OneDrive</span></div>
                                <div class="od-app-tile"><span style="font-size:22px;">📊</span><span>Power BI</span></div>
                                <div class="od-app-tile"><span style="font-size:22px;">📋</span><span>Planner</span></div>
                                <div class="od-app-tile"><span style="font-size:22px;">🏢</span><span>Org Explorer</span></div>
                            </div>
                        </div>
                    ` : ""}

                    ${this.renderOneDriveInnerLayout("outlook")}
                </div>
            </div>
        `;
    },

    renderSearchBar(scopeLabel) {
        const s = this.state;
        return `
            <div class="od-search-pill">
                <span>🔍</span>
                <input type="text" id="od-top-search-input" value="${s.searchQuery}" placeholder="Hledat dokumenty na OneDrive (Enter)...">
                <span class="od-search-scope" id="od-top-search-btn">${scopeLabel} ▾</span>
            </div>
        `;
    },

    renderOneDriveInnerLayout(hostType) {
        const s = this.state;
        const selectedFile = s.files.find(f => f.id === s.selectedFileId);

        const visibleFiles = s.files.filter(f => {
            if (s.activeSection === "my_files" && f.section !== "my_files" && f.section !== "both") return false;
            if (s.activeSection === "library" && s.activeLibraryName && f.lib !== s.activeLibraryName) return false;
            if (s.typeFilter !== "all" && f.type !== s.typeFilter) return false;
            if (s.searchQuery && !f.name.toLowerCase().includes(s.searchQuery.toLowerCase()) && !f.owner.toLowerCase().includes(s.searchQuery.toLowerCase())) return false;
            if (s.nameFilterQuery && !f.name.toLowerCase().includes(s.nameFilterQuery.toLowerCase()) && !f.owner.toLowerCase().includes(s.nameFilterQuery.toLowerCase())) return false;
            return true;
        });

        return `
            <aside class="od-left-sidebar">
                ${hostType === "teams" ? `<div style="font-weight:700; font-size:16px; color:#242424; margin:0 0 12px 6px;">OneDrive 🖥️</div>` : ""}

                <div style="position:relative;">
                    <button class="od-btn-create" id="od-btn-create-upload">
                        <span>＋</span> Vytvořit nebo nahrát
                    </button>
                    ${s.createMenuOpen ? `
                        <div class="od-context-menu" style="position:absolute; left:0; top:42px; width:215px;" onclick="event.stopPropagation()">
                            <div class="od-ctx-item" data-od-upload="folder"><span>📁 Složka</span></div>
                            <div class="od-ctx-sep"></div>
                            <div class="od-ctx-item" data-od-upload="files"><span>⬆️ <strong>Nahrát soubory</strong></span></div>
                            <div class="od-ctx-item" data-od-upload="folder_up"><span>📂 Nahrát složku</span></div>
                            <div class="od-ctx-sep"></div>
                            <div class="od-ctx-item" data-od-upload="word"><span>📘 Dokument Word</span></div>
                            <div class="od-ctx-item" data-od-upload="excel"><span>📗 Sešit Excel</span></div>
                            <div class="od-ctx-item" data-od-upload="ppt"><span>📙 Prezentace PowerPoint</span></div>
                        </div>
                    ` : ""}
                </div>

                <div style="font-size:12px; font-weight:600; color:#605e5c; padding:4px 12px 8px;">Jana Testovací</div>

                <div class="od-nav-link ${s.activeSection === "home" ? "active" : ""}" data-od-nav="home"><span>🏠</span><span>Domů</span></div>
                <div class="od-nav-link ${s.activeSection === "my_files" ? "active" : ""}" data-od-nav="my_files"><span>📁</span><span>Moje soubory</span></div>
                <div class="od-nav-link ${s.activeSection === "shared" ? "active" : ""}" data-od-nav="shared"><span>👥</span><span>Sdílené</span></div>
                <div class="od-nav-link ${s.activeSection === "favorites" ? "active" : ""}" data-od-nav="favorites"><span>⭐</span><span>Oblíbené</span></div>
                <div class="od-nav-link ${s.activeSection === "libraries_overview" ? "active" : ""}" data-od-nav="libraries_overview"><span>📚</span><span>Knihovny</span></div>
                <div class="od-nav-link" data-od-nav="trash"><span>🗑️</span><span>Koš</span></div>

                <div class="od-nav-section-title">Procházet soubory podle</div>
                <div class="od-nav-link" data-od-nav="people"><span>👤</span><span>Lidé</span></div>
                <div class="od-nav-link" data-od-nav="meetings"><span>📅</span><span>Schůzky</span></div>
                <div class="od-nav-link" data-od-nav="media"><span>🖼️</span><span>Multimédia</span></div>

                <div class="od-nav-section-title">Rychlý přístup</div>
                ${s.libraries.map(lib => `
                    <div class="od-nav-link ${s.activeSection === "library" && s.activeLibraryName === lib.name ? "active" : ""}" data-od-lib="${lib.name}">
                        <span class="od-lib-badge" style="background:${lib.color}">${lib.badge}</span>
                        <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${lib.name}</span>
                    </div>
                `).join("")}
            </aside>

            <main class="od-content-area" style="background:#ffffff; color:#242424;">
                ${s.activeSection === "home" ? `
                    <h3 style="font-size:15px; font-weight:600; margin-bottom:12px;">Pro vás</h3>
                    <div class="od-for-you-grid">
                        ${s.forYouCards.map(c => `
                            <div class="od-for-you-card">
                                <div style="display:flex; align-items:center; gap:8px; font-weight:600; font-size:13px;">
                                    <span>${c.icon}</span><span>${c.title}</span>
                                </div>
                                <div style="font-size:11.5px; color:#605e5c;">👤 ${c.author}<br><small>${c.time}</small></div>
                                <button class="od-card-open-btn" data-od-open-card="${c.title}">Otevřít</button>
                            </div>
                        `).join("")}
                    </div>
                ` : `
                    <h2 style="font-size:18px; font-weight:600; margin-bottom:12px;">
                        ${s.activeSection === "my_files" ? "Moje soubory ▾" : s.activeSection === "library" ? `📚 Knihovna: ${s.activeLibraryName}` : "Sdílené a firemní knihovny"}
                    </h2>
                `}

                <div class="od-command-bar">
                    <div class="od-cmd-left">
                        <button class="od-cmd-btn" id="od-cmd-open"><strong>${selectedFile?.icon || "📗"} Otevřít ▾</strong></button>
                        <button class="od-cmd-btn" id="od-cmd-share">📤 Sdílet</button>
                        <button class="od-cmd-btn" id="od-cmd-copy-link">🔗 Kopírovat odkaz</button>
                        <button class="od-cmd-btn" id="od-cmd-versions">🕒 Historie verzí</button>
                        <button class="od-cmd-btn" id="od-cmd-shortcut">↗️ Přidat zástupce ▾</button>
                        <button class="od-cmd-btn" id="od-cmd-download">⬇️ Stáhnout</button>
                        ${selectedFile?.local ? `<button class="od-cmd-btn" id="od-cmd-cloud" style="color:#0078d4; font-weight:600;">☁️ Uložit do OneDrive (cloud)</button>` : ""}
                    </div>
                    <div style="font-size:12px; color:#605e5c;">✕ Počet vybraných: ${selectedFile ? 1 : 0}</div>
                </div>

                <div class="od-type-pills">
                    <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                        <span style="font-weight:600; font-size:13px; margin-right:6px;">Poslední</span>
                        <button class="od-pill-btn ${s.typeFilter === "all" ? "active" : ""}" data-od-type="all">Vše</button>
                        <button class="od-pill-btn ${s.typeFilter === "word" ? "active" : ""}" data-od-type="word">📘 Word</button>
                        <button class="od-pill-btn ${s.typeFilter === "excel" ? "active" : ""}" data-od-type="excel">📗 Excel</button>
                        <button class="od-pill-btn ${s.typeFilter === "ppt" ? "active" : ""}" data-od-type="ppt">📙 PowerPoint</button>
                        <button class="od-pill-btn ${s.typeFilter === "pdf" ? "active" : ""}" data-od-type="pdf">📕 PDF</button>
                    </div>
                    <div>
                        <input type="text" id="od-filter-name-inp" value="${s.nameFilterQuery}" placeholder="Filtrovat podle jména nebo os..."
                               style="padding:5px 12px; border:1px solid #d2d0ce; border-radius:16px; font-size:12px; width:220px; outline:none;">
                    </div>
                </div>

                <table class="od-files-table">
                    <thead>
                        <tr>
                            <th style="width:32px;"></th>
                            <th>Název ▾</th>
                            <th>${s.activeSection === "my_files" ? "Změněno ↓" : "Otevřené"}</th>
                            <th>${s.activeSection === "my_files" ? "Autor změny" : "Vlastník"}</th>
                            ${s.activeSection === "my_files" ? `<th>Velikost</th><th>Sdílení</th>` : ""}
                            <th>Aktivita</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${visibleFiles.map(f => `
                            <tr class="od-file-row ${s.selectedFileId === f.id ? "selected" : ""}" data-od-file="${f.id}">
                                <td><span class="od-row-check">✓</span></td>
                                <td>
                                    <div style="display:flex; align-items:center; justify-content:space-between; gap:8px;">
                                        <div style="display:flex; align-items:center; gap:10px;">
                                            <span style="font-size:16px;">${f.icon}</span>
                                            <div>
                                                <div style="font-weight:500;">${f.name}</div>
                                                ${s.activeSection !== "my_files" ? `<div style="font-size:11px; color:#605e5c;">${f.lib}</div>` : ""}
                                            </div>
                                        </div>
                                        <div style="display:flex; align-items:center; gap:6px;">
                                            <button class="od-cmd-btn" data-od-dots="${f.id}" title="Další akce (•••)">•••</button>
                                            <button class="od-cmd-btn" data-od-quick-share="${f.id}" title="Sdílet odkaz">📤</button>
                                        </div>
                                    </div>
                                </td>
                                <td>${f.modified}</td>
                                <td>${f.owner}</td>
                                ${s.activeSection === "my_files" ? `<td>${f.size}</td><td>${f.sharing === "Sdíleno" ? "👥 Sdíleno" : "Soukromé"}</td>` : ""}
                                <td style="font-size:12px; color:#605e5c;">${f.activity}</td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            </main>
        `;
    },

    renderContextMenu() {
        const cm = this.state.contextMenu;
        return `
            <div class="od-context-menu" style="left:${cm.x}px; top:${cm.y}px;" onclick="event.stopPropagation()">
                <div class="od-ctx-item" data-od-ctx="open"><span>📗 Otevřít</span><span>▸</span></div>
                <div class="od-ctx-item" data-od-ctx="preview"><span>🔍 Náhled</span></div>
                <div class="od-ctx-sep"></div>
                <div class="od-ctx-item" data-od-ctx="share"><span>📤 <strong>Sdílet</strong></span></div>
                <div class="od-ctx-item" data-od-ctx="copy_link"><span>🔗 <strong>Kopírovat odkaz</strong></span></div>
                <div class="od-ctx-item" data-od-ctx="manage_access"><span>👤 Spravovat přístup</span></div>
                <div class="od-ctx-sep"></div>
                <div class="od-ctx-item" data-od-ctx="delete"><span>🗑️ Odstranit</span></div>
                <div class="od-ctx-item" data-od-ctx="favorite"><span>⭐ Přidat do oblíbených</span></div>
                <div class="od-ctx-item" data-od-ctx="shortcut"><span>↗️ Přidat zástupce</span><span>▸</span></div>
                <div class="od-ctx-sep"></div>
                <div class="od-ctx-item" data-od-ctx="download"><span>⬇️ Stáhnout</span></div>
                <div class="od-ctx-item" data-od-ctx="rename"><span>✏️ Přejmenovat</span></div>
                <div class="od-ctx-item" data-od-ctx="move"><span>📁 Přesunout do</span></div>
                <div class="od-ctx-item" data-od-ctx="copy"><span>📄 Kopírovat do</span></div>
                <div class="od-ctx-item" data-od-ctx="duplicate"><span>📑 Duplikovat</span></div>
                <div class="od-ctx-item" data-od-ctx="offline"><span>📥 Zpřístupnit offline</span></div>
                <div class="od-ctx-item" data-od-ctx="automate"><span>⚙️ Automatizovat</span><span>▸</span></div>
                <div class="od-ctx-sep"></div>
                <div class="od-ctx-item" data-od-ctx="versions"><span>🕒 <strong>Historie verzí</strong></span></div>
                <div class="od-ctx-sep"></div>
                <div class="od-ctx-item" data-od-ctx="details"><span>ℹ️ Podrobnosti</span></div>
            </div>
        `;
    },

    renderUploadModal() {
        return `
            <div class="od-modal-backdrop" onclick="event.stopPropagation()">
                <div class="od-modal-box" style="color:#242424;">
                    <h3 style="margin-bottom:8px;">⬆️ Uložit / Nahrát soubor na OneDrive</h3>
                    <p style="font-size:12.5px; color:#605e5c; margin-bottom:14px;">Vyber pracovní dokument ze svého počítače pro uložení do cloudu OneDrive:</p>
                    <div style="display:flex; flex-direction:column; gap:8px; margin-bottom:18px;">
                        <div class="od-nav-link active" data-od-confirm-upload="Rozpocet_projektu_2026.xlsx" style="border:1px solid #0078d4; background:#eff6fc;">
                            <span>📗</span>
                            <div><strong>Rozpocet_projektu_2026.xlsx</strong><div style="font-size:11px; color:#605e5c;">Pracovní tabulka • 420 kB</div></div>
                        </div>
                        <div class="od-nav-link" data-od-confirm-upload="Prezentace_strategicke_cile.pptx" style="border:1px solid #d2d0ce;">
                            <span>📙</span>
                            <div><strong>Prezentace_strategicke_cile.pptx</strong><div style="font-size:11px; color:#605e5c;">Prezentace • 1,8 MB</div></div>
                        </div>
                    </div>
                    <div style="display:flex; justify-content:flex-end;"><button class="od-card-open-btn" id="od-close-upload">Zrušit</button></div>
                </div>
            </div>
        `;
    },

    renderShareModal() {
        const s = this.state;
        const file = s.files.find(f => f.id === s.selectedFileId) || s.files[1];
        return `
            <div class="od-modal-backdrop" onclick="event.stopPropagation()">
                <div class="od-modal-box" style="color:#242424;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <h3 style="font-size:16px;">📤 Sdílet „${file.name}“</h3>
                        <span style="cursor:pointer;" id="od-close-share">✕</span>
                    </div>
                    <p style="font-size:12px; color:#605e5c; margin-bottom:10px;">Sdílením odkazu na OneDrive pracují všichni nad jedním originálem a <strong>nevznikají duplicity</strong>.</p>
                    <input type="text" value="Petr Vzorový (petr.vzorovy@example.com)" style="width:100%; padding:8px 10px; border:1px solid #8a8886; border-radius:4px; margin-bottom:12px; font-size:12.5px;">
                    <div style="margin-bottom:16px;">
                        <label style="font-size:12px; font-weight:600; display:block; margin-bottom:4px;">Oprávnění odkazu:</label>
                        <select id="od-share-perm-select" style="width:100%; padding:8px; border:1px solid #0078d4; border-radius:4px; font-weight:600;">
                            <option value="edit">✏️ Lidé ve vaší organizaci můžou upravovat</option>
                            <option value="view">👁️ Lidé ve vaší organizaci můžou jen zobrazit (Čtení)</option>
                        </select>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <button class="od-card-open-btn" id="od-confirm-copy-link">🔗 Kopírovat odkaz</button>
                        <button class="od-btn-create" id="od-confirm-send-share" style="margin:0;">Poslat odkaz</button>
                    </div>
                </div>
            </div>
        `;
    },

    renderVersionModal() {
        const s = this.state;
        const file = s.files.find(f => f.id === s.selectedFileId) || s.files[1];
        const versions = file.versions || [
            { ver: "Verze 2.0 (Aktuální)", date: "Dnes 11:20", author: "Jana Testovací", size: "13,1 kB" },
            { ver: "Verze 1.0 (Původní verze)", date: "Včera 17:05", author: "Jana Testovací", size: "12,4 kB" }
        ];
        return `
            <div class="od-modal-backdrop" onclick="event.stopPropagation()">
                <div class="od-modal-box" style="color:#242424;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <h3 style="font-size:16px;">🕒 Historie verzí – ${file.name}</h3>
                        <span style="cursor:pointer;" id="od-close-versions">✕</span>
                    </div>
                    <div style="display:flex; flex-direction:column; gap:8px; margin-bottom:16px;">
                        ${versions.map((v, idx) => `
                            <div style="padding:10px 12px; border:1px solid #e1dfdd; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
                                <div>
                                    <div style="font-weight:600; font-size:13px;">${v.ver}</div>
                                    <div style="font-size:11.5px; color:#605e5c;">${v.date} • ${v.author} • ${v.size}</div>
                                </div>
                                ${idx > 0 ? `<button class="od-card-open-btn" data-od-restore-ver="${v.ver}" style="border-color:#0078d4; color:#0078d4;">↺ Obnovit verzi</button>` : `<span style="font-size:11px; color:#107c10; font-weight:600;">Aktuální</span>`}
                            </div>
                        `).join("")}
                    </div>
                    <div style="display:flex; justify-content:flex-end;"><button class="od-card-open-btn" id="od-close-versions2">Zavřít</button></div>
                </div>
            </div>
        `;
    },

    rerenderCurrentHost() {
        if (this.state.environment === "teams" && window.SandboxModules.teams && window.SandboxModules.teams.state?.activeRail === "onedrive") {
            window.SandboxModules.teams.render();
        } else {
            this.render();
        }
    },

    bindEvents() {
        const c = this.container;
        const s = this.state;

        c.querySelector("#od-root")?.addEventListener("click", () => {
            if (s.contextMenu.open || s.createMenuOpen || s.outlookAppsFlyoutOpen) {
                s.contextMenu.open = false;
                s.createMenuOpen = false;
                s.outlookAppsFlyoutOpen = false;
                this.render();
            }
        });

        c.querySelectorAll("[data-launch-env]").forEach(el => {
            el.addEventListener("click", (e) => {
                e.stopPropagation();
                const env = el.getAttribute("data-launch-env");
                s.environment = env;
                if (env === "teams" && window.SandboxModules.teams) {
                    window.SandboxModules.teams.init(this.container, null, this.callbacks);
                    window.SandboxModules.teams.state.launchedFromDesktop = true;
                    window.SandboxModules.teams.state.appsFlyoutOpen = true; // Rovnou ukážeme ••• s OneDrive v Teams
                    window.SandboxModules.teams.render();
                    return;
                }
                if (env === "outlook" && !s.outlookOneDriveLoaded) {
                    s.outlookAppsFlyoutOpen = true;
                }
                this.render();
            });
        });

        c.querySelector("#od-ol-rail-apps")?.addEventListener("click", (e) => {
            e.stopPropagation();
            s.outlookAppsFlyoutOpen = !s.outlookAppsFlyoutOpen;
            this.render();
        });
        c.querySelector("#od-ol-rail-cloud")?.addEventListener("click", () => {
            s.outlookOneDriveLoaded = true;
            s.outlookAppsFlyoutOpen = false;
            this.render();
        });
        c.querySelector("#od-ol-pick-onedrive")?.addEventListener("click", () => {
            s.outlookOneDriveLoaded = true;
            s.outlookAppsFlyoutOpen = false;
            this.render();
            this.callbacks.onInfo("Aplikace OneDrive byla otevřena v prostředí Microsoft Outlook.");
        });

        this.bindEmbeddedEvents(c);
    },

    // Sdílené události pro vnitřek OneDrive (fungují jak v Chromu a Outlooku, tak uvnitř tvého teams.js!)
    bindEmbeddedEvents(c) {
        const s = this.state;

        c.querySelectorAll("[data-od-nav]").forEach(nav => {
            nav.addEventListener("click", () => {
                const target = nav.getAttribute("data-od-nav");
                s.activeSection = target;
                s.activeLibraryName = null;
                if (target === "libraries_overview") {
                    this.callbacks.onAction({ action: "onedrive_libraries", completed: true });
                } else if (!["home", "my_files", "shared", "favorites"].includes(target)) {
                    this.callbacks.onInfo("Tato sekce je v této ukázce nedostupná.");
                }
                this.rerenderCurrentHost();
            });
        });

        c.querySelectorAll("[data-od-lib]").forEach(libEl => {
            libEl.addEventListener("click", () => {
                s.activeSection = "library";
                s.activeLibraryName = libEl.getAttribute("data-od-lib");
                this.rerenderCurrentHost();
                this.callbacks.onAction({ action: "onedrive_libraries", completed: true });
                this.callbacks.onInfo(`Otevřena týmová knihovna: ${s.activeLibraryName}`);
            });
        });

        c.querySelector("#od-btn-create-upload")?.addEventListener("click", (e) => {
            e.stopPropagation();
            s.createMenuOpen = !s.createMenuOpen;
            this.rerenderCurrentHost();
        });

        c.querySelectorAll("[data-od-upload]").forEach(upItem => {
            upItem.addEventListener("click", (e) => {
                e.stopPropagation();
                s.createMenuOpen = false;
                s.uploadModalOpen = true;
                this.rerenderCurrentHost();
            });
        });

        c.querySelectorAll("[data-od-confirm-upload]").forEach(fileChoice => {
            fileChoice.addEventListener("click", () => {
                const fileName = fileChoice.getAttribute("data-od-confirm-upload");
                s.files.unshift({
                    id: "f_new_" + Date.now(),
                    name: fileName,
                    type: fileName.endsWith(".xlsx") ? "excel" : "ppt",
                    icon: fileName.endsWith(".xlsx") ? "📗" : "📙",
                    modified: "Právě teď",
                    owner: "Jana Testovací",
                    size: "420 kB",
                    sharing: "Soukromé",
                    activity: "Právě nahráno na OneDrive",
                    section: "both",
                    lib: s.activeLibraryName || "Moje soubory"
                });
                s.uploadModalOpen = false;
                this.rerenderCurrentHost();
                this.callbacks.onAction({ action: "onedrive_save_file", completed: true });
                this.callbacks.onInfo(`Soubor '${fileName}' byl úspěšně uložen na OneDrive.`);
            });
        });

        c.querySelector("#od-close-upload")?.addEventListener("click", () => {
            s.uploadModalOpen = false;
            this.rerenderCurrentHost();
        });

        c.querySelectorAll("[data-od-type]").forEach(pill => {
            pill.addEventListener("click", () => {
                s.typeFilter = pill.getAttribute("data-od-type");
                this.rerenderCurrentHost();
                if (s.typeFilter !== "all") {
                    this.callbacks.onAction({ action: "onedrive_search", completed: true });
                }
            });
        });

        const triggerSearch = (val) => {
            s.searchQuery = val;
            this.rerenderCurrentHost();
            if (val.trim().length > 0) {
                this.callbacks.onAction({ action: "onedrive_search", completed: true });
            }
        };

        c.querySelector("#od-top-search-input")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter") triggerSearch(e.target.value);
        });
        c.querySelector("#od-top-search-btn")?.addEventListener("click", () => {
            const val = c.querySelector("#od-top-search-input")?.value || "audit";
            triggerSearch(val);
        });

        c.querySelector("#od-filter-name-inp")?.addEventListener("input", (e) => {
            s.nameFilterQuery = e.target.value;
            this.rerenderCurrentHost();
            const inp = this.container.querySelector("#od-filter-name-inp");
            if (inp) {
                inp.focus();
                inp.setSelectionRange(inp.value.length, inp.value.length);
            }
            if (s.nameFilterQuery.trim().length >= 2) {
                this.callbacks.onAction({ action: "onedrive_search", completed: true });
            }
        });

        c.querySelectorAll("[data-od-file]").forEach(row => {
            const fId = row.getAttribute("data-od-file");
            row.addEventListener("click", () => {
                s.selectedFileId = fId;
                s.contextMenu.open = false;
                this.rerenderCurrentHost();
            });
            row.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                e.stopPropagation();
                s.selectedFileId = fId;
                this.openContextMenu(e.clientX, e.clientY, fId);
            });
        });

        c.querySelectorAll("[data-od-dots]").forEach(btn => {
            btn.addEventListener("click", (e) => {
                e.stopPropagation();
                const fId = btn.getAttribute("data-od-dots");
                s.selectedFileId = fId;
                const rect = btn.getBoundingClientRect();
                this.openContextMenu(rect.left - 140, rect.bottom + 4, fId);
            });
        });

        c.querySelectorAll("[data-od-quick-share]").forEach(btn => {
            btn.addEventListener("click", (e) => {
                e.stopPropagation();
                s.selectedFileId = btn.getAttribute("data-od-quick-share");
                s.shareModalOpen = true;
                this.rerenderCurrentHost();
            });
        });

        c.querySelector("#od-cmd-share")?.addEventListener("click", () => {
            s.shareModalOpen = true;
            this.rerenderCurrentHost();
        });
        c.querySelector("#od-cmd-copy-link")?.addEventListener("click", () => {
            this.completeShareTask("Odkaz na soubor byl zkopírován do schránky.");
        });
        c.querySelector("#od-cmd-versions")?.addEventListener("click", () => {
            s.versionModalOpen = true;
            this.rerenderCurrentHost();
        });
        c.querySelector("#od-cmd-open")?.addEventListener("click", () => {
            const f = s.files.find(x => x.id === s.selectedFileId);
            this.callbacks.onInfo(`Otevírám „${f ? f.name : "soubor"}“ v příslušné aplikaci Microsoft 365 (ukázka).`);
        });
        c.querySelector("#od-cmd-cloud")?.addEventListener("click", () => {
            const f = s.files.find(x => x.id === s.selectedFileId);
            if (!f || !f.local) return;
            f.local = false;
            f.activity = "Uloženo na OneDrive (cloud) · právě teď";
            this.rerenderCurrentHost();
            this.callbacks.onAction({ action: "od_cloud_only", completed: true });
            this.callbacks.onInfo(`Soubor „${f.name}“ se nyní ukládá do cloudu OneDrive, nikoli na lokální disk.`);
        });
        c.querySelector("#od-cmd-download")?.addEventListener("click", () => {
            this.callbacks.onInfo("Stahování je v této ukázce simulováno – soubor by se uložil do složky Stažené.");
        });
        c.querySelectorAll("[data-od-open-card]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.callbacks.onInfo(`Otevírám „${btn.getAttribute("data-od-open-card")}“ (ukázka).`);
            });
        });
        c.querySelectorAll(".od-modal-backdrop").forEach(bd => {
            bd.addEventListener("click", (e) => {
                if (e.target !== bd) return;
                s.uploadModalOpen = false;
                s.shareModalOpen = false;
                s.versionModalOpen = false;
                this.rerenderCurrentHost();
            });
        });
        c.querySelector("#od-cmd-shortcut")?.addEventListener("click", () => {
            this.callbacks.onAction({ action: "onedrive_libraries", completed: true });
            this.callbacks.onInfo("Zástupce na sdílenou knihovnu byl přidán.");
        });

        c.querySelectorAll("[data-od-ctx]").forEach(item => {
            item.addEventListener("click", (e) => {
                e.stopPropagation();
                const act = item.getAttribute("data-od-ctx");
                s.contextMenu.open = false;
                if (act === "share" || act === "manage_access") {
                    s.shareModalOpen = true;
                    this.rerenderCurrentHost();
                } else if (act === "copy_link") {
                    this.completeShareTask("Odkaz na dokument byl zkopírován do schránky.");
                } else if (act === "versions") {
                    s.versionModalOpen = true;
                    this.rerenderCurrentHost();
                } else if (act === "shortcut") {
                    this.rerenderCurrentHost();
                    this.callbacks.onAction({ action: "onedrive_libraries", completed: true });
                } else if (act === "copy" || act === "duplicate") {
                    this.rerenderCurrentHost();
                    this.callbacks.onMistake("Kopie souboru vytváří duplicitu. Místo kopírování použij sdílení odkazem na jeden originál.");
                } else {
                    const infoTexts = {
                        open: "Soubor se otevírá v novém okně (ukázka).",
                        preview: "Náhled souboru je zobrazen (ukázka).",
                        delete: "Soubor byl přesunut do Koše (ukázka).",
                        favorite: "Soubor byl přidán do oblíbených.",
                        download: "Stahování je v této ukázce simulováno.",
                        rename: "Přejmenování je v této ukázce pouze orientační.",
                        move: "Přesun je v této ukázce pouze orientační.",
                        details: "Podrobnosti souboru jsou zobrazeny v panelu (ukázka).",
                        offline: "Soubor bude dostupný offline (ukázka).",
                        automate: "Automatizace je v této ukázce pouze orientační."
                    };
                    this.rerenderCurrentHost();
                    this.callbacks.onInfo(infoTexts[act] || "Tato akce je v ukázce pouze orientační.");
                }
            });
        });

        c.querySelector("#od-confirm-copy-link")?.addEventListener("click", () => {
            s.shareModalOpen = false;
            this.completeShareTask("Odkaz s nastaveným oprávněním byl zkopírován.");
        });
        c.querySelector("#od-confirm-send-share")?.addEventListener("click", () => {
            s.shareModalOpen = false;
            this.completeShareTask("Dokument byl nasdílen přes odkaz bez vytváření duplicit.");
        });
        c.querySelector("#od-close-share")?.addEventListener("click", () => {
            s.shareModalOpen = false;
            this.rerenderCurrentHost();
        });

        c.querySelectorAll("[data-od-restore-ver]").forEach(btn => {
            btn.addEventListener("click", () => {
                const verName = btn.getAttribute("data-od-restore-ver");
                s.versionModalOpen = false;
                this.rerenderCurrentHost();
                this.callbacks.onAction({ action: "onedrive_versioning", completed: true });
                this.callbacks.onInfo(`Předchozí verze (${verName}) byla obnovena!`);
            });
        });
        c.querySelectorAll("#od-close-versions, #od-close-versions2").forEach(b => {
            b.addEventListener("click", () => {
                s.versionModalOpen = false;
                this.rerenderCurrentHost();
                this.callbacks.onAction({ action: "onedrive_versioning", completed: true });
            });
        });
    },

    openContextMenu(x, y, fileId) {
        const clampedX = Math.max(10, Math.min(x, window.innerWidth - 260));
        const clampedY = Math.max(10, Math.min(y, window.innerHeight - 420));
        this.state.contextMenu = { open: true, x: clampedX, y: clampedY, fileId };
        this.rerenderCurrentHost();
    },

    completeShareTask(msg) {
        const file = this.state.files.find(f => f.id === this.state.selectedFileId);
        if (file) file.sharing = "Sdíleno";
        this.rerenderCurrentHost();
        this.callbacks.onAction({ action: "onedrive_share_link", completed: true });
        this.callbacks.onInfo(msg);
    }
};