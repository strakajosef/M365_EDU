/**
 * ============================================================================
 * INTERAKTIVNÍ MODUL MICROSOFT ONENOTE (onenote.js) – PIXEL-PERFECT REPLIKA
 * ============================================================================
 */

window.SandboxModules = window.SandboxModules || {};

window.SandboxModules.onenote = {
    state: null,
    container: null,
    callbacks: null,

    init(container, mission, callbacks) {
        this.container = container;
        this.callbacks = callbacks;

        this.state = {
            activeRibbonTab: "home", // 'home' | 'insert'
            activeNotebook: "Projekt Testovací firma",
            activeSection: "Nový oddíl 3",
            activePage: "Nepojmenovaná stránka",
            searchScope: "Všechny poznámkové bloky",
            searchDropdownOpen: false,
            backstageOpen: false,
            ctxMenu: { open: false, x: 0, y: 0, type: null },
            canvasContent: `<p>Sem napište své poznámky, vložte soubor nebo vytvořte zaškrtávací seznam úkolů (Ctrl+1)...</p>
                            <p style="color:#605e5c; margin-top:12px;">💡 <em>Tip: Klikněte na oddíl, stránku nebo plátno pravým tlačítkem myši pro zobrazení kontextového menu.</em></p>`,

            sections: [
                { name: "Rychlé poznámky", color: "#0078d4" },
                { name: "Nový oddíl 1", color: "#107c10" },
                { name: "Nový oddíl 2", color: "#d83b01" },
                { name: "Nový oddíl 3", color: "#107c10" },
                { name: "Vlastní oddíl+", color: "#7719aa" }
            ],

            pages: [
                "Nepojmenovaná stránka",
                "OneNote0-Základní Zápisky, stránky a sekce",
                "OneNote0-Základní Přidávání textu",
                "OneNote1-Pokročilý Struktura bloků",
                "OneNote2-Specialista Propojení s Teams"
            ]
        };

        this.render();
    },

    render() {
        const s = this.state;

        this.container.innerHTML = `
            <div class="no-app" id="no-root">
                <!-- 1. FIALOVÝ TITLEBAR -->
                <header class="no-titlebar">
                    <div class="no-titlebar-left">
                        <span>🗂️ ↩ ↪</span>
                        <span>${s.activePage} - OneNote</span>
                    </div>
                    <div class="no-search-box" id="no-search-trigger" style="position:relative; cursor:pointer;">
                        <span>🔍</span>
                        <input type="text" value="${s.searchScope}" readonly style="cursor:pointer;">
                        <span style="font-size:10px;">▾</span>

                        ${s.searchDropdownOpen ? `
                            <div class="no-ctx-menu" style="position:absolute; left:0; top:30px; width:260px;" onclick="event.stopPropagation()">
                                <div class="no-ctx-row" data-no-searchscope="Najít na této stránce"><span>Najít na této stránce (Ctrl+F)</span></div>
                                <div class="no-ctx-row" data-no-searchscope="Tento oddíl"><span>Tento oddíl</span></div>
                                <div class="no-ctx-row" data-no-searchscope="Tento poznámkový blok"><span>Tento poznámkový blok</span></div>
                                <div class="no-ctx-row" data-no-searchscope="Všechny poznámkové bloky"><span>✓ Všechny poznámkové bloky</span></div>
                            </div>
                        ` : ""}
                    </div>
                    <div class="no-titlebar-right">
                        <button class="no-sticky-pill">📝 Sticky Notes</button>
                        <button class="no-share-pill" id="no-btn-share">Sdílet ▾</button>
                        <div style="width:24px; height:24px; border-radius:50%; background:#004578; color:white; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:600;">JT</div>
                    </div>
                </header>

                <!-- 2. PÁS KARET (RIBBON) -->
                <nav class="no-menubar">
                    <div class="no-menu-tab" id="no-file-btn" style="background:#7719aa; color:white; font-weight:600; border-radius:2px; height:24px;">Soubor</div>
                    <div class="no-menu-tab ${s.activeRibbonTab === "home" ? "active" : ""}" data-no-rtab="home">Domů</div>
                    <div class="no-menu-tab ${s.activeRibbonTab === "insert" ? "active" : ""}" data-no-rtab="insert">Vložení</div>
                    <div class="no-menu-tab ${s.activeRibbonTab === "draw" ? "active" : ""}" data-no-rtab="draw">Kreslení</div>
                    <div class="no-menu-tab ${s.activeRibbonTab === "history" ? "active" : ""}" data-no-rtab="history">Historie</div>
                    <div class="no-menu-tab ${s.activeRibbonTab === "review" ? "active" : ""}" data-no-rtab="review">Revize</div>
                    <div class="no-menu-tab ${s.activeRibbonTab === "view" ? "active" : ""}" data-no-rtab="view">Zobrazení</div>
                </nav>

                <div class="no-ribbon">
                    ${s.activeRibbonTab === "home" ? `
                        <div class="no-rgroup">
                            <div class="no-rgroup-content">
                                <button class="no-rbtn-lg" id="no-ribbon-paste"><span>📋</span><span>Vložit</span></button>
                            </div>
                            <div class="no-rgroup-title">Schránka</div>
                        </div>
                        <div class="no-rgroup">
                            <div class="no-rgroup-content" style="gap:6px;">
                                <select style="font-size:11px; height:24px; border:1px solid #d2d0ce;"><option>Calibri</option><option>Segoe UI</option></select>
                                <select style="font-size:11px; height:24px; border:1px solid #d2d0ce;"><option>11</option><option>12</option></select>
                                <button class="no-rbtn-lg" id="no-task-tag" style="background:#f3f2f1; border-color:#7719aa;" title="Úkol (Ctrl+1)">
                                    <span style="font-size:16px;">☑️</span><span>Úkol (Ctrl+1)</span>
                                </button>
                            </div>
                            <div class="no-rgroup-title">Základní text & Značky</div>
                        </div>
                        <div class="no-rgroup">
                            <div class="no-rgroup-content">
                                <button class="no-rbtn-lg" id="no-ribbon-meeting"><span>📅</span><span>Podrobnosti<br>schůzky</span></button>
                                <button class="no-rbtn-lg" id="no-ribbon-outlook"><span>✉️</span><span>Odeslat<br>stránku</span></button>
                                <button class="no-rbtn-lg" id="no-ribbon-teams"><span>🟣</span><span>Sdílet do<br>Teams</span></button>
                            </div>
                            <div class="no-rgroup-title">Propojení s Outlookem & Teams</div>
                        </div>
                    ` : s.activeRibbonTab === "insert" ? `
                        <!-- Karta VLOŽENÍ (Screen 10) -->
                        <div class="no-rgroup">
                            <div class="no-rgroup-content">
                                <button class="no-rbtn-lg" data-no-insert="table"><span>⊞</span><span>Tabulka</span></button>
                            </div>
                            <div class="no-rgroup-title">Tabulky</div>
                        </div>
                        <div class="no-rgroup">
                            <div class="no-rgroup-content">
                                <button class="no-rbtn-lg" data-no-insert="printout" style="background:#f3f2f1; border-color:#7719aa;"><span>📄</span><span>Výtisk<br>souboru</span></button>
                                <button class="no-rbtn-lg" data-no-insert="attachment" style="background:#f3f2f1; border-color:#7719aa;"><span>📎</span><span>Soubor jako<br>příloha</span></button>
                                <button class="no-rbtn-lg" data-no-insert="excel"><span>📊</span><span>Tabulka<br>Excelu</span></button>
                            </div>
                            <div class="no-rgroup-title">Soubory</div>
                        </div>
                        <div class="no-rgroup">
                            <div class="no-rgroup-content">
                                <button class="no-rbtn-lg" data-no-insert="picture"><span>🖼️</span><span>Obrázky</span></button>
                                <button class="no-rbtn-lg" data-no-insert="link"><span>🔗</span><span>Odkaz</span></button>
                            </div>
                            <div class="no-rgroup-title">Obrázky & Odkazy</div>
                        </div>
                    ` : this.renderSimpleRibbon(s.activeRibbonTab)}
                </div>

                <!-- 3. PRACOVNÍ PLOCHA -->
                <div class="no-workspace">
                    <aside class="no-notebook-pane">
                        <div class="no-nb-header" id="no-open-backstage">
                            <span>📖</span>
                            <span style="flex:1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${s.activeNotebook}</span>
                            <span>⌄</span>
                        </div>
                        ${s.sections.map(sec => `
                            <div class="no-section-tab ${s.activeSection === sec.name ? "active" : ""}"
                                 data-no-section="${sec.name}">
                                <div class="no-section-color" style="background:${sec.color};"></div>
                                <span style="flex:1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${sec.name}</span>
                            </div>
                        `).join("")}
                    </aside>

                    <aside class="no-pages-pane">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                            <button id="no-btn-add-page" style="width:100%; padding:5px 10px; background:#ffffff; border:1px solid #7719aa; border-radius:3px; color:#7719aa; font-weight:600; cursor:pointer;">+ Přidat stránku</button>
                        </div>
                        ${s.pages.map(pg => `
                            <div class="no-page-card ${s.activePage === pg ? "active" : ""}"
                                 data-no-page="${pg}">
                                <div style="font-size:12px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${pg}</div>
                            </div>
                        `).join("")}
                    </aside>

                    <main class="no-canvas" id="no-main-canvas">
                        <div class="no-canvas-header">
                            <span>úterý 20. října 2026 &nbsp;&nbsp; 15:03</span>
                            <span>Zápisky a poznámky OneNote</span>
                        </div>
                        <div style="font-size:18px; font-weight:600; margin-bottom:12px;" contenteditable="true">${s.activePage}</div>
                        <div style="min-height:350px; outline:none; font-size:13px; line-height:1.6;" contenteditable="true">
                            ${s.canvasContent}
                        </div>
                    </main>

                    ${s.backstageOpen ? this.renderBackstage() : ""}
                </div>

                ${s.ctxMenu.open ? this.renderContextMenu() : ""}
            </div>
        `;

        this.bindEvents();
    },

    renderSimpleRibbon(tab) {
        const groups = {
            draw: [["✏️", "Pero"], ["🖍️", "Zvýrazňovač"], ["🧽", "Guma"], ["📐", "Tvar"]],
            history: [["↶", "Zpět"], ["↷", "Znovu"], ["🕘", "Historie"]],
            review: [["🔤", "Pravopis"], ["🌐", "Přeložit"], ["💬", "Nový komentář"]],
            view: [["🔍", "Přiblížit"], ["📏", "Pravítko"], ["🌙", "Tmavý režim"]]
        };
        const buttons = groups[tab] || [];
        return `
            <div class="no-ribbon-groups">
                ${buttons.map(([icon, label], i) => `
                    <button class="no-rbtn-lg" data-no-deco="${tab}-${i}"><span>${icon}</span><span>${label}</span></button>
                `).join("")}
            </div>
        `;
    },

    renderBackstage() {
        return `
            <div class="no-backstage" onclick="event.stopPropagation()">
                <aside class="no-bs-sidebar">
                    <div class="no-bs-back" id="no-bs-back-btn">←</div>
                    <div class="no-bs-item active">🏠 Informace</div>
                    <div class="no-bs-item" data-no-bsaction="Nový">📂 Nový</div>
                    <div class="no-bs-item" data-no-bsaction="Otevřít">📁 Otevřít</div>
                    <div class="no-bs-item" id="no-bs-share-action">🔗 Sdílet blok</div>
                    <div class="no-bs-item" id="no-bs-teams-action">🟣 Sdílet do kanálu Teams</div>
                    <div class="no-bs-item" data-no-bsaction="Uložit kopii">💾 Uložit kopii</div>
                    <div class="no-bs-item" data-no-bsaction="Tisk">🖨️ Tisk</div>
                    <div class="no-bs-item" data-no-bsaction="Exportovat">📤 Exportovat</div>
                    <div class="no-bs-item" data-no-bsaction="Možnosti">⚙️ Možnosti</div>
                </aside>
                <main class="no-bs-content">
                    <h2 style="font-size:20px; font-weight:600; margin-bottom:18px;">Informace o poznámkovém bloku</h2>
                    <div style="display:flex; gap:16px; align-items:flex-start;">
                        <div style="background:#fff; border:1px solid #c8c6c4; padding:16px; border-radius:4px; width:220px; text-align:center;">
                            <span style="font-size:32px; color:#7719aa;">📓</span>
                            <div style="font-weight:600; margin-top:6px;">Projekt Testovací firma</div>
                            <div style="font-size:11px; color:#605e5c; margin-top:4px;">Uloženo na OneDrive / SharePoint</div>
                        </div>
                        <div>
                            <button id="no-bs-share-link" style="padding:6px 12px; background:#fff; border:1.5px solid #7719aa; border-radius:3px; font-weight:600; color:#7719aa; cursor:pointer;">Pozvěte ostatní do tohoto poznámkového bloku</button>
                            <button id="no-bs-teams-link" style="padding:6px 12px; margin-left:8px; background:#464eb8; border:none; border-radius:3px; font-weight:600; color:#fff; cursor:pointer;">🟣 Připnout do kanálu MS Teams</button>
                            <p style="color:#605e5c; margin-top:8px; font-size:11.5px;">Zde můžete spravovat strukturu bloků, oprávnění a synchronizaci s cloudem.</p>
                        </div>
                    </div>
                </main>
            </div>
        `;
    },

    renderContextMenu() {
        const cm = this.state.ctxMenu;
        return `
            <div class="no-ctx-menu" style="left:${Math.min(cm.x, window.innerWidth - 270)}px; top:${Math.min(cm.y, window.innerHeight - 240)}px;" onclick="event.stopPropagation()">
                ${cm.type === "section" ? `
                    <div class="no-ctx-row" data-no-ctxaction="new_sec"><span>➕ Nový oddíl</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="rename"><span>Přejmenovat</span></div>
                    <div class="no-ctx-sep"></div>
                    <div class="no-ctx-row" data-no-ctxaction="new_group"><span>Nová skupina oddílu</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="color"><span>Barva oddílu</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="move_up"><span>Přesunout nahoru</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="move_down"><span>Přesunout dolů</span></div>
                    <div class="no-ctx-sep"></div>
                    <div class="no-ctx-row" data-no-ctxaction="protect"><span>Chránit tento oddíl heslem...</span></div>
                ` : cm.type === "page" ? `
                    <div class="no-ctx-row" data-no-ctxaction="rename"><span>Přejmenovat</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="copilot_add"><span>✨ Přidat do nového zápisníku Copilota</span></div>
                    <div class="no-ctx-sep"></div>
                    <div class="no-ctx-row" data-no-ctxaction="copy_link"><span>Kopírovat odkaz na stránku</span></div>
                ` : `
                    <div class="no-ctx-row" data-no-ctxaction="format_text"><span>🔤 Formátovat text / Vložit odkaz</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="copilot_add"><span>✨ Přidat do nového zápisníku Copilota</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="ocr_text"><span>🔍 Kopírovat text z obrázku (OCR)</span></div>
                    <div class="no-ctx-sep"></div>
                    <div class="no-ctx-row" data-no-ctxaction="cut"><span>Vyjmout</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="copy"><span>Kopírovat</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="link"><span>Odkaz… (Ctrl+K)</span></div>
                    <div class="no-ctx-row" data-no-ctxaction="translate"><span>Přeložit</span></div>
                `}
            </div>
        `;
    },

    bindEvents() {
        const c = this.container;
        const s = this.state;

        c.querySelector("#no-root")?.addEventListener("click", () => {
            if (s.searchDropdownOpen || s.ctxMenu.open) {
                s.searchDropdownOpen = false;
                s.ctxMenu.open = false;
                this.render();
            }
        });

        // Přepínání karet Domů / Vložení
        c.querySelectorAll("[data-no-rtab]").forEach(tab => {
            tab.addEventListener("click", () => {
                s.activeRibbonTab = tab.getAttribute("data-no-rtab");
                this.render();
            });
        });

        // Akce na kartě Vložení (Příloha, Výtisk, Odkaz)
        c.querySelectorAll("[data-no-insert]").forEach(btn => {
            btn.addEventListener("click", () => {
                const type = btn.getAttribute("data-no-insert");
                if (type === "link") {
                    s.canvasContent += `<p>🔗 <a href="https://www.example.com/dokumentace" style="color:#0078d4;">Odkaz na projektovou dokumentaci</a></p>`;
                    this.render();
                    this.callbacks.onAction({ action: "format_text_link", completed: true });
                } else if (type === "printout" || type === "attachment") {
                    s.canvasContent += `<div style="padding:8px 12px; background:#f3f2f1; border:1px solid #d2d0ce; border-radius:4px; margin-top:8px; display:inline-block;">📎 Příloha: <strong>Specifikace_projektu.pdf</strong></div>`;
                    this.render();
                    this.callbacks.onAction({ action: "insert_files", completed: true });
                    this.callbacks.onInfo("Soubor byl úspěšně vložen na stránku OneNotu.");
                } else if (type === "table") {
                    s.canvasContent += `<table style="border-collapse:collapse; margin-top:8px;"><tr><td style="border:1px solid #888; padding:4px 10px;">Položka</td><td style="border:1px solid #888; padding:4px 10px;">Stav</td></tr><tr><td style="border:1px solid #888; padding:4px 10px;">&nbsp;</td><td style="border:1px solid #888; padding:4px 10px;">&nbsp;</td></tr></table>`;
                    this.render();
                    this.callbacks.onInfo("Tabulka byla vložena na stránku.");
                } else if (type === "picture") {
                    s.canvasContent += `<div style="width:160px; height:90px; border:1px dashed #888; margin-top:8px; display:flex; align-items:center; justify-content:center; color:#666;">🖼️ Obrázek (ukázka)</div>`;
                    this.render();
                    this.callbacks.onInfo("Obrázek byl vložen jako ukázka.");
                } else {
                    this.render();
                    this.callbacks.onInfo("Tabulku Excelu lze vložit po připojení k účtu Microsoft 365.");
                }
            });
        });

        // Vyhledávání
        c.querySelector("#no-search-trigger")?.addEventListener("click", (e) => {
            e.stopPropagation();
            s.searchDropdownOpen = !s.searchDropdownOpen;
            this.render();
        });

        c.querySelectorAll("[data-no-searchscope]").forEach(el => {
            el.addEventListener("click", () => {
                s.searchScope = el.getAttribute("data-no-searchscope");
                s.searchDropdownOpen = false;
                this.render();
                if (s.searchScope === "Všechny poznámkové bloky") {
                    this.callbacks.onAction({ action: "search_notebooks", completed: true });
                }
                this.callbacks.onInfo(`Rozsah hledání nastaven na: ${s.searchScope}`);
            });
        });

        // Otevření Backstage menu Soubor -> Informace (splní i úkol Struktura bloků)
        const openBackstage = () => {
            s.backstageOpen = true;
            this.render();
            this.callbacks.onAction({ action: "notebook_structure", completed: true });
        };
        c.querySelector("#no-file-btn")?.addEventListener("click", openBackstage);
        c.querySelector("#no-open-backstage")?.addEventListener("click", openBackstage);
        c.querySelector("#no-bs-back-btn")?.addEventListener("click", () => {
            s.backstageOpen = false;
            this.render();
        });

        // Oddíly vlevo
        c.querySelectorAll("[data-no-section]").forEach(el => {
            el.addEventListener("click", () => {
                const secName = el.getAttribute("data-no-section");
                if (secName === "Vlastní oddíl+") {
                    const newName = "Nový oddíl " + s.sections.length;
                    s.sections.splice(s.sections.length - 1, 0, { name: newName, color: "#0078d4" });
                    s.activeSection = newName;
                    this.render();
                    this.callbacks.onAction({ action: "create_section_page", completed: true });
                } else {
                    s.activeSection = secName;
                    this.render();
                }
            });
            el.addEventListener("contextmenu", (e) => {
                e.preventDefault(); e.stopPropagation();
                s.ctxMenu = { open: true, x: e.clientX, y: e.clientY, type: "section" };
                this.render();
            });
        });

        // Stránky
        c.querySelectorAll("[data-no-page]").forEach(el => {
            el.addEventListener("click", () => {
                s.activePage = el.getAttribute("data-no-page");
                this.render();
            });
            el.addEventListener("contextmenu", (e) => {
                e.preventDefault(); e.stopPropagation();
                s.ctxMenu = { open: true, x: e.clientX, y: e.clientY, type: "page" };
                this.render();
            });
        });

        c.querySelector("#no-main-canvas")?.addEventListener("contextmenu", (e) => {
            e.preventDefault(); e.stopPropagation();
            s.ctxMenu = { open: true, x: e.clientX, y: e.clientY, type: "canvas" };
            this.render();
        });

        c.querySelector("#no-btn-add-page")?.addEventListener("click", () => {
            const newPage = "Nová poznámka " + (s.pages.length + 1);
            s.pages.push(newPage);
            s.activePage = newPage;
            this.render();
            this.callbacks.onAction({ action: "create_section_page", completed: true });
        });

        c.querySelector("#no-task-tag")?.addEventListener("click", () => {
            s.canvasContent += `<p>☑️ <span contenteditable="true">Nový úkol ke splnění (Ctrl+1)</span></p>`;
            this.render();
            this.callbacks.onAction({ action: "organize_tasks", completed: true });
            this.callbacks.onInfo("Zaškrtávací seznam úkolů (Ctrl+1) byl úspěšně vložen.");
        });

        c.querySelector("#no-ribbon-meeting")?.addEventListener("click", () => {
            s.canvasContent += `<div style="border-left:3px solid #7719aa; padding-left:10px; margin-top:10px;"><strong>📅 Zápis ze schůzky:</strong> Porada týmu (19.10.2026)</div>`;
            this.render();
            this.callbacks.onAction({ action: "meeting_templates", completed: true });
            this.callbacks.onInfo("Podrobnosti schůzky z kalendáře byly vloženy do stránky.");
        });

        c.querySelector("#no-ribbon-outlook")?.addEventListener("click", () => {
            this.callbacks.onAction({ action: "outlook_integration", completed: true });
            this.callbacks.onInfo("Stránka OneNotu byla úspěšně odeslána do Outlooku.");
        });

        c.querySelectorAll("#no-ribbon-teams, #no-bs-teams-action").forEach(btn => {
            btn.addEventListener("click", () => {
                s.backstageOpen = false;
                this.render();
                this.callbacks.onAction({ action: "teams_integration", completed: true });
                this.callbacks.onInfo("Poznámkový blok byl propojen s kanálem v Microsoft Teams.");
            });
        });

        c.querySelector("#no-bs-teams-link")?.addEventListener("click", () => {
            s.backstageOpen = false;
            this.render();
            this.callbacks.onAction({ action: "teams_channel_pin", completed: true });
            this.callbacks.onInfo("Stránka byla připnuta do kanálu MS Teams.");
        });

        c.querySelectorAll("[data-no-deco]").forEach(el => {
            el.addEventListener("click", () => {
                this.callbacks.onInfo(`Příkaz '${el.textContent.trim()}' je v této ukázce pouze informativní.`);
            });
        });

        c.querySelectorAll("[data-no-bsaction]").forEach(el => {
            el.addEventListener("click", () => {
                const name = el.getAttribute("data-no-bsaction");
                this.callbacks.onInfo(`Akce '${name}' je v této ukázce pouze informativní.`);
            });
        });

        c.querySelectorAll("[data-no-ctxaction]").forEach(act => {
            act.addEventListener("click", (e) => {
                e.stopPropagation();
                const type = act.getAttribute("data-no-ctxaction");
                s.ctxMenu.open = false;
                this.render();

                if (type === "format_text") {
                    s.canvasContent += `<p><strong>Tučný text</strong> <span style="background:#ffff00;">zvýrazněný</span> <a href="https://www.example.com" style="color:#0078d4;">odkaz</a></p>`;
                    this.render();
                    this.callbacks.onAction({ action: "format_text_link", completed: true });
                } else if (type === "new_sec") {
                    this.callbacks.onAction({ action: "create_section_page", completed: true });
                } else {
                    this.callbacks.onInfo(`Akce '${type}' byla provedena.`);
                }
            });
        });

        c.querySelectorAll("#no-btn-share, #no-bs-share-link, #no-bs-share-action").forEach(b => {
            b.addEventListener("click", () => {
                s.backstageOpen = false;
                this.render();
                this.callbacks.onAction({ action: "share_notebook", completed: true });
                this.callbacks.onInfo("Pozvánka ke sdílení poznámkového bloku byla vygenerována.");
            });
        });
    }
};