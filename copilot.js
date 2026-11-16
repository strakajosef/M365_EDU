/**
 * ============================================================================
 * INTERAKTIVNÍ MODUL MICROSOFT COPILOT (copilot.js)
 * ============================================================================
 */

window.SandboxModules = window.SandboxModules || {};

window.SandboxModules.copilot = {
    state: null,
    container: null,
    callbacks: null,

    init(container, mission, callbacks) {
        this.container = container;
        this.callbacks = callbacks;

        this.state = {
            activeNav: "chat", // 'chat' | 'library' | 'agents'
            addContentMenuOpen: false,
            addOptionsOpen: false,
            agentsMenuOpen: false,
            moreChipsOpen: false,
            libFilter: "all",
            searchQuery: "",
            selectedAgent: null,
            messages: [],
            pinnedChats: [
                "Jana Testovací – DEVb",
                "Petr Vzorový"
            ],
            chatHistory: [
                "Chyba None při spojování řetězců",
                "Osnova školení Microsoft Teams",
                "Oprava typové chyby návratu",
                "Prioritní výpočet věku z data narození"
            ],
            libraryItems: [
                "Study Guide: Fill in the Blanks",
                "Study Guide: Quiz",
                "Mám výsledky auditu",
                "HTML šablonátor e-mailů (Ukázka) 2",
                "Souhrn digi učení týmu - výstupy a best practices"
            ],
            promptTemplates: [
                "Napiš shrnutí porady pro tým Firma Příklad s.r.o. v pěti bodech.",
                "Připrav e-mail pro klienta Jana Testovací (jana@example.com) o termínu schůzky.",
                "Vytvoř osnovu interního školení o Microsoft Teams pro 10 účastníků."
            ],
            agentsList: [
                "Audit Digi dovedností",
                "Lékový asistent",
                "Oracle",
                "Python",
                "Strážník Vzorový",
                "Power Automate",
                "Mia",
                "Service"
            ]
        };

        this.render();
    },

    render() {
        const s = this.state;

        this.container.innerHTML = `
            <div class="cp-app" id="cp-root">
                <div class="cp-workspace">
                    <!-- 1. Levý panel -->
                    <aside class="cp-sidebar">
                        <div class="cp-nav-item ${s.activeNav === "chat" && !s.selectedAgent ? "active" : ""}" data-cp-nav="chat">➕ Nový chat</div>
                        <div class="cp-nav-item ${s.activeNav === "search" ? "active" : ""}" data-cp-nav="search">🔍 Vyhledávání</div>
                        <div class="cp-nav-item ${s.activeNav === "library" ? "active" : ""}" data-cp-nav="library">📚 Knihovna</div>
                        <div class="cp-nav-item ${s.agentsMenuOpen || s.selectedAgent ? "active" : ""}" id="cp-agents-btn">🤖 Agenti ▾</div>
                        <div class="cp-nav-item ${s.activeNav === "notebooks" ? "active" : ""}" data-cp-nav="notebooks">📁 Poznámkové bloky</div>

                        <div class="cp-section-label">Připnuto</div>
                        ${s.pinnedChats.map(p => `<div class="cp-chat-history-item" data-cp-hist="${p}">📌 ${p}</div>`).join("")}

                        <div class="cp-section-label" style="margin-top:10px;">Chaty</div>
                        ${s.chatHistory.map(ch => `<div class="cp-chat-history-item" data-cp-hist="${ch}">💬 ${ch}</div>`).join("")}
                    </aside>

                    <!-- 2. Hlavní obsah podle záložky -->
                    ${s.activeNav === "library" ? this.renderLibrary() : s.activeNav === "search" ? this.renderSearch() : s.activeNav === "notebooks" ? this.renderNotebooks() : this.renderChatCanvas()}

                    <!-- Popup menu pro tlačítko '+' (Screen 46) -->
                    ${s.addContentMenuOpen ? `
                        <div class="cp-popup-menu" style="left:310px; bottom:135px;" onclick="event.stopPropagation()">
                            <div class="cp-popup-row" data-cp-action="header"><span>📎 Přidat obsah</span></div>
                            <div class="cp-popup-row" data-cp-action="upload"><span>⬆️ Nahrát obrázky a soubory</span></div>
                            <div class="cp-popup-row" data-cp-action="cloud"><span>☁️ Připojit cloudové soubory</span></div>
                            <div class="cp-popup-sep"></div>
                            <div class="cp-popup-row" data-cp-action="options"><span>⚙️ Přidání možnosti ${s.addOptionsOpen ? "▾" : "▸"}</span></div>
                            ${s.addOptionsOpen ? `
                                <div class="cp-popup-row cp-popup-sub" data-cp-action="opt_web"><span>🌐 Vyhledávání na webu</span></div>
                                <div class="cp-popup-row cp-popup-sub" data-cp-action="opt_image"><span>🎨 Vytvořit obrázek</span></div>
                                <div class="cp-popup-row cp-popup-sub" data-cp-action="opt_internal"><span>🏢 Interní zdroje firmy</span></div>
                            ` : ""}
                            <div class="cp-popup-row" data-cp-action="agent_chat"><span>🤖 Chatovat s agentem</span></div>
                        </div>
                    ` : ""}

                    <!-- Popup menu pro Agenty (Screen 48) -->
                    ${s.agentsMenuOpen ? `
                        <div class="cp-popup-menu" style="left:65px; top:128px;" onclick="event.stopPropagation()">
                            ${s.agentsList.map(ag => `<div class="cp-popup-row" data-cp-agent="${ag}"><span>🔹 ${ag}</span></div>`).join("")}
                            <div class="cp-popup-sep"></div>
                            <div class="cp-popup-row" data-cp-action="new_agent"><span>➕ Nový agent</span></div>
                            <div class="cp-popup-row" data-cp-action="more_agents"><span>🔎 Další agenti</span></div>
                        </div>
                    ` : ""}
                </div>
            </div>
        `;

        this.bindEvents();
    },

    renderChatCanvas() {
        const s = this.state;
        return `
            <main class="cp-main-canvas">
                <div class="cp-greeting" style="font-size:24px; font-weight:600; margin-bottom:20px;">
                    ${s.selectedAgent ? `🤖 Agent: ${s.selectedAgent}` : "S čím vám můžu pomoci?"}
                </div>

                ${s.messages.length > 0 ? `
                    <div style="width:100%; max-width:720px; margin-bottom:16px; display:flex; flex-direction:column; gap:8px;">
                        ${s.messages.map(m => `
                            <div style="padding:10px 14px; border-radius:10px; background:${m.role === "user" ? "#f3f2f1" : "#eef6fc"}; align-self:${m.role === "user" ? "flex-end" : "flex-start"}; max-width:85%;">
                                <strong>${m.role === "user" ? "Vy:" : "Copilot:"}</strong> ${m.text}
                            </div>
                        `).join("")}
                    </div>
                ` : ""}

                <div class="cp-input-container">
                    <div class="cp-input-row">
                        <button class="cp-plus-btn" id="cp-btn-plus" title="Přidat obsah">+</button>
                        <textarea class="cp-textarea" id="cp-prompt-input" rows="1" placeholder="Poslat zprávu Copilotovi (stiskni Enter pro odeslání)..."></textarea>
                        <button class="cp-plus-btn" id="cp-btn-send" title="Odeslat prompt" style="background:#0078d4; color:white;">➤</button>
                        <button class="cp-mic-btn" title="Hlasový vstup">🎤</button>
                    </div>
                </div>
                <div class="cp-chips-row">
                    <div class="cp-chip" data-cp-chip="refine">Upřesnit psaní</div>
                    <div class="cp-chip" data-cp-chip="draft">Koncept obsahu</div>
                    <div class="cp-chip" id="cp-chip-summarize" data-cp-chip="summarize">Shrnout</div>
                    <div class="cp-chip" data-cp-chip="find">Najít obsah</div>
                    <div class="cp-chip" id="cp-chip-more">•••</div>
                    ${s.moreChipsOpen ? `
                        <div class="cp-chip" data-cp-chip="data_analysis">Analyzovat data</div>
                        <div class="cp-chip" data-cp-chip="it_error">Najít návod k chybě</div>
                        <div class="cp-chip" data-cp-chip="html_app">Vytvořit HTML aplikaci</div>
                    ` : ""}
                </div>
            </main>
        `;
    },

    renderLibrary() {
        const s = this.state;
        const tabs = [["all", "Vše"], ["images", "Obrázky"], ["pages", "Stránky"], ["templates", "Šablony promptů"]];
        const tabsHtml = tabs.map(([key, label]) => `<span class="cp-lib-tab ${s.libFilter === key ? "active" : ""}" data-cp-libtab="${key}">${label}</span>`).join("");

        let body;
        if (s.libFilter === "templates") {
            body = `<div class="cp-lib-grid">
                ${s.promptTemplates.map(t => `
                    <div class="cp-lib-card" data-cp-template="${t}">
                        <div style="font-weight:600; font-size:13.5px;">📝 Šablona promptu</div>
                        <div style="font-size:12px; margin-top:6px;">${t}</div>
                    </div>
                `).join("")}
            </div>`;
        } else {
            const items = s.libFilter === "images" ? [] : s.libraryItems;
            body = `<div class="cp-lib-grid">
                ${items.length ? items.map(item => `
                    <div class="cp-lib-card" data-cp-libitem="${item}">
                        <div style="font-weight:600; font-size:13.5px;">📄 ${item}</div>
                        <div style="font-size:11px; color:#605e5c;">Uloženo v knihovně M365 Copilot</div>
                    </div>
                `).join("") : `<div style="color:#605e5c;">Žádné položky v této kategorii.</div>`}
            </div>`;
        }

        return `
            <div class="cp-library-view">
                <div class="cp-lib-title">Knihovna</div>
                <div class="cp-lib-tabs">${tabsHtml}</div>
                ${body}
            </div>
        `;
    },

    renderSearch() {
        const s = this.state;
        const q = s.searchQuery.toLowerCase();
        const all = [...s.pinnedChats, ...s.chatHistory];
        const found = all.filter(t => t.toLowerCase().includes(q));
        return `
            <main class="cp-main-canvas">
                <div style="font-size:24px; font-weight:600; margin-bottom:20px;">Vyhledávání</div>
                <input class="cp-textarea" id="cp-search-input" placeholder="Hledat v chatech..." value="${s.searchQuery}" style="width:100%; max-width:720px; border:1px solid #c8c6c4; border-radius:8px; padding:10px;">
                <div id="cp-search-results" style="width:100%; max-width:720px; margin-top:12px; display:flex; flex-direction:column; gap:6px;">
                    ${found.length ? found.map(t => `<div class="cp-chat-history-item" data-cp-hist="${t}">💬 ${t}</div>`).join("") : `<div style="color:#605e5c;">Nic nenalezeno.</div>`}
                </div>
            </main>
        `;
    },

    renderNotebooks() {
        return `
            <main class="cp-main-canvas">
                <div style="font-size:24px; font-weight:600; margin-bottom:20px;">Poznámkové bloky</div>
                <div style="width:100%; max-width:720px; display:flex; flex-direction:column; gap:6px;">
                    <div class="cp-chat-history-item" data-cp-notebook="Porady – Firma Příklad s.r.o.">📓 Porady – Firma Příklad s.r.o.</div>
                    <div class="cp-chat-history-item" data-cp-notebook="Osobní poznámky Jana Testovací">📓 Osobní poznámky Jana Testovací</div>
                </div>
            </main>
        `;
    },

    bindEvents() {
        const c = this.container;
        const s = this.state;

        c.querySelector("#cp-root")?.addEventListener("click", () => {
            if (s.addContentMenuOpen || s.agentsMenuOpen) {
                s.addContentMenuOpen = false;
                s.agentsMenuOpen = false;
                this.render();
            }
        });

        c.querySelectorAll("[data-cp-nav]").forEach(item => {
            item.addEventListener("click", () => {
                const nav = item.getAttribute("data-cp-nav");
                if (nav === "chat") s.messages = [];
                s.activeNav = nav;
                s.selectedAgent = null;
                s.addContentMenuOpen = false;
                s.agentsMenuOpen = false;
                this.render();
            });
        });

        // Kliknutí na připnutý nebo historický chat (otevře ukázkovou konverzaci)
        c.querySelectorAll("[data-cp-hist]").forEach(item => {
            item.addEventListener("click", (e) => {
                e.stopPropagation();
                const title = item.getAttribute("data-cp-hist");
                s.activeNav = "chat";
                s.selectedAgent = null;
                s.messages = [
                    { role: "user", text: title },
                    { role: "assistant", text: "Ukázková konverzace (fiktivní obsah) k tématu: " + title }
                ];
                this.render();
            });
        });

        c.querySelectorAll("[data-cp-notebook]").forEach(item => {
            item.addEventListener("click", () => {
                this.callbacks.onInfo(`Poznámkový blok otevřen: ${item.getAttribute("data-cp-notebook")}`);
            });
        });

        c.querySelector("#cp-search-input")?.addEventListener("input", (e) => {
            s.searchQuery = e.target.value;
            const list = c.querySelector("#cp-search-results");
            const q = s.searchQuery.toLowerCase();
            const found = [...s.pinnedChats, ...s.chatHistory].filter(t => t.toLowerCase().includes(q));
            list.innerHTML = found.length
                ? found.map(t => `<div class="cp-chat-history-item" data-cp-hist="${t}">💬 ${t}</div>`).join("")
                : `<div style="color:#605e5c;">Nic nenalezeno.</div>`;
            list.querySelectorAll("[data-cp-hist]").forEach(el => el.addEventListener("click", () => {
                s.activeNav = "chat";
                s.messages = [
                    { role: "user", text: el.getAttribute("data-cp-hist") },
                    { role: "assistant", text: "Ukázková konverzace (fiktivní obsah) k tématu: " + el.getAttribute("data-cp-hist") }
                ];
                this.render();
            }));
        });

        c.querySelector("#cp-chip-more")?.addEventListener("click", (e) => {
            e.stopPropagation();
            s.moreChipsOpen = !s.moreChipsOpen;
            this.render();
        });

        c.querySelectorAll("[data-cp-libtab]").forEach(tab => {
            tab.addEventListener("click", () => {
                s.libFilter = tab.getAttribute("data-cp-libtab");
                this.render();
                if (s.libFilter === "templates") {
                    this.callbacks.onAction({ action: "copilot_prompt_library", completed: true });
                }
            });
        });

        // Kliknutí na šablonu promptu vloží text do chatu
        c.querySelectorAll("[data-cp-template]").forEach(card => {
            card.addEventListener("click", () => {
                s.activeNav = "chat";
                s.messages = [];
                this.render();
                const inp = c.querySelector("#cp-prompt-input");
                if (inp) inp.value = card.getAttribute("data-cp-template");
                this.callbacks.onAction({ action: "copilot_prompt_library", completed: true });
                this.callbacks.onInfo("Šablona promptu byla vložena do chatu.");
            });
        });

        c.querySelector("#cp-agents-btn")?.addEventListener("click", (e) => {
            e.stopPropagation();
            s.agentsMenuOpen = !s.agentsMenuOpen;
            s.addContentMenuOpen = false;
            this.render();
        });

        c.querySelector("#cp-btn-plus")?.addEventListener("click", (e) => {
            e.stopPropagation();
            s.addContentMenuOpen = !s.addContentMenuOpen;
            s.agentsMenuOpen = false;
            this.render();
        });

        // Kliknutí na konkrétního Agenta v rozbalovacím seznamu
        c.querySelectorAll("[data-cp-agent]").forEach(agEl => {
            agEl.addEventListener("click", (e) => {
                e.stopPropagation();
                const agentName = agEl.getAttribute("data-cp-agent");
                s.selectedAgent = agentName;
                s.activeNav = "chat";
                s.agentsMenuOpen = false;
                this.render();
                this.callbacks.onAction({ action: "copilot_agent", completed: true });
                this.callbacks.onAction({ action: "copilot_action", completed: true });
                this.callbacks.onInfo(`Spuštěn specializovaný agent: ${agentName}`);
            });
        });

        // Kliknutí na akce v menu '+' nebo '➕ Nový agent'
        c.querySelectorAll("[data-cp-action]").forEach(act => {
            act.addEventListener("click", (e) => {
                e.stopPropagation();
                const type = act.getAttribute("data-cp-action");

                if (type === "options") {
                    s.addOptionsOpen = !s.addOptionsOpen;
                    this.render();
                    return;
                }

                s.addContentMenuOpen = false;
                s.agentsMenuOpen = false;
                this.render();

                if (type === "header") return;

                if (type === "upload" || type === "cloud") {
                    this.callbacks.onAction({ action: "copilot_upload", completed: true });
                } else if (type === "new_agent") {
                    this.callbacks.onAction({ action: "copilot_agent", completed: true });
                    this.callbacks.onAction({ action: "copilot_agent_create", completed: true });
                } else if (type === "agent_chat") {
                    this.callbacks.onAction({ action: "copilot_agent", completed: true });
                }
                this.callbacks.onAction({ action: "copilot_action", completed: true });
                this.callbacks.onInfo(type === "more_agents"
                    ? "Seznam dalších agentů není v této ukázce k dispozici."
                    : "Akce v Copilotu byla úspěšně provedena.");
            });
        });

        // Pilulky pod chatovacím polem (Shrnout, Koncept obsahu...)
        c.querySelectorAll("[data-cp-chip]").forEach(chip => {
            chip.addEventListener("click", () => {
                const chipType = chip.getAttribute("data-cp-chip");
                s.messages.push({ role: "user", text: `Použít rychlou akci: ${chip.textContent}` });
                s.messages.push({ role: "assistant", text: "Zde je vygenerovaný výstup podle vašeho požadavku..." });
                this.render();
                const actionMap = {
                    summarize: "copilot_summarize",
                    draft: "copilot_generate",
                    data_analysis: "copilot_data_analysis",
                    it_error: "copilot_it_error",
                    html_app: "copilot_html_app"
                };
                if (actionMap[chipType]) {
                    this.callbacks.onAction({ action: actionMap[chipType], completed: true });
                } else {
                    this.callbacks.onInfo("Rychlý prompt byl vložen do chatu.");
                }
            });
        });

        // Odeslání vlastního promptu tlačítkem nebo klávesou Enter
        const sendPrompt = () => {
            const inp = c.querySelector("#cp-prompt-input");
            const val = inp?.value.trim();
            if (!val) return;
            s.messages.push({ role: "user", text: val });
            s.messages.push({ role: "assistant", text: "Rozumím zadání. Zde je připravené shrnutí / HTML kód." });
            this.render();
            this.callbacks.onAction({ action: "copilot_summarize", completed: true });
            if (val.length >= 60) {
                this.callbacks.onAction({ action: "copilot_prompt_craft", completed: true });
            }
            if (/html/i.test(val)) {
                this.callbacks.onAction({ action: "copilot_html_app", completed: true });
            }
        };

        c.querySelector("#cp-btn-send")?.addEventListener("click", sendPrompt);
        c.querySelector("#cp-prompt-input")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                sendPrompt();
            }
        });

        // Kliknutí na dlaždici v Knihovně
        c.querySelectorAll("[data-cp-libitem]").forEach(card => {
            card.addEventListener("click", () => {
                const title = card.getAttribute("data-cp-libitem");
                this.callbacks.onAction({ action: "copilot_library", completed: true });
                this.callbacks.onInfo(`Otevřen dokument z knihovny: ${title}`);
            });
        });
    }
};