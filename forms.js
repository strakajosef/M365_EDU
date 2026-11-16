/**
 * ============================================================================
 * INTERAKTIVNÍ MODUL MICROSOFT FORMS (forms.js)
 * ============================================================================
 */

window.SandboxModules = window.SandboxModules || {};

window.SandboxModules.forms = {
    state: null,
    container: null,
    callbacks: null,

    init(container, mission, callbacks) {
        this.container = container;
        this.callbacks = callbacks;

        // Pokud je úkol "Vyplnění dotazníku", začínáme rovnou jako respondent
        const isRespondent = mission.targetCondition.action === "forms_fill_quiz";

        this.state = {
            view: isRespondent ? "respondent" : "home", // 'home' | 'editor' | 'respondent'
            settingsOpen: false,
            shareModalOpen: false,
            previewFrom: null, // 'editor' pokud je otevřen náhled z editoru
            ctxMenu: null,     // { x, y, index } pro kontextové menu otázky
            formTitle: isRespondent ? "Víte víc? Otestujte se v kvízu" : "Průzkum spokojenosti",
            questions: [
                { id: 1, type: "choice", title: "Jak jste byli spokojeni s průběhem akce?", options: ["Velmi spokojen(a)", "Spíše spokojen(a)", "Nespokojen(a)"] }
            ],
            answers: {} // Pro ukládání odpovědí respondenta
        };

        // Pro respondenta přidáme více testovacích otázek
        if (isRespondent) {
            this.state.questions = [
                { id: 1, type: "choice", title: "Kdy byla založena společnost Microsoft?", options: ["1975", "1985", "1995"] },
                { id: 2, type: "choice", title: "Který z těchto nástrojů neslouží primárně ke komunikaci?", options: ["MS Teams", "Skype", "MS Word"] },
                { id: 3, type: "choice", title: "Co znamená zkratka OneDrive?", options: ["Cloudové úložiště dat", "Herní konzole", "Grafický editor"] }
            ];
        }

        this.render();
    },

    render() {
        const s = this.state;

        this.container.innerHTML = `
            <div class="fm-app" id="fm-root">
                <header class="fm-header">
                    <div class="fm-header-left" data-fm-action="home">
                        <span style="font-size:18px;">F</span> Microsoft Forms
                    </div>
                    <div class="fm-header-right">
                        ${s.view === "editor" ? `
                            <button class="fm-btn-header" data-fm-action="style" title="Styl">🎨 Styl</button>
                            <button class="fm-btn-header" data-fm-action="settings" title="Nastavení">⚙️ Nastavení</button>
                            <button class="fm-btn-header" data-fm-action="preview" title="Náhled">👁️ Náhled</button>
                            <button class="fm-btn-share" data-fm-action="share">▷ Shromažďovat odpovědi</button>
                            <button class="fm-btn-header" data-fm-action="view_responses" title="Zobrazit odpovědi">📊 Zobrazit odpovědi</button>
                            <button class="fm-btn-header" data-fm-action="present" title="Prezentovat">🖥️ Prezentovat</button>
                        ` : ""}
                        <span style="margin-left:8px; font-weight:600; border:1px solid rgba(255,255,255,0.5); border-radius:50%; width:28px; height:28px; display:inline-flex; align-items:center; justify-content:center;">MS</span>
                    </div>
                </header>

                <main class="fm-body">
                    ${s.view === "home" ? this.renderHome() : ""}
                    ${s.view === "editor" ? this.renderEditor() : ""}
                    ${s.view === "respondent" ? this.renderRespondent() : ""}
                </main>

                ${s.ctxMenu ? this.renderContextMenu() : ""}

                <!-- Pravý panel Nastavení -->
                ${s.settingsOpen ? this.renderSettings() : ""}

                <!-- Modální okno Sdílení -->
                ${s.shareModalOpen ? this.renderShareModal() : ""}
            </div>
        `;

        this.bindEvents();
    },

    renderHome() {
        return `
            <div class="fm-home">
                <div class="fm-home-title">Vítejte, Jano Testovací</div>
                <div class="fm-templates">
                    <div class="fm-template-card" data-fm-action="new_form">
                        <span style="font-size:32px; color:#0f766e;">➕</span>
                        <strong style="color:#0f766e;">Nový formulář</strong>
                    </div>
                    <div class="fm-template-card" data-fm-action="new_quiz">
                        <span style="font-size:32px; color:#0f766e;">📋</span>
                        <strong style="color:#0f766e;">Nový kvíz</strong>
                    </div>
                    <div class="fm-template-card" data-fm-action="import">
                        <span style="font-size:32px;">📊</span>
                        <strong>Rychlý import</strong>
                    </div>
                </div>
                <h3>Moje formuláře</h3>
                <div class="fm-home-item" data-fm-action="open_form">
                    <div style="font-weight:600; margin-bottom:4px;">Dotazník spokojenosti IT 2026</div>
                    <div style="font-size:11px; color:#605e5c;">14 odpovědí • Změněno včera</div>
                </div>
            </div>
        `;
    },

    renderEditor() {
        const s = this.state;
        return `
            <div class="fm-editor-bg">
                <div class="fm-form-container" onclick="event.stopPropagation()">
                    <div class="fm-form-header-img">
                        <h1 class="fm-form-title" contenteditable="true">${s.formTitle}</h1>
                    </div>
                    <div class="fm-form-desc" contenteditable="true">Popis formuláře...</div>
                    
                    <div>
                        ${s.questions.map((q, i) => `
                            <div class="fm-question" data-fm-q="${i}">
                                <div class="fm-q-title"><span style="color:#0f766e;">${i+1}.</span> ${q.title}</div>
                                ${q.options.map(opt => `
                                    <div class="fm-q-option">
                                        <input type="radio" disabled>
                                        <input type="text" value="${opt}" style="border:none; outline:none; border-bottom:1px solid #edebe9; width:300px;">
                                    </div>
                                `).join("")}
                            </div>
                        `).join("")}
                    </div>

                    <div class="fm-add-new-bar">
                        <button class="fm-btn-add" data-fm-action="add_question">➕ Přidat nový</button>
                        <div style="margin-top:16px; font-weight:600; color:#323130; font-size:12px;">+ Rychlý start s</div>
                        <div class="fm-quick-start">
                            <button class="fm-qs-btn" data-fm-action="add_question">Volba</button>
                            <button class="fm-qs-btn" data-fm-action="add_question">Text</button>
                            <button class="fm-qs-btn" data-fm-action="add_question">Hodnocení</button>
                            <button class="fm-qs-btn" data-fm-action="add_question">Datum</button>
                            <button class="fm-qs-btn" data-fm-action="add_question">Net Promoter Score®</button>
                        </div>
                    </div>
                </div>
            </div>
        `;
    },

    renderRespondent() {
        const s = this.state;
        return `
            ${s.previewFrom ? `
                <div class="fm-preview-bar">
                    <span>Náhled formuláře – odpovědi se neodesílají.</span>
                    <button class="fm-btn-header" data-fm-action="back_editor">← Zpět do editoru</button>
                </div>` : ""}
            <div class="fm-editor-bg">
                <div class="fm-form-container">
                    <div class="fm-form-header-img" style="background:#464eb8;">
                        <h1 class="fm-form-title">${s.formTitle}</h1>
                    </div>
                    <div class="fm-form-desc">Dobrý den, jmenuji se Jana Testovací. Jakmile tento formulář odešlete, vlastník uvidí vaše jméno a e-mailovou adresu.</div>
                    
                    <div>
                        ${s.questions.map((q, i) => `
                            <div class="fm-question">
                                <div class="fm-q-title"><span style="color:#464eb8;">${i+1}.</span> ${q.title}</div>
                                ${q.options.map(opt => `
                                    <label class="fm-q-option fm-q-option-clickable">
                                        <input type="radio" name="q${q.id}" value="${opt}" class="fm-quiz-radio">
                                        <span>${opt}</span>
                                    </label>
                                `).join("")}
                            </div>
                        `).join("")}
                    </div>
                    
                    <div style="padding:24px 32px; background:#f8f9fa;">
                        <button class="fm-btn-share" id="fm-btn-submit-quiz" style="background:#464eb8; border-color:#464eb8;">Odeslat</button>
                    </div>
                </div>
            </div>
        `;
    },

    renderSettings() {
        return `
            <div class="fm-settings-panel" onclick="event.stopPropagation()">
                <div class="fm-settings-top">
                    <span>Nastavení</span>
                    <span style="cursor:pointer;" data-fm-action="close_settings">✕</span>
                </div>
                <div class="fm-settings-body">
                    <div class="fm-settings-group">
                        <strong>Kdo může vyplnit tento formulář</strong>
                        <label class="fm-settings-label"><input type="radio" name="who" data-fm-setting="who"> Odpovědět může kdokoli</label>
                        <label class="fm-settings-label"><input type="radio" name="who" checked data-fm-setting="who"> Odpovědět můžou jenom lidé v mé organizaci</label>
                        <div style="margin-left:24px;">
                            <label class="fm-settings-label"><input type="checkbox" checked data-fm-setting="record_name"> Zaznamenat jméno</label>
                            <label class="fm-settings-label"><input type="checkbox" checked data-fm-setting="one_response"> Jedna odpověď na osobu</label>
                        </div>
                    </div>
                    <div class="fm-settings-group">
                        <strong>Možnosti pro odpovědi</strong>
                        <label class="fm-settings-label"><input type="checkbox" checked data-fm-setting="accept"> Přijímat odpovědi</label>
                        <label class="fm-settings-label"><input type="checkbox" data-fm-setting="start_date"> Počáteční datum</label>
                        <label class="fm-settings-label"><input type="checkbox" data-fm-setting="end_date"> Koncové datum</label>
                        <label class="fm-settings-label"><input type="checkbox" data-fm-setting="shuffle"> Náhodné pořadí otázek</label>
                        <label class="fm-settings-label"><input type="checkbox" data-fm-setting="progress"> Zobrazit indikátor průběhu</label>
                    </div>
                    <div class="fm-settings-group">
                        <strong>Potvrzení o odpovědi</strong>
                        <label class="fm-settings-label"><input type="checkbox" data-fm-setting="email_receipt"> Povolit příjemcům uložit a upravit odpovědi</label>
                        <label class="fm-settings-label"><input type="checkbox" data-fm-setting="email_notification"> Dostávat e-mailové oznámení o každé odpovědi</label>
                    </div>
                </div>
            </div>
        `;
    },

    renderShareModal() {
        return `
            <div class="fm-modal-overlay">
                <div class="fm-share-modal" onclick="event.stopPropagation()">
                    <div class="fm-share-title">
                        <span>Poslat a shromáždit odpovědi</span>
                        <span style="cursor:pointer; color:#605e5c;" data-fm-action="close_share">✕</span>
                    </div>
                    <select style="width:100%; padding:8px; border:1px solid #8a8886; border-radius:4px; margin-bottom:16px;">
                        <option>Odpovědět můžou jenom lidé v organizaci Firma Příklad s.r.o.</option>
                        <option>Odpovědět může kdokoli</option>
                    </select>
                    
                    <div class="fm-share-link-box">
                        <input type="text" value="https://forms.example.com/r/ukazka-2026" readonly>
                        <button data-fm-share-method="copy">Kopírovat odkaz</button>
                    </div>

                    <div style="text-align:center; font-weight:600; color:#605e5c; margin-top:20px;">Nebo poslat pozvánku přes:</div>
                    <div class="fm-share-methods">
                        <div class="fm-share-icon" title="Outlook" data-fm-share-method="outlook">✉️</div>
                        <div class="fm-share-icon" title="Teams" data-fm-share-method="teams" style="color:#464eb8;">T</div>
                        <div class="fm-share-icon" title="QR kód" data-fm-share-method="qr">📱</div>
                    </div>
                    <div class="fm-share-send">
                        <button class="fm-btn-share" data-fm-share-method="send">Poslat</button>
                    </div>
                </div>
            </div>
        `;
    },

    renderContextMenu() {
        const m = this.state.ctxMenu;
        return `
            <div class="fm-ctx-menu" style="left:${m.x}px; top:${m.y}px;">
                <div class="fm-ctx-item" data-fm-ctx="dup">Duplikovat otázku</div>
                <div class="fm-ctx-item" data-fm-ctx="del">Odstranit otázku</div>
            </div>
        `;
    },

    nextId() {
        return Math.max(0, ...this.state.questions.map(q => q.id)) + 1;
    },

    bindEvents() {
        const c = this.container;
        const s = this.state;

        // Kliknutí mimo panel / modál je zavře
        c.querySelector("#fm-root")?.addEventListener("click", () => {
            if (s.settingsOpen || s.shareModalOpen || s.ctxMenu) {
                s.settingsOpen = false;
                s.shareModalOpen = false;
                s.ctxMenu = null;
                this.render();
            }
        });

        // Tlačítka pro akce
        c.querySelectorAll("[data-fm-action]").forEach(btn => {
            btn.addEventListener("click", (e) => {
                e.stopPropagation();
                const action = btn.getAttribute("data-fm-action");
                s.ctxMenu = null;

                if (action === "home") {
                    s.view = "home";
                    s.settingsOpen = false;
                    s.shareModalOpen = false;
                    s.previewFrom = null;
                    this.render();
                } else if (action === "new_form" || action === "new_quiz" || action === "open_form") {
                    s.view = "editor";
                    this.render();
                } else if (action === "import") {
                    this.callbacks.onInfo("Rychlý import není v této ukázce dostupný. Založ formulář přes „Nový formulář“.");
                } else if (action === "style") {
                    this.callbacks.onInfo("Motiv formuláře je v této ukázce pouze náhled, nic se nemění.");
                } else if (action === "present") {
                    this.callbacks.onInfo("Prezentace je v této ukázce pouze simulována.");
                } else if (action === "preview") {
                    s.settingsOpen = false;
                    s.shareModalOpen = false;
                    s.previewFrom = "editor";
                    s.view = "respondent";
                    this.render();
                } else if (action === "back_editor") {
                    s.previewFrom = null;
                    s.view = "editor";
                    this.render();
                } else if (action === "settings") {
                    s.settingsOpen = !s.settingsOpen;
                    s.shareModalOpen = false;
                    this.render();
                } else if (action === "close_settings") {
                    s.settingsOpen = false;
                    this.render();
                } else if (action === "share") {
                    s.shareModalOpen = true;
                    s.settingsOpen = false;
                    this.render();
                } else if (action === "close_share") {
                    s.shareModalOpen = false;
                    this.render();
                } else if (action === "view_responses") {
                    s.settingsOpen = false;
                    s.shareModalOpen = false;
                    this.render();
                    this.callbacks.onAction({ action: "forms_automation", completed: true });
                    this.callbacks.onInfo("Otevřeny výsledky dotazníku. Zatím nejsou žádné odpovědi.");
                } else if (action === "add_question") {
                    s.questions.push({ id: this.nextId(), type: "choice", title: "Nová otázka", options: ["Možnost 1", "Možnost 2"] });
                    this.render();
                    this.callbacks.onAction({ action: "forms_create_questions", completed: true });
                    this.callbacks.onInfo("Do formuláře byla přidána nová otázka.");
                }
            });
        });

        // Kontextové menu (pravé tlačítko) na otázkách v editoru
        c.querySelectorAll("[data-fm-q]").forEach(q => {
            q.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                e.stopPropagation();
                s.settingsOpen = false;
                s.shareModalOpen = false;
                s.ctxMenu = { x: e.clientX, y: e.clientY, index: Number(q.getAttribute("data-fm-q")) };
                this.render();
            });
        });

        c.querySelectorAll("[data-fm-ctx]").forEach(item => {
            item.addEventListener("click", (e) => {
                e.stopPropagation();
                const act = item.getAttribute("data-fm-ctx");
                const idx = s.ctxMenu ? s.ctxMenu.index : -1;
                s.ctxMenu = null;
                if (idx >= 0 && s.questions[idx]) {
                    if (act === "dup") {
                        const src = s.questions[idx];
                        s.questions.splice(idx + 1, 0, { ...src, id: this.nextId(), options: [...src.options] });
                    } else if (act === "del") {
                        s.questions.splice(idx, 1);
                    }
                }
                this.render();
            });
        });

        // Nastavení v administraci (úkol: Administrace dotazníku)
        c.querySelectorAll("[data-fm-setting]").forEach(chk => {
            chk.addEventListener("change", (e) => {
                e.stopPropagation();
                this.callbacks.onAction({ action: "forms_admin_settings", completed: true });
                // E-mailové oznámení je zároveň automatizace (úkol: Automatizace a rozesílka)
                if (chk.getAttribute("data-fm-setting") === "email_notification" && chk.checked) {
                    this.callbacks.onAction({ action: "forms_automation", completed: true });
                }
                this.callbacks.onInfo("Změna v nastavení formuláře byla uložena.");
            });
        });

        // Sdílení a shromažďování (úkol: Automatizace a rozesílka)
        c.querySelectorAll("[data-fm-share-method]").forEach(sm => {
            sm.addEventListener("click", (e) => {
                e.stopPropagation();
                s.shareModalOpen = false;
                this.render();
                this.callbacks.onAction({ action: "forms_automation", completed: true });
                this.callbacks.onInfo("Dotazník byl úspěšně připraven ke shromažďování odpovědí.");
            });
        });

        // Odeslání kvízu (úkol: Vyplnění dotazníku)
        c.querySelector("#fm-btn-submit-quiz")?.addEventListener("click", () => {
            if (s.previewFrom) {
                this.callbacks.onInfo("Náhled – odpověď se neodesílá.");
                return;
            }
            const checkedRadios = c.querySelectorAll(".fm-quiz-radio:checked");
            if (checkedRadios.length >= s.questions.length) {
                this.callbacks.onAction({ action: "forms_fill_quiz", completed: true });
            } else {
                this.callbacks.onMistake("Neodpověděl jsi na všechny otázky v kvízu. Vyber odpověď u všech otázek a zkus to znovu.");
            }
        });
    }
};
