/**
 * ============================================================================
 * INTERAKTIVNÍ MODUL MICROSOFT TO DO (todo.js)
 * ============================================================================
 * - Všechna jména, úkoly, seznamy a e-maily jsou smyšlené.
 * - Akce (action): todo_create_task, todo_set_due_date, todo_set_reminder,
 *   todo_set_recurrence, todo_share_list, todo_attach_file, todo_from_teams,
 *   todo_complete_task, todo_toggle_important, todo_add_to_my_day,
 *   todo_create_list, todo_move_task, todo_copy_task, todo_delete_task,
 *   todo_add_note, todo_open_view.
 * - Pravé tlačítko myši (contextmenu) na úkolech a e-mailech.
 */

window.SandboxModules = window.SandboxModules || {};

window.SandboxModules.todo = {
    state: null,
    container: null,
    callbacks: null,
    mission: null,

    TODAY: "2026-11-11",
    TOMORROW: "2026-11-12",
    NEXT_MONDAY: "2026-11-16",

    DUE_OPTS: [
        { v: "today", l: "Dnes (st)" },
        { v: "tomorrow", l: "Zítra (čt)" },
        { v: "next_week", l: "Příští týden (po)" },
        { v: "custom", l: "Vybrat datum" }
    ],
    REMINDER_OPTS: [
        { v: "later_today", l: "Později dnes 18:00" },
        { v: "tomorrow_9", l: "Zítra čt, 9:00" },
        { v: "next_week_9", l: "Příští týden po, 9:00" },
        { v: "custom", l: "Vybrat datum a čas" }
    ],
    REPEAT_OPTS: [
        { v: "daily", l: "Denně" },
        { v: "weekdays", l: "Pracovní dny" },
        { v: "weekly", l: "Týdně" },
        { v: "monthly", l: "Měsíčně" },
        { v: "yearly", l: "Ročně" },
        { v: "custom", l: "Vlastní", sep: true }
    ],

    init(container, mission, callbacks) {
        this.container = container;
        this.callbacks = callbacks;
        this.mission = mission;

        this.state = {
            view: "list",          // myday | important | planned | assigned | flagged | list
            listId: "tasks",
            layout: "grid",        // grid | list
            lists: [
                { id: "tasks", name: "Úkoly", icon: "⌂" },
                { id: "pyrtdm", name: "Tým PyRTDM", icon: "☰", section: "PyRTDM" },
                { id: "weekend", name: "Nákupy na víkend", icon: "☰" }
            ],
            tasks: [
                { id: "t1", title: "Připravit osnovu workshopu o OneNote", listId: "tasks", important: false, myDay: false, due: null, reminder: null, repeat: null, files: [], note: "", done: false, assignedBy: null },
                { id: "t2", title: "Zavolat dodavateli kancelářského papíru", listId: "tasks", important: false, myDay: false, due: "2026-11-09", reminder: null, repeat: null, files: [], note: "", done: false, assignedBy: null },
                { id: "t3", title: "Objednat catering na firemní akci", listId: "tasks", important: false, myDay: false, due: "2026-11-12", reminder: null, repeat: "weekly", files: [], note: "", done: false, assignedBy: null },
                { id: "t4", title: "Zkontrolovat rozpočet na 4. čtvrtletí", listId: "pyrtdm", important: true, myDay: true, due: "2026-11-11", reminder: null, repeat: null, files: [], note: "", done: false, assignedBy: "Martin Svoboda" },
                { id: "t5", title: "Aktualizovat seznam školicích kurzů", listId: "tasks", important: false, myDay: true, due: "2026-11-13", reminder: null, repeat: null, files: [], note: "", done: false, assignedBy: null },
                { id: "t6", title: "Koupit kávu a sušenky na pátek", listId: "weekend", important: false, myDay: false, due: null, reminder: null, repeat: null, files: [], note: "", done: true, assignedBy: null },
                { id: "t7", title: "Odeslat zápis z porady oddělení", listId: "tasks", important: true, myDay: false, due: null, reminder: null, repeat: null, files: [], note: "", done: false, assignedBy: null }
            ],
            emails: [
                { id: "e1", subject: "Podklady k náborovému workshopu", from: "Petra Nováková" },
                { id: "e2", subject: "Termín kontroly bezpečnostních pravidel", from: "Ondřej Kučera" }
            ],
            suggestions: [
                { id: "s1", source: "Teams", from: "Lucie Hejdová", text: "Mohl(a) bys připravit závěrečnou prezentaci pro tým?" },
                { id: "s2", source: "Teams", from: "Tomáš Vlček", text: "Pošli prosím aktualizované číslo účtu na služební cesty." }
            ],
            draft: { title: "", due: null, reminder: null, repeat: null },
            selectedTaskId: null,
            popover: null,      // { type: "due" | "reminder" | "repeat", taskId }
            modal: null,        // { type: "share" | "newlist" | "attach", ... }
            ctx: null,          // { x, y, kind: "task" | "email", id, sub }
            completedOpen: false
        };

        this.bindEvents();
        this.render();
    },

    emit(payload) {
        this.callbacks.onAction(payload);
    },

    find(id) {
        return this.state.tasks.find(t => t.id === id);
    },

    matchesView(t, view, listId) {
        switch (view) {
            case "myday": return t.myDay;
            case "important": return t.important;
            case "planned": return !!t.due;
            case "assigned": return !!t.assignedBy;
            case "list": return t.listId === listId;
            default: return false;
        }
    },

    countView(view, listId) {
        return this.state.tasks.filter(t => !t.done && this.matchesView(t, view, listId)).length;
    },

    labelDue(iso) {
        if (!iso) return "";
        if (iso === this.TODAY) return "Dnes";
        if (iso === this.TOMORROW) return "Zítra";
        const p = iso.split("-");
        return `${parseInt(p[2], 10)}. ${parseInt(p[1], 10)}.`;
    },

    labelOf(opts, value) {
        const o = opts.find(x => x.v === value);
        return o ? o.l : "Vlastní";
    },

    viewTitle() {
        const s = this.state;
        const titles = {
            myday: "☀ Můj den",
            important: "☆ Důležité",
            planned: "📅 Plánované úkoly",
            assigned: "👤 Mám přiřazeno",
            flagged: "⚑ E-mail označený příznakem"
        };
        if (s.view !== "list") return titles[s.view];
        const l = s.lists.find(x => x.id === s.listId);
        return `${l.icon} ${esc(l.name)}`;
    },

    render() {
        const s = this.state;
        this.container.innerHTML = `
            <div class="td-app" id="td-root">
                <header class="td-topbar">
                    <div class="td-brand"><span>▦</span> Outlook</div>
                    <div class="td-search">🔍 Hledat To Do</div>
                    <div class="td-top-icons"><span>🔔</span><span>⚙</span><span class="td-avatar">JS</span></div>
                </header>
                <div class="td-body">
                    ${this.renderSidebar()}
                    <main class="td-main">
                        ${this.renderMainHead()}
                        ${s.view === "flagged" ? "" : this.renderAddBar()}
                        ${s.view === "flagged" ? this.renderEmails() : this.renderContent()}
                    </main>
                    ${this.renderRightPane()}
                </div>
                ${this.renderPopover()}
                ${this.renderModal()}
                ${this.renderCtx()}
            </div>
        `;
    },

    renderSidebar() {
        const s = this.state;
        const nav = [
            ["myday", "☀", "Můj den", this.countView("myday")],
            ["important", "☆", "Důležité", this.countView("important")],
            ["planned", "📅", "Plánované úkoly", this.countView("planned")],
            ["assigned", "👤", "Mám přiřazeno", this.countView("assigned")],
            ["flagged", "⚑", "E-mail označený příznakem", s.emails.length]
        ];
        const navHtml = nav.map(([v, ic, label, n]) => `
            <button class="td-nav-item${s.view === v ? " active" : ""}" data-td-view="${v}">
                <span class="td-nav-ic">${ic}</span><span class="td-nav-label">${label}</span>
                ${n ? `<span class="td-nav-count">${n}</span>` : ""}
            </button>`).join("");
        const listHtml = (l) => `
            <button class="td-nav-item${s.view === "list" && s.listId === l.id ? " active" : ""}" data-td-list="${l.id}">
                <span class="td-nav-ic">${l.icon}</span><span class="td-nav-label">${esc(l.name)}</span>
                <span class="td-nav-count">${this.countView("list", l.id)}</span>
            </button>`;
        const plain = s.lists.filter(l => !l.section);
        const sectioned = s.lists.filter(l => l.section === "PyRTDM");
        return `
            <aside class="td-sidebar">
                <div class="td-side-top"><span>☰</span><span>⚙</span></div>
                ${navHtml}
                ${plain.map(listHtml).join("")}
                <div class="td-section-head">⌄ PyRTDM</div>
                ${sectioned.map(listHtml).join("")}
                <button class="td-nav-item td-new-list" data-td-act="new-list"><span class="td-nav-ic">+</span><span class="td-nav-label">Nový seznam</span></button>
            </aside>`;
    },

    renderMainHead() {
        const s = this.state;
        return `
            <div class="td-main-head">
                <div class="td-title-row">
                    <h1 class="td-title">${this.viewTitle()}</h1>
                    <button class="td-icon-btn" data-td-act="info-more" title="Další možnosti">⋯</button>
                    ${s.view === "list" ? `<button class="td-share-btn" data-td-act="share-list">👥 Sdílet</button>` : ""}
                </div>
                <div class="td-tabs-row">
                    <div class="td-tabs">
                        <button class="td-tab${s.layout === "grid" ? " active" : ""}" data-td-layout="grid">▦ Mřížka</button>
                        <button class="td-tab${s.layout === "list" ? " active" : ""}" data-td-layout="list">☰ Seznam</button>
                    </div>
                    <div class="td-sort">
                        <button class="td-sort-btn" data-td-act="info-sort">Řadit</button>
                        <button class="td-sort-btn" data-td-act="info-group">Skupina</button>
                    </div>
                </div>
            </div>`;
    },

    renderAddBar() {
        const d = this.state.draft;
        return `
            <div class="td-add-bar">
                <div class="td-add-row">
                    <span class="td-add-circle">○</span>
                    <input class="td-add-input" data-td-input="title" placeholder="Přidat úkol" value="${esc(d.title)}">
                    <button class="td-add-btn" data-td-act="add-task"${d.title.trim() ? "" : " disabled"}>Přidat</button>
                </div>
                <div class="td-add-tools">
                    <button class="td-tool${d.due ? " on" : ""}" data-td-pop="due" title="Přidat termín splnění">📅${d.due ? ` <small>${this.labelDue(d.due)}</small>` : ""}</button>
                    <button class="td-tool${d.reminder ? " on" : ""}" data-td-pop="reminder" title="Připomenutí">🔔${d.reminder ? ` <small>${this.labelOf(this.REMINDER_OPTS, d.reminder)}</small>` : ""}</button>
                    <button class="td-tool${d.repeat ? " on" : ""}" data-td-pop="repeat" title="Opakovat">🔁${d.repeat ? ` <small>${this.labelOf(this.REPEAT_OPTS, d.repeat)}</small>` : ""}</button>
                </div>
            </div>`;
    },

    renderContent() {
        const s = this.state;
        const matching = s.tasks.filter(t => this.matchesView(t, s.view, s.listId));
        const open = matching.filter(t => !t.done);
        const done = matching.filter(t => t.done);

        let body;
        if (!open.length) {
            body = `<div class="td-empty">Tady zatím nic není. Přidej první úkol nahoře.</div>`;
        } else if (s.view === "planned") {
            body = this.groupPlanned(open);
        } else {
            body = open.map(t => this.renderTask(t)).join("");
        }

        const doneHtml = done.length ? `
            <div class="td-done-section">
                <div class="td-done-head" data-td-act="toggle-completed">${s.completedOpen ? "⌄" : "›"} Dokončeno ${done.length}</div>
                ${s.completedOpen ? done.map(t => this.renderTask(t)).join("") : ""}
            </div>` : "";

        return `<div class="td-tasks td-layout-${s.layout}">${body}</div>${doneHtml}`;
    },

    groupPlanned(open) {
        const groups = { "Dříve": [], "Dnes": [], "Zítra": [], "Později": [] };
        open.forEach(t => {
            if (t.due < this.TODAY) groups["Dříve"].push(t);
            else if (t.due === this.TODAY) groups["Dnes"].push(t);
            else if (t.due === this.TOMORROW) groups["Zítra"].push(t);
            else groups["Později"].push(t);
        });
        return Object.keys(groups).filter(k => groups[k].length).map(k => `
            <div class="td-group-head">${k} <span>${groups[k].length}</span></div>
            ${groups[k].map(t => this.renderTask(t)).join("")}`).join("");
    },

    renderTask(t) {
        const sel = this.state.selectedTaskId === t.id;
        const overdue = t.due && t.due < this.TODAY && !t.done;
        const meta = [
            t.myDay ? `<span class="td-meta">☀ Můj den</span>` : "",
            t.due ? `<span class="td-due${overdue ? " td-overdue" : ""}">📅 ${this.labelDue(t.due)}</span>` : "",
            t.reminder ? `<span class="td-meta">🔔 ${this.labelOf(this.REMINDER_OPTS, t.reminder)}</span>` : "",
            t.repeat ? `<span class="td-meta">🔁 ${this.labelOf(this.REPEAT_OPTS, t.repeat)}</span>` : "",
            t.files.length ? `<span class="td-meta">📎 ${t.files.length}</span>` : "",
            t.assignedBy ? `<span class="td-meta">👤 Přiřadil(a) ${esc(t.assignedBy)}</span>` : ""
        ].join("");
        return `
            <div class="td-task${t.done ? " td-done" : ""}${sel ? " td-selected" : ""}" data-td-task="${t.id}">
                <button class="td-check" data-td-check="${t.id}" title="Označit jako dokončené">${t.done ? "✓" : "○"}</button>
                <div class="td-task-main">
                    <div class="td-task-title">${esc(t.title)}</div>
                    <div class="td-task-meta">${meta}</div>
                </div>
                <button class="td-star${t.important ? " on" : ""}" data-td-star="${t.id}" title="Označit jako důležité">${t.important ? "★" : "☆"}</button>
            </div>`;
    },

    renderEmails() {
        const s = this.state;
        if (!s.emails.length) return `<div class="td-empty">Žádné e-maily s příznakem.</div>`;
        return `<div class="td-tasks td-layout-list">${s.emails.map(e => `
            <div class="td-task td-email" data-td-email="${e.id}">
                <span class="td-flag">⚑</span>
                <div class="td-task-main">
                    <div class="td-task-title">${esc(e.subject)}</div>
                    <div class="td-task-meta"><span class="td-meta">✉ ${esc(e.from)}</span></div>
                </div>
                <button class="td-tiny-btn" data-td-email-add="${e.id}">Přidat do Úkolů</button>
            </div>`).join("")}</div>`;
    },

    renderRightPane() {
        const s = this.state;
        const t = s.selectedTaskId ? this.find(s.selectedTaskId) : null;
        if (t) {
            return `
                <aside class="td-detail">
                    <div class="td-detail-head">
                        <button class="td-check" data-td-check="${t.id}">${t.done ? "✓" : "○"}</button>
                        <span class="td-detail-title">${esc(t.title)}</span>
                        <button class="td-star${t.important ? " on" : ""}" data-td-star="${t.id}">${t.important ? "★" : "☆"}</button>
                    </div>
                    <div class="td-detail-row" data-td-myday="${t.id}">☀ ${t.myDay ? "Přidáno do Můj den" : "Přidat do zobrazení Můj den"}</div>
                    <div class="td-detail-row" data-td-pop="due" data-td-task="${t.id}">📅 Termín splnění <span>${t.due ? this.labelDue(t.due) : "Přidat"}</span></div>
                    <div class="td-detail-row" data-td-pop="reminder" data-td-task="${t.id}">🔔 Připomenutí <span>${t.reminder ? this.labelOf(this.REMINDER_OPTS, t.reminder) : "Přidat"}</span></div>
                    <div class="td-detail-row" data-td-pop="repeat" data-td-task="${t.id}">🔁 Opakovat <span>${t.repeat ? this.labelOf(this.REPEAT_OPTS, t.repeat) : "Nikdy"}</span></div>
                    <div class="td-detail-section">Přílohy</div>
                    ${t.files.map(f => `<div class="td-detail-file">📎 ${esc(f)}</div>`).join("")}
                    <button class="td-detail-btn" data-td-act="attach" data-td-task="${t.id}">📎 Přidat soubor</button>
                    <div class="td-detail-section">Poznámka</div>
                    <textarea class="td-note" data-td-input="note" data-td-note-task="${t.id}" placeholder="Přidat poznámku">${esc(t.note)}</textarea>
                    <button class="td-detail-btn td-danger" data-td-act="delete-task" data-td-task="${t.id}">🗑 Odstranit úkol</button>
                </aside>`;
        }
        if (s.view === "myday") {
            const recent = s.tasks.filter(x => !x.done).slice(-3).reverse();
            return `
                <aside class="td-right">
                    <div class="td-right-head">Návrhy</div>
                    ${s.suggestions.map(g => `
                        <div class="td-sugg">
                            <div class="td-sugg-src">${esc(g.source)} · ${esc(g.from)}</div>
                            <div class="td-sugg-text">${esc(g.text)}</div>
                            <button class="td-sugg-add" data-td-teams="${g.id}" title="Přidat do Úkolů">+</button>
                        </div>`).join("")}
                    <div class="td-right-head">Naposledy přidané</div>
                    ${recent.map(x => `<div class="td-recent">${esc(x.title)}</div>`).join("")}
                </aside>`;
        }
        return "";
    },

    renderPopover() {
        const p = this.state.popover;
        if (!p) return "";
        const opts = p.type === "due" ? this.DUE_OPTS : p.type === "reminder" ? this.REMINDER_OPTS : this.REPEAT_OPTS;
        const head = { due: "Přidat termín splnění", reminder: "Připomenutí", repeat: "Opakovat" }[p.type];
        return `
            <div class="td-backdrop" data-td-act="close-modal"></div>
            <div class="td-popover td-pop-${p.type}">
                <div class="td-pop-head">${head}</div>
                ${opts.map(o => `${o.sep ? `<div class="td-pop-sep"></div>` : ""}<button class="td-pop-item" data-td-pop-val="${o.v}">${o.l}</button>`).join("")}
            </div>`;
    },

    renderModal() {
        const m = this.state.modal;
        if (!m) return "";
        let inner = "";
        if (m.type === "share") {
            const l = this.state.lists.find(x => x.id === m.listId);
            inner = `
                <div class="td-modal-head">Sdílet seznam „${esc(l.name)}“ <button class="td-x" data-td-act="close-modal">✕</button></div>
                <label class="td-label">Komu chceš seznam sdílet?</label>
                <input class="td-input" data-td-input="share" value="${esc(m.value)}" placeholder="Jméno nebo e-mail kolegy">
                <label class="td-label">Oprávnění</label>
                <select class="td-select" data-td-input="perm">
                    <option value="edit"${m.perm === "edit" ? " selected" : ""}>Může upravovat</option>
                    <option value="view"${m.perm === "view" ? " selected" : ""}>Může zobrazit</option>
                </select>
                <div class="td-modal-actions">
                    <button class="td-btn-secondary" data-td-act="close-modal">Zrušit</button>
                    <button class="td-btn-primary" data-td-act="share-go">Sdílet</button>
                </div>`;
        } else if (m.type === "newlist") {
            inner = `
                <div class="td-modal-head">Nový seznam <button class="td-x" data-td-act="close-modal">✕</button></div>
                <input class="td-input" data-td-input="newlist" value="${esc(m.value)}" placeholder="Název seznamu">
                <div class="td-modal-actions">
                    <button class="td-btn-secondary" data-td-act="close-modal">Zrušit</button>
                    <button class="td-btn-primary" data-td-act="newlist-go">Vytvořit</button>
                </div>`;
        } else if (m.type === "attach") {
            const files = ["Zápis_z_porady.docx", "Rozpočet_Q4.xlsx", "Prezentace_workshop.pptx", "Osnova_workshop.docx"];
            inner = `
                <div class="td-modal-head">Přidat soubor k úkolu <button class="td-x" data-td-act="close-modal">✕</button></div>
                <label class="td-label">Název souboru</label>
                <input class="td-input" data-td-input="attach" value="${esc(m.value || "")}" placeholder="Např. Osnova_workshop.docx">
                <div class="td-modal-actions">
                    <button class="td-btn-secondary" data-td-act="close-modal">Zrušit</button>
                    <button class="td-btn-primary" data-td-act="attach-go">Přidat</button>
                </div>
                <label class="td-label">Nedávné soubory</label>
                ${files.map(f => `<button class="td-file-btn" data-td-attach-file="${esc(f)}">📎 ${esc(f)}</button>`).join("")}`;
        }
        return `<div class="td-backdrop" data-td-act="close-modal"></div><div class="td-modal">${inner}</div>`;
    },

    renderCtx() {
        const c = this.state.ctx;
        if (!c) return "";
        const style = `left:${c.x}px; top:${c.y}px;`;
        if (c.kind === "email") {
            return `
                <div class="td-ctx" style="${style}">
                    <div class="td-ctx-item" data-td-ctx="email-add">Přidat do Úkolů</div>
                    <div class="td-ctx-sep"></div>
                    <div class="td-ctx-item" data-td-ctx="email-unflag">Odebrat příznak</div>
                </div>`;
        }
        const t = this.find(c.id);
        if (!t) return "";
        const subList = (mode) => c.sub === mode ? `
            <div class="td-ctx-sub">
                ${this.state.lists.map(l => `<div class="td-ctx-item" data-td-ctx-list="${l.id}" data-td-mode="${mode}">${esc(l.name)}</div>`).join("")}
            </div>` : "";
        return `
            <div class="td-ctx" style="${style}">
                <div class="td-ctx-item" data-td-ctx="myday">Přidat do zobrazení Můj den</div>
                <div class="td-ctx-item" data-td-ctx="important">Označit jako důležité</div>
                <div class="td-ctx-item" data-td-ctx="done">Označit jako dokončené</div>
                <div class="td-ctx-sep"></div>
                <div class="td-ctx-item" data-td-ctx="due-today">Termín splnění: dnes</div>
                <div class="td-ctx-item" data-td-ctx="due-tomorrow">Termín splnění: zítra</div>
                <div class="td-ctx-item" data-td-ctx="due-clear">Odebrat termín splnění</div>
                <div class="td-ctx-sep"></div>
                <div class="td-ctx-item" data-td-ctx="new-list">Vytvořit z tohoto úkolu nový seznam…</div>
                <div class="td-ctx-item td-has-sub" data-td-sub="move">Přesunout úkol do… <span>›</span></div>
                ${subList("move")}
                <div class="td-ctx-item td-has-sub" data-td-sub="copy">Kopírovat úkol do… <span>›</span></div>
                ${subList("copy")}
                <div class="td-ctx-sep"></div>
                <div class="td-ctx-item td-ctx-danger" data-td-ctx="delete">Odstranit úkol</div>
            </div>`;
    },

    bindEvents() {
        const root = this.container;

        root.addEventListener("click", (e) => {
            const s = this.state;
            const hit = (sel) => e.target.closest(sel);
            let x;

            if ((x = hit("[data-td-sub]")) && s.ctx) {
                const mode = x.dataset.tdSub;
                s.ctx.sub = s.ctx.sub === mode ? null : mode;
                return this.render();
            }
            if ((x = hit("[data-td-ctx-list]")) && s.ctx) {
                this.moveOrCopy(x.dataset.tdMode, x.dataset.tdCtxList);
                return this.render();
            }
            if ((x = hit("[data-td-ctx]")) && s.ctx) {
                this.ctxAction(x.dataset.tdCtx);
                return this.render();
            }
            if ((x = hit("[data-td-check]"))) {
                const t = this.find(x.dataset.tdCheck);
                t.done = !t.done;
                this.emit({ action: "todo_complete_task", value: t.done, taskId: t.id });
                return this.render();
            }
            if ((x = hit("[data-td-star]"))) {
                const t = this.find(x.dataset.tdStar);
                t.important = !t.important;
                this.emit({ action: "todo_toggle_important", value: t.important, taskId: t.id });
                return this.render();
            }
            if ((x = hit("[data-td-myday]"))) {
                const t = this.find(x.dataset.tdMyday);
                t.myDay = !t.myDay;
                this.emit({ action: "todo_add_to_my_day", value: t.myDay, taskId: t.id });
                return this.render();
            }
            if ((x = hit("[data-td-pop-val]")) && s.popover) {
                const p = s.popover;
                s.popover = null;
                this.applyOption(p.type, x.dataset.tdPopVal, p.taskId);
                return this.render();
            }
            if ((x = hit("[data-td-pop]"))) {
                s.popover = { type: x.dataset.tdPop, taskId: x.dataset.tdTask || null };
                return this.render();
            }
            if ((x = hit("[data-td-attach-file]")) && s.modal) {
                const fileName = x.dataset.tdAttachFile;
                const t = this.find(s.modal.taskId);
                t.files.push(fileName);
                s.modal = null;
                this.emit({ action: "todo_attach_file", value: fileName, taskId: t.id });
                return this.render();
            }
            if ((x = hit("[data-td-teams]"))) {
                const g = s.suggestions.find(y => y.id === x.dataset.tdTeams);
                s.suggestions = s.suggestions.filter(y => y !== g);
                this.addTask(g.text, "tasks", {});
                this.emit({ action: "todo_from_teams", source: "teams", value: g.text });
                return this.render();
            }
            if ((x = hit("[data-td-email-add]"))) {
                const e1 = s.emails.find(y => y.id === x.dataset.tdEmailAdd);
                this.addTask(e1.subject, "tasks", {});
                this.emit({ action: "todo_from_teams", source: "email", value: e1.subject });
                return this.render();
            }
            if ((x = hit("[data-td-view]"))) {
                s.view = x.dataset.tdView;
                s.selectedTaskId = null;
                s.ctx = null;
                this.emit({ action: "todo_open_view", value: s.view });
                return this.render();
            }
            if ((x = hit("[data-td-list]"))) {
                s.view = "list";
                s.listId = x.dataset.tdList;
                s.selectedTaskId = null;
                s.ctx = null;
                this.emit({ action: "todo_open_view", value: "list:" + s.listId });
                return this.render();
            }
            if ((x = hit("[data-td-layout]"))) {
                s.layout = x.dataset.tdLayout;
                return this.render();
            }
            if ((x = hit("[data-td-act]"))) {
                this.handleAct(x.dataset.tdAct, x);
                return;
            }
            if ((x = hit(".td-task[data-td-task]"))) {
                s.selectedTaskId = s.selectedTaskId === x.dataset.tdTask ? null : x.dataset.tdTask;
                s.ctx = null;
                return this.render();
            }
            if (s.ctx) {
                s.ctx = null;
                this.render();
            }
        });

        root.addEventListener("contextmenu", (e) => {
            const taskEl = e.target.closest(".td-task[data-td-task]");
            const emailEl = e.target.closest(".td-email[data-td-email]");
            if (!taskEl && !emailEl) return;
            e.preventDefault();
            const s = this.state;
            s.ctx = taskEl
                ? { x: e.clientX, y: e.clientY, kind: "task", id: taskEl.dataset.tdTask, sub: null }
                : { x: e.clientX, y: e.clientY, kind: "email", id: emailEl.dataset.tdEmail, sub: null };
            this.render();
        });

        root.addEventListener("input", (e) => {
            const k = e.target.dataset && e.target.dataset.tdInput;
            const s = this.state;
            if (!k) return;
            if (k === "title") {
                s.draft.title = e.target.value;
                const btn = root.querySelector('[data-td-act="add-task"]');
                if (btn) btn.disabled = !e.target.value.trim();
            } else if ((k === "share" || k === "newlist" || k === "attach") && s.modal) {
                s.modal.value = e.target.value;
            }
        });

        root.addEventListener("change", (e) => {
            const k = e.target.dataset && e.target.dataset.tdInput;
            const s = this.state;
            if (k === "perm" && s.modal) {
                s.modal.perm = e.target.value;
            } else if (k === "note") {
                const t = this.find(e.target.dataset.tdNoteTask);
                t.note = e.target.value;
                this.emit({ action: "todo_add_note", value: e.target.value, taskId: t.id });
            }
        });

        root.addEventListener("keydown", (e) => {
            if (e.key !== "Enter" || !e.target.dataset) return;
            if (e.target.dataset.tdInput === "title") {
                e.preventDefault();
                this.createTask();
            } else if (e.target.dataset.tdInput === "attach") {
                e.preventDefault();
                this.handleAct("attach-go");
            } else if (e.target.dataset.tdInput === "share") {
                e.preventDefault();
                this.handleAct("share-go");
            } else if (e.target.dataset.tdInput === "newlist") {
                e.preventDefault();
                this.handleAct("newlist-go");
            }
        });
    },

    handleAct(act, el) {
        const s = this.state;
        switch (act) {
            case "close-modal":
                s.popover = null;
                s.modal = null;
                s.ctx = null;
                break;
            case "add-task":
                this.createTask();
                return;
            case "new-list":
                s.modal = { type: "newlist", value: "" };
                break;
            case "share-list":
                s.modal = { type: "share", listId: s.listId, value: "", perm: "edit" };
                break;
            case "share-go": {
                if (!s.modal || s.modal.type !== "share") return;
                const v = s.modal.value.trim();
                if (!v) {
                    this.callbacks.onMistake("Zadej jméno nebo e-mail kolegy, se kterým chceš seznam sdílet.");
                    return;
                }
                const listId = s.modal.listId;
                const perm = s.modal.perm;
                s.modal = null;
                this.emit({ action: "todo_share_list", listId: listId, value: v, permission: perm });
                break;
            }
            case "newlist-go": {
                if (!s.modal || s.modal.type !== "newlist") return;
                const name = s.modal.value.trim();
                if (!name) {
                    this.callbacks.onMistake("Seznam potřebuje název.");
                    return;
                }
                const taskId = s.modal.taskId || null;
                const newId = "l" + s.lists.length + Date.now() % 1000;
                s.lists.push({ id: newId, name: name, icon: "☰" });
                s.modal = null;
                if (taskId) {
                    this.find(taskId).listId = newId;
                }
                s.view = "list";
                s.listId = newId;
                this.emit({ action: "todo_create_list", value: name, taskId: taskId });
                break;
            }
            case "attach": {
                if (!el.dataset.tdTask) return;
                s.modal = { type: "attach", taskId: el.dataset.tdTask, value: "" };
                break;
            }
            case "attach-go": {
                if (!s.modal || s.modal.type !== "attach") return;
                const fileName = s.modal.value.trim();
                if (!fileName) {
                    this.callbacks.onMistake("Zadej název souboru, který chceš k úkolu připojit.");
                    return;
                }
                const t = this.find(s.modal.taskId);
                t.files.push(fileName);
                s.modal = null;
                this.emit({ action: "todo_attach_file", value: fileName, taskId: t.id });
                break;
            }
            case "delete-task": {
                const t = this.find(el.dataset.tdTask);
                s.tasks = s.tasks.filter(x => x !== t);
                s.selectedTaskId = null;
                this.emit({ action: "todo_delete_task", value: t.title, taskId: t.id });
                break;
            }
            case "toggle-completed":
                s.completedOpen = !s.completedOpen;
                break;
            case "info-more":
                this.callbacks.onInfo("Další možnosti seznamu jsou v této ukázce pouze náhled.");
                return;
            case "info-sort":
            case "info-group":
                this.callbacks.onInfo("Řazení a seskupení jsou v této ukázce pouze náhled.");
                return;
            default:
                return;
        }
        this.render();
    },

    applyOption(type, val, taskId) {
        const s = this.state;
        const t = taskId ? this.find(taskId) : null;
        const d = s.draft;
        let action;
        if (type === "due") {
            const iso = { today: this.TODAY, tomorrow: this.TOMORROW, next_week: this.NEXT_MONDAY, custom: this.NEXT_MONDAY }[val];
            if (t) t.due = iso; else d.due = iso;
            action = "todo_set_due_date";
        } else if (type === "reminder") {
            if (t) t.reminder = val; else d.reminder = val;
            action = "todo_set_reminder";
        } else {
            if (t) t.repeat = val; else d.repeat = val;
            action = "todo_set_recurrence";
        }
        this.emit({ action: action, value: val, taskId: taskId || null });
    },

    ctxAction(key) {
        const s = this.state;
        const c = s.ctx;
        s.ctx = null;
        if (c.kind === "email") {
            if (key === "email-add") {
                const em = s.emails.find(y => y.id === c.id);
                this.addTask(em.subject, "tasks", {});
                this.emit({ action: "todo_from_teams", source: "email", value: em.subject });
            } else if (key === "email-unflag") {
                s.emails = s.emails.filter(y => y.id !== c.id);
            }
            return;
        }
        const t = this.find(c.id);
        if (!t) return;
        switch (key) {
            case "myday":
                t.myDay = !t.myDay;
                this.emit({ action: "todo_add_to_my_day", value: t.myDay, taskId: t.id });
                break;
            case "important":
                t.important = true;
                this.emit({ action: "todo_toggle_important", value: true, taskId: t.id });
                break;
            case "done":
                t.done = true;
                this.emit({ action: "todo_complete_task", value: true, taskId: t.id });
                break;
            case "due-today":
                t.due = this.TODAY;
                this.emit({ action: "todo_set_due_date", value: "today", taskId: t.id });
                break;
            case "due-tomorrow":
                t.due = this.TOMORROW;
                this.emit({ action: "todo_set_due_date", value: "tomorrow", taskId: t.id });
                break;
            case "due-clear":
                t.due = null;
                this.emit({ action: "todo_set_due_date", value: "none", taskId: t.id });
                break;
            case "new-list":
                s.modal = { type: "newlist", value: t.title, taskId: t.id };
                break;
            case "delete":
                s.tasks = s.tasks.filter(x => x !== t);
                if (s.selectedTaskId === t.id) s.selectedTaskId = null;
                this.emit({ action: "todo_delete_task", value: t.title, taskId: t.id });
                break;
        }
    },

    moveOrCopy(mode, listId) {
        const s = this.state;
        const c = s.ctx;
        s.ctx = null;
        const t = this.find(c.id);
        if (!t) return;
        if (mode === "move") {
            t.listId = listId;
            this.emit({ action: "todo_move_task", value: listId, taskId: t.id });
        } else {
            this.addTask(t.title, listId, { important: t.important, due: t.due });
            this.emit({ action: "todo_copy_task", value: listId, taskId: t.id });
        }
    },

    addTask(title, listId, extra) {
        const s = this.state;
        s.tasks.push(Object.assign({
            id: "t" + Date.now() + Math.floor(Math.random() * 1000),
            title: title,
            listId: listId,
            important: false,
            myDay: false,
            due: null,
            reminder: null,
            repeat: null,
            files: [],
            note: "",
            done: false,
            assignedBy: null
        }, extra));
    },

    createTask() {
        const s = this.state;
        const d = s.draft;
        const title = d.title.trim();
        if (!title) {
            this.callbacks.onMistake("Název úkolu je prázdný. Napiš, co máš udělat.");
            return;
        }
        const listId = s.view === "list" ? s.listId : "tasks";
        this.addTask(title, listId, {
            important: s.view === "important",
            myDay: s.view === "myday",
            due: d.due,
            reminder: d.reminder,
            repeat: d.repeat
        });
        s.draft = { title: "", due: null, reminder: null, repeat: null };
        this.emit({ action: "todo_create_task", value: title, listId: listId });
        this.render();
    }
};

function esc(text) {
    return String(text == null ? "" : text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}
