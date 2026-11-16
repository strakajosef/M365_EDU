/**
 * ============================================================================
 * INTERAKTIVNÍ MODUL MICROSOFT PLANNER (planner.js)
 * ============================================================================
 */

window.SandboxModules = window.SandboxModules || {};

const PLANNER_DATA = {
    today: "2026-11-11",
    planName: "Tým – jarní kampaň",
    planTile: "TK",
    statuses: ["Nezahájeno", "Probíhá", "Dokončeno"],
    priorities: ["Naléhavé", "Důležitá", "Střední", "Nízká"],
    recurrences: ["Neopakuje se", "Denně", "Pracovní dny", "Týdně", "Měsíčně", "Ročně", "Vlastní"],
    reminders: ["Žádné", "V 9:00 v den termínu", "1 den předem", "1 hodinu předem"],
    members: ["Adam Kubíček", "Lenka Dvořáková", "Petr Hrubý"],
    files: [
        { name: "Harmonogram_HRS_2026.xlsx", icon: "📗" },
        { name: "Zápis_z_porady_07.docx", icon: "📘" },
        { name: "Prezentace_onboarding.pptx", icon: "📙" },
        { name: "Odkaz_na_dokumentaci.url", icon: "🔗" }
    ],
    teamsMessages: [
        { title: "Připravit podklady pro jarní setkání", from: "Lenka Dvořáková", text: "Mohl bys do pátku připravit podklady k jarnímu setkání?" },
        { title: "Zkontrolovat rozpočtový výhled", from: "Petr Hrubý", text: "Prosím zkontroluj rozpočtový výhled pro první čtvrtletí." }
    ],
    views: [
        { id: "grid", label: "Mřížka", icon: "▦" },
        { id: "panel", label: "Panel", icon: "▮" },
        { id: "timeline", label: "Časová osa", icon: "▭" },
        { id: "calendar", label: "Kalendář", icon: "▤" },
        { id: "charts", label: "Grafy", icon: "◔" }
    ]
};

const plEsc = (value) => String(value ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
}[c]));

const plInitials = (name) => name.split(" ").map((p) => p[0]).join("").slice(0, 2).toUpperCase();

const plFormatDate = (iso) => {
    if (!iso) return "";
    const [, m, d] = iso.split("-");
    return `${parseInt(d, 10)}. ${parseInt(m, 10)}.`;
};

const plOptions = (list, selected) => list
    .map((item) => `<option value="${plEsc(item)}" ${item === selected ? "selected" : ""}>${plEsc(item)}</option>`)
    .join("");

window.SandboxModules.planner = {
    state: null,
    container: null,
    callbacks: null,

    init(container, mission, callbacks) {
        this.container = container;
        this.callbacks = callbacks;

        this.state = {
            view: "panel",
            buckets: [
                { id: "b1", name: "Celoútvarové úkoly" },
                { id: "b2", name: "Rozvoj v týmu" },
                { id: "b3", name: "Ad hoc úkoly" },
                { id: "b4", name: "Sdílení informací do týmu" }
            ],
            tasks: [
                { id: "t1", bucket: "b1", title: "Příprava rozpočtu na Q1", status: "Probíhá", priority: "Důležitá", due: "2026-11-08", recurrence: "Neopakuje se", reminder: "Žádné", assignees: ["Adam Kubíček"], checklist: [{ text: "Sesbírat podklady", done: true }, { text: "Schválit vedením", done: false }], attachments: [{ name: "Harmonogram_HRS_2026.xlsx", icon: "📗" }], note: "" },
                { id: "t2", bucket: "b2", title: "Vyhodnocení průzkumu spokojenosti", status: "Probíhá", priority: "Střední", due: "2026-11-20", recurrence: "Neopakuje se", reminder: "Žádné", assignees: ["Lenka Dvořáková", "Petr Hrubý"], checklist: [{ text: "Export odpovědí", done: true }, { text: "Shrnutí pro vedení", done: false }, { text: "Rozeslání týmu", done: false }], attachments: [{ name: "Zápis_z_porady_07.docx", icon: "📘" }], note: "" },
                { id: "t3", bucket: "b2", title: "Předávání informací na týmovkách", status: "Nezahájeno", priority: "Nízká", due: "", recurrence: "Týdně", reminder: "Žádné", assignees: ["Adam Kubíček"], checklist: [], attachments: [], note: "Doplnit témata do seznamu." },
                { id: "t4", bucket: "b3", title: "Objednat občerstvení na setkání", status: "Nezahájeno", priority: "Střední", due: "", recurrence: "Neopakuje se", reminder: "Žádné", assignees: [], checklist: [], attachments: [], note: "" },
                { id: "t5", bucket: "b4", title: "Sdílení zápisu z porady", status: "Dokončeno", priority: "Nízká", due: "", recurrence: "Neopakuje se", reminder: "Žádné", assignees: ["Petr Hrubý"], checklist: [], attachments: [], note: "" }
            ],
            members: [...PLANNER_DATA.members],
            sharedWith: [],
            completedOpen: {},
            addingIn: null,
            ctx: null,
            detailId: null,
            detailTab: "details",
            memberListOpen: false,
            modal: null // 'share' | 'files' | 'teams'
        };

        this.bindEvents();
        this.render();
    },

    emit(payload) {
        this.callbacks.onAction(payload);
    },

    info(text) {
        this.callbacks.onInfo(text);
    },

    mistake(text) {
        this.callbacks.onMistake(text);
    },

    getTask(id) {
        return this.state.tasks.find((t) => t.id === id);
    },

    getBucketName(id) {
        return this.state.buckets.find((b) => b.id === id)?.name || "";
    },

    isOverdue(task) {
        return task.due && task.due < PLANNER_DATA.today && task.status !== "Dokončeno";
    },

    bindEvents() {
        this.container.onclick = (e) => this.handleClick(e);
        this.container.onchange = (e) => this.handleChange(e);
        this.container.oncontextmenu = (e) => this.handleContextMenu(e);
    },

    handleClick(e) {
        const s = this.state;
        const el = e.target.closest("[data-pl]");

        if (s.ctx && !e.target.closest(".pl-ctx")) {
            s.ctx = null;
            if (!el) {
                this.render();
                return;
            }
        }
        if (!el) return;

        const action = el.dataset.pl;
        const taskId = el.dataset.task || el.closest("[data-task]")?.dataset.task;

        switch (action) {
            case "info-generic":
                this.info("Tato funkce není v sandboxu k dispozici.");
                break;
            case "info-license":
                this.info("Licence Planneru není v sandboxu potřeba.");
                break;
            case "new-plan":
                this.info("Vytvoření nového plánu není v tomto sandboxu povoleno.");
                break;
            case "view":
                if (el.dataset.view !== "panel") {
                    this.info("V sandboxu je k dispozici pouze pohled Panel.");
                }
                break;
            case "task-open":
                s.detailId = el.dataset.task;
                s.detailTab = "details";
                s.memberListOpen = false;
                break;
            case "task-done": {
                const t = this.getTask(el.dataset.task);
                if (!t) break;
                t.status = t.status === "Dokončeno" ? "Nezahájeno" : "Dokončeno";
                this.emit({ action: "planner_set_status", status: t.status, taskId: t.id });
                break;
            }
            case "completed-toggle":
                s.completedOpen[el.dataset.bucket] = !s.completedOpen[el.dataset.bucket];
                break;
            case "add-open":
                s.addingIn = el.dataset.bucket;
                break;
            case "add-cancel":
                s.addingIn = null;
                break;
            case "add-submit":
                this.createTask(el.dataset.bucket, el.closest(".pl-add-form"));
                break;
            case "ctx-item":
                this.handleContextItem(el.dataset.act, el.dataset.task);
                break;
            case "detail-close":
            case "modal-close":
                if (s.modal) {
                    s.modal = null;
                } else {
                    s.detailId = null;
                    s.memberListOpen = false;
                }
                break;
            case "detail-tab":
                s.detailTab = el.dataset.tab;
                break;
            case "assign-toggle":
                s.memberListOpen = !s.memberListOpen;
                break;
            case "assign": {
                const t = this.getTask(taskId);
                if (t && !t.assignees.includes(el.dataset.name)) {
                    t.assignees.push(el.dataset.name);
                }
                s.memberListOpen = false;
                this.emit({ action: "planner_assign_user", value: el.dataset.name, taskId });
                break;
            }
            case "attach-open":
                s.modal = "files";
                break;
            case "attach-file": {
                const t = this.getTask(taskId);
                const file = PLANNER_DATA.files.find((f) => f.name === el.dataset.file);
                if (t && file) {
                    t.attachments.push({ name: file.name, icon: file.icon });
                }
                s.modal = null;
                this.emit({ action: "planner_attach_file", file: el.dataset.file, taskId });
                this.info("Příloha byla přidána k úkolu.");
                break;
            }
            case "share-open":
                s.modal = "share";
                break;
            case "share-submit":
                this.shareSubmit();
                break;
            case "teams-open":
                s.modal = "teams";
                break;
            case "teams-create":
                this.createFromTeams(Number(el.dataset.idx));
                break;
            default:
                return;
        }
        this.render();
    },

    handleChange(e) {
        const field = e.target.dataset.plChange;
        if (!field) return;

        const s = this.state;
        const taskId = e.target.closest("[data-task]")?.dataset.task;
        const t = this.getTask(taskId);
        if (!t) return;
        const value = e.target.type === "checkbox" ? e.target.checked : e.target.value;

        switch (field) {
            case "status":
                t.status = value;
                this.emit({ action: "planner_set_status", status: value, taskId });
                break;
            case "priority":
                t.priority = value;
                this.emit({ action: "planner_set_priority", priority: value, taskId });
                break;
            case "due":
                t.due = value;
                this.emit({ action: "planner_set_due", value, taskId });
                break;
            case "recurrence":
                t.recurrence = value;
                this.emit({ action: "planner_set_recurrence", recurrence: value, taskId });
                break;
            case "reminder":
                t.reminder = value;
                this.emit({ action: "planner_set_reminder", reminder: value, taskId });
                break;
            case "bucket":
                t.bucket = s.buckets.find((b) => b.name === value)?.id || t.bucket;
                this.emit({ action: "planner_move_bucket", bucket: value, taskId });
                break;
            case "check": {
                const item = t.checklist[Number(e.target.dataset.idx)];
                if (!item) return;
                item.done = value;
                this.emit({ action: "planner_checklist_toggle", value: item.text, done: value, taskId });
                break;
            }
            case "note":
                t.note = value;
                return;
            default:
                return;
        }
        this.render();
    },

    handleContextMenu(e) {
        const card = e.target.closest(".pl-card");
        if (!card) return;
        e.preventDefault();
        const rect = this.container.getBoundingClientRect();
        this.state.ctx = {
            x: Math.min(e.clientX - rect.left, rect.width - 230),
            y: Math.min(e.clientY - rect.top, rect.height - 230),
            taskId: card.dataset.task
        };
        this.render();
    },

    handleContextItem(act, taskId) {
        const s = this.state;
        s.ctx = null;
        if (act === "delete") {
            s.tasks = s.tasks.filter((t) => t.id !== taskId);
            if (s.detailId === taskId) s.detailId = null;
            this.emit({ action: "planner_delete_task", taskId });
            this.info("Úkol byl odstraněn.");
            return;
        }
        s.detailId = taskId;
        s.detailTab = act === "attach" ? "attachments" : "details";
        if (act === "reminder") this.info("Připomenutí nastav v poli Připomenutí.");
        if (act === "recurrence") this.info("Opakování nastav v poli Opakovat.");
    },

    createTask(bucketId, form) {
        if (!form) return;
        const title = form.querySelector('[data-pl-input="new-title"]').value.trim();
        const due = form.querySelector('[data-pl-input="new-due"]').value;
        const assignee = form.querySelector('[data-pl-input="new-assignee"]').value;

        if (!title) {
            this.mistake("Název úkolu je povinný. Zadej ho do pole „Zadejte název úkolu“.");
            return;
        }

        const task = {
            id: "t" + Date.now(),
            bucket: bucketId,
            title,
            status: "Nezahájeno",
            priority: "Střední",
            due,
            recurrence: "Neopakuje se",
            reminder: "Žádné",
            assignees: assignee ? [assignee] : [],
            checklist: [],
            attachments: [],
            note: ""
        };
        this.state.tasks.push(task);
        this.state.addingIn = null;
        this.emit({ action: "planner_create_task", value: title, bucket: this.getBucketName(bucketId), due, taskId: task.id });
    },

    shareSubmit() {
        const input = this.container.querySelector('[data-pl-input="member"]');
        const role = this.container.querySelector('[data-pl-input="role"]')?.value || "Člen";
        const member = input ? input.value.trim() : "";

        if (!member) {
            this.mistake("Zadej jméno nebo e-mail člena, kterého chceš do plánu přidat.");
            return;
        }
        this.state.sharedWith.push({ name: member, role });
        this.state.modal = null;
        this.emit({ action: "planner_share_plan", value: member, role });
        this.info(`Plán byl sdílen s: ${member}.`);
    },

    createFromTeams(index) {
        const msg = PLANNER_DATA.teamsMessages[index];
        if (!msg) return;
        const bucket = this.state.buckets.find((b) => b.name === "Ad hoc úkoly") || this.state.buckets[0];
        this.state.tasks.push({
            id: "t" + Date.now(),
            bucket: bucket.id,
            title: msg.title,
            status: "Nezahájeno",
            priority: "Střední",
            due: "",
            recurrence: "Neopakuje se",
            reminder: "Žádné",
            assignees: [],
            checklist: [],
            attachments: [{ name: "Zpráva z Teams: " + msg.from, icon: "💬" }],
            note: msg.text
        });
        this.state.modal = null;
        this.emit({ action: "planner_from_teams", source: "teams", value: msg.title });
        this.info("Úkol byl vytvořen ze zprávy z Teams.");
    },

    render() {
        this.container.innerHTML = `
            <div class="pl-app">
                ${this.renderTopBar()}
                <div class="pl-body">
                    ${this.renderNav()}
                    <main class="pl-main">
                        ${this.renderPlanHead()}
                        ${this.renderToolbar()}
                        <div class="pl-board">
                            ${this.state.buckets.map((b) => this.renderBucket(b)).join("")}
                        </div>
                    </main>
                </div>
                ${this.renderModals()}
                ${this.renderContextMenu()}
            </div>`;
    },

    renderTopBar() {
        return `
            <header class="pl-topbar">
                <button class="pl-waffle" data-pl="info-generic" title="Aplikace">⋮⋮⋮</button>
                <div class="pl-brand"><span class="pl-logo"></span><span>Planner</span></div>
                <div class="pl-top-right">
                    <button class="pl-top-link" data-pl="info-license">💎 Požádat o licenci</button>
                    <button class="pl-icon-btn" data-pl="info-generic" title="Sdílení">👥</button>
                    <button class="pl-icon-btn" data-pl="info-generic" title="Nastavení">⚙</button>
                    <button class="pl-icon-btn" data-pl="info-generic" title="Nápověda">?</button>
                    <span class="pl-avatar">TS</span>
                </div>
            </header>`;
    },

    renderNav() {
        return `
            <aside class="pl-nav">
                <div class="pl-create-wrap">
                    <button class="pl-btn-create" data-pl="new-plan">＋ Vytvořit plán</button>
                    <button class="pl-btn-create-arrow" data-pl="info-generic">▾</button>
                </div>
                <ul class="pl-nav-list">
                    <li data-pl="info-generic"><span class="pl-nav-icon">☼</span> Můj den</li>
                    <li data-pl="info-generic"><span class="pl-nav-icon">✓</span> Moje úkoly</li>
                    <li class="is-active"><span class="pl-nav-icon">▦</span> Moje plány</li>
                </ul>
                <div class="pl-pinned-title">Připnuto</div>
                <div class="pl-pinned-item" data-pl="info-generic">
                    <span class="pl-pin-tile">ZV</span> Zpětná vazba
                </div>
            </aside>`;
    },

    renderPlanHead() {
        return `
            <div class="pl-plan-head">
                <div class="pl-crumbs">
                    <span class="pl-crumb-link">▦ Moje plány</span>
                    <span class="pl-crumb-sep">›</span>
                    <span class="pl-plan-tile">${PLANNER_DATA.planTile}</span>
                    <strong>${plEsc(PLANNER_DATA.planName)}</strong>
                    <button class="pl-icon-btn" data-pl="info-generic" title="Více">⋯</button>
                </div>
                <button class="pl-btn-share" data-pl="share-open">👥 Sdílet ▾</button>
            </div>`;
    },

    renderToolbar() {
        const s = this.state;
        return `
            <div class="pl-toolbar">
                <div class="pl-views">
                    ${PLANNER_DATA.views.map((v) => `
                        <button class="pl-view ${v.id === s.view ? "is-active" : ""}" data-pl="view" data-view="${v.id}">
                            <span>${v.icon}</span> ${v.label}
                        </button>`).join("")}
                    <button class="pl-view pl-view-more" data-pl="info-generic">···</button>
                    <button class="pl-view pl-view-teams" data-pl="teams-open">💬 Z Teams zprávy</button>
                    <span class="pl-toolbar-sep"></span>
                    <button class="pl-icon-btn" data-pl="info-generic" title="Hledat">⌕</button>
                    <button class="pl-icon-btn" data-pl="info-generic" title="Filtr">⚲</button>
                </div>
                <div class="pl-toolbar-right">
                    <button class="pl-group-btn" data-pl="info-generic">☰ Seskupit podle: Kontejner ⌄</button>
                </div>
            </div>`;
    },

    renderBucket(bucket) {
        const s = this.state;
        const tasks = s.tasks.filter((t) => t.bucket === bucket.id);
        const active = tasks.filter((t) => t.status !== "Dokončeno");
        const done = tasks.filter((t) => t.status === "Dokončeno");
        const isOpen = !!s.completedOpen[bucket.id];

        return `
            <section class="pl-bucket">
                <header class="pl-bucket-head">
                    <h3>${plEsc(bucket.name)} (${active.length})</h3>
                    <button class="pl-icon-btn" data-pl="info-generic" title="Více">⋯</button>
                </header>
                <button class="pl-add-btn" data-pl="add-open" data-bucket="${bucket.id}">＋ Přidat úkol</button>
                ${s.addingIn === bucket.id ? this.renderAddForm(bucket) : ""}
                <div class="pl-cards">
                    ${active.map((t) => this.renderCard(t)).join("")}
                </div>
                <button class="pl-completed-toggle" data-pl="completed-toggle" data-bucket="${bucket.id}">
                    <span>${isOpen ? "⌃" : "⌄"}</span> Dokončené úkoly <span class="pl-count">${done.length}</span>
                </button>
                ${isOpen ? `<div class="pl-cards">${done.map((t) => this.renderCard(t)).join("")}</div>` : ""}
            </section>`;
    },

    renderAddForm(bucket) {
        return `
            <div class="pl-add-form" data-bucket="${bucket.id}">
                <input class="pl-input" type="text" data-pl-input="new-title" maxlength="120" placeholder="Zadejte název úkolu * (povinné)">
                <label class="pl-add-label">Nastavit termín splnění
                    <input class="pl-input" type="date" data-pl-input="new-due">
                </label>
                <select class="pl-input" data-pl-input="new-assignee">
                    <option value="">Přiřadit</option>
                    ${plOptions(this.state.members, "")}
                </select>
                <div class="pl-add-actions">
                    <button class="pl-btn-primary" data-pl="add-submit" data-bucket="${bucket.id}">Přidat úkol</button>
                    <button class="pl-btn-plain" data-pl="add-cancel">Zrušit</button>
                </div>
            </div>`;
    },

    renderCard(t) {
        const doneCount = t.checklist.filter((c) => c.done).length;
        const isDone = t.status === "Dokončeno";
        const urgent = t.priority === "Naléhavé";
        const important = t.priority === "Důležitá";
        const overdue = this.isOverdue(t);

        return `
            <article class="pl-card ${isDone ? "is-done" : ""}" data-pl="task-open" data-task="${t.id}">
                <div class="pl-card-top">
                    <button class="pl-circle ${isDone ? "is-checked" : ""}" data-pl="task-done" data-task="${t.id}" aria-label="Označit jako dokončený"></button>
                    <div class="pl-card-title">${plEsc(t.title)}</div>
                </div>
                ${t.note ? `<div class="pl-card-note">${plEsc(t.note)}</div>` : ""}
                ${t.attachments.length && t.attachments[0].icon !== "💬" ? `<div class="pl-card-file">${plEsc(t.attachments[0].icon)} ${plEsc(t.attachments[0].name)}</div>` : ""}
                <div class="pl-card-meta">
                    ${urgent ? `<span class="pl-prio pl-prio-urgent" title="Naléhavé">!!</span>` : ""}
                    ${important ? `<span class="pl-prio pl-prio-high" title="Důležitá">!</span>` : ""}
                    ${t.status === "Probíhá" ? `<span class="pl-status-dot" title="Probíhá"></span>` : ""}
                    ${t.attachments.length ? `<span class="pl-meta">📎 ${t.attachments.length}</span>` : ""}
                    ${t.checklist.length ? `<span class="pl-meta">☑ ${doneCount}/${t.checklist.length}</span>` : ""}
                </div>
                <div class="pl-card-foot">
                    ${t.due ? `<span class="pl-due ${overdue ? "is-overdue" : ""}">📅 ${plFormatDate(t.due)}</span>` : ""}
                    ${t.recurrence !== "Neopakuje se" ? `<span class="pl-recur">⟳ ${plEsc(t.recurrence)}</span>` : ""}
                    <span class="pl-spacer"></span>
                    ${t.assignees.length
                        ? t.assignees.map((n) => `<span class="pl-avatar-mini" title="${plEsc(n)}">${plInitials(n)}</span>`).join("")
                        : `<span class="pl-assign-icon" title="Přiřadit">👥</span>`}
                </div>
            </article>`;
    },

    renderModals() {
        const s = this.state;
        if (s.modal === "share") return this.renderShareModal();
        if (s.modal === "files") return this.renderFilesModal();
        if (s.modal === "teams") return this.renderTeamsModal();
        if (s.detailId && this.getTask(s.detailId)) {
            return this.renderDetail(this.getTask(s.detailId));
        }
        return "";
    },

    renderDetail(t) {
        const s = this.state;
        const doneCount = t.checklist.filter((c) => c.done).length;
        const bucketNames = s.buckets.map((b) => b.name);

        const details = `
            <div class="pl-fields">
                <label class="pl-field">Stav
                    <select class="pl-input" data-pl-change="status">${plOptions(PLANNER_DATA.statuses, t.status)}</select>
                </label>
                <label class="pl-field">Priorita
                    <select class="pl-input" data-pl-change="priority">${plOptions(PLANNER_DATA.priorities, t.priority)}</select>
                </label>
                <label class="pl-field">Termín splnění
                    <input class="pl-input" type="date" data-pl-change="due" value="${plEsc(t.due)}">
                </label>
                <label class="pl-field">Připomenutí
                    <select class="pl-input" data-pl-change="reminder">${plOptions(PLANNER_DATA.reminders, t.reminder)}</select>
                </label>
                <label class="pl-field">Opakovat
                    <select class="pl-input" data-pl-change="recurrence">${plOptions(PLANNER_DATA.recurrences, t.recurrence)}</select>
                </label>
                <label class="pl-field">Kontejner
                    <select class="pl-input" data-pl-change="bucket">${plOptions(bucketNames, this.getBucketName(t.bucket))}</select>
                </label>
            </div>
            <h4 class="pl-section-title">Kontrolní seznam (dokončené položky: ${doneCount} z(e) ${t.checklist.length})</h4>
            <ul class="pl-checklist">
                ${t.checklist.map((c, i) => `
                    <li><label><input type="checkbox" data-pl-change="check" data-idx="${i}" ${c.done ? "checked" : ""}> ${plEsc(c.text)}</label></li>`).join("") || "<li class=\"pl-muted\">Zatím žádné položky.</li>"}
            </ul>
            <h4 class="pl-section-title">Poznámky</h4>
            <textarea class="pl-input pl-notes" data-pl-change="note" placeholder="Přidat poznámky">${plEsc(t.note)}</textarea>`;

        const attachments = `
            <div class="pl-attach-list">
                ${t.attachments.map((a) => `
                    <div class="pl-attach-item"><span>${plEsc(a.icon)}</span><span class="pl-attach-name">${plEsc(a.name)}</span><button class="pl-icon-btn" data-pl="info-generic">⋯</button></div>`).join("") || "<div class=\"pl-muted\">Úkol zatím nemá žádné přílohy.</div>"}
            </div>
            <button class="pl-btn-outline" data-pl="attach-open">＋ Přidat přílohu</button>`;

        return `
            <div class="pl-modal-overlay">
                <div class="pl-modal pl-detail" data-task="${t.id}">
                    <div class="pl-detail-head">
                        <span class="pl-detail-plan">${plEsc(PLANNER_DATA.planName)}</span>
                        <div class="pl-detail-tools">
                            <button class="pl-icon-btn" data-pl="info-generic" title="Více">⋯</button>
                            <button class="pl-icon-btn" data-pl="info-generic" title="Zvětšit">⤢</button>
                            <button class="pl-icon-btn pl-chat-toggle" data-pl="info-generic" title="Chat">💬</button>
                            <button class="pl-icon-btn" data-pl="detail-close" title="Zavřít">✕</button>
                        </div>
                    </div>
                    <div class="pl-detail-body">
                        <div class="pl-detail-main">
                            <div class="pl-detail-title">
                                <span class="pl-circle ${t.status === "Dokončeno" ? "is-checked" : ""}"></span>
                                <h2>${plEsc(t.title)}</h2>
                            </div>
                            <div class="pl-detail-created">Vytvořil(a) jste: ${plFormatDate("2026-08-17")} · Naposledy změnil(a): ${plEsc(PLANNER_DATA.members[1])}, ${plFormatDate(PLANNER_DATA.today)}</div>
                            <div class="pl-detail-line"><span>🏷</span> <button class="pl-link-btn" data-pl="info-generic">Přidat popisek</button></div>
                            <div class="pl-detail-line">
                                <span>👥</span>
                                <button class="pl-link-btn" data-pl="assign-toggle">Přiřadit uživateli</button>
                                ${t.assignees.map((n) => `<span class="pl-chip">${plEsc(n)}</span>`).join("")}
                            </div>
                            ${s.memberListOpen ? `
                                <div class="pl-member-pick">
                                    ${s.members.map((m) => `<button class="pl-member-btn" data-pl="assign" data-name="${plEsc(m)}"><span class="pl-avatar-mini">${plInitials(m)}</span> ${plEsc(m)}</button>`).join("")}
                                </div>` : ""}
                            <div class="pl-tabs">
                                <button class="pl-tab ${s.detailTab === "details" ? "is-active" : ""}" data-pl="detail-tab" data-tab="details">▤ Podrobnosti o úkolu</button>
                                <button class="pl-tab ${s.detailTab === "attachments" ? "is-active" : ""}" data-pl="detail-tab" data-tab="attachments">📎 Přílohy (${t.attachments.length})</button>
                            </div>
                            ${s.detailTab === "details" ? details : attachments}
                        </div>
                        <aside class="pl-chat">
                            <h3>Chat s úkoly</h3>
                            <div class="pl-chat-empty">
                                <div class="pl-chat-icon">💬</div>
                                <strong>Zahájit konverzaci</strong>
                                <p>Použijte @zmínky k zapojení spolupracovníků a emoji reakce k rychlému sdílení názorů.</p>
                            </div>
                            <div class="pl-chat-input">Napište zprávu</div>
                        </aside>
                    </div>
                </div>
            </div>`;
    },

    renderShareModal() {
        const s = this.state;
        return `
            <div class="pl-modal-overlay">
                <div class="pl-modal pl-small">
                    <header class="pl-small-head"><h3>Sdílet plán</h3><button class="pl-icon-btn" data-pl="modal-close">✕</button></header>
                    <div class="pl-small-body">
                        <p>Přidej člena do plánu „${plEsc(PLANNER_DATA.planName)}“.</p>
                        <div class="pl-share-row">
                            <input class="pl-input" type="text" data-pl-input="member" placeholder="Zadejte jméno nebo e-mail">
                            <select class="pl-input" data-pl-input="role"><option>Člen</option><option>Vlastník</option></select>
                        </div>
                        <h4 class="pl-section-title">Lidé s přístupem</h4>
                        <ul class="pl-people">
                            ${s.members.map((m) => `<li><span class="pl-avatar-mini">${plInitials(m)}</span> ${plEsc(m)} <span class="pl-muted">Člen</span></li>`).join("")}
                            ${s.sharedWith.map((p) => `<li><span class="pl-avatar-mini">${plInitials(p.name)}</span> ${plEsc(p.name)} <span class="pl-muted">${plEsc(p.role)}</span></li>`).join("")}
                        </ul>
                    </div>
                    <footer class="pl-small-foot">
                        <button class="pl-btn-plain" data-pl="modal-close">Zrušit</button>
                        <button class="pl-btn-primary" data-pl="share-submit">Sdílet</button>
                    </footer>
                </div>
            </div>`;
    },

    renderFilesModal() {
        return `
            <div class="pl-modal-overlay">
                <div class="pl-modal pl-small">
                    <header class="pl-small-head"><h3>Přidat přílohu</h3><button class="pl-icon-btn" data-pl="modal-close">✕</button></header>
                    <div class="pl-small-body">
                        <ul class="pl-file-list">
                            ${PLANNER_DATA.files.map((f) => `
                                <li><button class="pl-file-btn" data-pl="attach-file" data-file="${plEsc(f.name)}" data-task="${this.state.detailId}">
                                    <span>${f.icon}</span> ${plEsc(f.name)}
                                </button></li>`).join("")}
                        </ul>
                    </div>
                    <footer class="pl-small-foot">
                        <button class="pl-btn-plain" data-pl="modal-close">Zrušit</button>
                    </footer>
                </div>
            </div>`;
    },

    renderTeamsModal() {
        return `
            <div class="pl-modal-overlay">
                <div class="pl-modal pl-small">
                    <header class="pl-small-head"><h3>Úkol z Teams zprávy</h3><button class="pl-icon-btn" data-pl="modal-close">✕</button></header>
                    <div class="pl-small-body">
                        <p class="pl-muted">Vyber zprávu z kanálu Teams, ze které chceš vytvořit úkol v plánu.</p>
                        <ul class="pl-file-list">
                            ${PLANNER_DATA.teamsMessages.map((m, i) => `
                                <li class="pl-teams-msg">
                                    <div><strong>${plEsc(m.from)}</strong><div>${plEsc(m.text)}</div></div>
                                    <button class="pl-btn-primary pl-btn-sm" data-pl="teams-create" data-idx="${i}">Vytvořit úkol</button>
                                </li>`).join("")}
                        </ul>
                    </div>
                    <footer class="pl-small-foot">
                        <button class="pl-btn-plain" data-pl="modal-close">Zavřít</button>
                    </footer>
                </div>
            </div>`;
    },

    renderContextMenu() {
        const ctx = this.state.ctx;
        if (!ctx) return "";
        return `
            <div class="pl-ctx" style="left:${ctx.x}px;top:${ctx.y}px">
                <button data-pl="ctx-item" data-act="open" data-task="${ctx.taskId}">Otevřít detail úkolu</button>
                <button data-pl="ctx-item" data-act="reminder" data-task="${ctx.taskId}">🔔 Nastavit připomenutí</button>
                <button data-pl="ctx-item" data-act="recurrence" data-task="${ctx.taskId}">⟳ Opakovat úkol</button>
                <button data-pl="ctx-item" data-act="attach" data-task="${ctx.taskId}">📎 Přidat přílohu</button>
                <hr>
                <button data-pl="ctx-item" data-act="delete" data-task="${ctx.taskId}">🗑 Odstranit úkol</button>
            </div>`;
    }
};
