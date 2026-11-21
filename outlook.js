/**
 * ============================================================================
 * INTERAKTIVNÍ MODUL MICROSOFT OUTLOOK (outlook.js) – PIXEL-PERFECT REPLIKA
 * ============================================================================
 */

window.SandboxModules = window.SandboxModules || {};

window.SandboxModules.outlook = {
    state: null,
    container: null,
    callbacks: null,

    init(container, mission, callbacks) {
        this.container = container;
        this.callbacks = callbacks;
        this.mission = mission;

        this.state = {
            helpOpen: false,
            helpMissionId: null, // Úkol, jehož nápověda se zobrazuje (null = aktivní úkol)
            helpShowHint: false,
            activeRail: "mail", // 'mail' | 'calendar'
            activeFolder: "Doručená pošta",
            favOpen: true,
            accOpen: true,
            foldersCollapsed: false,
            collapsedGroups: [],
            mailFilter: "all", // 'all' | 'unread'
            selectedMailId: null, // Výchozí prázdný stav jako na Screenu 01
            activeTab: "home", // 'home' | 'send' | 'folder' | 'view' | 'help'
            expandedFolders: {}, // Rozbalené složky v levém panelu (název -> bool)
            folderMails: {}, // Vygenerované e-maily ve složkách (naplní buildFolderMails)
            todoList: "Úkoly",
            todos: [
                { id: 1, title: "Zkontrolovat ticket INC-88412", done: false, list: "Důležité" },
                { id: 2, title: "Připravit podklady k reportu za září", done: false, list: "Můj den" },
                { id: 3, title: "Potvrdit termín nasazení na testovací prostředí", done: false, list: "Plánované" },
                { id: 4, title: "Dokončit školení práce s daty v Excelu", done: true, list: "Mám přiřazeno" },
                { id: 5, title: "Odpovědět na e-mail od Evy Vzorové", done: false, list: "E-mail s příznakem" }
            ],
            copilotMessages: [],

            // Backstage menu "Soubor" (Screen 06, 13)
            backstageOpen: false,
            backstageContext: "main", // 'main' | 'compose'

            // Modální okna
            optionsModalOpen: false,
            optionsTab: "general", // 'general' | 'mail'
            signatureEditorOpen: false,
            signatureText: "Jana Testovací\nSenior analytik provozu\nODDĚLENÍ DIGITÁLNÍ TRANSFORMACE",

            autoReplyModalOpen: false,
            autoReplyEnabled: false,
            autoReplyText: "Dobrý den, ahoj,\n\nv tuto chvíli jsem mimo kancelář a budu k dispozici od 30.10.2026. Zaslanému e-mailu se budu věnovat co nejdříve po svém návratu.\nV případě urgentního požadavku jsem k dispozici na služebním čísle.\n\nDěkuji, s pozdravem\n\nJana Testovací\nSenior analytik provozu",

            // Dialogy nad aplikací (pick/formulář) – viz renderDialog()
            dialog: null,
            attachments: [],
            searchTerm: "",
            roomBooked: null,

            composeWindowOpen: false,
            composeAppsOpen: false,
            composeTab: "message", // 'message' | 'insert' | 'options' | 'format' | 'review'
            myTemplatesOpen: false,
            composeSubject: "Šablona",
            saveAsDialogOpen: false,
            saveAsFileName: "Šablona.msg",
            saveAsFileType: "msg", // 'msg' | 'oft' | 'txt' | 'htm'

            meetingWindowOpen: false,
            meetingTab: "appointment", // 'appointment' | 'scheduling'
            selectedRoom: "",
            selectedTimeSlot: "11:35",
            autoPickMenuOpen: false,
            selectedCalendarEvent: null,

            taskWindowOpen: false,
            taskSubject: "",

            // Kontextové menu pravého tlačítka myši
            ctxMenu: { open: false, x: 0, y: 0, mailId: null, subMenu: null },

            // Fiktivní firemní složky (inspirováno rozložením ze Screenu 01)
            folders: [
                { name: "Doručená pošta", count: 2, sub: false },
                { name: "!Monitoring", count: 0, sub: true },
                { name: "Automatizace", count: 2, sub: true },
                { name: "Transformace CRM", count: 0, sub: true },
                { name: "WebDev", count: 0, sub: true },
                { name: "PyAnalytics", count: 0, sub: true },
                { name: "! Kampaně", count: 0, sub: true },
                { name: "sprava_kampani", count: 0, sub: true },
                { name: "Datový sklad", count: 0, sub: true },
                { name: "Bezpečnost", count: 0, sub: true },
                { name: "Reporting & BI", count: 0, sub: true },
                { name: "Archiv projektů", count: 0, sub: true },
                { name: "Fakturace_2026", count: 0, sub: true }
            ],

            // Fiktivní e-maily (se změněnými jmény a texty)
            mails: [
                {
                    id: 201,
                    group: "Dnes",
                    sender: "Karel Příklad (ServiceDesk)",
                    initials: "KP",
                    avatarColor: "#607d8b",
                    email: "servicedesk@firma-priklad.cz",
                    subject: "Ticket (INC-88412) Chyba synchronizace modulu plateb - vrácen kód 'NotFound'",
                    time: "11:46",
                    unread: false,
                    category: null,
                    flag: null,
                    body: `<p><strong>Karel Příklad</strong> aktualizoval(a) požadavek v systému ServiceDesk:</p>
                           <p style="margin:10px 0;"><a href="#" style="color:#0078d4;">INC-88412: Chyba synchronizace modulu plateb - vrácen kód 'NotFound'</a></p>
                           <p><strong>Změnil:</strong> Karel Příklad<br><strong>Cílová verze opravy:</strong> 2026.4.0</p>`
                },
                {
                    id: 202,
                    group: "Dnes",
                    sender: "Eva Vzorová",
                    initials: "EV",
                    avatarColor: "#0078d4",
                    email: "eva.vzorova@firma-priklad.cz",
                    subject: "BI Sklad - zpoždění ranního přenosu prodejních dat",
                    time: "08:24",
                    unread: true,
                    category: null,
                    flag: null,
                    body: `<p>Dobrý den,</p><p style="margin-top:8px;">dnešní ranní nahrávání dat do centrálního úložiště má zpoždění přibližně 40 minut z důvodu noční optimalizace databáze.</p>`
                },
                {
                    id: 203,
                    group: "Dnes",
                    sender: "notifikace@firma-priklad.cz",
                    initials: "NC",
                    avatarColor: "#038387",
                    email: "notifikace@firma-priklad.cz",
                    subject: "Přehled změn komponent v ePortálu",
                    time: "6:06",
                    unread: false,
                    category: null,
                    flag: null,
                    body: `<p>Automatický denní souhrn nasazených změn na testovacím prostředí ePortálu.</p>`
                },
                {
                    id: 204,
                    group: "Včera",
                    sender: "Tomáš Modelový",
                    initials: "TM",
                    avatarColor: "#498205",
                    email: "tomas.modelovy@firma-priklad.cz",
                    subject: "REL-4120 – Klientská zóna – úprava formulářů a šablon...",
                    time: "ne 23:05",
                    unread: false,
                    category: null,
                    flag: "done",
                    body: `<p>Ahoj, přikládám finální specifikaci pro nasazení balíčku REL-4120 do produkce.</p>`
                },
                {
                    id: 205,
                    group: "Minulý týden",
                    sender: "Radka Ukázková (Jira)",
                    initials: "RU",
                    avatarColor: "#e3008c",
                    email: "jira@firma-priklad.cz",
                    subject: "Jira Radka Ukázková vás zmínil(a) u úkolu US-90412...",
                    time: "čt 15.10",
                    unread: false,
                    category: { name: "Červená kategorie", color: "#d83b01", bg: "#f9d9d9" },
                    flag: null,
                    body: `<p>👤 <a href="#" style="color:#0078d4;">Radka Ukázková</a> <strong>vás zmínil(a) u úkolu</strong> 🔴 <a href="#" style="color:#0078d4;">US-90412</a></p>
                           <div style="border-left:3px solid #d2d0ce; padding-left:12px; margin-top:10px;">
                               <p style="font-size:14px; color:#0078d4;">Kontrola výpočtu marže v novém ceníku</p>
                               <p style="margin-top:6px;">Prosím o prověření výpočtu na straně frontendu i backendu. Díky!</p>
                               <p style="margin-top:6px;">FYI <a href="#" style="color:#0078d4;">Jana Testovací</a></p>
                           </div>`
                },
                {
                    id: 206,
                    group: "Minulý týden",
                    sender: "Michal Pokusný",
                    initials: "MP",
                    avatarColor: "#8764b8",
                    email: "michal.pokusny@firma-priklad.cz",
                    subject: "Michal Pokusný vám nasdílel(a) soubor Strategicky_plan_2026...",
                    time: "út 13.10",
                    unread: true,
                    category: null,
                    flag: null,
                    body: `<p>Kolega <strong>Michal Pokusný</strong> s vámi sdílel dokument v úložišti OneDrive.</p>`
                },
                {
                    id: 207,
                    group: "Před dvěma týdny",
                    sender: "Pavla Ilustrační",
                    initials: "PI",
                    avatarColor: "#ca5010",
                    email: "pavla.ilustracni@example.com",
                    subject: "Potvrzení - aktivace vzdělávací licence AkademiaPlus",
                    time: "07.10.2026",
                    unread: false,
                    category: null,
                    flag: null,
                    body: `<p>Dobrý den, vaše firemní vzdělávací licence byla úspěšně prodloužena na další období.</p>`
                },
                {
                    id: 208,
                    group: "Před třemi týdny",
                    sender: "Lenka Zkušební (ext) (Jira)",
                    initials: "LZ",
                    avatarColor: "#004578",
                    email: "jira@example.com",
                    subject: "Jira (TASK-30511) Modul Smlouvy 2.0 - po přidání přílohy...",
                    time: "02.10.2026",
                    unread: false,
                    category: null,
                    flag: null,
                    body: `<p>Automatická notifikace o změně stavu testovacího scénáře TASK-30511.</p>`
                }
            ]
        };

        if (mission.id === "m_forms_0") {
            this.state.mails.unshift({
                id: 209,
                group: "Dnes",
                sender: "Karel Příklad",
                initials: "KP",
                avatarColor: "#0078d4",
                email: "karel.priklad@firma-priklad.cz",
                subject: "Pozvánka k vyplnění znalostního kvízu v Microsoft Forms",
                time: "09:15",
                unread: true,
                category: null,
                flag: null,
                body: `<p>Dobrý den,</p>
                       <p style="margin:10px 0;">zvu tě k vyplnění krátkého znalostního kvízu. Otevři prosím formulář Microsoft Forms pomocí následujícího odkazu:</p>
                       <p style="margin:14px 0;"><a href="#" data-ol-forms-quiz-link style="color:#0078d4; text-decoration:underline; font-weight:600;">Otevřít kvíz v Microsoft Forms</a></p>
                       <p>Děkuji!</p>
                       <p style="margin-top:14px;">Karel</p>`
            });
        }

        this.state.folderMails = this.buildFolderMails();
        this.render();
    },

    // Obsah záložek pásu karet (Složka, Posílání a příjem, Zobrazení, Nápověda)
    ribbonTabs: {
        send: [
            { title: "Poslat a přijmout", buttons: [{ label: "Poslat a přijmout<br>ve všech složkách", icon: "🔄" }, { label: "Aktualizovat<br>složku", icon: "🔃" }, { label: "Poslat všechno", icon: "➤" }, { label: "Skupiny pro odeslání<br>a přijímání ▾", icon: "👥" }] },
            { title: "Stáhnout", buttons: [{ label: "Zobrazit<br>průběh", icon: "📊" }, { label: "Zrušit vše", icon: "✖️" }] },
            { title: "Server", buttons: [{ label: "Stáhnout<br>záhlaví", icon: "📥" }, { label: "Označit ke stažení ▾", icon: "🏷️" }, { label: "Zrušit označení<br>ke stažení ▾", icon: "🚫" }, { label: "Zpracovat označená<br>záhlaví ▾", icon: "📋" }] },
            { title: "Předvolby", buttons: [{ label: "Předvolby<br>stahování ▾", icon: "⚙️" }, { label: "Pracovat<br>offline", icon: "📴" }] }
        ],
        folder: [
            { title: "Nová", buttons: [{ label: "Nová<br>složka", icon: "📁" }, { label: "Nová složka<br>výsledků hledání", icon: "🔍" }] },
            { title: "Akce", buttons: [{ label: "Přejmenovat<br>složku", icon: "✏️" }, { label: "Přesunout<br>složku", icon: "📂" }, { label: "Odstranit<br>složku", icon: "🗑️" }] },
            { title: "Vyčistit", buttons: [{ label: "Označit vše<br>jako přečtené", icon: "✔️" }, { label: "Spustit<br>pravidla", icon: "⚡" }, { label: "Zobrazit všechny<br>složky od A do Z", icon: "🔤" }, { label: "Vyčistit<br>složku ▾", icon: "🧹" }, { label: "Odstranit vše", icon: "🗑️" }, { label: "Obnovit odstraněnou<br>poštu", icon: "↩️" }] },
            { title: "Oblíbené", buttons: [{ label: "Přidat k<br>oblíbeným", icon: "⭐" }] },
            { title: "Online zobrazení", buttons: [{ label: "Zobrazit<br>na serveru", icon: "🌐" }] },
            { title: "Vlastnosti", buttons: [{ label: "Zásady", icon: "📜" }, { label: "Nastavení automatické<br>archivace", icon: "🗄️" }, { label: "Oprávnění<br>složky", icon: "🔐" }, { label: "Vlastnosti<br>složky", icon: "ℹ️" }] }
        ],
        view: [
            { title: "Aktuální zobrazení", buttons: [{ label: "Změnit<br>zobrazení ▾", icon: "🔄" }, { label: "Nastavení<br>zobrazení", icon: "⚙️" }, { label: "Obnovit<br>zobrazení", icon: "↩️" }] },
            { title: "Zprávy", buttons: [{ label: "Zobrazit jako konverzace", icon: "☑️" }, { label: "Nastavení konverzace ▾", icon: "⚙️" }] },
            { title: "Prioritní doručená pošta", buttons: [{ label: "Zobrazit prioritní<br>doručenou poštu", icon: "📥" }] },
            { title: "Uspořádání", buttons: [{ label: "Náhled<br>zprávy ▾", icon: "🗂️" }, { label: "Datum", icon: "📅" }, { label: "Od", icon: "👤" }, { label: "Komu", icon: "👥" }, { label: "Kategorie", icon: "🎨" }, { label: "Obrátit pořadí řazení", icon: "↕️" }, { label: "Přidat sloupce", icon: "➕" }, { label: "Rozbalit či sbalit ▾", icon: "➖" }] },
            { title: "Rozložení", buttons: [{ label: "Zmenšit<br>mezery", icon: "↕️" }, { label: "Podokno<br>složek ▾", icon: "📂" }, { label: "Podokno<br>čtení ▾", icon: "📖" }, { label: "Panel<br>úkolů ▾", icon: "✔️" }] },
            { title: "Okno", buttons: [{ label: "Okno<br>Připomínky", icon: "🔔" }, { label: "Otevřít v<br>novém okně", icon: "🪟" }, { label: "Zavřít všechny<br>položky", icon: "✖️" }] },
            { title: "Asistivní čtečka", buttons: [{ label: "Asistivní<br>čtečka", icon: "🔊" }] }
        ],
        help: [
            { title: "Pomoc", buttons: [{ label: "Nápověda", icon: "❓", id: "ol-help-open" }, { label: "Co chcete<br>udělat?", icon: "💡" }] },
            { title: "Zpětná vazba", buttons: [{ label: "Odeslat<br>zpětnou vazbu", icon: "💬" }] },
            { title: "Informace", buttons: [{ label: "Možnosti", icon: "⚙️", id: "ol-help-options" }, { label: "Vyzkoušet<br>nový Outlook", icon: "🆕" }] }
        ]
    },

    // Vygeneruje 3–10 náhodných e-mailů pro každou složku. Seed je odvozen od názvu složky,
    // takže se seznam při každém překreslení nezmění.
    buildFolderMails() {
        const templates = [
            { sender: "Jan Ukázkový", subject: "Podklady k reportu za září", body: "<p>Dobrý den, v příloze posílám podklady k měsíčnímu reportu.</p>" },
            { sender: "Marta Testová", subject: "Zápis z porady oddělení", body: "<p>Ahoj, v příloze je zápis z dnešní porady včetně úkolů.</p>" },
            { sender: "Systém SharePoint", subject: "Dokument byl upraven", body: "<p>Dokument byl upraven uživatelem Petr Ukázka. Změny si zobrazte v knihovně.</p>" },
            { sender: "Ondřej Příklad", subject: "Re: Termín nasazení", body: "<p>Termín nasazení posouvám na úterý. Dej vědět, jestli to vyhovuje.</p>" },
            { sender: "Lucie Ilustrační", subject: "Pozvánka na školení Excel", body: "<p>Zveme vás na interní školení práce s daty v Excelu.</p>" },
            { sender: "Notifikace Jira", subject: "Změna stavu u úkolu TASK-30612", body: "<p>Úkol TASK-30612 přešel do stavu Ke kontrole.</p>" }
        ];
        const groups = ["Dnes", "Včera", "Minulý týden", "Před dvěma týdny", "Před třemi týdny"];
        const palette = ["#607d8b", "#0078d4", "#038387", "#498205", "#8764b8", "#ca5010"];
        const names = [...this.state.folders.map(f => f.name), "Odeslaná pošta"];
        let seed = 1;
        const rnd = () => {
            seed = (Math.imul(seed, 1103515245) + 12345) >>> 0;
            return seed / 4294967296;
        };
        const result = {};
        names.forEach((name, folderIdx) => {
            seed = [...name].reduce((a, ch) => (Math.imul(a, 31) + ch.charCodeAt(0)) >>> 0, 7);
            const count = 3 + Math.floor(rnd() * 8);
            result[name] = Array.from({ length: count }, (_, i) => {
                const t = templates[Math.floor(rnd() * templates.length)];
                const hh = String(8 + Math.floor(rnd() * 10)).padStart(2, "0");
                const mm = String(Math.floor(rnd() * 60)).padStart(2, "0");
                const slug = t.sender.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/\s+/g, ".");
                return {
                    id: 1001 + folderIdx * 100 + i,
                    group: groups[Math.floor(rnd() * groups.length)],
                    sender: t.sender,
                    initials: t.sender.split(" ").map(w => w[0]).join("").slice(0, 2).toUpperCase(),
                    avatarColor: palette[Math.floor(rnd() * palette.length)],
                    email: `${slug}@firma-priklad.cz`,
                    subject: t.subject,
                    time: `${hh}:${mm}`,
                    unread: rnd() < 0.3,
                    category: null,
                    flag: null,
                    body: t.body
                };
            });
        });
        return result;
    },

    // Najde e-mail v doručené poště i ve složkách
    findMail(id) {
        const s = this.state;
        return s.mails.find(m => m.id === id) || Object.values(s.folderMails).flat().find(m => m.id === id);
    },

    // E-maily aktuálně vybrané složky
    currentMails() {
        const s = this.state;
        if (s.activeFolder === "Doručená pošta") return s.mails;
        return s.folderMails[s.activeFolder] || [];
    },

    // Rozbalovací seznam "Všechny aplikace" v okně Nová zpráva
    renderComposeAppsMenu() {
        return `
            <div class="ol-apps-menu" id="ol-apps-menu">
                <div class="ol-apps-menu-title">Všechny aplikace</div>
                <button class="ol-apps-item" id="ol-app-my-templates"><span>📄</span><span>My Templates</span></button>
                <button class="ol-apps-item" disabled><span>📊</span><span>Viva Insights</span></button>
                <button class="ol-apps-item" disabled><span>📝</span><span>Forms</span></button>
                <button class="ol-apps-item" disabled><span>🔍</span><span>Zkontrolovat pravopis</span></button>
            </div>
        `;
    },

    // Postranní panel My Templates se šablonami zpráv
    renderMyTemplatesPanel() {
        const templates = [
            { name: "Mám zpoždění", preview: "Dobrý den, omlouvám se, s odpovědí mám zpoždění. Ozvu se do konce dne." },
            { name: "Oběd?", preview: "Ahoj, nechceš jít na oběd? Dej vědět, kdy se ti hodí." },
            { name: "Odpovím později", preview: "Děkuji za zprávu. Na podrobnou odpověď se podívám později." }
        ];
        return `
            <aside class="ol-mytpl-panel" id="ol-mytpl-panel">
                <header class="ol-mytpl-head">
                    <span>My Templates</span>
                    <span class="ol-mytpl-close" id="ol-close-my-templates" title="Zavřít">✕</span>
                </header>
                <button class="ol-mytpl-new" id="ol-mt-new">+ Šablona</button>
                <div class="ol-mytpl-list">
                    ${templates.map(t => `
                        <div class="ol-mytpl-item">
                            <div class="ol-mytpl-name">${t.name}</div>
                            <div class="ol-mytpl-preview">${t.preview}</div>
                        </div>
                    `).join("")}
                </div>
            </aside>
        `;
    },

    // Řádek složky v levém panelu s rozbalovacím šipkou a e-maily pod ním
    renderFolderRow(f) {
        const s = this.state;
        const expanded = !!s.expandedFolders[f.name];
        const mails = f.name === "Doručená pošta" ? s.mails : (s.folderMails[f.name] || []);
        return `
            <div class="ol-fitem ${f.sub ? "sub" : ""} ${s.activeFolder === f.name ? "active" : ""}" data-ol-folder="${f.name}">
                <span class="ol-fname"><span class="ol-fchev" data-ol-fexp="${f.name}">${expanded ? "∨" : "›"}</span>${f.name}</span>
                ${f.count ? `<span class="ol-fcount">${f.count}</span>` : ""}
            </div>
            ${expanded ? mails.map(m => `
                <div class="ol-fmail ${m.unread ? "unread" : ""} ${s.selectedMailId === m.id ? "selected" : ""}" data-ol-mail="${m.id}">
                    <span>${m.sender}</span>
                    <span class="ol-fmail-sub">${m.subject}</span>
                </div>
            `).join("") : ""}
        `;
    },

    // Pás karet podle aktivní záložky a sekce
    renderRibbon() {
        const s = this.state;
        if (s.activeRail !== "mail" && s.activeRail !== "calendar") return "";
        if (s.activeTab === "home") {
            return s.activeRail === "mail" ? this.renderMailRibbon() : this.renderCalendarRibbon();
        }
        const groups = this.ribbonTabs[s.activeTab] || [];
        return `
            <div class="ol-ribbon">
                ${groups.map(g => `
                    <div class="ol-rgroup">
                        <div class="ol-rgroup-content">
                            ${g.buttons.map(b => `
                                <button class="ol-rbtn-lg" ${b.id ? `id="${b.id}"` : ""}>
                                    <span class="ol-rbtn-ico">${b.icon}</span>
                                    <span>${b.label}</span>
                                </button>
                            `).join("")}
                        </div>
                        <div class="ol-rgroup-title">${g.title}</div>
                    </div>
                `).join("")}
            </div>
        `;
    },

    // Hlavní pracovní plocha podle levé lišty (pošta, kalendář, aplikace)
    renderWorkspaceView() {
        switch (this.state.activeRail) {
            case "calendar": return this.renderCalendarView();
            case "todo": return this.renderTodoView();
            case "copilot": return this.renderCopilotView();
            case "onedrive": return this.renderOneDriveView();
            case "people": return this.renderPeopleView();
            case "apps": return this.renderAppsView();
            default: return this.renderMailView();
        }
    },

    // Úkoly kurzu = všechny Outlook mise z GAME_DATA (sdílený seznam se zadáním)
    getCourseMissions() {
        const data = typeof GAME_DATA !== "undefined" ? GAME_DATA : null;
        const all = (data && data.missions) || [];
        return all.filter(m => m.appModule === "outlook");
    },

    isSolved(missionId) {
        return typeof this.callbacks?.isMissionSolved === "function" && this.callbacks.isMissionSolved(missionId);
    },

    // Kontextová nápověda: zadání a postupná nápověda aktuálního (nebo vybraného) úkolu
    renderHelpPanel() {
        const s = this.state;
        const course = this.getCourseMissions();
        const mission = course.find(m => m.id === s.helpMissionId) || this.mission;
        if (!mission) return "";
        return `
            <aside class="ol-help-panel">
                <div class="ol-help-header">
                    <span>Nápověda</span>
                    <button class="ol-help-close" id="ol-help-close" title="Zavřít">✕</button>
                </div>
                <div class="ol-help-body">
                    <div class="ol-help-label">Úkol</div>
                    <div class="ol-help-topic">${mission.topic}</div>
                    <div class="ol-help-label">Zadání</div>
                    <div class="ol-help-text">${mission.taskDescription}</div>
                    ${s.helpShowHint
                        ? `<div class="ol-help-label">Nápověda</div><div class="ol-help-text ol-help-hint">${mission.hint}</div>`
                        : `<button class="ol-help-hint-btn" id="ol-help-hint-toggle">Zobrazit nápovědu</button>`}
                    ${course.length > 1 ? `
                        <div class="ol-help-label" style="margin-top:14px;">Další úkoly kurzu</div>
                        ${course.filter(m => m.id !== mission.id).map(m => `
                            <div class="ol-help-link" data-ol-help-mission="${m.id}">${this.isSolved(m.id) ? "✓ " : ""}${m.topic}</div>
                        `).join("")}
                    ` : ""}
                </div>
            </aside>
        `;
    },

    renderTodoView() {
        const s = this.state;
        const lists = ["Můj den", "Důležité", "Plánované", "Mám přiřazeno", "E-mail s příznakem", "Úkoly", "Úkoly kurzu"];
        if (s.todoList === "Úkoly kurzu") {
            const missions = this.getCourseMissions();
            const solved = missions.filter(m => this.isSolved(m.id)).length;
            return `
                <section class="ol-app-view">
                    <aside class="ol-app-side">
                        <div class="ol-app-side-title">Úkoly</div>
                        ${lists.map(l => `<div class="ol-app-side-item ${s.todoList === l ? "active" : ""}" data-ol-todolist="${l}">${l}</div>`).join("")}
                    </aside>
                    <div class="ol-app-main">
                        <div class="ol-app-title">Úkoly kurzu (${solved}/${missions.length} splněno)</div>
                        <div style="font-size:12px; color:#605e5c; margin-bottom:10px;">Kliknutím na úkol zobrazíte jeho zadání a nápovědu. Splnění se eviduje jen během této relace.</div>
                        ${missions.map(m => `
                            <div class="ol-todo-item ${this.isSolved(m.id) ? "done" : ""}" data-ol-course-task="${m.id}" style="cursor:pointer;">
                                <span class="ol-todo-check">${this.isSolved(m.id) ? "✓" : ""}</span>
                                <span>${m.topic}</span>
                            </div>
                        `).join("")}
                    </div>
                </section>
            `;
        }
        const tasks = s.todos.filter(t => s.todoList === "Úkoly" || t.list === s.todoList);
        return `
            <section class="ol-app-view">
                <aside class="ol-app-side">
                    <div class="ol-app-side-title">Úkoly</div>
                    ${lists.map(l => `<div class="ol-app-side-item ${s.todoList === l ? "active" : ""}" data-ol-todolist="${l}">${l}</div>`).join("")}
                </aside>
                <div class="ol-app-main">
                    <div class="ol-app-title">${s.todoList}</div>
                    <div class="ol-todo-add">
                        <input type="text" id="ol-todo-input" placeholder="Přidat úkol">
                        <button id="ol-todo-add-btn">Přidat</button>
                    </div>
                    ${tasks.map(t => `
                        <div class="ol-todo-item ${t.done ? "done" : ""}" data-ol-todo="${t.id}">
                            <span class="ol-todo-check">${t.done ? "✓" : ""}</span>
                            <span>${t.title}</span>
                        </div>
                    `).join("")}
                </div>
            </section>
        `;
    },

    renderCopilotView() {
        const s = this.state;
        return `
            <section class="ol-app-view">
                <aside class="ol-app-side">
                    <div class="ol-app-side-item" id="ol-cop-new">+ Nová konverzace</div>
                    <div class="ol-app-side-title">Dnes</div>
                    <div class="ol-app-side-item active">Plán práce na týden</div>
                    <div class="ol-app-side-item">Shrnutí e-mailů</div>
                    <div class="ol-app-side-title">Agenti</div>
                    <div class="ol-app-side-item">Agent Excel</div>
                </aside>
                <div class="ol-app-main ol-cop-main">
                    ${s.copilotMessages.length === 0 ? `
                        <div class="ol-cop-hello">
                            <div class="ol-cop-logo">🌈</div>
                            <h2>Jak vám mohu pomoci?</h2>
                        </div>
                    ` : ""}
                    <div class="ol-cop-chat">
                        ${s.copilotMessages.map(m => `<div class="ol-cop-msg ${m.role}">${m.text}</div>`).join("")}
                    </div>
                    <div class="ol-cop-input">
                        <input type="text" id="ol-cop-input" placeholder="Zeptejte se na cokoliv">
                        <button id="ol-cop-send">➤</button>
                    </div>
                </div>
            </section>
        `;
    },

    renderOneDriveView() {
        const files = [
            ["📁", "Dokumenty projektu", "—", "Před 2 dny"],
            ["📁", "Sdílené s vámi", "—", "Včera"],
            ["📊", "Strategicky_plan_2026.xlsx", "2,4 MB", "út 13.10"],
            ["📝", "Zápis z porady.docx", "128 kB", "čt 15.10"],
            ["📄", "Prezentace_Q4.pptx", "5,1 MB", "07.10.2026"],
            ["📕", "Návod_Outlook.pdf", "890 kB", "02.10.2026"]
        ];
        return `
            <section class="ol-app-view">
                <aside class="ol-app-side">
                    <div class="ol-app-side-title">OneDrive</div>
                    <div class="ol-app-side-item active">Moje soubory</div>
                    <div class="ol-app-side-item">Sdíleno se mnou</div>
                    <div class="ol-app-side-item">Nedávné</div>
                    <div class="ol-app-side-item">Koš</div>
                </aside>
                <div class="ol-app-main">
                    <div class="ol-app-title">Moje soubory</div>
                    <table class="ol-app-table">
                        <thead><tr><th>Název</th><th>Velikost</th><th>Změněno</th></tr></thead>
                        <tbody>
                            ${files.map(f => `<tr><td>${f[0]} ${f[1]}</td><td>${f[2]}</td><td>${f[3]}</td></tr>`).join("")}
                        </tbody>
                    </table>
                </div>
            </section>
        `;
    },

    renderPeopleView() {
        const people = [
            ["Eva Vzorová", "eva.vzorova@firma-priklad.cz", "BI analytik"],
            ["Michal Pokusný", "michal.pokusny@firma-priklad.cz", "Projektový manažer"],
            ["Karel Příklad", "servicedesk@firma-priklad.cz", "ServiceDesk"],
            ["Radka Ukázková", "jira@firma-priklad.cz", "Vývojářka"],
            ["Tomáš Modelový", "tomas.modelovy@firma-priklad.cz", "Business analytik"]
        ];
        return `
            <section class="ol-app-view">
                <aside class="ol-app-side">
                    <div class="ol-app-side-title">Lidé</div>
                    <div class="ol-app-side-item active">Kontakty</div>
                    <div class="ol-app-side-item">Oblíbené</div>
                </aside>
                <div class="ol-app-main">
                    <div class="ol-app-title">Kontakty</div>
                    <table class="ol-app-table">
                        <thead><tr><th>Jméno</th><th>E-mail</th><th>Funkce</th></tr></thead>
                        <tbody>
                            ${people.map(p => `<tr><td>${p[0]}</td><td>${p[1]}</td><td>${p[2]}</td></tr>`).join("")}
                        </tbody>
                    </table>
                </div>
            </section>
        `;
    },

    // Spouštěč aplikací (ikona "Další aplikace" v levé liště)
    renderAppsView() {
        const apps = [
            { rail: "mail", icon: "✉️", name: "Pošta" },
            { rail: "calendar", icon: "📅", name: "Kalendář" },
            { rail: "people", icon: "👥", name: "Lidé" },
            { rail: "todo", icon: "✔️", name: "Úkoly" },
            { rail: "copilot", icon: "🌈", name: "Copilot" },
            { rail: "onedrive", icon: "☁️", name: "OneDrive" }
        ];
        return `
            <section class="ol-app-view">
                <div class="ol-app-main">
                    <div class="ol-app-title">Aplikace</div>
                    <div class="ol-tile-grid">
                        ${apps.map(a => `
                            <div class="ol-tile" data-ol-rail="${a.rail}">
                                <div class="ol-tile-ico">${a.icon}</div>
                                <div>${a.name}</div>
                            </div>
                        `).join("")}
                    </div>
                </div>
            </section>
        `;
    },

    render() {
        const s = this.state;

        this.container.innerHTML = `
            <div class="ol-app" id="ol-root">
                <!-- 1. HORNÍ MODRÁ LIŠTA OKNA -->
                <header class="ol-titlebar">
                    <div class="ol-titlebar-left">
                        <span class="ol-qaccess-icon" title="Automatické ukládání">📬</span>
                        <span class="ol-qaccess-icon" title="Odeslat a přijmout vše">🔄</span>
                        <span class="ol-qaccess-icon" title="Zpět">↩️</span>
                        <span class="ol-qaccess-icon">▾</span>
                    </div>
                    <div class="ol-search-box">
                        <span>🔍</span>
                        <input type="text" id="ol-search-input" placeholder="Hledat" value="${this.state.searchTerm}">
                    </div>
                    <div class="ol-titlebar-right">
                        <div class="ol-avatar-top" title="Jana Testovací (jana.testovaci@firma-priklad.cz)">JT</div>
                        <span class="ol-qaccess-icon">🔔</span>
                        <div class="ol-win-controls">
                            <span>─</span><span>☐</span><span>✕</span>
                        </div>
                    </div>
                </header>

                <!-- 2. HORNÍ MENU KARTY (Soubor | Domů | Posílání a příjem | Složka | Zobrazení | Nápověda) -->
                <nav class="ol-menubar">
                    <div class="ol-menubar-left">
                        <div class="ol-menu-item file-btn" id="ol-btn-file-main">Soubor</div>
                        <div class="ol-menu-item ${s.activeTab === "home" ? "active" : ""}" data-ol-tab="home">Domů</div>
                        <div class="ol-menu-item ${s.activeTab === "send" ? "active" : ""}" data-ol-tab="send">Posílání a příjem</div>
                        <div class="ol-menu-item ${s.activeTab === "folder" ? "active" : ""}" data-ol-tab="folder">Složka</div>
                        <div class="ol-menu-item ${s.activeTab === "view" ? "active" : ""}" data-ol-tab="view">Zobrazení</div>
                        <div class="ol-menu-item ${s.activeTab === "help" ? "active" : ""}" data-ol-tab="help">Nápověda</div>
                    </div>
                    <div class="ol-menubar-right">
                        <span>Vyzkoušejte nový Outlook <span class="ol-toggle-switch"></span></span>
                        <span class="ol-copilot-pill">🌈 Copilot ▾</span>
                    </div>
                </nav>

                <!-- 3. PÁS KARET (RIBBON) PODLE SEKCE -->
                ${this.renderRibbon()}

                <!-- 4. HLAVNÍ PRACOVNÍ PLOCHA -->
                <div class="ol-workspace">
                    <!-- Levá svislá lišta ikon (Pošta, Kalendář, Lidé, To-Do) -->
                    <aside class="ol-appbar">
                        <div class="ol-appbar-btn ${s.activeRail === "mail" ? "active" : ""}" data-ol-rail="mail" title="Pošta">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="5" width="18" height="14" rx="1"/><path d="M3 7l9 6 9-6"/></svg>
                        </div>
                        <div class="ol-appbar-btn ${s.activeRail === "calendar" ? "active" : ""}" data-ol-rail="calendar" title="Kalendář">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="4" width="18" height="17" rx="1"/><path d="M16 2v4M8 2v4M3 9h18"/></svg>
                        </div>
                        <div class="ol-appbar-btn ${s.activeRail === "people" ? "active" : ""}" data-ol-rail="people" title="Lidé">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="9" cy="7" r="4"/><path d="M3 20v-2a4 4 0 014-4h4a4 4 0 014 4v2"/></svg>
                        </div>
                        <div class="ol-appbar-btn ${s.activeRail === "todo" ? "active" : ""}" data-ol-rail="todo" title="To Do / Úkoly">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="2.2"><path d="M20 6L9 17l-5-5"/></svg>
                        </div>
                        <div class="ol-appbar-btn ${s.activeRail === "copilot" ? "active" : ""}" data-ol-rail="copilot" title="Copilot">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#464eb8" stroke-width="1.8"><path d="M12 3l2.2 5.8L20 11l-5.8 2.2L12 19l-2.2-5.8L4 11l5.8-2.2z"/></svg>
                        </div>
                        <div class="ol-appbar-btn ${s.activeRail === "onedrive" ? "active" : ""}" data-ol-rail="onedrive" title="OneDrive">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.8"><path d="M7 18a4 4 0 010-8 6 6 0 0111.5 1.5A3.5 3.5 0 0118 18z"/></svg>
                        </div>
                        <div class="ol-appbar-btn ${s.activeRail === "apps" ? "active" : ""}" data-ol-rail="apps" title="Další aplikace">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="4" y="4" width="6" height="6"/><rect x="14" y="4" width="6" height="6"/><rect x="4" y="14" width="6" height="6"/><rect x="14" y="14" width="6" height="6"/></svg>
                        </div>
                    </aside>

                    ${this.renderWorkspaceView()}

                    <!-- BACKSTAGE MENU "SOUBOR" (Screen 06, 13) -->
                    ${s.backstageOpen ? this.renderBackstage() : ""}

                    <!-- OKNO: NOVÁ ZPRÁVA (Screen 11) -->
                    ${s.composeWindowOpen ? this.renderComposeWindow() : ""}
                    ${s.composeWindowOpen && s.composeAppsOpen ? this.renderComposeAppsMenu() : ""}
                    ${s.composeWindowOpen && s.myTemplatesOpen ? this.renderMyTemplatesPanel() : ""}

                    <!-- OKNO: NOVÁ SCHŮZKA + MÍSTNOSTI + POMOCNÍK PRO PLÁNOVÁNÍ (Screen 03, 04, 05) -->
                    ${s.meetingWindowOpen ? this.renderMeetingWindow() : ""}

                    <!-- DETAIL UDÁLOSTI V KALENDÁŘI -->
                    ${s.selectedCalendarEvent ? this.renderCalendarEventDetails() : ""}

                    <!-- OKNO: ÚKOL Z E-MAILU (Screen 16) -->
                    ${s.taskWindowOpen ? this.renderTaskWindow() : ""}

                    <!-- MODÁL: MOŽNOSTI APLIKACE OUTLOOK (Screen 06, 07) -->
                    ${s.optionsModalOpen ? this.renderOptionsModal() : ""}

                    <!-- PANEL KONTEXTOVÉ NÁPOVĚDY -->
                    ${s.helpOpen ? this.renderHelpPanel() : ""}

                    <!-- MODÁL: AUTOMATICKÉ ODPOVĚDI (Screen 08) -->
                    ${s.autoReplyModalOpen ? this.renderAutoReplyModal() : ""}

                    <!-- DIALOG: ULOŽIT JAKO ŠABLONU *.OFT (Screen 14) -->
                    ${s.saveAsDialogOpen ? this.renderSaveAsDialog() : ""}
                </div>

                <!-- KONTEXTOVÉ MENU NA PRAVÉ TLAČÍTKO MYŠI (Screen 09, 10, 12, 15) -->
                ${s.ctxMenu.open ? this.renderContextMenu() : ""}

                <!-- SPODNÍ STAVOVÝ ŘÁDEK -->
                <footer class="ol-statusbar">
                    <div>Položky: 23 &nbsp;&nbsp; Nepřečteno: 2</div>
                    <div style="display:flex; gap:18px;">
                        <span>Všechny složky jsou aktuální.</span>
                        <span>Připojený přes poskytovatele: Microsoft Exchange</span>
                        <span>➖ ━━━●━━━ ➕ 100 %</span>
                    </div>
                </footer>
            </div>
        `;

        this.bindEvents();
    },

    // ========================================================================
    // PÁSY KARET (RIBBONS) – POŠTA & KALENDÁŘ (Screen 01, 02)
    // ========================================================================
    renderMailRibbon() {
        return `
            <div class="ol-ribbon">
                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg" id="ol-btn-new-mail">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#107c10" stroke-width="1.8"><rect x="3" y="5" width="18" height="14" rx="1"/><path d="M3 7l9 6 9-6M18 3v6M15 6h6"/></svg>
                            <span>Nový<br>e-mail</span>
                        </button>
                        <button class="ol-rbtn-lg" id="ol-btn-new-items">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.6"><rect x="4" y="4" width="16" height="16" rx="1"/><path d="M8 9h8M8 13h6"/></svg>
                            <span>Nové<br>položky ▾</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Nové</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <div class="ol-rstack">
                            <button class="ol-rbtn-sm" title="Ignorovat"><span style="color:#a4262c;">🚫</span></button>
                            <button class="ol-rbtn-sm" title="Vyčistit"><span>🧹 ▾</span></button>
                            <button class="ol-rbtn-sm" title="Nevyžádaná pošta"><span style="color:#a4262c;">👤🚫 ▾</span></button>
                        </div>
                        <button class="ol-rbtn-lg" id="ol-btn-del-mail">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.6"><path d="M3 6h18M8 6V4h8v2M6 6v14h12V6"/></svg>
                            <span>Odstranit</span>
                        </button>
                        <button class="ol-rbtn-lg" id="ol-btn-archive">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#107c10" stroke-width="1.7"><rect x="3" y="4" width="18" height="4"/><path d="M5 8v12h14V8M10 12h4"/></svg>
                            <span>Archivovat</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Odstranit</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <div class="ol-rstack">
                            <button class="ol-rbtn-sm"><span style="color:#7719aa;">↩</span> Odpovědět</button>
                            <button class="ol-rbtn-sm"><span style="color:#7719aa;">⟪</span> Odpovědět všem</button>
                            <button class="ol-rbtn-sm"><span style="color:#0078d4;">→</span> Přeposlat</button>
                        </div>
                        <div class="ol-rstack">
                            <button class="ol-rbtn-sm" title="Odpovědět schůzkou">📅</button>
                            <button class="ol-rbtn-sm" title="IM">💬 ▾</button>
                        </div>
                    </div>
                    <div class="ol-rgroup-title">Odpovědět</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#464eb8" stroke-width="1.8"><rect x="3" y="6" width="12" height="12" rx="2"/><circle cx="17" cy="9" r="2.5"/><path d="M15 18h6v-3a2 2 0 00-2-2h-2"/></svg>
                            <span>Nasdílet<br>do Teams</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Teams</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg">
                            <svg viewBox="0 0 24 24" fill="#eaa300" stroke="#b87400" stroke-width="1.4"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
                            <span>Rychlé<br>kroky ▾</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Rychlé kroky ↘</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg" id="ol-ribbon-rules">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.7"><path d="M3 7h18v13H3zM3 7l4-3h6l2 3"/></svg>
                            <span>Přesunout<br>▾</span>
                        </button>
                        <button class="ol-rbtn-lg" id="ol-ribbon-tags">
                            <svg viewBox="0 0 24 24" fill="#e81123" stroke="#a4262c" stroke-width="1.5"><path d="M5 3v18M5 4h13l-3 4 3 4H5"/></svg>
                            <span>Značky<br>▾</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Přesunout & Značky</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <div class="ol-rstack">
                            <button class="ol-rbtn-sm">👥 Nová skupina</button>
                            <button class="ol-rbtn-sm">🔎 Procházet skupiny</button>
                        </div>
                    </div>
                    <div class="ol-rgroup-title">Skupiny</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <div class="ol-rstack">
                            <input type="text" placeholder="Hledat lidi" style="height:20px; font-size:11px; padding:0 6px; border:1px solid #8a8886; width:105px;">
                            <button class="ol-rbtn-sm">📖 Adresář</button>
                            <button class="ol-rbtn-sm" id="ol-btn-filter-adv">⏳ Filtrovat e-maily ▾</button>
                        </div>
                    </div>
                    <div class="ol-rgroup-title">Najít</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.7"><path d="M4 18L9 6l5 12M6 14h6M17 8c1.5 1.5 1.5 6.5 0 8M20 5c3 3 3 11 0 14"/></svg>
                            <span>Číst<br>nahlas</span>
                        </button>
                        <button class="ol-rbtn-lg">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.7"><path d="M4 5h8M8 3v2M5 5c0 5 3 8 6 9M12 19l4-9 4 9M13.5 16h5"/></svg>
                            <span>Přeložit<br>▾</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Řeč a Jazyk</div>
                </div>
            </div>
        `;
    },

    renderCalendarRibbon() {
        return `
            <div class="ol-ribbon">
                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg" id="ol-btn-new-meeting">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.7"><rect x="3" y="4" width="18" height="17" rx="1"/><path d="M16 2v4M8 2v4M3 9h18"/></svg>
                            <span>Nová<br>událost</span>
                        </button>
                        <button class="ol-rbtn-lg" id="ol-btn-new-meeting2">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.7"><rect x="3" y="4" width="14" height="15" rx="1"/><circle cx="18" cy="14" r="3"/><path d="M15 21v-1a3 3 0 016 0v1"/></svg>
                            <span>Nová<br>schůzka</span>
                        </button>
                        <button class="ol-rbtn-lg">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#605e5c" stroke-width="1.7"><circle cx="12" cy="10" r="5"/><path d="M10 15v3h4v-3"/></svg>
                            <span>Přidat čas na<br>intenzivní práci</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Nové</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg">
                            <svg viewBox="0 0 24 24" fill="#464eb8" stroke="#464eb8" stroke-width="1.5"><rect x="3" y="7" width="12" height="10" rx="2"/><polygon points="16 10 21 7 21 17 16 14"/></svg>
                            <span>Okamžitá<br>schůzka</span>
                        </button>
                        <button class="ol-rbtn-lg" id="ol-btn-teams-meeting">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#464eb8" stroke-width="1.8"><rect x="3" y="5" width="13" height="13" rx="2"/><path d="M7 10h5M9.5 10v5"/></svg>
                            <span>Schůzka v<br>Teams ▾</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Schůzka Teams</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.7"><rect x="3" y="4" width="18" height="16"/><path d="M9 14l-3-3 3-3M6 11h9"/></svg>
                            <span>Dnes</span>
                        </button>
                        <button class="ol-rbtn-lg">
                            <svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.7"><rect x="3" y="4" width="18" height="16"/><path d="M14 8l3 3-3 3M7 11h10"/></svg>
                            <span>Příštích<br>7 dní</span>
                        </button>
                    </div>
                    <div class="ol-rgroup-title">Přejít na ↘</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg"><svg viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.6"><rect x="4" y="4" width="16" height="16"/><path d="M4 9h16"/></svg><span>Den</span></button>
                        <button class="ol-rbtn-lg"><svg viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.6"><rect x="3" y="4" width="18" height="16"/><path d="M3 9h18M9 9v11M15 9v11"/></svg><span>Pracovní<br>týden</span></button>
                        <button class="ol-rbtn-lg active"><svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.8"><rect x="3" y="4" width="18" height="16"/><path d="M3 9h18M8 9v11M13 9v11M18 9v11"/></svg><span>Týden</span></button>
                        <button class="ol-rbtn-lg"><svg viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.6"><rect x="3" y="4" width="18" height="16"/><path d="M3 9h18M3 14h18"/></svg><span>Měsíc</span></button>
                        <button class="ol-rbtn-lg"><svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.6"><rect x="3" y="4" width="18" height="16"/><path d="M7 12h10M7 16h6"/></svg><span>Zobrazení<br>časového plánu</span></button>
                    </div>
                    <div class="ol-rgroup-title">Uspořádat ↘</div>
                </div>

                <div class="ol-rgroup">
                    <div class="ol-rgroup-content">
                        <button class="ol-rbtn-lg"><svg viewBox="0 0 24 24" fill="none" stroke="#107c10" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg><span>Přidat<br>kalendář ▾</span></button>
                        <button class="ol-rbtn-lg" id="ol-btn-share-cal"><svg viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.6"><circle cx="12" cy="12" r="8"/></svg><span>Sdílet<br>kalendář ▾</span></button>
                        <button class="ol-rbtn-lg" id="ol-btn-bookings"><svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="1.6"><rect x="4" y="5" width="16" height="15"/><path d="M4 10h16M9 3v4M15 3v4"/></svg><span>Bookings<br>stránka ▾</span></button>
                    </div>
                    <div class="ol-rgroup-title">Spravovat kalendáře</div>
                </div>
            </div>
        `;
    },

    // ========================================================================
    // 4-SLOUPCOVÝ POHLED POŠTY (Screen 01, 09, 10, 12, 15)
    // ========================================================================
    renderMailView() {
        const s = this.state;
        const term = s.searchTerm.toLowerCase();
        const visibleMails = this.currentMails().filter(m => (s.mailFilter === "all" || m.unread)
            && (!term || m.sender.toLowerCase().includes(term) || m.subject.toLowerCase().includes(term)));
        const activeMail = this.findMail(s.selectedMailId);
        const groups = ["Dnes", "Včera", "Minulý týden", "Před dvěma týdny", "Před třemi týdny"];

        return `
            <!-- 1. Strom složek -->
            <aside class="ol-folderpane">
                <div class="ol-fsection-head" data-ol-fsec="fav"><span>${s.favOpen ? "∨" : "›"}</span> Oblíbené <span style="margin-left:auto; font-weight:400;">‹</span></div>
                ${s.favOpen ? `<div class="ol-fitem" data-ol-folder="Odeslaná pošta"><span>Odeslaná pošta</span></div>` : ""}

                <div class="ol-fsection-head" style="margin-top:10px;" data-ol-fsec="acc"><span>${s.accOpen ? "∨" : "›"}</span> jana.testovaci@firma-priklad...</div>
                ${s.accOpen ? s.folders.filter(f => !f.sub || !s.foldersCollapsed).map(f => this.renderFolderRow(f)).join("") : ""}
            </aside>

            <!-- 2. Seznam zpráv seskupený podle dnů -->
            <section class="ol-msglist">
                <div class="ol-msglist-header">
                    <div class="ol-pivots">
                        <div class="ol-pivot ${s.mailFilter === "all" ? "active" : ""}" data-ol-filter="all">Všechny</div>
                        <div class="ol-pivot ${s.mailFilter === "unread" ? "active" : ""}" data-ol-filter="unread">Nepřečtené</div>
                    </div>
                    <div class="ol-sort-label">Podle: Datum ∨ &nbsp; ↑</div>
                </div>

                <div class="ol-msglist-scroll">
                    ${groups.map(g => {
                        const rows = visibleMails.filter(m => m.group === g);
                        if (!rows.length) return "";
                        return `
                            <div class="ol-group-bar" data-ol-group="${g}"><span>${s.collapsedGroups.includes(g) ? "›" : "∨"}</span> ${g}</div>${s.collapsedGroups.includes(g) ? "" : rows.map(m => `
                                <div class="ol-msg-row ${m.unread ? "unread" : ""} ${s.selectedMailId === m.id ? "selected" : ""}"
                                     data-ol-mail="${m.id}">
                                    <div class="ol-msg-top">
                                        <span>${m.sender}</span>
                                        <span style="font-size:11px; color:#107c10;">${m.flag === "done" ? "✓" : ""}${m.flag === "followup" ? "🚩" : ""}</span>
                                    </div>
                                    <div class="ol-msg-sub">
                                        <span class="ol-msg-sub-text">${m.subject}</span>
                                        <span class="ol-msg-time">${m.time}</span>
                                    </div>
                                    ${m.category ? `<div class="ol-cat-tag" style="border-left-color:${m.category.color}; background:${m.category.bg};">${m.category.name}</div>` : ""}
                                </div>
                            `).join("")}
                        `;
                    }).join("")}
                </div>
            </section>

            <!-- 3. Podokno čtení (Prázdné se symbolickou obálkou NEBO detail zprávy) -->
            <section class="ol-readpane">
                ${!activeMail ? `
                    <div class="ol-read-empty">
                        <svg viewBox="0 0 64 44" fill="none"><rect x="2" y="2" width="60" height="40" stroke="#8a8886" stroke-width="2"/><path d="M2 4l30 22L62 4" stroke="#8a8886" stroke-width="2"/></svg>
                        <div style="font-size:15px; color:#605e5c;">Vyberte položku, kterou si chcete přečíst.</div>
                        <div style="color:#0078d4; font-size:12px; margin-top:6px; cursor:pointer;" id="ol-show-preview-link">
                            Sem klikněte, pokud chcete, aby se vždy zobrazoval náhled zpráv.
                        </div>
                    </div>
                ` : `
                    <div class="ol-read-content">
                        <div class="ol-read-subject">
                            <span>${activeMail.subject}</span>
                            <span style="background:#0078d4; color:white; padding:4px 7px; border-radius:3px; font-size:12px;">📝</span>
                        </div>

                        <div class="ol-read-meta">
                            <div class="ol-read-avatar" style="background:${activeMail.avatarColor}">${activeMail.initials}</div>
                            <div>
                                <div style="font-size:13px; color:#201f1e;">${activeMail.sender} &lt;${activeMail.email}&gt;</div>
                                <div style="font-size:11.5px; color:#605e5c; margin-top:2px;">Komu &nbsp; 🟢 Jana Testovací</div>
                            </div>
                            <div class="ol-read-actions">
                                <button class="ol-read-actbtn" title="Reagovat">😊</button>
                                <button class="ol-read-actbtn" title="Odpovědět">↩</button>
                                <button class="ol-read-actbtn" title="Odpovědět všem">⟪</button>
                                <button class="ol-read-actbtn" title="Přeposlat">→</button>
                                <button class="ol-read-actbtn" title="Další akce">⋯</button>
                            </div>
                        </div>

                        ${activeMail.category ? `<div class="ol-cat-tag" style="margin-bottom:8px; border-left-color:${activeMail.category.color}; background:${activeMail.category.bg};">${activeMail.category.name}</div>` : ""}

                        <div class="ol-infobar">
                            ⓘ Pokud se vyskytly potíže se zobrazením této zprávy, kliknutím sem ji zobrazíte ve webovém prohlížeči.<br>
                            Kliknutím sem stáhnete obrázky. Za účelem ochrany vašeho soukromí zabránila aplikace Outlook stažení některých obrázků v této zprávě.
                        </div>
                        <div style="color:#0078d4; font-size:11.5px; margin-bottom:14px; display:flex; gap:18px;">
                            <span>🌐 Přeložit zprávu do: Čeština</span>
                            <span>Nikdy nepřekládat z: Angličtina</span>
                        </div>

                        <div style="font-size:12.5px; line-height:1.55; color:#201f1e; padding-top:8px; border-top:1px solid #edebe9;">
                            ${activeMail.body}
                        </div>
                    </div>
                `}
            </section>

            <!-- 4. Pravý postranní panel s mini-kalendářem a agendou (Screen 01) -->
            <aside class="ol-todopane">
                <div class="ol-minical-header">
                    <span>‹</span>
                    <span>říjen 2026</span>
                    <span>› &nbsp; ✕</span>
                </div>
                <table class="ol-minical-table">
                    <thead>
                        <tr><th>#</th><th>PO</th><th>ÚT</th><th>ST</th><th>ČT</th><th>PÁ</th><th>SO</th><th>NE</th></tr>
                    </thead>
                    <tbody>
                        <tr><td class="wknum">40</td><td class="muted">28</td><td class="muted">29</td><td class="muted">30</td><td>1</td><td>2</td><td>3</td><td>4</td></tr>
                        <tr><td class="wknum">41</td><td>5</td><td>6</td><td>7</td><td>8</td><td>9</td><td>10</td><td>11</td></tr>
                        <tr><td class="wknum">42</td><td>12</td><td>13</td><td>14</td><td>15</td><td>16</td><td>17</td><td>18</td></tr>
                        <tr><td class="wknum">43</td><td class="today">19</td><td>20</td><td>21</td><td>22</td><td>23</td><td>24</td><td>25</td></tr>
                        <tr><td class="wknum">44</td><td>26</td><td>27</td><td>28</td><td>29</td><td>30</td><td>31</td><td class="muted">1</td></tr>
                    </tbody>
                </table>

                <div class="ol-agenda-group">Dnes</div>
                <div class="ol-agenda-row">
                    <div class="ol-agenda-time">Celý den</div>
                    <div class="ol-agenda-card" data-ol-calendar-event="Práce z domova (VPN)" data-ol-event-time="Celý den" data-ol-event-date="Dnes">🏢 <strong>Práce z domova (VPN)</strong></div>
                </div>
                <div class="ol-agenda-row">
                    <div class="ol-agenda-time">11:00</div>
                    <div class="ol-agenda-card yellow-solid" data-ol-calendar-event="AkademiaPlus" data-ol-event-time="11:00" data-ol-event-date="Dnes">AkademiaPlus</div>
                </div>
                <div class="ol-agenda-row">
                    <div class="ol-agenda-time">13:05</div>
                    <div class="ol-agenda-card" data-ol-calendar-event="porada oddělení" data-ol-event-time="13:05" data-ol-event-date="Dnes"><strong>porada oddělení</strong><br><span style="color:#605e5c;">Schůzka Microsoft Teams</span></div>
                </div>
                <div class="ol-agenda-row">
                    <div class="ol-agenda-time">15:00</div>
                    <div class="ol-agenda-card" data-ol-calendar-event="Nasazení na testovací prostředí" data-ol-event-time="15:00" data-ol-event-date="Dnes"><strong>Nasazení na testovací...</strong><br><span style="color:#605e5c;">Schůzka Microsoft Teams</span></div>
                </div>

                <div class="ol-agenda-group">Zítra</div>
                <div class="ol-agenda-row">
                    <div class="ol-agenda-time">Celý den</div>
                    <div class="ol-agenda-card" data-ol-calendar-event="Práce z domova (VPN)" data-ol-event-time="Celý den" data-ol-event-date="Zítra">🏢 <strong>Práce z domova (VPN)</strong></div>
                </div>
                <div class="ol-agenda-row">
                    <div class="ol-agenda-time">9:05</div>
                    <div class="ol-agenda-card" data-ol-calendar-event="Analytici CRM - řešitelská schůzka" data-ol-event-time="9:05" data-ol-event-date="Zítra"><strong>Analytici CRM - řešitelská...</strong><br><span style="color:#605e5c;">Schůzka Microsoft Teams</span></div>
                </div>
                <div class="ol-agenda-row">
                    <div class="ol-agenda-time">10:05</div>
                    <div class="ol-agenda-card yellow-solid" data-ol-calendar-event="Knihovna Online" data-ol-event-time="10:05" data-ol-event-date="Zítra">Knihovna Online</div>
                </div>
            </aside>
        `;
    },

    // ========================================================================
    // POHLED KALENDÁŘE (Screen 02)
    // ========================================================================
    renderCalendarView() {
        const days = [
            { name: "pondělí", num: "19", icon: "🏠", active: true },
            { name: "úterý", num: "20", icon: "🏠", active: false },
            { name: "středa", num: "21", icon: "🏠", active: false },
            { name: "čtvrtek", num: "22", icon: "🏢", active: false },
            { name: "pátek", num: "23", icon: "🏠", active: false },
            { name: "sobota", num: "24", icon: "", active: false },
            { name: "neděle", num: "25", icon: "", active: false }
        ];
        const hours = ["8:00", "9:00", "10:00", "11:00", "12:00", "13:00", "14:00"];

        return `
            <div class="ol-calview">
                <aside class="ol-cal-sidebar">
                    <div class="ol-minical-header"><span>‹</span><span>říjen 2026</span><span>›</span></div>
                    <table class="ol-minical-table">
                        <thead><tr><th>#</th><th>PO</th><th>ÚT</th><th>ST</th><th>ČT</th><th>PÁ</th><th>SO</th><th>NE</th></tr></thead>
                        <tbody>
                            <tr><td class="wknum">42</td><td>12</td><td>13</td><td>14</td><td>15</td><td>16</td><td>17</td><td>18</td></tr>
                            <tr style="background:#cde6f7;"><td class="wknum">43</td><td class="today">19</td><td>20</td><td>21</td><td>22</td><td>23</td><td>24</td><td>25</td></tr>
                            <tr><td class="wknum">44</td><td>26</td><td>27</td><td>28</td><td>29</td><td>30</td><td>31</td><td class="muted">1</td></tr>
                        </tbody>
                    </table>

                    <div style="margin-top:16px; font-weight:600;">∨ ☐ Vlastní kalendáře</div>
                    <div style="background:#e8b3b9; padding:5px 8px; margin-top:4px; font-weight:600;">☑ Kalendář</div>
                    <div style="padding:5px 8px;">☐ Narozeniny</div>
                    <div style="margin-top:12px; font-weight:600;">∨ ☐ Tým: Eva Vzorová</div>
                </aside>

                <section class="ol-cal-main">
                    <div class="ol-cal-top-nav">
                        <button class="ol-form-btn">Dnes</button>
                        <span>‹ &nbsp; ›</span>
                        <span style="font-size:16px; font-weight:600; color:#201f1e;">19 – 25. října 2026</span>
                        <span style="margin-left:auto; font-size:12px;">📅 Týden ∨</span>
                    </div>

                    <div class="ol-cal-weekgrid">
                        <!-- Záhlaví dnů -->
                        <div style="border-bottom:1px solid #e1dfdd;"></div>
                        ${days.map(d => `
                            <div class="ol-cal-dayhead ${d.active ? "active-day" : ""}">
                                <div class="ol-cal-dayname">${d.name}</div>
                                <div class="ol-cal-daynum"><span><strong>${d.num}</strong></span><span>${d.icon}</span></div>
                            </div>
                        `).join("")}

                        <!-- Řádek celodenních událostí -->
                        <div style="font-size:10px; color:#605e5c; text-align:right; padding:4px;">Celý</div>
                        <div class="ol-cal-allday-banner">← Od data: 16. X. &nbsp;&nbsp;&nbsp;&nbsp; Práce z domova (VPN)</div>
                        <div style="border-left:1px solid #e1dfdd;"></div>
                        <div class="ol-cal-allday-banner" style="grid-column: 6 / 9;">Práce z domova (VPN) &nbsp;&nbsp; Do data: 27. X. →</div>

                        <!-- Časové sloty -->
                        ${hours.map(h => `
                            <div style="font-size:10.5px; color:#605e5c; text-align:right; padding:4px 6px; border-bottom:1px solid #edebe9;">${h}</div>${days.map((d, idx) => `
                                <div class="ol-cal-slot ${idx >= 5 ? "dotted" : ""} ${d.active ? "active-col" : ""}" data-ol-cal-cell="${h}">
                                    ${idx === 0 && h === "11:00" ? `<div class="ol-cal-block yellow" data-ol-calendar-event="AkademiaPlus" data-ol-event-time="11:00" data-ol-event-date="pondělí 19. října 2026"><strong>AkademiaPlus</strong></div>` : ""}
                                    ${idx === 0 && h === "13:00" ? `<div class="ol-cal-block rose" data-ol-calendar-event="porada oddělení" data-ol-event-time="13:00" data-ol-event-date="pondělí 19. října 2026"><strong>porada oddělení</strong><br>Schůzka Microsoft Teams</div>` : ""}
                                    ${idx === 1 && h === "9:00" ? `<div class="ol-cal-block rose" data-ol-calendar-event="Analytici CRM" data-ol-event-time="9:00" data-ol-event-date="úterý 20. října 2026"><strong>Analytici CRM</strong></div>` : ""}
                                    ${idx === 1 && h === "10:00" ? `<div class="ol-cal-block yellow" data-ol-calendar-event="Knihovna Online" data-ol-event-time="10:00" data-ol-event-date="úterý 20. října 2026"><strong>Knihovna Online</strong> 🔒</div>` : ""}
                                    ${idx === 2 && h === "10:00" ? `<div class="ol-cal-block rose" data-ol-calendar-event="PROJEKT PHOENIX" data-ol-event-time="10:00" data-ol-event-date="středa 21. října 2026"><strong>PROJEKT PHOENIX</strong><br>Schůzka Teams; Eva Vzorová</div>` : ""}
                                    ${idx === 2 && h === "11:00" ? `<div class="ol-cal-block yellow" data-ol-calendar-event="Angličtina" data-ol-event-time="11:00" data-ol-event-date="středa 21. října 2026"><strong>Angličtina</strong></div>` : ""}
                                    ${idx === 2 && h === "13:00" ? `<div class="ol-cal-block green" data-ol-calendar-event="REVIZE - kontrola norem" data-ol-event-time="13:00" data-ol-event-date="středa 21. října 2026"><strong>REVIZE - kontrola norem</strong></div>` : ""}
                                    ${idx === 4 && h === "8:00" ? `<div class="ol-cal-block rose" data-ol-calendar-event="Workshop" data-ol-event-time="8:00" data-ol-event-date="pátek 23. října 2026"><strong>Workshop</strong></div>` : ""}
                                </div>
                            `).join("")}
                        `).join("")}
                    </div>
                </section>
            </div>
        `;
    },

    renderCalendarEventDetails() {
        const event = this.state.selectedCalendarEvent;
        return `
            <div class="ol-modal-layer" onclick="event.stopPropagation()">
                <section class="ol-win-frame" role="dialog" aria-modal="true" aria-labelledby="ol-calendar-event-title">
                    <header class="ol-titlebar">
                        <div class="ol-titlebar-left">
                            <span>📅</span>
                            <span>${event.title} - Schůzka</span>
                        </div>
                        <div class="ol-win-controls"><span style="cursor:pointer;" id="ol-close-calendar-event">✕</span></div>
                    </header>
                    <nav class="ol-menubar">
                        <div class="ol-menubar-left">
                            <div class="ol-menu-item file-btn">Soubor</div>
                            <div class="ol-menu-item active">Schůzka</div>
                            <div class="ol-menu-item">Vložení</div>
                            <div class="ol-menu-item">Nápověda</div>
                        </div>
                        <div class="ol-menubar-right"><span class="ol-copilot-pill">🌈 Copilot ▾</span></div>
                    </nav>
                    <div class="ol-ribbon">
                        <div class="ol-rgroup">
                            <div class="ol-rgroup-content">
                                <button class="ol-rbtn-lg"><span>📅</span><span>Otevřít<br>kalendář</span></button>
                            </div>
                            <div class="ol-rgroup-title">Akce</div>
                        </div>
                        <div class="ol-rgroup">
                            <div class="ol-rgroup-content">
                                <button class="ol-rbtn-lg"><span>🔔</span><span>Připomenutí</span></button>
                                <button class="ol-rbtn-lg"><span>🔁</span><span>Opakování</span></button>
                            </div>
                            <div class="ol-rgroup-title">Možnosti</div>
                        </div>
                    </div>
                    <div style="flex:1; overflow:auto; padding:24px 32px; color:#201f1e;">
                        <h1 id="ol-calendar-event-title" style="font-size:22px; font-weight:400; margin-bottom:22px;">${event.title}</h1>
                        <div style="max-width:760px; border-top:1px solid #edebe9;">
                            <div class="ol-field-line">
                                <span style="width:118px; color:#605e5c;">Datum</span>
                                <span>${event.date}</span>
                            </div>
                            <div class="ol-field-line">
                                <span style="width:118px; color:#605e5c;">Čas</span>
                                <span>${event.time}</span>
                            </div>
                            <div class="ol-field-line">
                                <span style="width:118px; color:#605e5c;">Kalendář</span>
                                <span>Jana Testovací</span>
                            </div>
                        </div>
                    </div>
                </section>
            </div>
        `;
    },

    // ========================================================================
    // KONTEXTOVÉ MENU PRAVÉHO TLAČÍTKA MYŠI (Screen 09, 10, 12, 15)
    // ========================================================================
    renderContextMenu() {
        const cm = this.state.ctxMenu;
        const targetMail = this.state.mails.find(m => m.id === cm.mailId) || this.state.mails[0];

        return `
            <div class="ol-ctx-popup" style="left:${cm.x}px; top:${cm.y}px;" onclick="event.stopPropagation()">
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon">📄</span><span>Kopírovat</span></div></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon">🖨️</span><span>Rychlý tisk</span></div></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#7719aa;">↩</span><span>Odpovědět</span></div></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#7719aa;">⟪</span><span>Odpovědět všem</span></div></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#0078d4;">→</span><span>Přeposlat</span></div></div>
                <div class="ol-ctx-sep"></div>
                <div class="ol-ctx-item" data-ol-ctx="unread"><div class="ol-ctx-item-left"><span class="ol-ctx-icon">✉️</span><span>Označit jako nepřečtené</span></div></div>

                <!-- Zařadit do kategorií ▸ (Screen 10) -->
                <div class="ol-ctx-item" data-ol-sub="categories">
                    <div class="ol-ctx-item-left"><span class="ol-ctx-icon">⊞</span><span>Zařadit do kategorií</span></div><span>›</span>
                    ${cm.subMenu === "categories" ? `
                        <div class="ol-ctx-sub">
                            <div class="ol-ctx-item" data-ol-cat="clear"><span>Vymazat všechny kategorie</span></div>
                            <div class="ol-ctx-sep"></div>
                            <div class="ol-ctx-item" data-ol-cat="Červená kategorie|#d83b01|#f9d9d9"><div class="ol-ctx-item-left"><span class="ol-cat-box" style="background:#e85c5c;"></span><span>Červená kategorie</span></div></div>
                            <div class="ol-ctx-item" data-ol-cat="Modrá kategorie|#0078d4|#deecf9"><div class="ol-ctx-item-left"><span class="ol-cat-box" style="background:#3a96dd;"></span><span>Modrá kategorie</span></div></div>
                            <div class="ol-ctx-item" data-ol-cat="Nachová kategorie|#8764b8|#efebf7"><div class="ol-ctx-item-left"><span class="ol-cat-box" style="background:#9b87da;"></span><span>Nachová kategorie</span></div></div>
                            <div class="ol-ctx-item" data-ol-cat="Oranžová kategorie|#ca5010|#fdeadb"><div class="ol-ctx-item-left"><span class="ol-cat-box" style="background:#f77f3c;"></span><span>Oranžová kategorie</span></div></div>
                            <div class="ol-ctx-item" data-ol-cat="Prioritní spis|#004578|#cfe4f5"><div class="ol-ctx-item-left"><span class="ol-cat-box" style="background:#2b658e;"></span><span>Prioritní spis</span></div></div>
                            <div class="ol-ctx-item" data-ol-cat="Zelená kategorie|#107c10|#dff6dd"><div class="ol-ctx-item-left"><span class="ol-cat-box" style="background:#47b869;"></span><span>Zelená kategorie</span></div></div>
                            <div class="ol-ctx-item" data-ol-cat="Žlutá kategorie|#eaa300|#fff4ce"><div class="ol-ctx-item-left"><span class="ol-cat-box" style="background:#fae34c;"></span><span>Žlutá kategorie</span></div></div>
                        </div>
                    ` : ""}
                </div>

                <!-- Zpracovat ▸ (Screen 15) -->
                <div class="ol-ctx-item" data-ol-sub="followup">
                    <div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#e81123;">🚩</span><span>Zpracovat</span></div><span>›</span>
                    ${cm.subMenu === "followup" ? `
                        <div class="ol-ctx-sub">
                            <div class="ol-ctx-item" data-ol-follow="today"><div class="ol-ctx-item-left"><span style="color:#e81123;">🚩</span><span>Dnes</span></div></div>
                            <div class="ol-ctx-item" data-ol-follow="tomorrow"><div class="ol-ctx-item-left"><span style="color:#e81123;">🚩</span><span>Zítra</span></div></div>
                            <div class="ol-ctx-item" data-ol-follow="week"><div class="ol-ctx-item-left"><span style="color:#e81123;">🚩</span><span>Tento týden</span></div></div>
                            <div class="ol-ctx-item" data-ol-follow="nextweek"><div class="ol-ctx-item-left"><span style="color:#e81123;">🚩</span><span>Příští týden</span></div></div>
                            <div class="ol-ctx-sep"></div>
                            <div class="ol-ctx-item" data-ol-follow="reminder"><div class="ol-ctx-item-left"><span>🔔</span><span>Přidat připomenutí...</span></div></div>
                            <div class="ol-ctx-item" data-ol-follow="done"><div class="ol-ctx-item-left"><span style="color:#107c10;">✓</span><span>Označit jako splněné</span></div></div>
                        </div>
                    ` : ""}
                </div>

                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon">🛡️</span><span>Přiřadit zásadu</span></div><span>›</span></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon">🔎</span><span>Najít související</span></div><span>›</span></div>
                <div class="ol-ctx-sep"></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#eaa300;">⚡</span><span>Rychlé kroky</span></div><span>›</span></div>

                <!-- Pravidla ▸ (Screen 12) -->
                <div class="ol-ctx-item" data-ol-sub="rules">
                    <div class="ol-ctx-item-left"><span class="ol-ctx-icon">📋</span><span>Pravidla</span></div><span>›</span>
                    ${cm.subMenu === "rules" ? `
                        <div class="ol-ctx-sub" style="width:285px;">
                            <div class="ol-ctx-item" data-ol-rule="always_move"><span>Vždy přesunout zprávy od: ${targetMail.sender}</span></div>
                            <div class="ol-ctx-item" data-ol-rule="create_rule"><div class="ol-ctx-item-left"><span>✉️</span><span>Vytvořit pravidlo...</span></div></div>
                            <div class="ol-ctx-item" data-ol-rule="manage_rules"><div class="ol-ctx-item-left"><span>⚙️</span><span>Spravovat pravidla a upozornění...</span></div></div>
                        </div>
                    ` : ""}
                </div>

                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon">📁</span><span>Přesunout</span></div><span>›</span></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#7719aa;">🅽</span><span>Poslat do OneNotu</span></div></div>
                <div class="ol-ctx-sep"></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#a4262c;">🚫</span><span>Ignorovat</span></div></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#a4262c;">👤</span><span>Nevyžádaná pošta</span></div><span>›</span></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon">🗑️</span><span>Odstranit</span></div></div>
                <div class="ol-ctx-item"><div class="ol-ctx-item-left"><span class="ol-ctx-icon" style="color:#107c10;">📥</span><span>Archivovat...</span></div></div>
            </div>
        `;
    },

    // ========================================================================
    // BACKSTAGE MENU "SOUBOR" (Screen 06 a Screen 13)
    // ========================================================================
    renderBackstage() {
        const s = this.state;
        const isCompose = s.backstageContext === "compose";

        return `
            <div class="ol-backstage" onclick="event.stopPropagation()">
                <aside class="ol-bs-sidebar">
                    <div class="ol-bs-backbtn" id="ol-bs-back">←</div>
                    <div class="ol-bs-navitem active">🏠 Informace</div>
                    ${isCompose ? `
                        <div class="ol-bs-navitem" id="ol-bs-save-as">💾 Uložit</div>
                        <div class="ol-bs-navitem" id="ol-bs-save-as2">✏️ Uložit jako</div>
                        <div class="ol-bs-navitem">🖨️ Tisk</div>
                        <div class="ol-bs-navitem" id="ol-bs-back2">📁 Zavřít</div>
                    ` : `
                        <div class="ol-bs-navitem">📂 Otevřít a exportovat</div>
                        <div class="ol-bs-navitem" id="ol-bs-save-as">💾 Uložit jako</div>
                        <div class="ol-bs-navitem">📎 Uložit přílohy</div>
                        <div class="ol-bs-navitem">🖨️ Tisk</div>
                    `}
                    <div style="margin-top:auto; border-top:1px solid #d2d0ce; padding-top:6px;">
                        <div class="ol-bs-navitem">👤 Účet Office</div>
                        <div class="ol-bs-navitem" id="ol-bs-open-options">⚙️ Možnosti</div>
                    </div>
                </aside>

                <main class="ol-bs-main">
                    ${isCompose ? `
                        <!-- Screen 13: Backstage otevřené zprávy / šablony -->
                        <h2 style="font-size:19px; font-weight:600; margin-bottom:18px;">Šablona</h2>
                        <div class="ol-bs-row">
                            <button class="ol-bs-squarebtn">
                                <span style="font-size:24px;">🔒</span>
                                <span>Šifrovat<br>▾</span>
                            </button>
                            <div>
                                <div style="font-size:14px; font-weight:600;">Šifrovat tuto položku</div>
                                <div style="color:#605e5c; margin-top:4px;">Umožňuje nastavit omezení pro tuto položku. Příjemcům můžete například znemožnit přeposlání této e-mailové zprávy dalším lidem.</div>
                            </div>
                        </div>
                        <div class="ol-bs-row">
                            <button class="ol-bs-squarebtn">
                                <span style="font-size:24px; color:#0078d4;">📥</span>
                                <span>Přesunout<br>do složky ▾</span>
                            </button>
                            <div>
                                <div style="font-size:14px; font-weight:600;">Přesunout položku do jiné složky</div>
                                <div style="color:#605e5c; margin-top:4px;">Přesunutí nebo zkopírování této položky do jiné složky<br>• Aktuální složka: Odchozí pošta</div>
                            </div>
                        </div>
                        <div class="ol-bs-row">
                            <button class="ol-bs-squarebtn">
                                <span style="font-size:24px; color:#0078d4;">📋</span>
                                <span>Vlastnosti</span>
                            </button>
                            <div>
                                <div style="font-size:14px; font-weight:600;">Vlastnosti</div>
                                <div style="color:#605e5c; margin-top:4px;">Nastavení a zobrazení upřesňujících možností a vlastností této položky</div>
                            </div>
                        </div>
                    ` : `
                        <!-- Screen 06: Informace o účtu -->
                        <h2 style="font-size:19px; font-weight:600; margin-bottom:14px;">Informace o účtu</h2>
                        <div style="border:1px solid #8a8886; background:#ffffff; padding:6px 12px; width:430px; display:flex; align-items:center; gap:10px; margin-bottom:8px;">
                            <span style="background:#0078d4; color:white; font-weight:700; padding:4px 8px; border-radius:2px;">E</span>
                            <div>
                                <div style="font-weight:600;">jana.testovaci@firma-priklad.cz</div>
                                <div style="font-size:11px; color:#605e5c;">Microsoft Exchange</div>
                            </div>
                            <span style="margin-left:auto;">∨</span>
                        </div>
                        <button class="ol-form-btn" style="margin-bottom:22px;">➕ Přidat účet</button>

                        <div class="ol-bs-row">
                            <button class="ol-bs-squarebtn">
                                <span style="font-size:24px; color:#0078d4;">👤⚙️</span>
                                <span>Nastavení<br>účtu ▾</span>
                            </button>
                            <div>
                                <div style="font-size:14px; font-weight:600;">Nastavení účtu</div>
                                <div style="color:#605e5c; margin-top:4px;">Umožňuje změnit nastavení tohoto účtu nebo nastavit další připojení.<br>• Přistupujte k tomuto účtu na webu: <a href="#" style="color:#0078d4;">https://outlook.example.com/owa/firma-priklad.cz/</a></div>
                            </div>
                        </div>

                        <div class="ol-bs-row">
                            <button class="ol-bs-squarebtn" id="ol-bs-autoreply">
                                <span style="font-size:24px; color:#a4262c;">🚪←</span>
                                <span>Automatické<br>odpovědi</span>
                            </button>
                            <div>
                                <div style="font-size:14px; font-weight:600;">Automatické odpovědi (Mimo kancelář)</div>
                                <div style="color:#605e5c; margin-top:4px;">Pomocí automatických odpovědí můžete ostatním oznamovat, že jste mimo kancelář, na dovolené nebo že momentálně nemůžete na e-mailové zprávy odpovídat.</div>
                            </div>
                        </div>

                        <div class="ol-bs-row">
                            <button class="ol-bs-squarebtn">
                                <span style="font-size:24px;">✉️🧹</span>
                                <span>Nástroje<br>▾</span>
                            </button>
                            <div>
                                <div style="font-size:14px; font-weight:600;">Nastavení poštovní schránky</div>
                                <div style="color:#605e5c; margin-top:4px;">Můžete spravovat velikost poštovní schránky vyprázdněním odstraněných položek a provedením archivace.</div>
                                <div class="ol-quota-bar"><div class="ol-quota-fill"></div></div>
                                <div style="font-size:11px; color:#323130;">• 90,7 GB volného místa z 99 GB</div>
                            </div>
                        </div>

                        <div class="ol-bs-row">
                            <button class="ol-bs-squarebtn" id="ol-bs-rules">
                                <span style="font-size:24px; color:#0078d4;">⚙️🔔</span>
                                <span>Spravovat pravidla<br>a upozornění</span>
                            </button>
                            <div>
                                <div style="font-size:14px; font-weight:600;">Pravidla a upozornění</div>
                                <div style="color:#605e5c; margin-top:4px;">Pravidla a upozornění vám usnadní uspořádávání příchozích e-mailových zpráv a umožní vám dostávat aktualizace v případě přidání, změny nebo odebrání položek.</div>
                            </div>
                        </div>
                    `}
                </main>
            </div>
        `;
    },

    // ========================================================================
    // OKNO NOVÁ ZPRÁVA (Screen 11)
    // ========================================================================
    // Pás karet okna nové zprávy – obsah se mění podle zvolené karty (Zpráva, Vložení, Možnosti, Formátování textu, Revize)
    renderComposeRibbon() {
        const s = this.state;
        const tabs = {
            message: [
                { title: "Schránka", buttons: [{ label: "Vložit<br>▾", icon: "📋" }, { label: "Vyjmout", icon: "✂️" }, { label: "Kopírovat", icon: "📄" }, { label: "Kopírovat<br>formát", icon: "🖌️" }] },
                { title: "Základní text", buttons: [{ label: "Písmo<br>▾", icon: "🔤" }, { label: "Odstavec<br>▾", icon: "¶" }, { label: "Styly<br>▾", icon: "🅰️" }] },
                { title: "Jména", buttons: [{ label: "Adresář", icon: "📖" }, { label: "Kontrola<br>jmen", icon: "👤✓" }] },
                { title: "Zahrnout", buttons: [{ label: "Připojit<br>soubor ▾", icon: "📎", id: "ol-btn-attach" }, { label: "Odkaz ▾", icon: "🔗" }, { label: "Podpis ▾", icon: "✒️" }] },
                { title: "Značky", buttons: [{ label: "Zpracovat<br>▾", icon: "🚩" }, { label: "Vysoká<br>důležitost", icon: "❗" }, { label: "Nízká<br>důležitost", icon: "❕" }, { label: "Přiřadit<br>zásadu ▾", icon: "🏷️" }] },
                { title: "Hlas", buttons: [{ label: "Diktovat", icon: "🎙️", id: "ol-btn-dictate" }] },
                { title: "Aplikace", buttons: [{ label: "Všechny<br>aplikace ▾", icon: "🧩", id: "ol-btn-all-apps", active: s.composeAppsOpen }] },
                { title: "Citlivost", buttons: [{ label: "Citlivost ▾", icon: "🛡️" }] },
                { title: "Editor", buttons: [{ label: "Editor", icon: "📝" }] },
                { title: "Asistivní nástroje", buttons: [{ label: "Asistivní<br>čtečka", icon: "🔊" }] },
                { title: "Najít čas", buttons: [{ label: "Nový Plánovač<br>průkazů", icon: "📅" }] }
            ],
            insert: [
                { title: "Zahrnout", buttons: [{ label: "Připojit<br>soubor ▾", icon: "📎", id: "ol-btn-attach" }, { label: "Položka aplikace<br>Outlook", icon: "✉️" }, { label: "Hlasování", icon: "📊" }, { label: "Vizitka ▾", icon: "📇" }, { label: "Kalendář", icon: "📅" }, { label: "Podpis ▾", icon: "✒️" }] },
                { title: "Tabulky", buttons: [{ label: "Tabulka ▾", icon: "▦" }] },
                { title: "Ilustrace", buttons: [{ label: "Obrázky ▾", icon: "🖼️" }, { label: "Obrazce ▾", icon: "🔷" }, { label: "Ikony", icon: "🌿" }, { label: "3D modely ▾", icon: "🧊" }, { label: "SmartArt", icon: "📊" }, { label: "Graf", icon: "📈" }, { label: "Snímek<br>obrazovky ▾", icon: "📸" }] },
                { title: "Odkazy", buttons: [{ label: "Odkaz ▾", icon: "🔗" }, { label: "Záložka", icon: "🔖" }] },
                { title: "Text", buttons: [{ label: "Textové<br>pole ▾", icon: "🅰️" }, { label: "Rychlé<br>části ▾", icon: "🧩" }, { label: "WordArt ▾", icon: "🅦" }, { label: "Datum a čas", icon: "🗓️" }, { label: "Objekt", icon: "📦" }] },
                { title: "Symboly", buttons: [{ label: "Rovnice ▾", icon: "π" }, { label: "Symbol ▾", icon: "Ω" }, { label: "Vodorovná<br>čára", icon: "─" }] }
            ],
            options: [
                { title: "Motivy", buttons: [{ label: "Motivy ▾", icon: "🎨" }, { label: "Barvy ▾", icon: "🟦" }, { label: "Písma ▾", icon: "🅰️" }, { label: "Efekty ▾", icon: "🔵" }] },
                { title: "Zobrazit pole", buttons: [{ label: "Skrytá", icon: "👁️" }, { label: "Od", icon: "👤" }] },
                { title: "Šifrovat", buttons: [{ label: "Šifrovat ▾", icon: "🔒" }] },
                { title: "Sledování", buttons: [{ label: "Použít hlasovací<br>tlačítka ▾", icon: "✅" }, { label: "Požadovat oznámení<br>o doručení", icon: "☐" }, { label: "Požadovat oznámení<br>o přečtení", icon: "☐" }] },
                { title: "Další možnosti", buttons: [{ label: "Uložit odeslanou<br>položku do ▾", icon: "💾" }, { label: "Odložit<br>doručení", icon: "🕒" }, { label: "Odpovědi posílat<br>na adresu", icon: "👤" }] }
            ],
            format: [
                { title: "Formát", buttons: [{ label: "HTML", icon: "Aa" }, { label: "Prostý text", icon: "Aa" }, { label: "Formát RTF", icon: "Aa" }] },
                { title: "Písmo", buttons: [{ label: "Písmo<br>▾", icon: "🔤" }, { label: "Tučně", icon: "B" }, { label: "Kurzíva", icon: "I" }, { label: "Podtržení", icon: "U" }, { label: "Barva<br>písma ▾", icon: "A" }] },
                { title: "Odstavec", buttons: [{ label: "Odrážky ▾", icon: "•" }, { label: "Číslování ▾", icon: "1." }, { label: "Zarovnání ▾", icon: "≡" }, { label: "Odsazení ▾", icon: "⇥" }] },
                { title: "Styly", buttons: [{ label: "Normální", icon: "A" }, { label: "Bez mezer", icon: "A" }, { label: "Nadpis 1", icon: "A" }, { label: "Nadpis 2", icon: "A" }, { label: "Název", icon: "A" }, { label: "Změnit<br>styly ▾", icon: "🅰️" }] },
                { title: "Úpravy", buttons: [{ label: "Najít ▾", icon: "🔍" }, { label: "Nahradit", icon: "↔️" }, { label: "Vybrat ▾", icon: "↖️" }] },
                { title: "Lupa", buttons: [{ label: "Lupa", icon: "🔎" }] }
            ],
            review: [
                { title: "Kontrola", buttons: [{ label: "Pravopis<br>a gramatika", icon: "✔️" }, { label: "Výzkumník", icon: "📚" }] },
                { title: "Jazyk", buttons: [{ label: "Přeložit ▾", icon: "🌐" }] },
                { title: "Komentáře", buttons: [{ label: "Nový<br>komentář", icon: "💬" }] },
                { title: "Sledování", buttons: [{ label: "Sledování<br>změn", icon: "📝" }] }
            ]
        };
        const groups = tabs[s.composeTab] || tabs.message;
        return `
            <div class="ol-ribbon">
                ${groups.map(g => `
                    <div class="ol-rgroup">
                        <div class="ol-rgroup-content">
                            ${g.buttons.map(b => `
                                <button class="ol-rbtn-lg ${b.active ? "active" : ""}" ${b.id ? `id="${b.id}"` : ""}>
                                    <span class="ol-rbtn-ico">${b.icon}</span>
                                    <span>${b.label}</span>
                                </button>
                            `).join("")}
                        </div>
                        <div class="ol-rgroup-title">${g.title}</div>
                    </div>
                `).join("")}
            </div>
        `;
    },

    renderComposeWindow() {
        const s = this.state;
        return `
            <div class="ol-modal-layer" onclick="event.stopPropagation()">
                <div class="ol-win-frame">
                    <header class="ol-titlebar">
                        <div class="ol-titlebar-left">
                            <span>📬 💾 ↩️</span>
                            <span>${s.composeSubject || "Bez názvu"} - Zpráva (HTML)</span>
                        </div>
                        <div class="ol-win-controls"><span style="cursor:pointer;" id="ol-close-compose">✕</span></div>
                    </header>

                    <nav class="ol-menubar">
                        <div class="ol-menubar-left">
                            <div class="ol-menu-item file-btn" id="ol-btn-file-compose">Soubor</div>
                            ${[["message", "Zpráva"], ["insert", "Vložení"], ["options", "Možnosti"], ["format", "Formátování textu"], ["review", "Revize"]].map(([key, label]) => `
                                <div class="ol-menu-item ${s.composeTab === key ? "active" : ""}" data-ol-ctab="${key}">${label}</div>
                            `).join("")}
                            <div class="ol-menu-item">Nápověda</div>
                        </div>
                        <div class="ol-menubar-right"><span class="ol-copilot-pill">🌈 Copilot ▾</span></div>
                    </nav>

                    ${this.renderComposeRibbon()}

                    <div style="padding:12px 18px; display:flex; gap:14px; border-bottom:1px solid #d2d0ce;">
                        <button class="ol-send-tall" id="ol-send-compose">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.7"><polygon points="2 3 22 12 2 21 6 12 2 3"/><path d="M6 12h16"/></svg>
                            <span>Poslat</span>
                        </button>
                        <div style="flex:1;">
                            <div class="ol-field-line"><button class="ol-pill-btn" style="min-width:74px;">Komu</button><input type="text" style="border:none; outline:none; flex:1;"></div>
                            <div class="ol-field-line"><button class="ol-pill-btn" style="min-width:74px;">Kopie</button><input type="text" style="border:none; outline:none; flex:1;"></div>
                            <div class="ol-field-line" style="border-bottom:none;">
                                <span style="width:74px; color:#605e5c; padding-left:6px;">Předmět</span>
                                <input type="text" style="border:none; outline:none; flex:1; font-weight:600;" value="${s.composeSubject}">
                                <span style="font-size:11px; color:#323130;">🛡️ Skupina\Skupina-Viditelna ∨</span>
                            </div>
                        </div>
                    </div>

                    <textarea style="flex:1; border:none; outline:none; padding:16px 20px; font-family:inherit; font-size:13px;" placeholder="Napište text zprávy nebo klikněte vlevo nahoře na Soubor -> Uložit jako pro uložení šablony (*.oft)..."></textarea>
                </div>
            </div>
        `;
    },

    // ========================================================================
    // DIALOG "ULOŽIT JAKO" PRO ŠABLONU *.OFT (Screen 14)
    // ========================================================================
    renderSaveAsDialog() {
        const s = this.state;
        const fakeFolders = ["Archiv_2025", "Copilot_Prompty", "Denní_kontrola", "Tabulky_Excel", "Knihy", "Moje_soubory", "Poznámky"];

        return `
            <div class="ol-modal-layer" style="z-index:1600;" onclick="event.stopPropagation()">
                <div class="ol-dialog-frame" style="width:680px;">
                    <div class="ol-dialog-top">
                        <span>💾 Uložit jako</span>
                        <span style="cursor:pointer;" id="ol-close-saveas">✕</span>
                    </div>
                    <div style="padding:8px 12px; border-bottom:1px solid #edebe9; display:flex; gap:8px; align-items:center;">
                        <span>← → ↑</span>
                        <div style="flex:1; border:1px solid #d2d0ce; padding:3px 8px; background:#fff;">📁 Jana – Ukázková firma ▸ Dokumenty</div>
                        <input type="text" placeholder="Prohledat: Dokumenty" style="width:160px; padding:3px 6px; border:1px solid #d2d0ce;">
                    </div>

                    <div style="display:flex; height:240px; border-bottom:1px solid #d2d0ce;">
                        <div style="width:155px; border-right:1px solid #edebe9; padding:8px; color:#323130;">
                            <div style="padding:4px;">🏠 Domů</div>
                            <div style="padding:4px;">🖼️ Galerie</div>
                            <div style="padding:4px; background:#e1dfdd; font-weight:600;">☁️ OneDrive - Ukázková firma</div>
                            <div style="padding:4px; margin-top:6px;">⬇️ Stažené soubory</div>
                            <div style="padding:4px;">📄 Dokumenty</div>
                        </div>
                        <div style="flex:1; padding:8px 12px; overflow-y:auto;">
                            <table style="width:100%; border-collapse:collapse; font-size:11.5px;">
                                <thead>
                                    <tr style="color:#605e5c; border-bottom:1px solid #edebe9; text-align:left;">
                                        <th style="padding:4px;">Název</th><th>Stav</th><th>Datum změny</th><th>Typ</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${fakeFolders.map(fn => `
                                        <tr>
                                            <td style="padding:4px;">📁 ${fn}</td>
                                            <td style="color:#0078d4;">☁️</td>
                                            <td style="color:#605e5c;">19.10.2026 9:45</td>
                                            <td style="color:#605e5c;">Složka souborů</td>
                                        </tr>
                                    `).join("")}
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <div style="padding:12px 16px; background:#f8f8f8;">
                        <div style="display:flex; align-items:center; margin-bottom:8px;">
                            <label style="width:110px; text-align:right; padding-right:10px;">Název souboru:</label>
                            <input type="text" style="flex:1; padding:4px 6px; border:1px solid #8a8886;" value="${s.saveAsFileName}">
                        </div>
                        <div style="display:flex; align-items:center; margin-bottom:12px;">
                            <label style="width:110px; text-align:right; padding-right:10px;">Uložit jako typ:</label>
                            <select id="ol-saveas-type" style="flex:1; padding:4px 6px; border:1.5px solid #0078d4; font-weight:600;">
                                <option value="msg" ${s.saveAsFileType === "msg" ? "selected" : ""}>Formát zpráv aplikace Outlook - Unicode (*.msg)</option>
                                <option value="oft" ${s.saveAsFileType === "oft" ? "selected" : ""}>Šablona Outlook (*.oft)</option>
                                <option value="txt" ${s.saveAsFileType === "txt" ? "selected" : ""}>Pouze text (*.txt)</option>
                                <option value="htm" ${s.saveAsFileType === "htm" ? "selected" : ""}>HTML (*.htm;*.html)</option>
                            </select>
                        </div>
                        <div style="display:flex; justify-content:flex-end; gap:8px;">
                            <button class="ol-pill-btn" id="ol-confirm-saveas" style="min-width:85px; border:1.5px solid #0078d4; font-weight:600;">Uložit</button>
                            <button class="ol-pill-btn" id="ol-close-saveas2" style="min-width:85px;">Zrušit</button>
                        </div>
                    </div>
                </div>
            </div>
        `;
    },

    // ========================================================================
    // OKNO NOVÁ SCHŮZKA + MÍSTNOSTI + POMOCNÍK PRO PLÁNOVÁNÍ (Screen 03, 04, 05)
    // ========================================================================
    renderMeetingWindow() {
        const s = this.state;
        const rooms = [
            { id: "Zasedacka_Praha_1NP_A_Alpy_A117", cap: "6", loc: "CZ Praha" },
            { id: "Zasedacka_Praha_1NP_A_Balt_A109", cap: "8", loc: "CZ Praha" },
            { id: "Zasedacka_Praha_1NP_A_ Tatry_A121 (Teams)", cap: "30", loc: "CZ Praha" },
            { id: "Zasedacka_Praha_1NP_A_Dunaj_A118", cap: "12", loc: "CZ Praha" },
            { id: "Zasedacka_Praha_1NP_Silentbox_Sever", cap: "2", loc: "CZ Praha" },
            { id: "Zasedacka_Praha_1NP_Silentbox_Jih", cap: "2", loc: "CZ Praha" }
        ];

        return `
            <div class="ol-modal-layer" onclick="event.stopPropagation()">
                <div class="ol-win-frame">
                    <header class="ol-titlebar">
                        <div class="ol-titlebar-left">
                            <span>📬 💾 ↩️</span>
                            <span>Bez názvu - Schůzka</span>
                        </div>
                        <div class="ol-win-controls"><span style="cursor:pointer;" id="ol-close-meeting">✕</span></div>
                    </header>

                    <nav class="ol-menubar">
                        <div class="ol-menubar-left">
                            <div class="ol-menu-item file-btn">Soubor</div>
                            <div class="ol-menu-item ${s.meetingTab === "appointment" ? "active" : ""}" data-ol-mtab="appointment">Schůzka</div>
                            <div class="ol-menu-item ${s.meetingTab === "scheduling" ? "active" : ""}" data-ol-mtab="scheduling">Pomocník pro plánování</div>
                            <div class="ol-menu-item">Vložení</div>
                            <div class="ol-menu-item">Formátování textu</div>
                        </div>
                        <div class="ol-menubar-right"><span class="ol-copilot-pill">🌈 Copilot ▾</span></div>
                    </nav>

                    <!-- Pás karet okna schůzky -->
                    <div class="ol-ribbon">
                        ${s.meetingTab === "appointment" ? `
                            <div class="ol-rgroup">
                                <div class="ol-rgroup-content">
                                    <button class="ol-rbtn-lg"><span>🗑️</span><span>Odstranit</span></button>
                                </div>
                                <div class="ol-rgroup-title">Akce</div>
                            </div>
                            <div class="ol-rgroup">
                                <div class="ol-rgroup-content">
                                    <button class="ol-rbtn-lg"><span style="color:#464eb8; font-size:18px;">🟣</span><span>Připojit se ke<br>schůzce Teams</span></button>
                                    <button class="ol-rbtn-lg"><span>🎛️</span><span>Možnosti<br>schůzky</span></button>
                                    <button class="ol-rbtn-lg"><span style="color:#e81123;">✕</span><span>Nehostovat<br>online</span></button>
                                </div>
                                <div class="ol-rgroup-title">Schůzka Teams</div>
                            </div>
                            <div class="ol-rgroup">
                                <div class="ol-rgroup-content">
                                    <button class="ol-rbtn-lg"><span style="color:#7719aa; font-size:18px;">🅽</span><span>Poslat do<br>OneNotu</span></button>
                                </div>
                                <div class="ol-rgroup-title">OneNote</div>
                            </div>
                            <div class="ol-rgroup">
                                <div class="ol-rgroup-content">
                                    <button class="ol-rbtn-lg"><span>📅✕</span><span>Zrušit<br>pozvání</span></button>
                                    <div class="ol-rstack">
                                        <button class="ol-rbtn-sm">📖 Adresář</button>
                                        <button class="ol-rbtn-sm">👤✓ Kontrola jmen</button>
                                        <button class="ol-rbtn-sm">📩 Možnosti odpovědi ▾</button>
                                    </div>
                                </div>
                                <div class="ol-rgroup-title">Účastníci</div>
                            </div>
                            <div class="ol-rgroup">
                                <div class="ol-rgroup-content">
                                    <div class="ol-rstack">
                                        <div style="font-size:11px;">Zobrazit jako: <select style="font-size:11px;"><option>🟦 Nemám čas</option></select></div>
                                        <div style="font-size:11px; margin-top:3px;">Připomenutí: <select style="font-size:11px;"><option>15 min.</option></select></div>
                                    </div>
                                    <button class="ol-rbtn-lg"><span>🔄</span><span>Opakování</span></button>
                                </div>
                                <div class="ol-rgroup-title">Možnosti</div>
                            </div>
                            <div class="ol-rgroup">
                                <div class="ol-rgroup-content">
                                    <button class="ol-rbtn-lg" id="ol-switch-to-sched"><span>📊</span><span>Pomocník pro<br>plánování</span></button>
                                    <button class="ol-rbtn-lg"><span>🗓️</span><span>Nový Plánovací<br>průzkum</span></button>
                                </div>
                                <div class="ol-rgroup-title">Najít čas</div>
                            </div>
                        ` : `
                            <!-- Screen 05: Pás karet Pomocníka pro plánování -->
                            <div class="ol-rgroup">
                                <div class="ol-rgroup-content" style="position:relative;">
                                    <button class="ol-rbtn-lg"><span style="color:#107c10; font-size:18px;">🔄</span><span>Aktualizovat<br>dostupnost</span></button>
                                    <button class="ol-rbtn-lg active" id="ol-btn-autopick"><span style="font-size:18px;">🔍→</span><span>Automatický<br>termín ▾</span></button>
                                    ${s.autoPickMenuOpen ? `
                                        <div class="ol-ctx-popup" style="position:absolute; left:55px; top:62px; width:215px;">
                                            <div class="ol-ctx-item" data-ol-autopick="11:35"><span>Povinné osoby</span></div>
                                            <div class="ol-ctx-item" data-ol-autopick="11:35"><span>Povinné osoby a jeden zdroj</span></div>
                                            <div class="ol-ctx-item" data-ol-autopick="11:35"><span>✓ Všechny osoby a jeden zdroj</span></div>
                                            <div class="ol-ctx-item" data-ol-autopick="11:35"><span>Všechny osoby a zdroje</span></div>
                                        </div>
                                    ` : ""}
                                </div>
                                <div class="ol-rgroup-title">Plánování</div>
                            </div>
                            <div class="ol-rgroup">
                                <div class="ol-rgroup-content">
                                    <button class="ol-rbtn-lg"><span>👤+</span><span>Přidat<br>účastníky</span></button>
                                    <button class="ol-rbtn-lg"><span>🚪+</span><span>Přidat<br>místnosti</span></button>
                                    <button class="ol-rbtn-lg"><span>🚪🔍</span><span>Vyhledávač<br>místností</span></button>
                                </div>
                                <div class="ol-rgroup-title">Účastníci</div>
                            </div>
                        `}
                    </div>

                    <!-- Tělo okna schůzky -->
                    <div style="display:flex; flex:1; overflow:hidden;">
                        <div style="flex:1; padding:12px 18px; overflow-y:auto; display:flex; flex-direction:column;">
                            <div style="font-size:11.5px; color:#323130; margin-bottom:8px;">ⓘ Toto pozvání na schůzku jste ještě neodeslali.</div>

                            <div style="display:flex; gap:14px; align-items:flex-start; border-bottom:1px solid #d2d0ce; padding-bottom:12px;">
                                <button class="ol-send-tall" id="ol-send-meeting">
                                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#323130" stroke-width="1.7"><polygon points="2 3 22 12 2 21 6 12 2 3"/><path d="M6 12h16"/></svg>
                                    <span>Poslat</span>
                                </button>

                                <div style="flex:1;">
                                    ${s.meetingTab === "appointment" ? `
                                        <div class="ol-field-line" style="border-bottom-color:#0078d4;">
                                            <span style="width:118px; text-align:center; color:#605e5c;">Název</span>
                                            <input type="text" style="border:none; outline:none; flex:1;" placeholder="Zadejte název schůzky...">
                                        </div>
                                        <div class="ol-field-line">
                                            <button class="ol-pill-btn">Povinní účastníci</button>
                                            <input type="text" style="border:none; outline:none; flex:1;" value="Eva Vzorová; Karel Příklad">
                                        </div>
                                        <div class="ol-field-line">
                                            <button class="ol-pill-btn">Nepovinní účastníci</button>
                                            <input type="text" style="border:none; outline:none; flex:1;">
                                        </div>
                                    ` : ""}

                                    <div class="ol-field-line" style="border-bottom:none; padding:6px 0;">
                                        <span style="width:118px; text-align:center; color:#605e5c;">Čas zahájení</span>
                                        <input type="text" value="po 19.10.2026 📅" style="width:115px; padding:3px 6px; border:1px solid #8a8886;">
                                        <select style="padding:3px 6px; border:1px solid #8a8886;">
                                            <option value="11:35" ${s.selectedTimeSlot === "11:35" ? "selected" : ""}>11:35</option>
                                            <option value="13:00" ${s.selectedTimeSlot === "13:00" ? "selected" : ""}>13:00</option>
                                        </select>
                                        <label style="margin-left:8px;"><input type="checkbox"> Celý den</label>
                                        <label style="margin-left:8px;"><input type="checkbox"> 🌐 Časová pásma</label>
                                    </div>
                                    <div class="ol-field-line" style="border-bottom:none; padding:2px 0 6px;">
                                        <span style="width:118px; text-align:center; color:#605e5c;">Čas ukončení</span>
                                        <input type="text" value="po 19.10.2026 📅" style="width:115px; padding:3px 6px; border:1px solid #8a8886;">
                                        <select style="padding:3px 6px; border:1px solid #8a8886;">
                                            <option>12:00</option>
                                        </select>
                                        <span style="color:#0078d4; margin-left:8px; cursor:pointer;">🔄 Nastavit jako opakované</span>
                                    </div>

                                    ${s.meetingTab === "appointment" ? `
                                        <div class="ol-field-line">
                                            <button class="ol-pill-btn">Místo</button>
                                            <input type="text" style="border:none; outline:none; flex:1; text-decoration:underline;" value="${s.selectedRoom || "Schůzka Microsoft Teams"}">
                                            <button class="ol-pill-btn" style="min-width:90px;">🚪🔍 Místnosti</button>
                                        </div>
                                    ` : ""}
                                </div>
                            </div>

                            ${s.meetingTab === "appointment" ? `
                                <!-- Screen 04: Text pozvánky Teams dole -->
                                <div style="padding-top:16px; color:#201f1e;">
                                    <div style="border-top:1px solid #605e5c; width:340px; margin-bottom:14px;"></div>
                                    <h3 style="font-size:16px; margin-bottom:8px;">Schůzka v Microsoft Teams</h3>
                                    <div style="font-size:14px; font-weight:600; margin-bottom:8px;">
                                        Připojit se: <a href="#" style="color:#464eb8;">https://teams.example.com/meet/0000000000000?p=ukazka</a>
                                    </div>
                                    <div style="color:#323130; font-size:11.5px; line-height:1.6;">
                                        ID schůzky: 339 724 107 071 576 5<br>
                                        Přístupový kód: BE7Lv3Pw
                                    </div>
                                    <div style="border-top:1px solid #d2d0ce; margin-top:14px; padding-top:8px; font-size:11px;">
                                        <a href="#" style="color:#464eb8;">Potřebujete pomoc?</a> | <a href="#" style="color:#464eb8;">Systémová reference</a>
                                    </div>
                                </div>
                            ` : `
                                <!-- Screen 05: Tabulka Pomocníka pro plánování -->
                                <div style="margin-top:10px; flex:1; display:flex; flex-direction:column;">
                                    <table class="ol-sched-grid">
                                        <thead>
                                            <tr style="background:#faf9f8;">
                                                <th style="width:210px;">Všichni účastníci</th>
                                                <th>9:00</th>
                                                <th>10:00</th>
                                                <th>11:00</th>
                                                <th style="color:#0078d4;">11:35 – 12:00 (Volno)</th>
                                                <th>13:00</th>
                                                <th>14:00</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <tr>
                                                <td><strong>∨ Povinný účastník</strong></td>
                                                <td class="ol-sched-bar-tentative"></td>
                                                <td class="ol-sched-bar-tentative"></td>
                                                <td class="ol-sched-bar-tentative"></td>
                                                <td class="ol-sched-slot-free ${s.selectedTimeSlot === "11:35" ? "selected" : ""}" data-ol-pickslot="11:35">Volno</td>
                                                <td class="ol-sched-bar-busy"></td>
                                                <td class="ol-sched-bar-tentative"></td>
                                            </tr>
                                            <tr>
                                                <td>☑ 🟢 <u>Jana Testovací</u> 🏠</td>
                                                <td class="ol-sched-bar-tentative"></td>
                                                <td class="ol-sched-bar-tentative"></td>
                                                <td class="ol-sched-bar-tentative"></td>
                                                <td class="ol-sched-slot-free ${s.selectedTimeSlot === "11:35" ? "selected" : ""}" data-ol-pickslot="11:35">✓ Vybrat 11:35</td>
                                                <td class="ol-sched-bar-busy">porada</td>
                                                <td class="ol-sched-bar-tentative"></td>
                                            </tr>
                                            <tr>
                                                <td>☑ 🟢 <u>Eva Vzorová</u></td>
                                                <td></td>
                                                <td class="ol-sched-bar-busy">Nemá čas</td>
                                                <td></td>
                                                <td class="ol-sched-slot-free ${s.selectedTimeSlot === "11:35" ? "selected" : ""}" data-ol-pickslot="11:35">✓ K dispozici</td>
                                                <td class="ol-sched-bar-busy">Nemá čas</td>
                                                <td></td>
                                            </tr>
                                            <tr>
                                                <td><strong>∨ Zdroj (místnost nebo zařízení)</strong></td>
                                                <td></td><td></td><td></td>
                                                <td class="ol-sched-slot-free ${s.selectedTimeSlot === "11:35" ? "selected" : ""}" data-ol-pickslot="11:35">${s.selectedRoom ? "✓ Volná" : "Vyber místnost"}</td>
                                                <td></td><td></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    <div style="margin-top:auto; padding-top:8px; font-size:11px; color:#605e5c; display:flex; gap:14px;">
                                        <span>🟦 Nemám čas</span>
                                        <span>▨ Nezávazně</span>
                                        <span>🟪 Mimo kancelář</span>
                                        <span>░ Práce jinde</span>
                                    </div>
                                </div>
                            `}
                        </div>

                        <!-- Pravý panel Místnosti (Screen 04) -->
                        ${s.meetingTab === "appointment" ? `
                            <aside class="ol-rooms-panel">
                                <div style="display:flex; justify-content:space-between; align-items:center; font-size:14px; margin-bottom:10px;">
                                    <span>Místnosti</span>
                                    <span style="font-size:12px; color:#605e5c;">∨ &nbsp; ✕</span>
                                </div>
                                <div style="display:flex; justify-content:space-between; font-size:11px; color:#323130; margin-bottom:3px;">
                                    <span>Budova</span><span style="color:#605e5c;">Vymazat filtry ∨</span>
                                </div>
                                <select style="width:100%; padding:5px; border:1px solid #d2d0ce; margin-bottom:10px;">
                                    <option>Seznam_mistnosti_Praha</option>
                                </select>
                                <div style="font-weight:600; margin-bottom:8px; font-size:11.5px;">Konferenční místnosti v: Seznam_mistnosti_Praha</div>
                                ${rooms.map(r => `
                                    <div class="ol-room-item ${s.selectedRoom === r.id ? "selected" : ""}" data-ol-room="${r.id}">
                                        <div class="ol-room-circle">🚪</div>
                                        <div style="font-size:11px; line-height:1.35;">
                                            <div style="color:#201f1e; font-weight:600;">${r.id}</div>
                                            <div style="color:#107c10;">🕒 K dispozici &nbsp;<span style="color:#605e5c;">👤 ${r.cap} 📍 ${r.loc} 🔊 🖥️ 📹</span></div>
                                        </div>
                                    </div>
                                `).join("")}
                            </aside>
                        ` : ""}
                    </div>
                </div>
            </div>
        `;
    },

    // ========================================================================
    // OKNO ÚKOLU Z E-MAILU (Screen 16)
    // ========================================================================
    renderTaskWindow() {
        const s = this.state;
        return `
            <div class="ol-modal-layer" onclick="event.stopPropagation()">
                <div class="ol-win-frame" style="height:88%; width:92%;">
                    <header class="ol-titlebar">
                        <div class="ol-titlebar-left">
                            <span>📬 💾 ↩️</span>
                            <span>${s.taskSubject} - Úkol</span>
                        </div>
                        <div class="ol-win-controls"><span style="cursor:pointer;" id="ol-close-task">✕</span></div>
                    </header>

                    <nav class="ol-menubar">
                        <div class="ol-menubar-left">
                            <div class="ol-menu-item file-btn">Soubor</div>
                            <div class="ol-menu-item active">Úkol</div>
                            <div class="ol-menu-item">Vložení</div>
                            <div class="ol-menu-item">Formátování textu</div>
                        </div>
                    </nav>

                    <div class="ol-ribbon">
                        <div class="ol-rgroup">
                            <div class="ol-rgroup-content">
                                <button class="ol-rbtn-lg" id="ol-save-task"><span style="color:#0078d4; font-size:18px;">💾→</span><span>Uložit<br>a zavřít</span></button>
                                <button class="ol-rbtn-lg"><span>🗑️</span><span>Odstranit</span></button>
                                <button class="ol-rbtn-lg"><span>✉️→</span><span>Přeposlat</span></button>
                                <button class="ol-rbtn-lg"><span style="color:#7719aa;">🅽</span><span>Poslat do<br>OneNotu</span></button>
                            </div>
                            <div class="ol-rgroup-title">Akce</div>
                        </div>
                        <div class="ol-rgroup">
                            <div class="ol-rgroup-content">
                                <button class="ol-rbtn-lg active"><span>📋</span><span>Úkol</span></button>
                                <button class="ol-rbtn-lg"><span>📑</span><span>Podrobnosti</span></button>
                            </div>
                            <div class="ol-rgroup-title">Zobrazit</div>
                        </div>
                        <div class="ol-rgroup">
                            <div class="ol-rgroup-content">
                                <button class="ol-rbtn-lg" id="ol-save-task2"><span style="color:#107c10; font-size:18px;">✓</span><span>Označit jako<br>dokončené</span></button>
                                <button class="ol-rbtn-lg"><span>📋👤</span><span>Přiřadit<br>úkol</span></button>
                            </div>
                            <div class="ol-rgroup-title">Spravovat úkol</div>
                        </div>
                    </div>

                    <div style="padding:14px 22px; border-bottom:1px solid #d2d0ce;">
                        <div class="ol-field-line" style="border-bottom-color:#0078d4;">
                            <span style="width:95px; color:#605e5c;">Předmět</span>
                            <input type="text" style="border:none; outline:none; flex:1; font-size:14px;" value="${s.taskSubject}">
                        </div>
                        <div style="display:grid; grid-template-columns: 240px 240px 1fr; gap:8px; margin-top:10px; font-size:11.5px;">
                            <div>Datum zahájení: <input type="text" value="Nikdy 📅" style="width:110px; padding:2px;"></div>
                            <div>Stav: <select style="width:120px; padding:2px;"><option>Nezahájeno</option><option>Probíhá</option></select></div>
                            <div></div>
                            <div>Termín splnění: &nbsp;<input type="text" value="Zítra 📅" style="width:110px; padding:2px;"></div>
                            <div>Priorita: <select style="width:120px; padding:2px;"><option>Střední</option><option>Vysoká</option></select></div>
                            <div>Hotovo (%): <input type="text" value="0%" style="width:45px; padding:2px;"></div>
                        </div>
                    </div>

                    <div style="padding:16px 22px; font-size:12px; line-height:1.5; color:#201f1e;">
                        <div><strong>Od:</strong> Karel Příklad (ServiceDesk) &lt;servicedesk@firma-priklad.cz&gt;</div>
                        <div><strong>Odesláno:</strong> pondělí 19. října 2026 11:46</div>
                        <div><strong>Komu:</strong> Jana Testovací</div>
                        <div><strong>Předmět:</strong> ${s.taskSubject}</div>
                    </div>
                </div>
            </div>
        `;
    },

    // ========================================================================
    // MODÁL: MOŽNOSTI APLIKACE OUTLOOK (Screen 06, 07)
    // ========================================================================
    renderOptionsModal() {
        const s = this.state;
        const leftTabs = [
            { id: "general", label: "Obecné" },
            { id: "mail", label: "Pošta" },
            { id: "cal", label: "Kalendář" },
            { id: "groups", label: "Skupiny" },
            { id: "people", label: "Lidé" },
            { id: "tasks", label: "Úkoly" },
            { id: "search", label: "Hledání" },
            { id: "lang", label: "Jazyk" },
            { id: "access", label: "Přístupnost" },
            { id: "adv", label: "Upřesnit" },
            { id: "ribbon", label: "Přizpůsobit pás karet" },
            { id: "qat", label: "Panel nástrojů Rychlý přístup" },
            { id: "addins", label: "Doplňky" },
            { id: "trust", label: "Centrum zabezpečení" }
        ];

        return `
            <div class="ol-modal-layer" onclick="event.stopPropagation()">
                <div class="ol-dialog-frame" style="width:940px; height:680px;">
                    <div class="ol-dialog-top">
                        <span>Možnosti aplikace Outlook</span>
                        <span>? &nbsp;&nbsp; <span style="cursor:pointer;" id="ol-close-options">✕</span></span>
                    </div>

                    <div style="display:flex; flex:1; overflow:hidden;">
                        <aside style="width:165px; background:#f3f2f1; border-right:1px solid #d2d0ce; padding:4px 0;">
                            ${leftTabs.map(t => `
                                <div style="padding:6px 12px; cursor:pointer; border:1px solid ${s.optionsTab === t.id ? "#323130" : "transparent"}; background:${s.optionsTab === t.id ? "#ffffff" : "transparent"}; font-weight:${s.optionsTab === t.id ? "600" : "400"};"
                                     data-ol-opt="${t.id}">
                                    ${t.label}
                                </div>
                            `).join("")}
                        </aside>

                        <div style="flex:1; padding:16px 22px; overflow-y:auto;">
                            ${s.optionsTab === "general" ? `
                                <!-- Screen 06: Obecné -->
                                <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px;">
                                    <span style="font-size:22px; color:#0078d4;">📋⚙️</span>
                                    <span style="font-size:13px;">Obecné možnosti pro práci s Outlookem</span>
                                </div>
                                <div style="font-weight:600; color:#323130; border-bottom:1px solid #edebe9; padding-bottom:3px; margin-bottom:8px;">Možnosti cloudového úložiště</div>
                                <label style="display:block; margin-bottom:14px;"><input type="checkbox" checked> Ukládat nastavení Outlooku do cloudu (vyžaduje restartování aplikace Outlook)</label>

                                <div style="font-weight:600; color:#323130; border-bottom:1px solid #edebe9; padding-bottom:3px; margin-bottom:8px;">Možnosti uživatelského rozhraní</div>
                                <label style="display:block; margin-bottom:4px;"><input type="radio" checked> Optimalizovat pro vzhled</label>
                                <label style="display:block; margin-bottom:8px;"><input type="checkbox" checked> Při výběru zobrazovat minipanel nástrojů</label>

                                <div style="font-weight:600; color:#323130; border-bottom:1px solid #edebe9; padding-bottom:3px; margin:14px 0 8px;">Vlastní nastavení Microsoft Office</div>
                                <div style="margin-bottom:6px;">Uživatelské jméno: <input type="text" value="Jana Testovací" style="margin-left:12px; padding:3px 6px; width:190px; border:1px solid #8a8886;"></div>
                                <div style="margin-bottom:10px;">Iniciály: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" value="SM" style="margin-left:12px; padding:3px 6px; width:60px; border:1px solid #8a8886;"></div>
                                <div style="margin-bottom:6px;">Motiv Office: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<select style="margin-left:12px; padding:3px 6px; width:140px;"><option>Barevný</option><option>Tmavě šedý</option></select></div>
                            ` : s.optionsTab !== "mail" ? this.renderOptionsExtra(s.optionsTab) : `
                                <!-- Screen 07: Pošta -->
                                <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px;">
                                    <span style="font-size:22px; color:#323130;">✉️</span>
                                    <span style="font-size:13px;">Změna nastavení pro vytvářené a přijímané zprávy</span>
                                </div>
                                <div style="font-weight:600; color:#323130; border-bottom:1px solid #edebe9; padding-bottom:3px; margin-bottom:10px;">Vytváření zpráv</div>

                                <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom:1px solid #edebe9;">
                                    <div>📝 Vytvářet zprávy v tomto formátu: <select style="padding:2px 6px;"><option>HTML</option></select></div>
                                    <button class="ol-pill-btn">Možnosti editoru...</button>
                                </div>

                                <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-bottom:1px solid #edebe9;">
                                    <div>✒️ Vytváření nebo změna podpisů zpráv</div>
                                    <button class="ol-pill-btn" id="ol-btn-open-sig" style="border:2px solid #201f1e; font-weight:600;">Podpisy...</button>
                                </div>

                                ${s.signatureEditorOpen ? `
                                    <div style="background:#f3f9fd; border:1px solid #0078d4; padding:10px; margin:8px 0; border-radius:3px;">
                                        <div style="font-weight:600; margin-bottom:4px;">Upravit podpis (Jana Testovací):</div>
                                        <textarea id="ol-sig-textarea" style="width:100%; height:64px; padding:6px; font-family:inherit; font-size:12px;">${s.signatureText}</textarea>
                                        <button class="ol-pill-btn" id="ol-save-sig-btn" style="margin-top:6px; background:#0078d4; color:white; border-color:#0078d4;">Uložit podpis</button>
                                    </div>
                                ` : ""}

                                <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-bottom:1px solid #edebe9;">
                                    <div>🔤 Pomocí šablon e-mailu můžete změnit výchozí písmo, styly, barvy a pozadí.</div>
                                    <button class="ol-pill-btn" id="ol-btn-email-templates">Šablony e-mailů a písma...</button>
                                </div>

                                <div style="font-weight:600; color:#323130; border-bottom:1px solid #edebe9; padding-bottom:3px; margin:14px 0 8px;">Doručení zprávy</div>
                                <label style="display:block; margin-bottom:4px;"><input type="checkbox" checked> Přehrávat zvuk</label>
                                <label style="display:block; margin-bottom:4px;"><input type="checkbox" checked> Zobrazovat upozornění na ploše</label>
                            `}
                        </div>
                    </div>

                    <div style="padding:10px 16px; border-top:1px solid #d2d0ce; display:flex; justify-content:flex-end; gap:8px; background:#f8f8f8;">
                        <button class="ol-pill-btn" id="ol-ok-options" style="min-width:80px; border:1.5px solid #201f1e;">OK</button>
                        <button class="ol-pill-btn" id="ol-close-options2" style="min-width:80px;">Zrušit</button>
                    </div>
                </div>
            </div>
        `;
    },

    renderOptionsExtra(id) {
        const h = t => `<div style="font-weight:600; color:#323130; border-bottom:1px solid #edebe9; padding-bottom:3px; margin:14px 0 8px;">${t}</div>`;
        const cb = (t, checked = true) => `<label style="display:block; margin-bottom:5px;"><input type="checkbox" ${checked ? "checked" : ""}> ${t}</label>`;
        const rb = (t, checked = false, name = "g") => `<label style="display:block; margin-bottom:5px;"><input type="radio" name="${name}" ${checked ? "checked" : ""}> ${t}</label>`;
        const row = (label, input) => `<div style="margin-bottom:8px;">${label}<span style="margin-left:10px;">${input}</span></div>`;
        const sel = (opts, w = 160) => `<select style="padding:2px 6px; width:${w}px;">${opts.map(o => `<option>${o}</option>`).join("")}</select>`;
        const head = (icon, t) => `<div style="display:flex; align-items:center; gap:10px; margin-bottom:14px;"><span style="font-size:20px;">${icon}</span><span style="font-size:13px;">${t}</span></div>`;
        const btn = (t) => `<button class="ol-pill-btn" style="margin-left:6px;">${t}</button>`;

        switch (id) {
            case "cal":
                return `${head("📅", "Možnosti kalendáře")}
                    ${h("Pracovní doba")}
                    ${row("Pracovní hodiny:", `<input type="text" value="8:00 – 17:00" style="width:140px; padding:2px 6px;">`)}
                    ${row("Pracovní dny:", `<span>Po – Pá</span>`)}
                    ${row("První den v týdnu:", sel(["pondělí", "neděle", "sobota"]))}
                    ${row("První týden v roce:", sel(["Začíná 1. ledna", "První čtyřdenní týden"]))}
                    ${h("Možnosti kalendáře")}
                    ${cb("Automaticky přidávat pozvánky na schůzky online", false)}
                    ${cb("Zkrátit události a schůzky, které jsou v kalendáři", true)}
                    ${row("Výchozí délka schůzky:", sel(["30 minut", "15 minut", "60 minut"], 120))}
                    ${row("Výchozí připomenutí:", sel(["15 minut", "5 minut", "1 hodina"], 120))}
                    ${cb("Zobrazovat čísla týdnů", false)}
                    ${h("Ostatní")}
                    ${cb("Zobrazit návrhy aktivit v kalendáři", true)}`;
            case "groups":
                return `${head("👥", "Možnosti skupin")}
                    ${h("Možnosti uživatelského rozhraní")}
                    ${cb("Zobrazovat nápis Skupiny v navigačním podokně", true)}
                    ${cb("Zobrazovat hlavičky skupin v seznamu zpráv", true)}
                    ${h("Oznámení")}
                    ${cb("Upozornit mě na nové členy skupiny", false)}
                    ${h("Možnosti příloh")}
                    ${rb("Příloha se uloží do skupiny, pokud je to možné", true, "att")}
                    ${rb("Příloha se vždy odešle e-mailem", false, "att")}`;
            case "people":
                return `${head("🧑‍🤝‍🧑", "Možnosti kontaktů a lidí")}
                    ${h("Zobrazení jmen")}
                    ${row("Pořadí jmen:", sel(["Křestní jméno, příjmení", "Příjmení, křestní jméno"], 200))}
                    ${h("Registr kontaktů")}
                    ${row("Výchozí registr kontaktů:", sel(["Kontakty Outlooku"], 200))}
                    ${h("Stav online a fotografie")}
                    ${cb("Zobrazovat fotografie kontaktů", true)}
                    ${cb("Zobrazovat stav online u kontaktů", true)}`;
            case "tasks":
                return `${head("✔️", "Možnosti úkolů")}
                    ${h("Možnosti Mého dne a úkolů")}
                    ${row("Výchozí čas připomenutí:", `<input type="text" value="10:00" style="width:80px; padding:2px 6px;">`)}
                    ${cb("Zobrazit úkoly v zobrazení Můj den", true)}
                    ${h("Barvy")}
                    ${row("Barva úkolů po termínu:", `<span style="display:inline-block; width:40px; height:14px; background:#a4262c; vertical-align:middle;"></span>`)}
                    ${row("Barva dokončených úkolů:", `<span style="display:inline-block; width:40px; height:14px; background:#8a8886; vertical-align:middle;"></span>`)}
                    ${h("Tlačítko Rychlé kliknutí")}
                    ${row("Výchozí kategorie:", sel(["Žádná", "Modrá kategorie", "Červená kategorie"], 160))}
                    ${h("Pracovní čas")}
                    ${row("Hodin za den:", `<input type="text" value="8" style="width:50px; padding:2px 6px;">`)}
                    ${row("Hodin za týden:", `<input type="text" value="40" style="width:50px; padding:2px 6px;">`)}`;
            case "search":
                return `${head("🔍", "Možnosti hledání")}
                    ${h("Zdroje")}
                    ${cb("Hledat v e-mailech", true)}
                    ${cb("Hledat v kalendáři", true)}
                    ${cb("Hledat v kontaktech", true)}
                    ${h("Výsledky")}
                    ${rb("Hledat ve všech složkách", true, "scope")}
                    ${rb("Hledat jen v aktuální složce", false, "scope")}
                    ${h("Zvýraznění")}
                    ${cb("Zvýraznit výsledky hledání", true)}`;
            case "lang":
                return `${head("🌐", "Jazykové možnosti")}
                    ${h("Jazyk uživatelského rozhraní")}
                    <div style="display:flex; gap:8px; align-items:flex-start;">
                        <div style="flex:1; border:1px solid #8a8886; background:#fff; padding:4px;">
                            <div style="padding:3px 6px; background:#deecf9;">Čeština (výchozí)</div>
                            <div style="padding:3px 6px;">English (United States)</div>
                        </div>
                        <div style="display:flex; flex-direction:column; gap:4px;"><button class="ol-pill-btn">▲</button><button class="ol-pill-btn">▼</button></div>
                    </div>
                    ${h("Jazyky pro vytváření a kontrolu pravopisu")}
                    <div style="border:1px solid #8a8886; background:#fff; padding:4px; min-height:60px;">
                        <div style="padding:3px 6px;">Čeština</div>
                        <div style="padding:3px 6px;">English (United States)</div>
                    </div>
                    ${cb("Ukládat jazykové nastavení do cloudu", true)}`;
            case "access":
                return `${head("♿", "Možnosti usnadnění přístupu")}
                    ${h("Kontrola přístupnosti")}
                    ${cb("Kontrolovat přístupnost při vytváření zprávy", true)}
                    ${rb("Zobrazit chyby a upozornění", true, "acc")}
                    ${rb("Zobrazit jen chyby", false, "acc")}
                    ${h("Zvuk")}
                    ${cb("Přehrávat zvuk při chybě přístupnosti", false)}
                    ${h("Zobrazení")}
                    ${cb("Zvýraznit zaměření klávesnicí", true)}
                    ${btn("Šablony...")}`;
            case "adv":
                return `${head("⚙️", "Upřesnit")}
                    ${h("Podokno pro čtení")}
                    ${cb("Označit zprávy jako přečtené při zobrazení v podokně pro čtení", true)}
                    ${h("Zobrazení aplikace")}
                    ${cb("Zobrazit aplikace v Outlooku", true)}
                    ${h("Spuštění")}
                    ${row("Výchozí složka po spuštění:", sel(["Doručená pošta", "Koncepty"], 160))}
                    ${h("Automatické archivování")}
                    ${cb("Automaticky archivovat staré položky", false)}
                    ${h("Soubory a prohlížeč")}
                    ${cb("Otevírat odkazy v novém okně prohlížeče", true)}
                    ${h("Připomenutí")}
                    ${cb("Přehrát zvuk připomenutí", true)}
                    ${cb("Zobrazit připomenutí na obrazovce", true)}
                    ${h("Export")}
                    ${btn("Exportovat...")}`;
            case "ribbon":
                return `${head("🎛️", "Přizpůsobení pásu karet")}
                    <div style="display:flex; gap:10px;">
                        <div style="flex:1; border:1px solid #8a8886; background:#fff; padding:6px; min-height:300px; font-size:12px;">
                            <div style="font-weight:600; margin-bottom:4px;">Karty pásu</div>
                            ${[["Domů", true], ["Posílání a příjem", true], ["Složka", true], ["Zobrazení", true], ["Doplňky", true], ["Nápověda", true], ["Vývojář", false]].map(([t, c]) => `<label style="display:block; margin:3px 0;"><input type="checkbox" ${c ? "checked" : ""}> ${t}</label>`).join("")}
                        </div>
                        <div style="display:flex; flex-direction:column; gap:6px; justify-content:center;"><button class="ol-pill-btn">Přidat &gt;&gt;</button><button class="ol-pill-btn">&lt;&lt; Odebrat</button></div>
                        <div style="flex:1; border:1px solid #8a8886; background:#fff; padding:6px; font-size:12px;">
                            <div style="font-weight:600; margin-bottom:4px;">Vlastní karty</div>
                        </div>
                    </div>
                    <div style="margin-top:8px; display:flex; gap:6px;"><button class="ol-pill-btn">Nová karta</button><button class="ol-pill-btn">Nová skupina</button><button class="ol-pill-btn">Přejmenovat...</button></div>`;
            case "qat":
                return `${head("⚡", "Panel nástrojů Rychlý přístup")}
                    ${h("Přizpůsobit panel nástrojů")}
                    ${row("Vyberte příkazy z:", sel(["Běžné příkazy", "Všechny příkazy"], 200))}
                    ${cb("Uložit", true)}${cb("Zpět", true)}${cb("Opakovat", false)}`;
            case "addins":
                return `${head("🧩", "Doplňky")}
                    ${h("Aktivní doplňky aplikace")}
                    <div style="border:1px solid #d2d0ce; padding:6px; font-size:12px;">Microsoft Copilot<br>Microsoft To Do<br>Lidé</div>
                    ${h("Spravovat")}
                    ${row("Typ doplňku:", sel(["Doplňky COM", "Doplňky Outlooku"], 200))}
                    ${btn("Přejít...")}`;
            case "trust":
                return `${head("🛡️", "Centrum zabezpečení")}
                    ${h("Zabezpečení")}
                    ${cb("Blokovat externí obsah v e-mailech", true)}
                    ${cb("Zobrazovat varování u podezřelých odkazů", true)}
                    ${h("Ochrana před nevyžádanou poštou")}
                    ${rb("Standardní", true, "junk")}
                    ${rb("Vysoká", false, "junk")}
                    ${btn("Nastavení centra zabezpečení...")}`;
            default:
                return "";
        }
    },

    // ========================================================================
    renderAutoReplyModal() {
        const s = this.state;
        return `
            <div class="ol-modal-layer" onclick="event.stopPropagation()">
                <div class="ol-dialog-frame" style="width:575px;">
                    <div class="ol-dialog-top">
                        <span>Automatické odpovědi – jana.testovaci@firma-priklad.cz</span>
                        <span style="cursor:pointer;" id="ol-close-autoreply">✕</span>
                    </div>

                    <div style="padding:14px 18px; background:#f8f8f8;">
                        <div style="margin-bottom:6px;">
                            <label style="cursor:pointer; font-weight:600;">
                                <input type="radio" name="ar_mode" value="off" ${!s.autoReplyEnabled ? "checked" : ""}> Neposílat automatické odpovědi
                            </label>
                        </div>
                        <div style="margin-bottom:12px;">
                            <label style="cursor:pointer; font-weight:600;">
                                <input type="radio" name="ar_mode" value="on" id="ol-radio-autoreply-on" ${s.autoReplyEnabled ? "checked" : ""}> Posílat automatické odpovědi
                            </label>
                            <div style="padding-left:22px; margin-top:6px; color:#605e5c;">
                                <label><input type="checkbox" checked> Posílat jenom během tohoto časového intervalu:</label>
                                <div style="display:flex; align-items:center; gap:8px; margin-top:6px;">
                                    <span style="width:85px;">Počáteční čas:</span>
                                    <input type="text" value="po 19.10.2026" style="width:125px; padding:2px 6px; border:1px solid #c8c6c4;">
                                    <input type="text" value="11:00" style="width:70px; padding:2px 6px; border:1px solid #c8c6c4;">
                                </div>
                                <div style="display:flex; align-items:center; gap:8px; margin-top:4px;">
                                    <span style="width:85px;">Koncový čas:</span>
                                    <input type="text" value="út 20.10.2026" style="width:125px; padding:2px 6px; border:1px solid #c8c6c4;">
                                    <input type="text" value="11:00" style="width:70px; padding:2px 6px; border:1px solid #c8c6c4;">
                                </div>
                            </div>
                        </div>

                        <div style="color:#605e5c; font-size:11.5px; margin-bottom:6px;">Automaticky odpovědět každému odesílateli jednou pomocí následujících zpráv:</div>

                        <div style="display:flex; border-bottom:1px solid #c8c6c4;">
                            <div style="background:#ffffff; border:1px solid #c8c6c4; border-bottom:none; padding:5px 12px; font-weight:600;">👥⬅ Uvnitř organizace</div>
                            <div style="background:#edebe9; border:1px solid #c8c6c4; border-bottom:none; padding:5px 12px; color:#605e5c;">🌍⬅ Mimo organizaci (vypnuto)</div>
                        </div>

                        <div style="background:#ffffff; border:1px solid #c8c6c4; border-top:none; padding:6px;">
                            <div style="display:flex; gap:10px; padding:4px 6px; border-bottom:1px solid #edebe9; color:#605e5c; font-size:11.5px;">
                                <select style="font-size:11px;"><option>Tahoma</option></select>
                                <select style="font-size:11px;"><option>10</option></select>
                                <b>B</b><i>I</i><u>U</u><span>≣</span>
                            </div>
                            <textarea style="width:100%; height:145px; border:none; outline:none; padding:8px; font-family:Tahoma, sans-serif; font-size:12px; resize:none;">${s.autoReplyText}</textarea>
                        </div>

                        <div style="display:flex; justify-content:space-between; align-items:center; margin-top:14px;">
                            <button class="ol-pill-btn" style="min-width:85px;">Pravidla...</button>
                            <div style="display:flex; gap:8px;">
                                <button class="ol-pill-btn" id="ol-save-autoreply" style="min-width:85px; border:1.5px solid #0078d4; font-weight:600;">OK</button>
                                <button class="ol-pill-btn" id="ol-close-autoreply2" style="min-width:85px;">Zrušit</button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    },

    // ========================================================================
    // OBSLUHA UDÁLOSTÍ A VYHODNOCENÍ SANDBOX MISÍ
    // ========================================================================
    bindEvents() {
        const c = this.container;
        const s = this.state;

        c.querySelector("#ol-root")?.addEventListener("click", () => {
            if (s.ctxMenu.open || s.autoPickMenuOpen) {
                s.ctxMenu.open = false;
                s.autoPickMenuOpen = false;
                this.render();
            }
        });

        c.querySelectorAll("[data-ol-rail]").forEach(btn => {
            btn.addEventListener("click", () => {
                s.activeRail = btn.getAttribute("data-ol-rail");
                this.render();
            });
        });

        c.querySelectorAll("[data-ol-tab]").forEach(t => {
            t.addEventListener("click", () => {
                s.activeTab = t.getAttribute("data-ol-tab");
                if (s.activeRail !== "mail" && s.activeRail !== "calendar") s.activeRail = "mail";
                this.render();
            });
        });

        c.querySelectorAll("[data-ol-folder]").forEach(f => {
            f.addEventListener("click", () => {
                const name = f.getAttribute("data-ol-folder");
                s.activeFolder = name;
                s.activeRail = "mail";
                s.expandedFolders[name] = true;
                this.render();
            });
        });

        c.querySelectorAll("[data-ol-fexp]").forEach(ch => {
            ch.addEventListener("click", (e) => {
                e.stopPropagation();
                const name = ch.getAttribute("data-ol-fexp");
                s.expandedFolders[name] = !s.expandedFolders[name];
                this.render();
            });
        });

        c.querySelectorAll("[data-ol-course-task]").forEach(t => {
            t.addEventListener("click", () => {
                s.helpMissionId = t.getAttribute("data-ol-course-task");
                s.helpShowHint = false;
                s.helpOpen = true;
                this.render();
            });
        });

        c.querySelector("#ol-help-open")?.addEventListener("click", () => {
            s.helpOpen = !s.helpOpen;
            s.helpMissionId = null;
            s.helpShowHint = false;
            this.render();
        });
        c.querySelector("#ol-help-close")?.addEventListener("click", () => {
            s.helpOpen = false;
            this.render();
        });
        c.querySelector("#ol-help-hint-toggle")?.addEventListener("click", () => {
            s.helpShowHint = true;
            this.render();
        });
        c.querySelectorAll("[data-ol-help-mission]").forEach(h => {
            h.addEventListener("click", () => {
                s.helpMissionId = h.getAttribute("data-ol-help-mission");
                s.helpShowHint = false;
                this.render();
            });
        });

        c.querySelectorAll("[data-ol-todolist]").forEach(l => {
            l.addEventListener("click", () => {
                s.todoList = l.getAttribute("data-ol-todolist");
                this.render();
            });
        });

        c.querySelectorAll("[data-ol-todo]").forEach(t => {
            t.addEventListener("click", () => {
                const item = s.todos.find(x => String(x.id) === t.getAttribute("data-ol-todo"));
                if (item) item.done = !item.done;
                this.render();
            });
        });

        const addTodo = () => {
            const input = c.querySelector("#ol-todo-input");
            const text = input ? input.value.trim() : "";
            if (!text) return;
            s.todos.push({ id: Date.now(), title: text, done: false });
            this.render();
        };
        c.querySelector("#ol-todo-add-btn")?.addEventListener("click", addTodo);
        c.querySelector("#ol-todo-input")?.addEventListener("keydown", (e) => { if (e.key === "Enter") addTodo(); });

        const sendCopilot = () => {
            const input = c.querySelector("#ol-cop-input");
            const text = input ? input.value.trim() : "";
            if (!text) return;
            s.copilotMessages.push({ role: "user", text });
            s.copilotMessages.push({ role: "bot", text: "Děkuji za dotaz. V ukázkovém prostředí Copilot odpovídá jen orientačně." });
            this.render();
        };
        c.querySelector("#ol-cop-send")?.addEventListener("click", sendCopilot);
        c.querySelector("#ol-cop-input")?.addEventListener("keydown", (e) => { if (e.key === "Enter") sendCopilot(); });
        c.querySelector("#ol-cop-new")?.addEventListener("click", () => {
            s.copilotMessages = [];
            this.render();
        });

        c.querySelector("#ol-help-options")?.addEventListener("click", () => {
            s.optionsModalOpen = true;
            this.render();
        });

        c.querySelectorAll("[data-ol-filter]").forEach(p => {
            p.addEventListener("click", () => {
                s.mailFilter = p.getAttribute("data-ol-filter");
                this.render();
            });
        });

        c.querySelector("#ol-show-preview-link")?.addEventListener("click", () => {
            s.selectedMailId = s.mails[0].id;
            this.render();
        });

        c.querySelector("[data-ol-forms-quiz-link]")?.addEventListener("click", (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.callbacks.onNavigate("forms");
        });

        // Levé a pravé tlačítko myši na e-mailech
        c.querySelectorAll("[data-ol-mail]").forEach(row => {
            const id = parseInt(row.getAttribute("data-ol-mail"), 10);
            row.addEventListener("click", () => {
                s.selectedMailId = id;
                s.ctxMenu.open = false;
                this.render();
            });
            row.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                e.stopPropagation();
                s.selectedMailId = id;
                const clampedX = Math.max(10, Math.min(e.clientX, window.innerWidth - 490));
                const clampedY = Math.max(10, Math.min(e.clientY, window.innerHeight - 380));
                s.ctxMenu = { open: true, x: clampedX, y: clampedY, mailId: id, subMenu: null };
                this.render();
            });
        });

        // Podnabídky kontextového menu
        c.querySelectorAll("[data-ol-sub]").forEach(item => {
            item.addEventListener("mouseenter", () => {
                s.ctxMenu.subMenu = item.getAttribute("data-ol-sub");
                this.render();
            });
            item.addEventListener("click", (e) => {
                e.stopPropagation();
                s.ctxMenu.subMenu = item.getAttribute("data-ol-sub");
                this.render();
            });
        });

        // Zařazení do kategorie (Screen 10)
        c.querySelectorAll("[data-ol-cat]").forEach(el => {
            el.addEventListener("click", (e) => {
                e.stopPropagation();
                const raw = el.getAttribute("data-ol-cat");
                const mail = this.findMail(s.ctxMenu.mailId);
                if (mail) {
                    if (raw === "clear") mail.category = null;
                    else {
                        const [name, color, bg] = raw.split("|");
                        mail.category = { name, color, bg };
                    }
                }
                s.ctxMenu.open = false;
                this.render();
                this.callbacks.onAction({ action: "outlook_context_workflow", completed: true });
            });
        });

        // Zpracovat / Připomenutí -> Otevře okno Úkolu (Screen 15, 16)
        c.querySelectorAll("[data-ol-follow]").forEach(el => {
            el.addEventListener("click", (e) => {
                e.stopPropagation();
                const mode = el.getAttribute("data-ol-follow");
                const mail = this.findMail(s.ctxMenu.mailId) || s.mails[0];
                mail.flag = mode === "done" ? "done" : "followup";
                s.ctxMenu.open = false;
                if (mode !== "done") {
                    s.taskSubject = mail.subject;
                    s.taskWindowOpen = true;
                }
                this.render();
                this.callbacks.onAction({ action: "outlook_context_workflow", completed: true });
            });
        });

        c.querySelectorAll("#ol-save-task, #ol-save-task2").forEach(b => {
            b.addEventListener("click", () => {
                s.taskWindowOpen = false;
                this.render();
                this.callbacks.onAction({ action: "outlook_context_workflow", completed: true });
            });
        });
        c.querySelector("#ol-close-task")?.addEventListener("click", () => { s.taskWindowOpen = false; this.render(); });

        // Pravidla (Screen 12)
        c.querySelectorAll("[data-ol-rule]").forEach(el => {
            el.addEventListener("click", (e) => {
                e.stopPropagation();
                s.ctxMenu.open = false;
                this.render();
                this.callbacks.onAction({ action: "outlook_context_workflow", completed: true });
                this.callbacks.onInfo("Pravidlo pro automatické zpracování e-mailů bylo uloženo.");
            });
        });

        // Backstage menu SOUBOR (Screen 06, 13)
        c.querySelector("#ol-btn-file-main")?.addEventListener("click", () => {
            s.backstageOpen = true;
            s.backstageContext = "main";
            this.render();
        });
        c.querySelector("#ol-btn-file-compose")?.addEventListener("click", () => {
            s.backstageOpen = true;
            s.backstageContext = "compose";
            this.render();
        });
        c.querySelectorAll("#ol-bs-back, #ol-bs-back2").forEach(b => {
            b.addEventListener("click", () => { s.backstageOpen = false; this.render(); });
        });

        // Automatické odpovědi (Screen 08)
        c.querySelector("#ol-bs-autoreply")?.addEventListener("click", () => {
            s.backstageOpen = false;
            s.autoReplyModalOpen = true;
            this.render();
        });
        c.querySelector("#ol-radio-autoreply-on")?.addEventListener("change", () => {
            s.autoReplyEnabled = true;
        });
        c.querySelector("#ol-save-autoreply")?.addEventListener("click", () => {
            const isChecked = c.querySelector("#ol-radio-autoreply-on")?.checked;
            s.autoReplyEnabled = Boolean(isChecked);
            s.autoReplyModalOpen = false;
            this.render();
            if (s.autoReplyEnabled) {
                this.callbacks.onAction({ action: "outlook_autoreply", status: "enabled" });
                this.callbacks.onAction({ action: "save_signature", value: s.signatureText });
            } else {
                this.callbacks.onMistake("Přepni v okně přepínač na 'Posílat automatické odpovědi' a klikni na OK!");
            }
        });
        c.querySelectorAll("#ol-close-autoreply, #ol-close-autoreply2").forEach(b => {
            b.addEventListener("click", () => { s.autoReplyModalOpen = false; this.render(); });
        });

        // Možnosti -> Pošta -> Podpisy (Screen 06, 07)
        c.querySelector("#ol-bs-open-options")?.addEventListener("click", () => {
            s.backstageOpen = false;
            s.optionsModalOpen = true;
            this.render();
        });
        c.querySelectorAll("[data-ol-opt]").forEach(tab => {
            tab.addEventListener("click", () => {
                const id = tab.getAttribute("data-ol-opt");
                if (id) {
                    s.optionsTab = id;
                    this.render();
                }
            });
        });
        c.querySelector("#ol-btn-open-sig")?.addEventListener("click", () => {
            s.signatureEditorOpen = true;
            this.render();
        });
        c.querySelector("#ol-save-sig-btn")?.addEventListener("click", () => {
            const val = c.querySelector("#ol-sig-textarea")?.value.trim() || "Jana Testovací";
            s.signatureText = val;
            s.signatureEditorOpen = false;
            s.optionsModalOpen = false;
            this.render();
            this.callbacks.onAction({ action: "save_signature", value: val });
            this.callbacks.onAction({ action: "outlook_autoreply", status: "enabled" });
        });
        c.querySelectorAll("#ol-close-options, #ol-close-options2, #ol-ok-options").forEach(b => {
            b.addEventListener("click", () => { s.optionsModalOpen = false; this.render(); });
        });

        // Nový e-mail + Soubor -> Uložit jako -> Šablona *.oft (Screen 11, 13, 14)
        c.querySelector("#ol-btn-new-mail")?.addEventListener("click", () => {
            s.composeWindowOpen = true;
            this.render();
        });
        c.querySelector("#ol-close-compose")?.addEventListener("click", () => {
            s.composeWindowOpen = false;
            this.render();
        });
        c.querySelectorAll("#ol-bs-save-as, #ol-bs-save-as2").forEach(b => {
            b.addEventListener("click", () => {
                s.backstageOpen = false;
                s.saveAsDialogOpen = true;
                this.render();
            });
        });
        c.querySelector("#ol-saveas-type")?.addEventListener("change", (e) => {
            s.saveAsFileType = e.target.value;
            s.saveAsFileName = s.saveAsFileType === "oft" ? "Šablona.oft" : "Šablona.msg";
            this.render();
        });
        c.querySelector("#ol-confirm-saveas")?.addEventListener("click", () => {
            const chosen = c.querySelector("#ol-saveas-type")?.value;
            if (chosen === "oft") {
                s.saveAsDialogOpen = false;
                s.composeWindowOpen = false;
                this.render();
                this.callbacks.onAction({ action: "outlook_save_template", fileType: "oft" });
            } else {
                this.callbacks.onMistake("V rozbalovacím seznamu 'Uložit jako typ' vyber 'Šablona Outlook (*.oft)'!");
            }
        });
        c.querySelectorAll("#ol-close-saveas, #ol-close-saveas2").forEach(b => {
            b.addEventListener("click", () => { s.saveAsDialogOpen = false; this.render(); });
        });

        // Kalendář -> Nová schůzka + Místnosti + Pomocník pro plánování (Screen 02, 03, 04, 05)
        c.querySelectorAll("#ol-btn-new-meeting, #ol-btn-new-meeting2, #ol-btn-teams-meeting, [data-ol-cal-cell]").forEach(el => {
            el.addEventListener("click", () => {
                s.meetingWindowOpen = true;
                this.render();
            });
        });
        c.querySelector("#ol-close-meeting")?.addEventListener("click", () => {
            s.meetingWindowOpen = false;
            this.render();
        });
        c.querySelector("#ol-close-calendar-event")?.addEventListener("click", () => {
            s.selectedCalendarEvent = null;
            this.render();
        });
        c.querySelectorAll("[data-ol-mtab]").forEach(t => {
            t.addEventListener("click", () => {
                s.meetingTab = t.getAttribute("data-ol-mtab");
                this.render();
            });
        });
        c.querySelector("#ol-switch-to-sched")?.addEventListener("click", () => {
            s.meetingTab = "scheduling";
            this.render();
        });
        c.querySelectorAll("[data-ol-room]").forEach(r => {
            r.addEventListener("click", () => {
                s.selectedRoom = r.getAttribute("data-ol-room");
                this.render();
            });
        });
        c.querySelectorAll("[data-ol-pickslot]").forEach(sl => {
            sl.addEventListener("click", () => {
                s.selectedTimeSlot = sl.getAttribute("data-ol-pickslot");
                this.render();
            });
        });
        c.querySelector("#ol-btn-autopick")?.addEventListener("click", (e) => {
            e.stopPropagation();
            s.autoPickMenuOpen = !s.autoPickMenuOpen;
            this.render();
        });
        c.querySelectorAll("[data-ol-autopick]").forEach(ap => {
            ap.addEventListener("click", (e) => {
                e.stopPropagation();
                s.selectedTimeSlot = "11:35";
                s.autoPickMenuOpen = false;
                this.render();
            });
        });
        c.querySelector("#ol-send-meeting")?.addEventListener("click", () => {
            if (s.selectedRoom && s.selectedTimeSlot === "11:35") {
                s.meetingWindowOpen = false;
                this.render();
                this.callbacks.onAction({ action: "outlook_schedule_meeting", completed: true });
            } else if (!s.selectedRoom) {
                this.callbacks.onMistake("Nejprve na kartě Schůzka v pravém panelu 'Místnosti' klikni na volnou zasedací místnost!");
            } else {
                this.callbacks.onMistake("Přepni na 'Pomocník pro plánování' a vyber volný termín 11:35 – 12:00!");
            }
        });

        const emit = (action, extra = {}) => this.callbacks.onAction({ action, completed: true, ...extra });

        c.querySelectorAll("[data-ol-cat]").forEach(el => {
            el.addEventListener("click", () => emit("outlook_categorize"));
        });
        c.querySelectorAll('[data-ol-rule="create_rule"]').forEach(el => {
            el.addEventListener("click", () => emit("outlook_create_rule"));
        });
        c.querySelectorAll("#ol-save-task, #ol-save-task2").forEach(b => {
            b.addEventListener("click", () => emit("outlook_create_task"));
        });
        c.querySelector("#ol-btn-archive")?.addEventListener("click", () => emit("outlook_archive"));
        c.querySelector("#ol-search-input")?.addEventListener("keydown", (e) => {
            if (e.key !== "Enter") return;
            const term = e.target.value.trim();
            if (term) emit("outlook_search", { term });
        });
        c.querySelector("#ol-btn-attach")?.addEventListener("click", () => emit("outlook_attach_file"));
        c.querySelector("#ol-btn-dictate")?.addEventListener("click", () => emit("outlook_dictate"));
        c.querySelectorAll("[data-ol-ctab]").forEach(el => el.addEventListener("click", () => {
            s.composeTab = el.dataset.olCtab;
            this.render();
        }));
        c.querySelector("#ol-btn-all-apps")?.addEventListener("click", () => {
            s.composeAppsOpen = !s.composeAppsOpen;
            this.render();
        });
        c.querySelector("#ol-app-my-templates")?.addEventListener("click", () => {
            s.composeAppsOpen = false;
            s.myTemplatesOpen = true;
            this.render();
            this.callbacks.onAction({ action: "outlook_view_templates", completed: true });
        });
        c.querySelector("#ol-close-my-templates")?.addEventListener("click", () => {
            s.myTemplatesOpen = false;
            this.render();
        });
        c.querySelector("#ol-mt-new")?.addEventListener("click", () => {
            this.callbacks.onInfo("Šablona se vytvoří až po uložení v sandboxu – tato ukázka slouží k procvičení zobrazení.");
        });
        c.querySelector("#ol-btn-email-templates")?.addEventListener("click", () => emit("outlook_email_templates"));
        c.querySelector("#ol-btn-filter-adv")?.addEventListener("click", () => emit("outlook_advanced_filter"));
        c.querySelector("#ol-btn-bookings")?.addEventListener("click", () => emit("outlook_bookings"));
        c.querySelector("#ol-btn-share-cal")?.addEventListener("click", () => emit("outlook_share_calendar"));
        c.querySelectorAll("[data-ol-calendar-event]").forEach(el => {
            el.addEventListener("click", (e) => {
                e.stopPropagation();
                s.selectedCalendarEvent = {
                    title: el.dataset.olCalendarEvent,
                    time: el.dataset.olEventTime,
                    date: el.dataset.olEventDate
                };
                this.render();
                emit("outlook_calendar_open");
            });
        });
    }
};