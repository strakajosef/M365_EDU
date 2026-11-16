/**
 * ============================================================================
 * INTERAKTIVNÍ MODUL MICROSOFT TEAMS (teams.js) – DLE SCREENŮ 01–16 + ONEDRIVE
 * ============================================================================
 * Podporuje:
 * - Pravé tlačítko myši (contextmenu) i tři tečky (•••) na chatech a kanálech
 * - Filtry Nepřečtené / Kanály / Chaty + Rychlá zobrazení + Oblíbené
 * - Rozšířený editor formátování zpráv (Screen 08)
 * - Vysouvací panel aplikací ••• včetně vyhledání a spuštění OneDrive v záložce Teams
 * - Kalendář + Plánovač dostupnosti + Anketu o čase (Screen 03–06)
 * - Okamžitou schůzku -> Lobby -> Aktivní hovor -> Sdílení PowerPoint Live a Záznam (Screen 09–13)
 */

// Bezpečné vytvoření objektu, pokud ještě neexistuje
window.SandboxModules = window.SandboxModules || {};

function ensureStylesheet(id, href) {
    if (document.getElementById(id)) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = href;
    link.id = id;
    document.head.appendChild(link);
}

function ensureEmbeddedOneDriveStylesheet() {
    ensureStylesheet("dynamic-css-onedrive-teams", "onedrive_teams.css");
}

window.SandboxModules.teams = {
    state: null,
    container: null,
    callbacks: null,

    init(container, mission, callbacks) {
        // Teams se může spustit z plochy Windows 11 i během jiné mise, takže si teams.css načte sám
        ensureStylesheet("dynamic-css-teams", "teams.css");
        this.container = container;
        this.callbacks = callbacks;
        this.mission = mission;

        // Výchozí stav s anonymizovanými firemními daty
        this.state = {
            activeRail: "chat", // 'activity' | 'chat' | 'calendar' | 'onedrive'
            chatFilter: "all",  // 'all' | 'unread' | 'channels' | 'chats'
            selectedItemType: "channel", // 'channel' | 'chat'
            selectedItemId: "ch_teams",
            appsFlyoutOpen: false,
            appSearchQuery: "",
            composerExpanded: false,
            launchedFromDesktop: false, // Pokud jsme přišli z plochy Windows 11 v úkolu OneDrive

            // Kontextové menu (pravé myšítko nebo •••)
            contextMenu: { open: false, x: 0, y: 0, targetType: null, targetId: null },

            // Kalendář a schůzky
            newEventModalOpen: false,
            eventTab: "details", // 'details' | 'scheduler' | 'poll'
            selectedFreeSlot: null,
            meetNowPopupOpen: false,
            meetingName: "Schůzka týmu",
            meetingPhase: null, // null | 'lobby' | 'active'
            meetingDropdown: null, // null | 'share' | 'more'
            isRecording: false,
            sharedPpt: null,
            micOn: false,
            camOn: false,
            threads: {},
            reactKey: null,
            meetingLocked: false,
            presentersOrganizer: false,
            handRaised: false,
            participants: ["Vy (Organizátor)", "Pavla Příkladová", "Zbyněk Vzorový", "Tomáš Příkladný"],
            mutedParticipants: [],
            peoplePanelOpen: false,
            breakoutPanelOpen: false,
            breakoutRooms: [],
            transcriptOn: false,
            recordingReady: false,
            eventDetailId: null,
            searchQuery: "",
            searchOpen: false,
            attachOpen: false,
            files: [
                { id: "file_rozpocet", name: "Rozpocet_Demo_2026.xlsx", location: "Dokumenty týmu" },
                { id: "file_plan", name: "Plan_projektu_Demo.docx", location: "Dokumenty týmu" }
            ],
            customSections: [],
            sectionPickerFor: null,
            collapsedTeams: {},
            collapsedSections: {},
            convTab: "chat", // 'chat' | 'shared' | 'notes'
            threadKey: null,
            editKey: null,
            msgReplyDraft: "",
            pinnedMessages: [],
            savedMessages: [],
            pinnedApps: [],
            membersTeamId: null,
            teamMembers: {
                team_edu: [
                    { name: "Jana Testovací", role: "vlastník" },
                    { name: "Lukáš Vzorový", role: "člen" },
                    { name: "Marek Testovací", role: "člen" }
                ]
            },
            allDirectory: ["Petra Ukázková", "Filip Ukázkový", "Klára Demová", "Zbyněk Vzorový"],

            // Seznam aktivit (Screen 02)
            activities: [
                { id: "act1", author: "Jana Testovací", action: "publikoval(a) novou zprávu", preview: "Dotaz na testovací licenci pro nový kurz...", location: "Firma Demo s.r.o. > Obecné", time: "16:02", targetCh: "ch_obecne" },
                { id: "act2", author: "Petra Ukázková +2", action: "zareagoval(a) na vaši zprávu", preview: "Nemá tu někdo volnou testovací licenci?", location: "Skupina Demo Analytici", time: "13:31", targetChat: "chat_hrs" },
                { id: "act3", author: "Tomáš Příkladný", action: "vás zmínil(a) @", preview: "Prosím o kontrolu nasazení nového modulu", location: "Projekt Alfa", time: "11:15", targetChat: "chat_rtdm" }
            ],

            // Soukromé a skupinové chaty v sekci Oblíbené (Screen 01, 07, 15)
            chats: [
                { id: "chat_rtdm", name: "Projekt Alfa – Vývoj", initials: "PA", unread: false, muted: false, section: "favorites", messages: [
                    { author: "Tomáš Příkladný", time: "10:14", text: "Ahoj, tak včera večer se podařilo nasadit opravu databáze a zatím to běží stabilně.", mine: false },
                    { author: "Já", time: "10:18", text: "Paráda, díky za info! Konečně se tým v klidu vyspí.", mine: true, reactions: "❤️ 2" }
                ]},
                { id: "chat_hrs", name: "Tým Analytiků Demo", initials: "TA", unread: true, muted: false, section: "favorites", messages: [
                    { author: "Petra Ukázková", time: "12:50", text: "Kolegové, posílám aktualizovanou tabulku výsledků za tento kvartál.", mine: false },
                    { author: "Já", time: "13:03", text: "Díky Jano, mrknu na to hned po poradě.", mine: true, reactions: "👍 1" },
                    { author: "Lukáš Vzorový", time: "13:05", text: "Můžeme k tomu dnes udělat krátkou okamžitou schůzku?", mine: false }
                ]},
                { id: "chat_klima", name: "Klára Demová", initials: "KD", unread: false, muted: false, section: "chats", messages: [
                    { author: "Klára Demová", time: "Včera", text: "Ahoj, pošleš mi prosím odkaz na ten záznam ze školení?", mine: false }
                ]},
                { id: "chat_marek", name: "Marek Testovací", initials: "MT", unread: false, muted: false, section: "chats", messages: [
                    { author: "Marek Testovací", time: "Včera", text: "Zítra mám Home Office, budu na Teams.", mine: false }
                ]}
            ],

            // Týmy a kanály (Screen 01)
            teamsList: [
                {
                    id: "team_edu",
                    name: "Akademie Vzor (_int_demo)",
                    icon: "🎓",
                    channels: [
                        { id: "ch_obecne", name: "Obecné", unread: false, muted: false, posts: [
                            { author: "Jana Testovací", time: "16:02", title: "Vítejte v kanálu Akademie Vzor", body: "Zde najdete základní organizační informace ke školením." }
                        ]},
                        { id: "ch_copilot", name: "01 - Copilot", unread: false, muted: false, posts: [
                            { author: "Lektor Demo", time: "14:10", title: "Tipy na nejlepší prompty", body: "Vyzkoušejte shrnutí dlouhých dokumentů jedním příkazem." }
                        ]},
                        { id: "ch_outlook", name: "02 - Outlook", unread: false, muted: false, posts: [
                            { author: "Lektor Vzor", time: "11:00", title: "Pravidla a podpisy", body: "Jak si nastavit automatický podpis a třídění pošty." }
                        ]},
                        { id: "ch_teams", name: "03 - Teams", unread: true, muted: false, hasVideo: true, posts: [
                            {
                                author: "Lektor Vzor",
                                time: "09:30",
                                title: "Chat a komunikace v Teams – Co si ukážeme",
                                body: "• Individuální a skupinový chat<br>• Reakce na zprávy 👍 ❤️ 😊<br>• Odpovídání ve vlákně<br>• Označení kolegy pomocí @mention<br>• Editace a mazání vlastních zpráv",
                                videoTitle: "Ukazka_zaznamu_schuzky.mp4"
                            }
                        ]},
                        { id: "ch_kampane", name: "99 - Hlučné notifikace (Kampaně)", unread: true, muted: false, posts: [
                            { author: "Automatický Bot", time: "16:10", title: "Systémové hlášení #402", body: "Tento kanál generuje desítky upozornění. Klikněte na něj pravým tlačítkem myši (nebo •••) a zvolte Ztlumit!" }
                        ]}
                    ]
                },
                {
                    id: "team_inovace",
                    name: "Inovační komunita Demo 🚀 (_guest)",
                    icon: "💡",
                    channels: [
                        { id: "ch_ino_1", name: "01 💡 Nápady a testování", unread: false, muted: false, posts: [
                            { author: "Filip Ukázkový", time: "Včera", title: "Testování Power Automate", body: "Sdílejte zde své nápady na automatizaci rutinních úkolů." }
                        ]}
                    ]
                }
            ]
        };

        this.render();
    },

    render() {
        const s = this.state;

        this.container.innerHTML = `
            <div class="tm-app" id="tm-root">
                <!-- HORNÍ LIŠTA TEAMS -->
                <header class="tm-titlebar">
                    <div class="tm-titlebar-left">
                        <span>⬅ ➡</span>
                    </div>
                    <div class="tm-search-box">
                        <span>🔍</span>
                        <input type="text" id="tm-top-search" placeholder="Hledat (Ctrl+E) nebo stisknutím Ctrl+G přejít na kanál">
                    </div>
                    <div class="tm-titlebar-right">
                        ${s.launchedFromDesktop ? `
                            <button id="tm-back-to-win11" style="background:#3b3b3b; color:#fff; border:1px solid #666; border-radius:4px; padding:3px 10px; cursor:pointer; font-size:11.5px; margin-right:8px;">
                                🪟 Plocha Windows 11
                            </button>
                        ` : ""}
                        <span style="color:#adadad; cursor:pointer;">•••</span>
                        <div class="tm-user-avatar" title="Přihlášený uživatel">
                            JT<span class="tm-status-dot"></span>
                        </div>
                    </div>
                </header>

                <!-- HLAVNÍ TĚLO -->
                <div class="tm-body">
                    <!-- 1. LEVÝ PANEL IKON (RAIL) -->
                    <nav class="tm-rail">
                        <div class="tm-rail-item ${s.activeRail === "activity" ? "active" : ""}" data-rail="activity" title="Aktivita">
                            🔔<span class="tm-rail-badge">1</span>
                        </div>
                        <div class="tm-rail-item ${s.activeRail === "chat" ? "active" : ""}" data-rail="chat" title="Chat a Kanály">
                            💬<span class="tm-rail-badge">1</span>
                        </div>
                        <div class="tm-rail-item ${s.activeRail === "calendar" ? "active" : ""}" data-rail="calendar" title="Kalendář">
                            📅
                        </div>
                        <div class="tm-rail-item" data-rail="calls" title="Hovory">📞</div>
                        <div class="tm-rail-item ${s.activeRail === "onedrive" ? "active" : ""}" data-rail="onedrive" title="OneDrive (Klikni pro otevření v Teams)">
                            ☁️
                        </div>
                        <div class="tm-rail-item ${s.appsFlyoutOpen ? "active" : ""}" id="tm-btn-apps-flyout" title="Další přidané aplikace (•••)">
                            •••
                        </div>
                    </nav>

                    <!-- VYSOUVACÍ MENU APLIKACÍ (Screen 16 + OneDrive dle Screen 03) -->
                    ${s.appsFlyoutOpen ? this.renderAppsFlyout() : ""}

                    <!-- 2. STŘEDNÍ A PRAVÁ PLOCHA PODLE ZVOLENÉ SEKCE -->
                    ${s.activeRail === "activity" ? this.renderActivityView() : ""}
                    ${s.activeRail === "chat" ? this.renderChatAndChannelsView() : ""}
                    ${s.activeRail === "calendar" ? this.renderCalendarView() : ""}
                    ${s.activeRail === "onedrive" ? this.renderOneDriveInsideTeams() : ""}
                </div>

                <!-- KONTEXTOVÉ MENU (Pravé myšítko / Tři tečky - Screen 14, 15) -->
                ${s.contextMenu.open ? this.renderContextMenu() : ""}

                <!-- MODÁL: NOVÁ UDÁLOST / PLÁNOVAČ / ANKETA (Screen 04, 05, 06) -->
                ${s.newEventModalOpen ? this.renderNewEventModal() : ""}

                <!-- MODÁL: OKAMŽITÉ ZAHÁJENÍ SCHŮZKY (Screen 09) -->
                ${s.meetNowPopupOpen ? this.renderMeetNowPopup() : ""}

                <!-- OBRAZOVKA LOBBY A BĚŽÍCÍ SCHŮZKY (Screen 10, 11, 12, 13) -->
                ${s.meetingPhase ? this.renderMeetingStage() : ""}

                ${this.renderExtraOverlays()}
            </div>
        `;

        this.bindEvents();

        // Pokud je v Teams právě otevřená záložka OneDrive, navážeme i všechny události OneDrive!
        if (s.activeRail === "onedrive" && window.SandboxModules.onedrive) {
            window.SandboxModules.onedrive.bindEmbeddedEvents(this.container);
        }
    },

    // ========================================================================
    // VYKRESLENÍ ONEDRIVE PŘÍMO UVNITŘ PRAVÉ ČÁSTI TEAMS
    // ========================================================================
    renderOneDriveInsideTeams() {
        if (window.SandboxModules.onedrive) {
            // Pokud ještě OneDrive nebyl inicializován, inicializujeme jeho stav
            if (!window.SandboxModules.onedrive.state) {
                window.SandboxModules.onedrive.initStateOnly(this.container, this.callbacks);
            }
            ensureEmbeddedOneDriveStylesheet();
            return window.SandboxModules.onedrive.renderEmbeddedOneDrive("teams");
        }
        return `<div style="padding:30px; color:#fff;">Načítání aplikace OneDrive v Teams...</div>`;
    },

    renderAppsFlyout() {
        const s = this.state;
        const apps = [
            { id: "onedrive", name: "OneDrive", color: "#0078d4", icon: "☁️" },
            { id: "onenote", name: "OneNote", color: "#7719aa", icon: "N" },
            { id: "shifts", name: "Směny", color: "#464eb8", icon: "🕒" },
            { id: "dynamics", name: "Dynamics 365", color: "#002050", icon: "D" },
            { id: "approvals", name: "Schválení", color: "#0078d4", icon: "✓" },
            { id: "workflows", name: "Pracovní postupy", color: "#0066ff", icon: "⚡" },
            { id: "sharepoint", name: "SharePoint", color: "#038387", icon: "S" },
            { id: "powerbi", name: "Power BI", color: "#f2c811", icon: "📊" }
        ];
        const filtered = apps.filter(a => !s.appSearchQuery || a.name.toLowerCase().includes(s.appSearchQuery.toLowerCase()));

        return `
            <div class="tm-apps-flyout" onclick="event.stopPropagation()">
                <input type="text" class="tm-apps-search" id="tm-apps-search-inp" value="${s.appSearchQuery}" placeholder="🔍 Hledat aplikace (např. oned)">
                ${filtered.map(a => `
                    <div class="tm-app-row" data-launch-app="${a.id}" style="${a.id === "onedrive" ? "border:1px solid #0078d4; background:rgba(0,120,212,0.18); border-radius:6px;" : ""}">
                        <span class="tm-app-icon" style="background:${a.color}">${a.icon}</span>
                        <span><strong>${a.name}</strong></span>
                    </div>
                `).join("")}
            </div>
        `;
    },

    renderActivityView() {
        const s = this.state;
        return `
            <aside class="tm-list-pane">
                <div class="tm-pane-header">
                    <h2>Aktivita</h2>
                    <div class="tm-pane-actions"><button class="tm-icon-btn" title="Nastavení oznámení" data-dots-type="notif" data-dots-id="activity">•••</button><button class="tm-icon-btn" title="Hledat" data-open-search="1">🔍</button></div>
                </div>
                <div class="tm-filter-pills">
                    <span class="tm-pill active">Nepřečtené</span>
                    <span class="tm-pill">@Zmínky</span>
                    <span class="tm-pill">Reakce</span>
                </div>
                <div class="tm-tree-scroll">
                    ${s.activities.map((act, i) => `
                        <div class="tm-activity-card ${i === 0 ? "active" : ""}" data-act-click="${act.id}">
                            <div class="tm-mini-avatar">👤</div>
                            <div>
                                <div><strong>${act.author}</strong> ${act.action}</div>
                                <div style="color:#d1d1d1; font-size:12px; margin-top:2px;">${act.preview}</div>
                                <div class="tm-act-meta">${act.location} • ${act.time}</div>
                            </div>
                        </div>
                    `).join("")}
                </div>
            </aside>
            <section class="tm-main-pane">
                <div class="tm-conv-header">
                    <div class="tm-conv-title">📢 Přehled aktivity – Klikni na položku vlevo pro přechod do vlákna</div>
                </div>
                <div class="tm-messages-area">
                    <div class="tm-channel-post">
                        <div class="tm-post-header">
                            <span class="tm-mini-avatar">JT</span>
                            <strong>Jana Testovací</strong> <small style="color:#999;">16:02</small>
                        </div>
                        <div class="tm-post-title">Dotaz na testovací licenci pro nový kurz</div>
                        <p style="color:#d6d6d6;">Dobrý den, má někdo k dispozici testovací licenci, kterou letos nevyužijete?</p>
                        <div class="tm-post-reply-bar">↩ Odpovědět ve vlákně</div>
                    </div>
                </div>
            </section>
        `;
    },

    renderChatAndChannelsView() {
        const s = this.state;
        const showChats = s.chatFilter === "all" || s.chatFilter === "chats" || s.chatFilter === "unread";
        const showChannels = s.chatFilter === "all" || s.chatFilter === "channels" || s.chatFilter === "unread";

        let activeTitle = "";
        let isChannel = s.selectedItemType === "channel";
        let activeObj = null;

        if (isChannel) {
            for (const tm of s.teamsList) {
                const found = tm.channels.find(c => c.id === s.selectedItemId);
                if (found) { activeObj = found; activeTitle = found.name; break; }
            }
        } else {
            activeObj = s.chats.find(c => c.id === s.selectedItemId) || s.chats[0];
            activeTitle = activeObj.name;
        }

        return `
            <aside class="tm-list-pane">
                <div class="tm-pane-header">
                    <h2>Chat</h2>
                    <div class="tm-pane-actions">
                        <button class="tm-icon-btn" title="Další možnosti" data-info="Další možnosti seznamu chatů jsou v této ukázce omezeny.">•••</button>
                        <button class="tm-icon-btn" title="Hledat" data-open-search="1">🔍</button>
                        <button class="tm-icon-btn" title="Nový chat" data-info="Vytvoření nového chatu není v této ukázce potřeba.">📝</button>
                    </div>
                </div>

                <div class="tm-filter-pills">
                    <span class="tm-pill ${s.chatFilter === "unread" ? "active" : ""}" data-filter="unread">Nepřečtené</span>
                    <span class="tm-pill ${s.chatFilter === "channels" ? "active" : ""}" data-filter="channels">Kanály</span>
                    <span class="tm-pill ${s.chatFilter === "chats" ? "active" : ""}" data-filter="chats">Chaty</span>
                    ${s.chatFilter !== "all" ? `<span class="tm-pill" data-filter="all" title="Zrušit filtr">✕</span>` : ""}
                </div>

                <div class="tm-tree-scroll">
                    <div class="tm-section-label" data-toggle-section="quick">${s.collapsedSections.quick ? "▸" : "▾"} Rychlá zobrazení</div>
                    ${s.collapsedSections.quick ? "" : `
                    <div class="tm-tree-item" data-quick="copilot" data-info="Copilot v této ukázce pouze shrnuje vybraný chat.">
                        <div class="tm-item-left"><span class="tm-mini-avatar" style="background:#464eb8;">✨</span><span>Copilot</span></div>
                    </div>
                    <div class="tm-tree-item ${s.convTab === "saved" ? "active" : ""}" data-quick="saved">
                        <div class="tm-item-left"><span class="tm-mini-avatar" style="background:#333;">🔖</span><span>Uloženo</span></div>
                    </div>
                    `}

                    ${showChats ? `
                        ${this.renderChatSection("favorites", "Oblíbené")}
                        ${this.renderChatSection("chats", "Chaty")}
                    ` : ""}

                    ${showChannels ? `
                        <div class="tm-section-label" data-toggle-section="teams" style="margin-top:8px;">${s.collapsedSections.teams ? "▸" : "▾"} Týmy a kanály</div>
                        ${s.collapsedSections.teams ? "" : s.teamsList.map(tm => `
                            <div class="tm-tree-item tm-team-header" data-toggle-team="${tm.id}" data-ctx-type="team" data-ctx-id="${tm.id}" style="font-weight:700; color:#e0e0e0;">
                                <div class="tm-item-left">
                                    <span>${s.collapsedTeams[tm.id] ? "▸" : "▾"}</span>
                                    <span class="tm-team-square">${tm.icon}</span>
                                    <span>${tm.name}</span>
                                </div>
                                <button class="tm-hover-dots" data-dots-type="team" data-dots-id="${tm.id}">•••</button>
                            </div>
                            ${s.collapsedTeams[tm.id] ? "" : tm.channels.filter(ch => s.chatFilter !== "unread" || ch.unread).map(ch => `
                                <div class="tm-tree-item tm-channel-indent ${isChannel && s.selectedItemId === ch.id ? "active" : ""} ${ch.unread ? "unread" : ""}"
                                     data-select-channel="${ch.id}"
                                     data-ctx-type="channel"
                                     data-ctx-id="${ch.id}">
                                    <div class="tm-item-left">
                                        <span>${ch.name} ${ch.muted ? "🔇" : ""}</span>
                                    </div>
                                    <button class="tm-hover-dots" data-dots-type="channel" data-dots-id="${ch.id}">•••</button>
                                </div>
                            `).join("")}
                        `).join("")}
                    ` : ""}
                </div>
            </aside>

            <section class="tm-main-pane">
                <div class="tm-conv-header">
                    <div class="tm-conv-tabs">
                        <div class="tm-conv-title">
                            <span>${isChannel ? "📢" : "💬"} ${activeTitle}${activeObj?.muted ? "🔇 (Ztlumeno)" : ""}</span>
                        </div>
                        <span class="tm-tab-link ${s.convTab === "chat" ? "active" : ""}" data-tab="chat">${isChannel ? "Příspěvky" : "Chat"}</span>
                        <span class="tm-tab-link ${s.convTab === "shared" ? "active" : ""}" data-tab="shared">Sdíleno</span>
                        <span class="tm-tab-link ${s.convTab === "notes" ? "active" : ""}" data-tab="notes">Poznámky</span>
                    </div>
                    <div class="tm-conv-actions">
                        <button class="tm-btn-dark" id="tm-btn-meet-now">📹 Okamžitá schůzka ▾</button>
                        <button class="tm-btn-dark" id="tm-header-dots" data-dots-type="${isChannel ? "channel" : "chat"}" data-dots-id="${activeObj?.id}">•••</button>
                    </div>
                </div>

                <div class="tm-messages-area">
                    ${s.convTab === "shared" ? this.renderSharedTab() : s.convTab === "notes" ? `<div class="tm-empty-tab">📝 Poznámky k tomuto chatu zatím nejsou.</div>` : s.convTab === "saved" ? this.renderSavedList() : ""}
                    ${s.convTab !== "chat" ? "" : isChannel ? activeObj.posts.map((p, pi) => this.renderPost(activeObj, p, pi)).join("") : activeObj.messages.map((m, mi) => this.renderChatMessage(activeObj.id, m, mi)).join("")}
                    ${s.convTab === "chat" && s.threadKey && s.threadKey.startsWith(activeObj.id + "|") ? this.renderThreadPanel(activeObj) : ""}
                </div>

                <div class="tm-composer-wrap">
                    <div class="tm-composer-box ${s.composerExpanded ? "expanded" : ""}">
                        ${s.composerExpanded ? `
                            <div class="tm-format-toolbar">
                                <span><b>B</b></span>
                                <span><i>I</i></span>
                                <span><u>U</u></span>
                                <span><s>S</s></span>
                                <span>•≡</span>
                                <span>1≡</span>
                                <span>🖍️</span>
                                <span>❞ Citace</span>
                                <span>🔗 Odkaz</span>
                                <span>&lt;/&gt; Kód</span>
                                <span style="margin-left:auto;" id="tm-close-format" title="Zavřít panel formátování">🗑️</span>
                            </div>
                        ` : ""}
                        <div class="tm-composer-row">
                            <input type="text" class="tm-composer-input" id="tm-msg-input" placeholder="${isChannel ? "Publikovat v kanálu nebo odpovědět..." : "Napište zprávu..."}">
                            <div class="tm-composer-tools">
                                <span id="tm-toggle-format" title="Zobrazit možnosti formátování (Screen 08)" style="${s.composerExpanded ? "color:#7f85f5; font-weight:700;" : ""}">A🖌️</span>
                                <span title="Emotikony" data-info="Emotikony jsou v této ukázce pouze vizuální.">😊</span>
                                <span id="tm-attach" title="Přiložit soubor">📎</span>
                                <span title="Loop komponenta" data-info="Komponenty Loop nejsou v této ukázce potřeba.">➰</span>
                                <span title="Další aplikace" data-info="Další aplikace v psacím poli nejsou v této ukázce potřeba.">+</span>
                                <span id="tm-send-msg" style="color:#7f85f5; font-weight:700;" title="Odeslat">➤</span>
                            </div>
                        </div>
                    </div>
                </div>
            </section>
        `;
    },

    renderCalendarView() {
        const days = ["12 Po", "13 Út", "14 St", "15 Čt", "16 Pá (Dnes)"];
        const hours = ["09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00"];
        return `
            <aside class="tm-list-pane" style="width:220px; padding:14px;">
                <h2 style="font-size:18px; margin-bottom:14px;">Kalendář</h2>
                <div style="background:#242424; padding:12px; border-radius:8px; text-align:center; margin-bottom:16px;">
                    <strong>Říjen 2026</strong>
                    <p style="font-size:11px; color:#999; margin-top:6px;">42. týden (12. – 18. říjen)</p>
                </div>
                <div style="font-size:12px; color:#adadad;">
                    <p>☑️ Můj kalendář</p>
                    <p style="margin-top:6px;">☑️ Firemní akce</p>
                </div>
            </aside>
            <section class="tm-cal-grid-wrap">
                <div class="tm-cal-toolbar">
                    <div><strong>12. – 18. říjen 2026</strong> (Pracovní týden)</div>
                    <div style="display:flex; gap:8px;">
                        <button class="tm-btn-dark" id="tm-cal-meet-now">📹 Okamžitá schůzka</button>
                        <button class="tm-btn-purple" id="tm-cal-new-event">➕ Nová událost</button>
                    </div>
                </div>
                <div class="tm-cal-week">
                    <div></div>
                    ${days.map((d, idx) => `<div class="tm-cal-day-head ${idx === 4 ? "today" : ""}">${d}</div>`).join("")}
                    ${hours.map(h => `
                        <div style="padding:8px; font-size:11px; color:#888; text-align:right;">${h}</div>
                        ${days.map((d, dIdx) => `
                            <div class="tm-cal-slot" data-cal-slot="${d}${h}">
                                ${dIdx === 1 && h === "10:00" ? `<div class="tm-cal-event" data-ev-id="ev_standup">Pravidelný Standup</div>` : ""}
                                ${dIdx === 4 && h === "11:00" ? `<div class="tm-cal-event" data-ev-id="ev_skoleni">Školení M365</div>` : ""}
                            </div>
                        `).join("")}
                    `).join("")}
                </div>
            </section>
        `;
    },

    renderContextMenu() {
        const cm = this.state.contextMenu;
        const items = {
            chat: [
                ["open_window", "🗗 Otevřít v novém okně"],
                ["mark_unread", "✉️ Označit jako nepřečtené"],
                ["move_section", "⭐ Přejít do oddílu ▸"],
                ["remove_section", "⊖ Odebrat z oddílu"],
                ["mute", "🔕 Ztlumit / Vypnout oznámení"],
                ["copy_link", "🔗 Kopírovat odkaz"],
                ["hide", "🚫 Skrýt / Opustit"]
            ],
            channel: [
                ["open_window", "🗗 Otevřít v novém okně"],
                ["mark_unread", "✉️ Označit jako nepřečtené"],
                ["mute", "🔕 Ztlumit / Vypnout oznámení"],
                ["manage_apps", "⊞ Spravovat aplikace"],
                ["workflows", "⚡ Pracovní postupy"],
                ["copy_link", "🔗 Kopírovat odkaz"]
            ],
            team: [
                ["manage_members", "👥 Správa členů týmu"],
                ["copy_link", "🔗 Kopírovat odkaz"],
                ["hide", "🚫 Opustit tým"]
            ],
            notif: [
                ["notif_settings", "🔔 Nastavení oznámení"],
                ["notif_mute_all", "🔕 Ztlumit všechna oznámení"]
            ],
            msg: [
                ["thread_reply", "↩ Odpovědět ve vlákně"],
                ["react", "😊 Reagovat"],
                ["edit_msg", "✏️ Upravit zprávu"],
                ["save_msg", "🔖 Uložit zprávu"],
                ["pin_msg", "📌 Připnout do chatu"],
                ["copy_link", "🔗 Kopírovat odkaz"]
            ]
        };
        const list = items[cm.targetType] || items.chat;
        return `
            <div class="tm-ctx-popup" style="left:${cm.x}px; top:${cm.y}px;" onclick="event.stopPropagation()">
                ${list.map(([key, label], i) => `
                    ${cm.targetType === "channel" && i === 1 ? `<div class="tm-ctx-divider"></div>` : ""}
                    ${cm.targetType === "msg" && i === 2 ? `<div class="tm-ctx-divider"></div>` : ""}
                    <div class="tm-ctx-item" data-ctx-action="${key}"><span>${label}</span></div>
                `).join("")}
            </div>
        `;
    },

    renderNewEventModal() {
        const s = this.state;
        return `
            <div class="tm-modal-backdrop" onclick="event.stopPropagation()">
                <div class="tm-event-modal">
                    <div class="tm-event-top">
                        <div style="display:flex; gap:10px; align-items:center;">
                            <strong>Nová událost</strong>
                            <button class="tm-btn-dark ${s.eventTab === "details" ? "active" : ""}" data-ev-tab="details">📅 Událost</button>
                            <button class="tm-btn-dark ${s.eventTab === "scheduler" ? "active" : ""}" data-ev-tab="scheduler">📊 Plánovač (Dostupnost)</button>
                            <button class="tm-btn-dark ${s.eventTab === "poll" ? "active" : ""}" data-ev-tab="poll">🗳️ Anketa o čase</button>
                        </div>
                        <div style="display:flex; gap:8px;">
                            <button class="tm-btn-purple" id="tm-save-event">💾 Poslat / Uložit</button>
                            <button class="tm-btn-dark" id="tm-close-event">✕</button>
                        </div>
                    </div>

                    <div class="tm-event-body">
                        <div class="tm-event-form">
                            <input type="text" class="tm-input-line" id="tm-ev-title" placeholder="Přidat název schůzky" value="Plánování projektu">
                            <input type="text" class="tm-input-line" value="Účastníci: Pavla Příkladová, Zbyněk Vzorový (Klikni na Plánovač nahoře pro kontrolu volného času)">

                            ${s.eventTab === "details" ? `
                                <input type="text" class="tm-input-line" value="Středa 21.10.2026 | Vyber volný čas přes Plánovač nahoře">
                                <input type="text" class="tm-input-line" placeholder="Přidat místnost nebo polohu (např. Zasedačka A)">
                                <textarea class="tm-input-line" style="height:120px;" placeholder="Přidat agendu a popis schůzky..."></textarea>
                            ` : ""}

                            ${s.eventTab === "scheduler" ? `
                                <div style="background:#141414; padding:12px; border-radius:8px; border:1px solid #333;">
                                    <p style="margin-bottom:8px; color:#d1d1d1;"><strong>Plánovač dostupnosti (Screen 06):</strong> Červená pole znamenají, že kolegové nemají čas. Klikni na modře ohraničený sloupec, kdy mají všichni volno!</p>
                                    <table class="tm-scheduler-table">
                                        <thead>
                                            <tr>
                                                <th>Účastník</th>
                                                <th>10:00 – 11:00</th>
                                                <th>11:00 – 11:30 (Volno)</th>
                                                <th>12:00 – 14:00</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <tr>
                                                <td>👤 Vy (Organizátor)</td>
                                                <td class="tm-slot-busy">Obsazeno</td>
                                                <td class="tm-slot-free ${s.selectedFreeSlot === "11:00" ? "selected" : ""}" data-pick-slot="11:00">✅ K dispozici (Klikni)</td>
                                                <td class="tm-slot-busy">Obsazeno</td>
                                            </tr>
                                            <tr>
                                                <td>👤 Pavla Příkladová</td>
                                                <td class="tm-slot-busy">Nemá čas</td>
                                                <td class="tm-slot-free ${s.selectedFreeSlot === "11:00" ? "selected" : ""}" data-pick-slot="11:00">✅ K dispozici (Klikni)</td>
                                                <td class="tm-slot-busy">Nemá čas</td>
                                            </tr>
                                            <tr>
                                                <td>👤 Zbyněk Vzorový</td>
                                                <td class="tm-slot-busy">Nemá čas</td>
                                                <td class="tm-slot-free ${s.selectedFreeSlot === "11:00" ? "selected" : ""}" data-pick-slot="11:00">✅ K dispozici (Klikni)</td>
                                                <td class="tm-slot-busy">Mimo kancelář</td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                            ` : ""}

                            ${s.eventTab === "poll" ? `
                                <div style="background:#141414; padding:14px; border-radius:8px;">
                                    <h4>Vytvořte hlasování o nejlepším čase schůzky (Screen 05):</h4>
                                    <div style="margin-top:8px; padding:8px; background:#292929; border-radius:4px;">☑️ St 21.10.2026 11:05 - 11:30 (Všichni 3 dostupní)</div>
                                    <div style="margin-top:6px; padding:8px; background:#292929; border-radius:4px;">☑️ Čt 22.10.2026 12:05 - 12:30 (2 dostupní, 1 obsazený)</div>
                                </div>
                            ` : ""}
                        </div>
                    </div>
                </div>
            </div>
        `;
    },

    renderMeetNowPopup() {
        return `
            <div class="tm-modal-backdrop" onclick="event.stopPropagation()">
                <div style="background:#292929; border:1px solid #444; border-radius:10px; width:420px; padding:22px;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:14px;">
                        <h3 style="font-size:16px;">Okamžité zahájení schůzky</h3>
                        <span style="cursor:pointer;" id="tm-close-meetnow">✕</span>
                    </div>
                    <label style="font-size:12px; color:#adadad;">Název schůzky</label>
                    <input type="text" class="tm-input-line" style="width:100%; margin:6px 0 18px;" value="${this.state.meetingName}">
                    <div style="display:flex; gap:10px; justify-content:center;">
                        <button class="tm-btn-purple" id="tm-go-lobby">📹 Zahájit schůzku</button>
                        <button class="tm-btn-dark" id="tm-close-meetnow2">🔗 Získat odkaz ke sdílení</button>
                    </div>
                </div>
            </div>
        `;
    },

    renderMeetingStage() {
        const s = this.state;

        if (s.meetingPhase === "lobby") {
            return `
                <div class="tm-meeting-stage">
                    <div class="tm-meeting-bar">
                        <strong>Připojení ke schůzce | ${s.meetingName}</strong>
                        <button class="tm-btn-dark" id="tm-leave-call">✕ Zavřít</button>
                    </div>
                    <div style="flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:16px;">
                        <div style="font-size:32px;">👥</div>
                        <h2>Schůzka v Microsoft Teams</h2>
                        <div style="display:flex; gap:20px; margin-top:10px;">
                            <div style="width:320px; height:190px; background:#292929; border-radius:8px; border:1px solid #444; display:flex; flex-direction:column; justify-content:space-between; padding:14px;">
                                <div style="margin:auto; color:#999;">📷 Náhled kamery s pozadím</div>
                                <div style="display:flex; justify-content:space-between; font-size:12px;">
                                    <span>📹 Kamera zapnuta</span>
                                    <span>✨ Efekty a avatary</span>
                                </div>
                            </div>
                            <div style="width:280px; background:#1f1f1f; border:1px solid #3d3d3d; border-radius:8px; padding:16px; display:flex; flex-direction:column; gap:12px;">
                                <strong>🔘 Zvuk počítače</strong>
                                <div>🎙️ Mikrofon (Intel Smart Sound)</div>
                                <div>🔊 Reproduktory (Realtek Audio)</div>
                                <div style="margin-top:auto; display:flex; gap:10px; justify-content:flex-end;">
                                    <button class="tm-btn-dark" id="tm-leave-call2">Zrušit</button>
                                    <button class="tm-btn-purple" id="tm-join-active-call">Připojit se</button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        }

        return `
            <div class="tm-meeting-stage">
                <div class="tm-meeting-bar">
                    <div>
                        <span>🛡️ 01:25</span>
                        ${s.isRecording ? `<span style="margin-left:12px; color:#f25022; font-weight:700;">🔴 Nahrává se záznam a přepis schůzky</span>` : ""}
                        ${s.sharedPpt ? `<span style="margin-left:12px; color:#7f85f5; font-weight:700;">📊 PowerPoint Live: ${s.sharedPpt}</span>` : ""}
                    </div>
                    <div class="tm-meeting-controls">
                        <button class="tm-ctrl-btn ${s.micOn ? "active" : ""}" id="tm-ctrl-mic"><span>🎙️</span>${s.micOn ? "Mikrofon zap." : "Mikrofon vyp."}</button>
                        <button class="tm-ctrl-btn ${s.camOn ? "active" : ""}" id="tm-ctrl-cam"><span>📹</span>${s.camOn ? "Kamera zap." : "Kamera vyp."}</button>
                        <button class="tm-ctrl-btn ${s.meetingDropdown === "share" ? "active" : ""}" id="tm-meet-share"><span>⬆️</span>Sdílet</button>
                        <button class="tm-ctrl-btn ${s.handRaised ? "active" : ""}" id="tm-ctrl-hand"><span>✋</span>${s.handRaised ? "Spustit ruku" : "Zvednout ruku"}</button>
                        <button class="tm-ctrl-btn" id="tm-ctrl-react"><span>😊</span>Reagovat</button>
                        <button class="tm-ctrl-btn" id="tm-ctrl-chat"><span>💬</span>Chat</button>
                        <button class="tm-ctrl-btn ${s.peoplePanelOpen ? "active" : ""}" id="tm-ctrl-people"><span>👥</span>Lidé</button>
                        <button class="tm-ctrl-btn ${s.meetingDropdown === "more" ? "active" : ""}" id="tm-meet-more"><span>•••</span>Více</button>
                        <button class="tm-btn-leave" id="tm-leave-call">📞 Odejít</button>
                    </div>
                </div>

                ${s.meetingDropdown === "more" ? `
                    <div class="tm-meeting-dropdown">
                        <div class="tm-more-grid">
                            <div class="tm-more-tile" data-meet-action="record">⏺️<br>Záznam</div>
                            <div class="tm-more-tile" data-meet-action="transcribe">📝<br>Přepsat</div>
                            <div class="tm-more-tile" data-meet-action="timer">⏱️<br>Časovač</div>
                            <div class="tm-more-tile" data-meet-action="notes">📄<br>Poznámky</div>
                            <div class="tm-more-tile" data-meet-action="captions">💬<br>Titulky</div>
                            <div class="tm-more-tile" data-meet-action="breakout">🚪<br>Místnosti</div>
                        </div>
                        <div class="tm-ctx-divider"></div>
                        <div class="tm-ctx-item">⚙️ Nastavení zařízení</div>
                    </div>
                ` : ""}

                ${s.meetingDropdown === "share" ? `
                    <div class="tm-meeting-dropdown" style="width:340px;">
                        <h4 style="margin-bottom:8px;">Sdílet obsah</h4>
                        <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:12px;">
                            <div class="tm-more-tile" data-share-screen="desktop">🖥️ Celá obrazovka</div>
                            <div class="tm-more-tile" data-share-screen="window">🗔 Okno (6)</div>
                        </div>
                        <div class="tm-ppt-item" data-share-screen="whiteboard">🟦 <strong>Microsoft Whiteboard</strong> (Spolupráce)</div>
                        <div style="margin-top:10px; font-weight:700; font-size:12px; color:#adadad;">PowerPoint Live (Interaktivní prezentace):</div>
                        <div class="tm-ppt-item" data-share-ppt="Audit_digi_prezentace_2026.pptx">📊 Audit_digi_prezentace_2026.pptx</div>
                        <div class="tm-ppt-item" data-share-ppt="Strategie_M365_Q4.pptx">📊 Strategie_M365_Q4.pptx</div>
                    </div>
                ` : ""}

                <div class="tm-meeting-video">
                    <h3>📹 Aktivní schůzka Microsoft Teams</h3>
                    <p style="color:#adadad; margin-top:8px;">Použij horní ovládací lištu schůzky (např. <strong>⬆️ Sdílet</strong> pro PowerPoint Live nebo <strong>••• Více</strong> pro spuštění Záznamu či Místností).</p>
                </div>
            </div>
        `;
    },

    bindEvents() {
        const c = this.container;
        const s = this.state;

        c.querySelector("#tm-root")?.addEventListener("click", () => {
            if (s.contextMenu.open || s.appsFlyoutOpen) {
                s.contextMenu.open = false;
                s.appsFlyoutOpen = false;
                this.render();
            }
        });

        // Tlačítko návratu na plochu Windows 11 (pokud jsme v úkolu OneDrive)
        c.querySelector("#tm-back-to-win11")?.addEventListener("click", (e) => {
            e.stopPropagation();
            if (window.SandboxModules.onedrive) {
                window.SandboxModules.onedrive.state.environment = "desktop";
                window.SandboxModules.onedrive.render();
            }
        });

        // 1. Přepínání hlavní levé lišty (Rail) – NYNÍ VČETNĚ ONEDRIVE!
        c.querySelectorAll("[data-rail]").forEach(btn => {
            btn.addEventListener("click", (e) => {
                e.stopPropagation();
                const r = btn.getAttribute("data-rail");
                if (r === "activity" || r === "chat" || r === "calendar" || r === "onedrive") {
                    s.activeRail = r;
                    s.appsFlyoutOpen = false;
                    this.render();
                } else {
                    this.callbacks.onInfo("Tato sekce slouží jako přehled. Tvůj úkol najdeš v Chat/Týmy, Kalendáři nebo v OneDrive.");
                }
            });
        });

        // Otevření nabídky aplikací ••• v Railu (Screen 16 + vyhledání OneDrive dle Screen 03)
        c.querySelector("#tm-btn-apps-flyout")?.addEventListener("click", (e) => {
            e.stopPropagation();
            s.appsFlyoutOpen = !s.appsFlyoutOpen;
            this.render();
        });

        c.querySelector("#tm-apps-search-inp")?.addEventListener("input", (e) => {
            s.appSearchQuery = e.target.value;
            this.render();
            const inp = this.container.querySelector("#tm-apps-search-inp");
            if (inp) {
                inp.focus();
                inp.setSelectionRange(inp.value.length, inp.value.length);
            }
        });

        c.querySelectorAll("[data-launch-app]").forEach(row => {
            row.addEventListener("click", () => {
                const appId = row.getAttribute("data-launch-app");
                s.appsFlyoutOpen = false;
                if (appId === "onedrive") {
                    s.activeRail = "onedrive";
                    this.render();
                    this.callbacks.onInfo("OneDrive byl otevřen přímo v záložce Microsoft Teams.");
                    return;
                }
                this.render();
                this.callbacks.onAction({ action: "open_teams_app", appId: appId });
                this.callbacks.onInfo(`Otevřel(a) jsi aplikaci '${appId}' v prostředí Teams.`);
            });
        });

        // Horní vyhledávání v Teams (pokud je aktivní OneDrive, hledá i v souborech OneDrive)
        c.querySelector("#tm-top-search")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && s.activeRail === "onedrive" && window.SandboxModules.onedrive) {
                window.SandboxModules.onedrive.state.searchQuery = e.target.value;
                this.render();
                this.callbacks.onAction({ action: "onedrive_search", completed: true });
            }
        });

        // 2. Pilulkové filtry (Nepřečtené | Kanály | Chaty)
        c.querySelectorAll("[data-filter]").forEach(pill => {
            pill.addEventListener("click", () => {
                s.chatFilter = pill.getAttribute("data-filter");
                this.render();
            });
        });

        // 3. Výběr Chatu nebo Kanálu levým tlačítkem
        c.querySelectorAll("[data-select-chat]").forEach(el => {
            el.addEventListener("click", () => {
                s.selectedItemType = "chat";
                s.selectedItemId = el.getAttribute("data-select-chat");
                s.contextMenu.open = false;
                this.render();
            });
        });
        c.querySelectorAll("[data-select-channel]").forEach(el => {
            el.addEventListener("click", () => {
                s.selectedItemType = "channel";
                s.selectedItemId = el.getAttribute("data-select-channel");
                s.contextMenu.open = false;
                this.render();
            });
        });

        // 4. PRAVÉ TLAČÍTKO MYŠI (contextmenu) A TŘI TEČKY (•••) NA CHATECH A KANÁLECH
        c.querySelectorAll("[data-ctx-id]").forEach(item => {
            item.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.openContextMenu(e.clientX, e.clientY, item.getAttribute("data-ctx-type"), item.getAttribute("data-ctx-id"));
            });
        });

        c.querySelectorAll("[data-dots-id]").forEach(btn => {
            btn.addEventListener("click", (e) => {
                e.stopPropagation();
                const rect = btn.getBoundingClientRect();
                this.openContextMenu(rect.left - 180, rect.bottom + 4, btn.getAttribute("data-dots-type"), btn.getAttribute("data-dots-id"));
            });
        });

        c.querySelectorAll("[data-ctx-action]").forEach(actEl => {
            actEl.addEventListener("click", (e) => {
                e.stopPropagation();
                const actionName = actEl.getAttribute("data-ctx-action");
                const { targetType, targetId } = s.contextMenu;
                s.contextMenu.open = false;
                this.handleContextMenuAction(actionName, targetType, targetId);
            });
        });

        // 5. Rozšířený editor formátování a odeslání zprávy
        c.querySelector("#tm-toggle-format")?.addEventListener("click", () => {
            s.composerExpanded = !s.composerExpanded;
            this.render();
            if (s.composerExpanded) {
                this.callbacks.onAction({ action: "expand_formatter", value: true });
            }
        });
        c.querySelector("#tm-close-format")?.addEventListener("click", () => {
            s.composerExpanded = false;
            this.render();
        });
        c.querySelector("#tm-send-msg")?.addEventListener("click", () => this.sendMessage());
        c.querySelector("#tm-msg-input")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter") this.sendMessage();
        });

        // 6. Kalendář a Plánovač
        c.querySelector("#tm-cal-new-event")?.addEventListener("click", () => {
            s.newEventModalOpen = true;
            s.eventTab = "details";
            this.render();
        });
        c.querySelectorAll("[data-cal-slot]").forEach(slot => {
            slot.addEventListener("click", () => {
                s.newEventModalOpen = true;
                s.eventTab = "details";
                this.render();
            });
        });
        c.querySelector("#tm-close-event")?.addEventListener("click", () => {
            s.newEventModalOpen = false;
            this.render();
        });
        c.querySelectorAll("[data-ev-tab]").forEach(btn => {
            btn.addEventListener("click", () => {
                s.eventTab = btn.getAttribute("data-ev-tab");
                this.render();
            });
        });
        c.querySelectorAll("[data-pick-slot]").forEach(cell => {
            cell.addEventListener("click", () => {
                s.selectedFreeSlot = cell.getAttribute("data-pick-slot");
                this.render();
            });
        });
        c.querySelector("#tm-save-event")?.addEventListener("click", () => {
            if (s.selectedFreeSlot === "11:00") {
                s.newEventModalOpen = false;
                this.render();
                this.callbacks.onAction({ action: "schedule_meeting_free_slot", slot: "11:00" });
            } else {
                this.callbacks.onMistake("Než schůzku uložíš, přepni nahoře na '📊 Plánovač (Dostupnost)' a vyber volný čas (11:00 – 11:30), kdy mají všichni účastníci čas!");
            }
        });

        // 7. Okamžitá schůzka a ovládání hovoru
        c.querySelector("#tm-btn-meet-now")?.addEventListener("click", () => { s.meetNowPopupOpen = true; this.render(); });
        c.querySelector("#tm-cal-meet-now")?.addEventListener("click", () => { s.meetNowPopupOpen = true; this.render(); });
        c.querySelector("#tm-close-meetnow")?.addEventListener("click", () => { s.meetNowPopupOpen = false; this.render(); });
        c.querySelector("#tm-close-meetnow2")?.addEventListener("click", () => { s.meetNowPopupOpen = false; this.render(); });
        c.querySelector("#tm-go-lobby")?.addEventListener("click", () => {
            s.meetNowPopupOpen = false;
            s.meetingPhase = "lobby";
            this.render();
        });
        c.querySelector("#tm-join-active-call")?.addEventListener("click", () => {
            s.meetingPhase = "active";
            this.render();
        });
        c.querySelectorAll("#tm-leave-call, #tm-leave-call2").forEach(b => {
            b.addEventListener("click", () => {
                s.meetingPhase = null;
                s.meetingDropdown = null;
                this.render();
            });
        });

        c.querySelector("#tm-meet-more")?.addEventListener("click", () => {
            s.meetingDropdown = s.meetingDropdown === "more" ? null : "more";
            this.render();
        });
        c.querySelector("#tm-meet-share")?.addEventListener("click", () => {
            s.meetingDropdown = s.meetingDropdown === "share" ? null : "share";
            this.render();
        });

        c.querySelectorAll("[data-meet-action]").forEach(tile => {
            tile.addEventListener("click", () => {
                const act = tile.getAttribute("data-meet-action");
                s.meetingDropdown = null;
                if (act === "record" || act === "transcribe") {
                    s.isRecording = true;
                    this.render();
                    this.callbacks.onAction({ action: "meeting_feature", feature: "record" });
                } else if (act === "breakout") {
                    this.render();
                    this.callbacks.onAction({ action: "meeting_feature", feature: "breakout" });
                } else {
                    this.render();
                    this.callbacks.onInfo(`Aktivoval(a) jsi funkci '${act}'.`);
                }
            });
        });

        c.querySelectorAll("[data-share-ppt]").forEach(item => {
            item.addEventListener("click", () => {
                const pptName = item.getAttribute("data-share-ppt");
                s.sharedPpt = pptName;
                s.meetingDropdown = null;
                this.render();
                this.callbacks.onAction({ action: "share_ppt_live", status: "started" });
            });
        });
        c.querySelectorAll("[data-share-screen]").forEach(item => {
            item.addEventListener("click", () => {
                s.meetingDropdown = null;
                this.render();
                this.callbacks.onMistake("Zvolil(a) jsi sdílení celé obrazovky/okna. Úkolem je nasdílet interaktivní prezentaci přes sekci 'PowerPoint Live' dole v nabídce Sdílet!");
            });
        });

        this.bindExtraEvents(c);
    },

    bindExtraEvents(c) {
        const s = this.state;

        c.querySelectorAll("[data-toggle-section]").forEach(el => {
            el.addEventListener("click", () => {
                const id = el.getAttribute("data-toggle-section");
                s.collapsedSections[id] = !s.collapsedSections[id];
                this.render();
            });
        });
        c.querySelectorAll("[data-toggle-team]").forEach(el => {
            el.addEventListener("click", (e) => {
                if (e.target.closest(".tm-hover-dots")) return;
                const id = el.getAttribute("data-toggle-team");
                s.collapsedTeams[id] = !s.collapsedTeams[id];
                this.render();
            });
        });
        c.querySelectorAll("[data-info]").forEach(el => {
            if (el.hasAttribute("data-select-chat") || el.hasAttribute("data-select-channel")) return;
            el.addEventListener("click", () => this.callbacks.onInfo(el.getAttribute("data-info")));
        });
        c.querySelectorAll("[data-quick]").forEach(el => {
            el.addEventListener("click", () => {
                const kind = el.getAttribute("data-quick");
                if (kind === "saved") {
                    s.convTab = "saved";
                    this.render();
                } else {
                    this.callbacks.onInfo("Copilot v této ukázce pouze shrnuje vybraný chat.");
                }
            });
        });
        c.querySelectorAll("[data-tab]").forEach(el => {
            el.addEventListener("click", () => {
                s.convTab = el.getAttribute("data-tab");
                this.render();
            });
        });
        c.querySelectorAll("[data-open-search]").forEach(el => {
            el.addEventListener("click", () => {
                s.searchOpen = true;
                this.render();
                this.container.querySelector("#tm-search-inp")?.focus();
            });
        });
        c.querySelector("#tm-search-inp")?.addEventListener("input", (e) => {
            s.searchQuery = e.target.value;
            const pos = e.target.selectionStart;
            this.render();
            const inp = this.container.querySelector("#tm-search-inp");
            if (inp) { inp.focus(); inp.setSelectionRange(pos, pos); }
            if (s.searchQuery.trim()) this.callbacks.onAction({ action: "search_teams", query: s.searchQuery.trim() });
        });
        c.querySelectorAll("[data-close-overlay]").forEach(el => {
            el.addEventListener("click", () => {
                const which = el.getAttribute("data-close-overlay");
                if (which === "search") s.searchOpen = false;
                if (which === "members") s.membersTeamId = null;
                if (which === "people") s.peoplePanelOpen = false;
                if (which === "breakout") s.breakoutPanelOpen = false;
                if (which === "attach") s.attachOpen = false;
                this.render();
            });
        });
        c.querySelectorAll("[data-unsave]").forEach(el => {
            el.addEventListener("click", () => {
                s.savedMessages.splice(Number(el.getAttribute("data-unsave")), 1);
                this.render();
            });
        });

        c.querySelectorAll("[data-reply], [data-react-open], [data-edit-open], [data-save-msg], [data-pin-msg]").forEach(el => {
            el.addEventListener("click", (e) => {
                e.stopPropagation();
                const key = el.getAttribute("data-reply") || el.getAttribute("data-react-open") || el.getAttribute("data-edit-open") || el.getAttribute("data-save-msg") || el.getAttribute("data-pin-msg");
                if (el.hasAttribute("data-reply")) {
                    s.threadKey = key;
                    this.render();
                    this.callbacks.onAction({ action: "thread_reply", targetId: key });
                } else if (el.hasAttribute("data-react-open")) {
                    s.reactKey = s.reactKey === key ? null : key;
                    this.render();
                } else if (el.hasAttribute("data-edit-open")) {
                    s.editKey = key;
                    this.render();
                } else if (el.hasAttribute("data-save-msg")) {
                    this.saveMessage(key);
                } else {
                    this.pinMessage(key);
                }
            });
        });
        c.querySelectorAll("[data-react]").forEach(el => {
            el.addEventListener("click", () => {
                const key = el.getAttribute("data-react");
                const emoji = el.getAttribute("data-emoji");
                const m = this.findMessage(key);
                if (m) {
                    m.reactions = m.reactions || [];
                    const r = m.reactions.find(x => x.emoji === emoji);
                    if (r) r.count++; else m.reactions.push({ emoji, count: 1 });
                }
                s.reactKey = null;
                this.render();
                this.callbacks.onAction({ action: "react_message", targetId: key, emoji });
            });
        });
        c.querySelector("[data-edit-save]")?.addEventListener("click", (e) => {
            const key = e.currentTarget.getAttribute("data-edit-save");
            const val = this.container.querySelector("#tm-edit-inp")?.value.trim();
            const m = this.findMessage(key);
            if (m && val) {
                m.text = val;
                s.editKey = null;
                this.render();
                this.callbacks.onAction({ action: "edit_message", targetId: key });
            }
        });
        c.querySelector("[data-edit-cancel]")?.addEventListener("click", () => {
            s.editKey = null;
            this.render();
        });
        c.querySelector("#tm-edit-inp")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter") c.querySelector("[data-edit-save]")?.click();
        });

        c.querySelector("[data-thread-close]")?.addEventListener("click", () => {
            s.threadKey = null;
            this.render();
        });
        c.querySelectorAll("[data-thread-key]").forEach(el => {
            el.addEventListener("click", () => {
                s.threadKey = el.getAttribute("data-thread-key");
                this.render();
                this.callbacks.onAction({ action: "thread_reply", targetId: s.threadKey });
            });
        });
        const sendThread = () => {
            const inp = this.container.querySelector("#tm-thread-inp");
            const txt = inp?.value.trim();
            if (!txt || !s.threadKey) return;
            s.threads = s.threads || {};
            (s.threads[s.threadKey] = s.threads[s.threadKey] || []).push({ author: "Já", text: txt });
            this.render();
            this.callbacks.onAction({ action: "thread_reply", targetId: s.threadKey, text: txt });
        };
        c.querySelector("#tm-thread-send")?.addEventListener("click", sendThread);
        c.querySelector("#tm-thread-inp")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter") sendThread();
        });

        c.querySelector("#tm-play-video")?.addEventListener("click", () => {
            this.callbacks.onAction({ action: "play_recording", status: "started" });
            this.callbacks.onInfo("Přehrávání záznamu schůzky spuštěno.");
        });
        c.querySelector("#tm-attach")?.addEventListener("click", () => {
            s.attachOpen = !s.attachOpen;
            this.render();
        });
        c.querySelector("[data-attach-file]")?.addEventListener("click", () => {
            s.attachOpen = false;
            this.render();
            this.callbacks.onAction({ action: "share_file", source: "onedrive" });
        });
        c.querySelector("#tm-share-file")?.addEventListener("click", () => {
            this.callbacks.onAction({ action: "share_file", source: "shared_tab" });
            this.callbacks.onInfo("Soubor byl sdílen v konverzaci.");
        });

        c.querySelectorAll("[data-member-role]").forEach(el => {
            el.addEventListener("click", () => {
                this.callbacks.onAction({ action: "manage_members", teamId: s.membersTeamId, change: "role" });
            });
        });
        c.querySelectorAll("[data-breakout-assign]").forEach(el => {
            el.addEventListener("click", () => {
                this.callbacks.onAction({ action: "meeting_feature", feature: "breakout_assign" });
            });
        });

        c.querySelector("#tm-ctrl-mic")?.addEventListener("click", () => {
            s.micOn = !s.micOn;
            this.render();
            this.callbacks.onAction({ action: "toggle_mic", value: s.micOn });
        });
        c.querySelector("#tm-ctrl-cam")?.addEventListener("click", () => {
            s.camOn = !s.camOn;
            this.render();
            this.callbacks.onAction({ action: "toggle_camera", value: s.camOn });
        });
        c.querySelector("#tm-ctrl-hand")?.addEventListener("click", () => {
            s.handRaised = !s.handRaised;
            this.render();
            this.callbacks.onAction({ action: "raise_hand", value: s.handRaised });
        });
        c.querySelector("#tm-ctrl-react")?.addEventListener("click", () => {
            this.callbacks.onAction({ action: "meeting_reaction", reaction: "clap" });
            this.callbacks.onInfo("Reakce odeslána do schůzky.");
        });
        c.querySelector("#tm-ctrl-chat")?.addEventListener("click", () => {
            this.callbacks.onAction({ action: "meeting_chat_open" });
            this.callbacks.onInfo("Chat schůzky otevřen.");
        });
        c.querySelector("#tm-ctrl-people")?.addEventListener("click", () => {
            s.peoplePanelOpen = !s.peoplePanelOpen;
            this.render();
            if (s.peoplePanelOpen) this.callbacks.onAction({ action: "meeting_feature", feature: "people" });
        });
    },

    findMessage(key) {
        const [chatId, mi] = key.split("|");
        const chat = this.state.chats.find(c => c.id === chatId);
        return chat ? chat.messages[Number(mi)] : null;
    },

    saveMessage(key) {
        const [chatId, mi] = key.split("|");
        const chat = this.state.chats.find(c => c.id === chatId);
        const m = this.findMessage(key);
        if (!m) return;
        m.saved = true;
        this.state.savedMessages = this.state.savedMessages || [];
        this.state.savedMessages.push({ author: m.author, time: m.time, text: m.text, chatName: chat.name });
        this.render();
        this.callbacks.onAction({ action: "save_message", targetId: key });
    },

    pinMessage(key) {
        const m = this.findMessage(key);
        if (!m) return;
        m.pinned = true;
        this.state.pinnedMessages = this.state.pinnedMessages || [];
        this.state.pinnedMessages.push(key);
        this.render();
        this.callbacks.onAction({ action: "pin_message", targetId: key });
    },

    openContextMenu(x, y, targetType, targetId) {
        const clampedX = Math.max(10, Math.min(x, window.innerWidth - 250));
        const clampedY = Math.max(10, Math.min(y, window.innerHeight - 320));
        this.state.contextMenu = { open: true, x: clampedX, y: clampedY, targetType, targetId };
        this.render();
    },

    handleContextMenuAction(actionName, targetType, targetId) {
        const s = this.state;

        if (targetType === "msg") {
            if (actionName === "thread_reply") {
                s.threadKey = targetId;
                this.render();
                this.callbacks.onAction({ action: "thread_reply", targetId });
            } else if (actionName === "react") {
                s.reactKey = targetId;
                this.render();
            } else if (actionName === "edit_msg") {
                s.editKey = targetId;
                this.render();
            } else if (actionName === "save_msg") {
                this.saveMessage(targetId);
            } else if (actionName === "pin_msg") {
                this.pinMessage(targetId);
            } else {
                this.render();
                this.callbacks.onInfo("Odkaz na zprávu byl zkopírován do schránky.");
            }
            return;
        }

        if (targetType === "team") {
            if (actionName === "manage_members") {
                s.membersTeamId = targetId;
                this.render();
                this.callbacks.onAction({ action: "manage_members", teamId: targetId });
            } else {
                this.render();
                this.callbacks.onInfo("Tato akce v ukázce není potřeba.");
            }
            return;
        }

        if (targetType === "notif") {
            this.render();
            if (actionName === "notif_settings") {
                this.callbacks.onAction({ action: "notif_settings" });
            } else if (actionName === "notif_mute_all") {
                this.callbacks.onAction({ action: "notif_mute_all" });
            }
            return;
        }

        if (actionName === "open_window" || actionName === "copy_link" || actionName === "hide" || actionName === "remove_section" || actionName === "manage_apps" || actionName === "workflows") {
            this.render();
            this.callbacks.onInfo("Tato položka je v této ukázce pouze informativní.");
            return;
        }

        let targetObj = null;
        if (targetType === "chat") {
            targetObj = this.state.chats.find(c => c.id === targetId);
        } else {
            for (const tm of this.state.teamsList) {
                const found = tm.channels.find(c => c.id === targetId);
                if (found) { targetObj = found; break; }
            }
        }
        if (!targetObj) return;

        if (actionName === "mute") {
            targetObj.muted = !targetObj.muted;
            this.render();
            this.callbacks.onAction({
                action: "mute_item",
                targetId: targetId,
                value: targetObj.muted ? "off" : "all"
            });
            if (targetId !== "ch_kampane" && targetId !== "chat_hrs") {
                this.callbacks.onMistake(`Ztlumil(a) jsi '${targetObj.name}', zkontroluj v zadání nahoře, který konkrétní kanál či chat máš ztlumit!`);
            }
        } else if (actionName === "mark_unread") {
            targetObj.unread = true;
            this.render();
            this.callbacks.onInfo(`Položka '${targetObj.name}' byla označena jako nepřečtená.`);
        } else if (actionName === "move_section") {
            this.render();
            this.callbacks.onAction({ action: "organize_section", targetId: targetId });
            this.callbacks.onInfo(`Položka '${targetObj.name}' byla přesunuta do vybraného oddílu.`);
        } else {
            this.render();
            this.callbacks.onMistake(`Zvolil(a) jsi '${actionName}' u '${targetObj.name}', ale tato akce nesplňuje cíl aktuální mise.`);
        }
    },

    renderChatSection(id, label) {
        const s = this.state;
        let list = s.chats.filter(c => c.section === id);
        if (s.chatFilter === "unread") list = list.filter(c => c.unread);
        if (!list.length) return "";
        const collapsed = s.collapsedSections[id];
        return `
            <div class="tm-section-label" data-toggle-section="${id}" style="margin-top:8px;">${collapsed ? "▸" : "▾"} ${label}</div>
            ${collapsed ? "" : list.map(c => `
                <div class="tm-tree-item tm-chat-item ${s.selectedItemType === "chat" && s.selectedItemId === c.id ? "active" : ""} ${c.unread ? "unread" : ""}"
                     data-select-chat="${c.id}" data-ctx-type="chat" data-ctx-id="${c.id}">
                    <div class="tm-item-left">
                        <span class="tm-mini-avatar">${c.initials}</span>
                        <span>${c.name} ${c.muted ? "🔇" : ""}</span>
                    </div>
                    <button class="tm-hover-dots" data-dots-type="chat" data-dots-id="${c.id}">•••</button>
                </div>
            `).join("")}
        `;
    },

    renderSavedList() {
        const saved = this.state.savedMessages || [];
        if (!saved.length) return `<div class="tm-empty-tab">🔖 Zatím nemáš uložené žádné zprávy.</div>`;
        return `<div class="tm-saved-list">${saved.map((m, i) => `
            <div class="tm-saved-item">
                <div class="tm-saved-meta">${m.author} · ${m.time} · ${m.chatName}</div>
                <div>${m.text}</div>
                <button class="tm-link-btn" data-unsave="${i}">Odebrat z uložených</button>
            </div>`).join("")}</div>`;
    },

    renderPost(team, p, pi) {
        const key = `${team.id}|${pi}`;
        const reply = this.state.threadKey === key;
        return `
            <div class="tm-channel-post">
                <div class="tm-post-head">
                    <span class="tm-mini-avatar">${p.author.split(" ").map(w => w[0]).join("").slice(0, 2)}</span>
                    <strong>${p.author}</strong> <span class="tm-muted">${p.time}</span>
                </div>
                <div class="tm-post-title">${p.title}</div>
                <div class="tm-post-body">${p.body}</div>
                ${p.videoTitle ? `<div class="tm-video-placeholder" id="tm-play-video">▶ ${p.videoTitle}</div>` : ""}
                <div class="tm-post-reply-bar" data-thread-key="${key}">↩ Odpovědět ve vlákně</div>
            </div>`;
    },

    renderChatMessage(chatId, m, mi) {
        const s = this.state;
        const key = `${chatId}|${mi}`;
        const editing = s.editKey === key;
        const reacting = s.reactKey === key;
        const reactions = (m.reactions || []).map(r => `<span class="tm-reaction-pill">${r.emoji} ${r.count}</span>`).join("");
        return `
            <div class="tm-chat-row ${m.mine ? "mine" : ""}" data-ctx-type="msg" data-ctx-id="${key}">
                <div class="tm-msg-meta">${m.author} · ${m.time}${m.saved ? " · 🔖" : ""}${m.pinned ? " · 📌" : ""}</div>
                ${editing ? `
                    <div class="tm-edit-box">
                        <input id="tm-edit-inp" value="${m.text}" />
                        <button class="tm-btn-purple" data-edit-save="${key}">Uložit</button>
                        <button class="tm-link-btn" data-edit-cancel="1">Zrušit</button>
                    </div>` : `
                    <div class="tm-chat-bubble">${m.text}</div>`}
                ${reactions ? `<div class="tm-reactions">${reactions}</div>` : ""}
                ${reacting ? `
                    <div class="tm-react-bar">
                        ${["👍", "❤️", "😂", "😮"].map(e => `<button class="tm-react-emoji" data-react="${key}" data-emoji="${e}">${e}</button>`).join("")}
                    </div>` : ""}
                <div class="tm-msg-actions">
                    <button class="tm-link-btn" data-react-open="${key}">😊 Reagovat</button>
                    <button class="tm-link-btn" data-reply="${key}">↩ Odpovědět</button>
                    ${m.mine ? `<button class="tm-link-btn" data-edit-open="${key}">✏️ Upravit</button>` : ""}
                    <button class="tm-link-btn" data-save-msg="${key}">🔖 Uložit</button>
                    <button class="tm-link-btn" data-pin-msg="${key}">📌 Připnout</button>
                    <button class="tm-hover-dots tm-msg-dots" data-dots-type="msg" data-dots-id="${key}">•••</button>
                </div>
            </div>`;
    },

    renderThreadPanel(activeObj) {
        const key = this.state.threadKey;
        const replies = (this.state.threads && this.state.threads[key]) || [];
        return `
            <div class="tm-thread-panel">
                <div class="tm-thread-head">
                    <strong>Vlákno</strong>
                    <button class="tm-link-btn" data-thread-close="1">✕</button>
                </div>
                <div class="tm-thread-list">
                    ${replies.map(r => `<div class="tm-thread-reply"><strong>${r.author}</strong>: ${r.text}</div>`).join("") || `<div class="tm-muted">Zatím žádné odpovědi.</div>`}
                </div>
                <div class="tm-thread-compose">
                    <input id="tm-thread-inp" placeholder="Odpovědět ve vlákně…" />
                    <button class="tm-btn-purple" id="tm-thread-send">Odeslat</button>
                </div>
            </div>`;
    },

    renderSharedTab() {
        const files = this.state.files || [];
        return `
            <div class="tm-shared-list">
                <div class="tm-shared-head"><strong>Sdílené soubory</strong>
                    <button class="tm-btn-purple" id="tm-share-file">📤 Sdílet soubor</button>
                </div>
                ${files.map(f => `<div class="tm-shared-item">📄 ${f.name} <span class="tm-muted">${f.owner || ""}</span></div>`).join("") || `<div class="tm-muted">Žádné sdílené soubory.</div>`}
            </div>`;
    },

    renderExtraOverlays() {
        const s = this.state;
        let html = "";
        if (s.searchOpen) {
            html += `
                <div class="tm-overlay-panel tm-search-overlay">
                    <input id="tm-search-inp" placeholder="Hledat v Teams…" value="${s.searchQuery || ""}" />
                    <div class="tm-search-results">${this.renderSearchResults()}</div>
                    <button class="tm-link-btn" data-close-overlay="search">Zavřít</button>
                </div>`;
        }
        if (s.membersTeamId) {
            const team = s.teamsList.find(t => t.id === s.membersTeamId);
            const members = (s.teamMembers && s.teamMembers[s.membersTeamId]) || [];
            html += `
                <div class="tm-overlay-panel tm-members-panel">
                    <div class="tm-panel-head"><strong>Členové: ${team ? team.name : ""}</strong>
                        <button class="tm-link-btn" data-close-overlay="members">✕</button></div>
                    ${members.map((m, i) => `
                        <div class="tm-member-row"><span>${m.name} <span class="tm-muted">${m.role}</span></span>
                            <button class="tm-link-btn" data-member-role="${i}">Změnit roli</button></div>`).join("")}
                </div>`;
        }
        if (s.peoplePanelOpen) {
            html += `
                <div class="tm-overlay-panel tm-people-panel">
                    <div class="tm-panel-head"><strong>Účastníci (${s.participants ? s.participants.length : 0})</strong>
                        <button class="tm-link-btn" data-close-overlay="people">✕</button></div>
                    ${(s.participants || []).map(p => `<div class="tm-member-row">${p}</div>`).join("")}
                </div>`;
        }
        if (s.breakoutPanelOpen) {
            html += `
                <div class="tm-overlay-panel tm-breakout-panel">
                    <div class="tm-panel-head"><strong>Oddělené místnosti</strong>
                        <button class="tm-link-btn" data-close-overlay="breakout">✕</button></div>
                    ${(s.breakoutRooms || []).map((r, i) => `<div class="tm-member-row">${r.name} <button class="tm-link-btn" data-breakout-assign="${i}">Přiřadit účastníky</button></div>`).join("")}
                </div>`;
        }
        if (s.attachOpen) {
            html += `
                <div class="tm-overlay-panel tm-attach-panel">
                    <div class="tm-panel-head"><strong>Přiložit</strong>
                        <button class="tm-link-btn" data-close-overlay="attach">✕</button></div>
                    <button class="tm-link-btn" data-attach-file="1">📄 Soubor z OneDrive</button>
                </div>`;
        }
        return html;
    },

    renderSearchResults() {
        const q = (this.state.searchQuery || "").toLowerCase();
        if (!q) return `<div class="tm-muted">Zadej hledaný výraz.</div>`;
        const hits = [];
        this.state.chats.forEach(c => c.messages.forEach(m => {
            if (m.text.toLowerCase().includes(q)) hits.push(`💬 ${c.name}: ${m.text}`);
        }));
        this.state.teamsList.forEach(t => t.channels.forEach(ch => ch.posts.forEach(p => {
            if ((p.title + p.body).toLowerCase().includes(q)) hits.push(`📢 ${t.name} / ${ch.name}: ${p.title}`);
        })));
        return hits.length ? hits.map(h => `<div class="tm-search-hit">${h}</div>`).join("") : `<div class="tm-muted">Nic nenalezeno.</div>`;
    },

    sendMessage() {
        const inp = this.container.querySelector("#tm-msg-input");
        if (!inp || !inp.value.trim()) return;
        const txt = inp.value.trim();

        if (this.state.selectedItemType === "chat") {
            const ch = this.state.chats.find(c => c.id === this.state.selectedItemId);
            ch.messages.push({ author: "Já", time: "Právě teď", text: txt, mine: true });
        } else {
            for (const tm of this.state.teamsList) {
                const found = tm.channels.find(c => c.id === this.state.selectedItemId);
                if (found) {
                    found.posts.push({ author: "Já", time: "Právě teď", title: "Nový příspěvek", body: txt });
                    break;
                }
            }
        }
        this.render();
    }
};