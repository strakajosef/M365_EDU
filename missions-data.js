/**
 * ============================================================================
 * DATOVÁ VRSTVA APLIKACE (missions-data.js)
 * - taskDescription: Podrobné zadání úkolu a kontext situace (zobrazuje se nahoře v liště)
 * - hint: Konkrétní návod jak úkol udělat (zobrazí se ve Volném tréninku po najetí myší na moderátora)
 * ============================================================================
 */

const GAME_DATA = {
    // 1. Konfigurace soutěže
    config: {
        maxQuestionsPerGame: 5,
        difficulties: {
            0: { name: "Základní", points: 100 },
            1: { name: "Pokročilý", points: 200 },
            2: { name: "Specialista", points: 300 },
            3: { name: "Expert", points: 400 }
        }
    },

    // 2. Aktivní kategorie na herním plánu Riskuj
    categories: [
        { id: "copilot", name: "Copilot", color: "#6264A7", iconPlaceholder: "https://placehold.co/80x80/6264A7/ffffff?text=AI" },
        { id: "onedrive", name: "OneDrive & SP", color: "#0078D4", iconPlaceholder: "https://placehold.co/80x80/0078D4/ffffff?text=OD" },
        { id: "teams", name: "Teams", color: "#464EB8", iconPlaceholder: "https://placehold.co/80x80/464EB8/ffffff?text=TM" },
        { id: "outlook", name: "Outlook", color: "#005A9E", iconPlaceholder: "https://placehold.co/80x80/005A9E/ffffff?text=OL" },
        { id: "onenote", name: "OneNote", color: "#7719AA", iconPlaceholder: "https://placehold.co/80x80/7719AA/ffffff?text=ON" },
        { id: "forms", name: "Forms", color: "#0f766e", iconPlaceholder: "https://placehold.co/80x80/0f766e/ffffff?text=FM" },
        { id: "todo", name: "To Do", color: "#2564CF", iconPlaceholder: "https://placehold.co/80x80/2564CF/ffffff?text=TD" },
        { id: "planner", name: "Planner", color: "#31752F", iconPlaceholder: "https://placehold.co/80x80/31752F/ffffff?text=PL" }
    ],

    // 3. Kompletní katalog znalostí
    curriculumCatalog: [
        { app: "Clipchamp", level: 0, skill: "Vyhledat a spustit video" },
        { app: "Clipchamp", level: 2, skill: "Nahrání obrazovky či záznam videa" },
        { app: "Copilot", level: 0, skill: "Shrnutí textu, generování textu, šablony a knihovny promptů" },
        { app: "Copilot", level: 1, skill: "Tvorba agentů, dohledání návodu k IT chybě, analýza dat" },
        { app: "Copilot", level: 2, skill: "Vytváření jednoduchých HTML aplikací" },
        { app: "Forms", level: 0, skill: "Vyplnění dotazníku, který dostanu" },
        { app: "Forms", level: 2, skill: "Vytvoření otázek v dotazníku, administrace a nastavení" },
        { app: "Forms", level: 3, skill: "Automatizace aktivity spojených s dotazníkem - rozesílka, sběr, vyhodnocení" },
        { app: "Loop", level: 0, skill: "Loop stránka - základní využití a vyhledávání (workspace/stránka)" },
        { app: "Loop", level: 1, skill: "Loop komponenta v Teams, oprávnění, šablony, spolupráce v reálném čase" },
        { app: "Lists", level: 2, skill: "Vytvoření seznamu, oprávnění, upozornění, šablony, filtry, verzování" },
        { app: "Lists", level: 3, skill: "Automatizace aktivity spojených se seznamem" },
        { app: "OneDrive", level: 0, skill: "Uložení souboru, vyhledávání, nic na lokálním disku C:, sdílení odkazů bez duplicit" },
        { app: "OneDrive", level: 1, skill: "Work/school knihovny, Verzování" },
        { app: "OneNote", level: 0, skill: "Zápisky, stránky a sekce, formátování, vyhledávání poznámek" },
        { app: "OneNote", level: 1, skill: "Struktura bloků, organizace úkolů, vkládání souborů" },
        { app: "OneNote", level: 2, skill: "Sdílení bloku, šablony meetingů, propojení s Outlookem a Teams" },
        { app: "Outlook", level: 0, skill: "Přikládání souborů, kalendář a místnosti, třídění, automatické odpovědi, podpis, šablony" },
        { app: "Outlook", level: 1, skill: "Pravidla (Rules), View Templates, úkol z e-mailu, Diktování" },
        { app: "Outlook", level: 2, skill: "Pokročilé filtry, Bookings, správa a sdílení kalendářů" },
        { app: "Outlook", level: 3, skill: "Archivace a údržba" },
        { app: "To Do", level: 1, skill: "Vytvoření úkolu, připomenutí, Můj den, sdílené seznamy, přílohy, úkol z Teams zprávy nebo e-mailu" },
        { app: "Planner", level: 2, skill: "Vytvoření úkolu, připomenutí, opakované úkoly, sdílení plánu, přílohy, úkol z Teams zprávy" },
        { app: "Power Automate", level: 3, skill: "Znalost a aktivní používání aplikace Power Automate" },
        { app: "SharePoint", level: 0, skill: "Uložení souboru, dohledání dokumentů, sdílení linků a oprávnění" },
        { app: "SharePoint", level: 2, skill: "Webová část SharePoint - vytváření článků a odkazů" },
        { app: "SharePoint", level: 3, skill: "Automatizace aktivity spojených se soubory" },
        { app: "Teams", level: 0, skill: "Chat (vlákna, reakce), vyhledávání, týmy a kanály, oddíly, připnutí zpráv, správa oznámení, nahrání schůzky" },
        { app: "Teams", level: 1, skill: "PowerPoint Live, aktivní řízení online schůzky, záznam schůzky" },
        { app: "Teams", level: 2, skill: "Aplikace v Teams, Breakout rooms, správa týmů a členů" },
        { app: "Teams", level: 3, skill: "Moderování a nastavení velkých schůzek" }
    ],

    // 4. Hlášky moderátora
    moderatorQuotes: {
        menu: [
            "Vítej v M365 EduAréně! Vyber si svou úroveň a spusť soutěž Riskuj, nebo trénuj v Sandboxu.",
            "Jsem připraven! Zvol si hru a ukaž, jak ovládáš aplikace Microsoft 365."
        ],
        gameStart: [
            "Soutěž RISKUJ začíná! Vyber si kategorii a bodovou hodnotu. Praktické úkoly jsou na čas!",
            "Klikni na libovolné pole na herním plánu a vyřeš praktický scénář přímo v simulátoru aplikace."
        ],
        freeTraining: [
            "Jsi ve volném tréninku! Přečti si podrobné zadání nahoře. Když si nebudeš vědět rady, najeď na mě myší a ukážu ti návod!",
            "Tréninkový režim bez časového limitu! Zkus úkol splnit podle zadání, nebo na mě najeď myší pro nápovědu."
        ],
        beforeSandbox: [
            "Praktická mise v otevřeném Sandboxu! Přečti si pozorně zadání nahoře, časovač právě běží!",
            "Teď žádná teorie – ukaž mi přímo v aplikaci, jak bys tuto situaci vyřešil(a)!"
        ],
        sandboxMistake: [
            "Pozor, tahle akce k cíli nevede! Ztrácíš jeden pokus, podívej se znovu na zadání.",
            "Tady nejsi správně. Přemýšlej, kde se tato funkce v aplikaci nachází."
        ],
        correctAnswer: [
            "Skvělá práce! Úkol jsi zvládl(a) na jedničku.",
            "Přesně tak! Jde vidět, že v M365 nejsi žádný nováček."
        ],
        wrongAnswer: [
            "Bohužel, tentokrát to nevyšlo a body se odečítají. Příště už budeš vědět!",
            "Čas vypršel nebo došly pokusy! Tohle téma si pak vyzkoušej ve Volném tréninku."
        ],
        gameOver: [
            "Soutěž je u konce! Mrkni na své skóre a doporučení, na co se zaměřit."
        ]
    },

    // 5. ÚKOLY PRO HERNÍ PLÁN A VOLNÝ TRÉNINK
    missions: [
        // ================= SLOUPEC 1: MICROSOFT COPILOT =================
        {
            id: "m_copilot_0",
            categoryId: "copilot",
            difficulty: 0,
            type: "sandbox",
            appModule: "copilot",
            topic: "Shrnutí textu a knihovna promptů",
            taskDescription: "Představ si, že jsi dostal(a) dlouhý zápis z celofiremní porady a potřebuješ z něj během chvilky vytáhnout jen ty nejdůležitější body, případně si otevřít dříve uložený výstup. Požádej Copilota o rychlé shrnutí textu nebo otevři uložený materiál.",
            hint: "V rozhraní Copilot klikni pod chatovacím polem na rychlou pilulku 'Shrnout' (nebo napiš požadavek do chatu a odešli ho), případně v levém panelu otevři '📚 Knihovna' a klikni na uložený dokument.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: ["copilot_summarize", "copilot_library"], completed: true }
        },
        {
            id: "m_copilot_1",
            categoryId: "copilot",
            difficulty: 1,
            type: "sandbox",
            appModule: "copilot",
            topic: "Tvorba agentů a dohledání návodu k IT chybě",
            taskDescription: "Řešíš specifický firemní úkol – například audit digitálních dovedností, orientaci v lékovém katalogu nebo opakující se chybu v kódu – a běžný obecný chat ti nestačí. Aktivuj v Copilotu specializovaného firemního Agenta nebo založ nového.",
            hint: "V levém navigačním panelu klikni na položku '🤖 Agenti ▾' a z rozbalovací nabídky vyber některého ze specializovaných asistentů (např. 'Audit Digi dovedností') nebo klikni na '➕ Nový agent'.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: ["copilot_agent", "copilot_action"], completed: true }
        },
        {
            id: "m_copilot_2",
            categoryId: "copilot",
            difficulty: 2,
            type: "sandbox",
            appModule: "copilot",
            topic: "Připojení obsahu, analýza dat a tvorba HTML aplikací",
            taskDescription: "Máš připravený externí soubor s daty na disku nebo v cloudu a potřebuješ, aby z něj Copilot udělal hloubkovou analýzu nebo vygeneroval šablonu. Připoj externí soubor přímo do konverzace s Copilotem.",
            hint: "V chatovacím poli klikni vlevo na kulaté tlačítko '+' (Přidat obsah) a z nabídky vyber '⬆️ Nahrát obrázky a soubory' nebo '☁️ Připojit cloudové soubory'.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: ["copilot_upload", "copilot_action"], completed: true }
        },
        {
            id: "m_copilot_3",
            categoryId: "copilot",
            difficulty: 3,
            type: "sandbox",
            appModule: "copilot",
            topic: "Práce s výstupy a šablonami v Knihovně Copilot",
            taskDescription: "V minulém týdnu jsi s pomocí Copilota vytvořil(a) HTML šablonátor e-mailů a studijní přehledy pro tým. Nyní se k těmto vygenerovaným materiálům potřebuješ vrátit a otevřít je bez nového psaní promptů.",
            hint: "Otevři v levém navigačním panelu záložku '📚 Knihovna' a klikni na libovolnou kartu uloženého dokumentu (např. 'HTML šablonátor e-mailů').",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "copilot_library", completed: true }
        },
        {
            id: "m_copilot_4",
            categoryId: "copilot",
            difficulty: 0,
            type: "sandbox",
            appModule: "copilot",
            topic: "Generování textu",
            taskDescription: "Potřebuješ rychle připravit text pozvánky na týmovou schůzku. Nech Copilot vytvořit koncept obsahu.",
            hint: "Pod chatovacím polem klikni na pilulku 'Koncept obsahu'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "copilot_generate", completed: true }
        },
        {
            id: "m_copilot_5",
            categoryId: "copilot",
            difficulty: 0,
            type: "sandbox",
            appModule: "copilot",
            topic: "Možnosti využití Copilota – šablony a knihovny promptů",
            taskDescription: "Chceš zjistit, jaké hotové šablony promptů má Copilot k dispozici, abys je nemusel(a) psát od nuly. Otevři knihovnu promptů.",
            hint: "V levém panelu nebo v nabídce Copilota otevři 'Knihovnu' a podívej se na šablony promptů.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "copilot_prompt_library", completed: true }
        },
        {
            id: "m_copilot_6",
            categoryId: "copilot",
            difficulty: 0,
            type: "sandbox",
            appModule: "copilot",
            topic: "Formulace zadání (promptů)",
            taskDescription: "Zadání typu 'napiš něco o projektu' dává špatné výsledky. Napiš konkrétní zadání pro Copilot s cílem, kontextem a požadovaným formátem (alespoň 60 znaků) a odešli ho.",
            hint: "Uveď, co chceš získat, pro koho je text, jaký má být formát a délka. Odeslat můžeš tlačítkem nebo klávesou Enter.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "copilot_prompt_craft", completed: true }
        },
        {
            id: "m_copilot_7",
            categoryId: "copilot",
            difficulty: 1,
            type: "sandbox",
            appModule: "copilot",
            topic: "Analýza dat",
            taskDescription: "Máš tabulku s prodeji za poslední čtvrtletí a potřebuješ rychle zjistit trendy. Požádej Copilota o analýzu dat.",
            hint: "Pod chatovacím polem klikni na pilulku 'Analyzovat data'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "copilot_data_analysis", completed: true }
        },
        {
            id: "m_copilot_8",
            categoryId: "copilot",
            difficulty: 1,
            type: "sandbox",
            appModule: "copilot",
            topic: "Dohledání návodu k IT chybě",
            taskDescription: "Na notebooku se ti zobrazila chybová hláška, kterou neznáš. Zjisti pomocí Copilota postup řešení.",
            hint: "Pod chatovacím polem klikni na pilulku 'Najít návod k chybě'. Další pilulky najdeš přes tlačítko '•••'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "copilot_it_error", completed: true }
        },
        {
            id: "m_copilot_9",
            categoryId: "copilot",
            difficulty: 1,
            type: "sandbox",
            appModule: "copilot",
            topic: "Tvorba agentů",
            taskDescription: "Chceš si vytvořit vlastního agenta, který bude odpovídat na dotazy k interním postupům tvého týmu. Založ nového agenta.",
            hint: "V levém panelu klikni na 'Agenti' a vyber 'Nový agent'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "copilot_agent_create", completed: true }
        },
        {
            id: "m_copilot_10",
            categoryId: "copilot",
            difficulty: 2,
            type: "sandbox",
            appModule: "copilot",
            topic: "Vytváření jednoduchých HTML aplikací",
            taskDescription: "Potřebuješ jednoduchou webovou stránku s formulářem pro evidenci návštěv. Požádej Copilota o vytvoření HTML aplikace.",
            hint: "Pod chatovacím polem klikni na pilulku 'Vytvořit HTML aplikaci', nebo do zadání napiš slovo 'HTML'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "copilot_html_app", completed: true }
        },

        // ================= SLOUPEC 2: MICROSOFT ONEDRIVE & SHAREPOINT =================
        {
            id: "m_onedrive_0",
            categoryId: "onedrive",
            difficulty: 0,
            type: "sandbox",
            appModule: "onedrive",
            topic: "OneDrive 0-Základní: Uložení souboru",
            taskDescription: "Na svém počítači jsi dokončil(a) důležitou tabulku s rozpočtem projektu. Aby o ni firma nepřišla při poruše počítače (nic neukládáme pouze lokálně na disk C:), ulož tento pracovní dokument bezpečně do svého cloudového úložiště OneDrive.",
            hint: "Na ploše Windows 11 si otevři OneDrive (přes Teams, Outlook nebo Chrome). Poté vlevo nahoře klikni na modré tlačítko '+ Vytvořit nebo nahrát' -> 'Nahrát soubory' a vyber dokument.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "onedrive_save_file", completed: true }
        },
        {
            id: "m_onedrive_1",
            categoryId: "onedrive",
            difficulty: 0,
            type: "sandbox",
            appModule: "onedrive",
            topic: "OneDrive 0-Základní: Vyhledávání dokumentů na OneDrive",
            taskDescription: "Ve svém cloudovém úložišti máš spoustu dokumentů a potřebuješ během porady rychle dohledat konkrétní tabulku nebo PDF dokument. Použij vyhledávání nebo rychlé filtrování souborů v OneDrive.",
            hint: "Otevři OneDrive a vyhledej dokument: použij horní vyhledávací pole (a stiskni Enter), filtrovací pilulky typu souboru (Word, Excel, PDF) nebo pole 'Filtrovat podle jména'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "onedrive_search", completed: true }
        },
        {
            id: "m_onedrive_2",
            categoryId: "onedrive",
            difficulty: 0,
            type: "sandbox",
            appModule: "onedrive",
            topic: "OneDrive 0-Základní: Sdílení odkazů a oprávnění (nevytvářím duplicity)",
            taskDescription: "Potřebuješ s kolegyní Pavlou spolupracovat na jednom dokumentu. Místo posílání souboru v příloze e-mailu (čímž by vznikaly duplicitní verze) jí přímo z OneDrive nasdílej odkaz s nastaveným oprávněním.",
            hint: "V otevřeném OneDrive klikni pravým tlačítkem myši (nebo přes ikonu 📤 / •••) na libovolný soubor, vyber 'Sdílet' nebo 'Kopírovat odkaz' a v dialogu potvrď 'Poslat odkaz' či 'Kopírovat odkaz'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "onedrive_share_link", completed: true }
        },
        {
            id: "m_onedrive_3",
            categoryId: "onedrive",
            difficulty: 1,
            type: "sandbox",
            appModule: "onedrive",
            topic: "OneDrive 1-Pokročilý: Work/school knihovny",
            taskDescription: "Hledáš dokumenty, které nejsou uložené u tebe v soukromých souborech, ale patří celému projektovému týmu či oddělení. Přejdi v rozhraní OneDrive do sdílených firemních knihoven dokumentů.",
            hint: "V levém panelu OneDrive klikni na sekci '📚 Knihovny' nebo v části 'Rychlý přístup' dole otevři některou ze sdílených týmových knihoven (např. DIGI HRS nebo Česko-Slovensko SPOLU).",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "onedrive_libraries", completed: true }
        },
        {
            id: "m_onedrive_4",
            categoryId: "onedrive",
            difficulty: 1,
            type: "sandbox",
            appModule: "onedrive",
            topic: "OneDrive 1-Pokročilý: Verzování",
            taskDescription: "Kolega omylem přepsal důležitá data ve sdílené tabulce a změny uložil. Naštěstí je soubor na OneDrive – podívej se do přehledu předchozích verzí souboru a obnov starší funkční verzi.",
            hint: "Klikni pravým tlačítkem myši (nebo na ••• či v horní liště) na soubor v OneDrive, vyber '🕒 Historie verzí' a v otevřeném okně klikni na '↺ Obnovit verzi'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "onedrive_versioning", completed: true }
        },
        {
            id: "m_onedrive_5",
            categoryId: "onedrive",
            difficulty: 0,
            type: "sandbox",
            appModule: "onedrive",
            topic: "OneDrive 0-Základní: Uložení dokumentu do OneDrive (cloud)",
            taskDescription: "Dokument Plan_obce_2026.docx máš uložený lokálně v C:\\Dokumenty. Ulož ho do OneDrive tak, aby se ukládal do cloudu a nebyl pouze na disku.",
            hint: "1) Vyber soubor Plan_obce_2026.docx v seznamu. 2) V příkazové liště klikni na 'Uložit do OneDrive (cloud)'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "od_cloud_only", completed: true }
        },

        // ================= SLOUPEC 3: MICROSOFT TEAMS =================
        {
            id: "m_teams_0",
            categoryId: "teams",
            difficulty: 0,
            type: "sandbox",
            appModule: "teams",
            topic: "Správa oznámení a orientace v týmech a kanálech",
            taskDescription: "Potřebuješ se soustředit na práci, ale v jednom z týmových kanálů věnovaných marketingovým kampaním neustále vyskakují rušivé notifikace. Najdi tento kanál v seznamu týmů a ztlum jeho oznámení.",
            hint: "V levém panelu pod týmem najdi kanál s hlučnými notifikacemi (Kampaně), klikni na něj pravým tlačítkem myši (nebo na •••) a zvol možnost pro vypnutí/ztlumení oznámení.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: {
                action: "mute_item",
                targetId: "ch_kampane",
                value: "off"
            }
        },
        {
            id: "m_teams_1",
            categoryId: "teams",
            difficulty: 1,
            type: "sandbox",
            appModule: "teams",
            topic: "PowerPoint Live – interaktivní prezentace v online schůzce",
            taskDescription: "Vedeš online schůzku a potřebuješ účastníkům promítnout prezentaci tak, abys stále viděl(a) chat i své poznámky přednášejícího a kolegové si mohli sami překlikávat snímky. Spusť schůzku a nasdílej prezentaci přes PowerPoint Live.",
            hint: "Klikni vpravo nahoře na '📹 Okamžitá schůzka' -> 'Zahájit schůzku' -> 'Připojit se'. V horní liště hovoru klikni na '⬆️ Sdílet' a v sekci PowerPoint Live vyber prezentaci.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "share_ppt_live",
                status: "started"
            }
        },
        {
            id: "m_teams_2",
            categoryId: "teams",
            difficulty: 2,
            type: "sandbox",
            appModule: "teams",
            topic: "Plánování schůzky a kontrola dostupnosti účastníků",
            taskDescription: "Chceš svolat týmovou poradu v Teams, ale potřebuješ mít jistotu, že pozvaní kolegové nemají v danou dobu jiný hovor. Otevři kalendář v Teams, zkontroluj dostupnost účastníků v plánovači a ulož schůzku na volný čas.",
            hint: "V levé liště přepni na '📅 Kalendář', klikni na '➕ Nová událost', nahoře v okně přepni na '📊 Plánovač (Dostupnost)', vyber volný časový slot '11:00 – 11:30' a klikni na '💾 Poslat / Uložit'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "schedule_meeting_free_slot",
                slot: "11:00"
            }
        },
        {
            id: "m_teams_3",
            categoryId: "teams",
            difficulty: 3,
            type: "sandbox",
            appModule: "teams",
            topic: "Moderování velkých schůzek – Záznam a Skupinové místnosti",
            taskDescription: "Moderuješ velký online workshop v Teams. Potřebuješ buď spustit nahrávání a přepis schůzky pro nepřítomné kolegy, anebo rozdělit účastníky do menších diskuzních skupin (Breakout Rooms).",
            hint: "Spusť '📹 Okamžitá schůzka' -> 'Zahájit schůzku' -> 'Připojit se'. V horní liště běžící schůzky klikni na '••• Více' (nebo ikonu místností) a zvol '⏺️ Záznam', '📝 Přepsat' nebo '🚪 Místnosti'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "meeting_feature",
                feature: ["record", "breakout"]
            }
        },
        {
            id: "m_teams_4",
            categoryId: "teams",
            difficulty: 0,
            type: "sandbox",
            appModule: "teams",
            topic: "Vlákna a reakce v chatu",
            taskDescription: "Kolega ve skupinovém chatu napsal dotaz. Otevři vlákno pod jeho zprávou a odpověz v něm.",
            hint: "Najdi zprávu v chatu a klikni na ikonu odpovědi nebo použij pravé tlačítko myši na zprávu.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "thread_reply"
            }
        },
        {
            id: "m_teams_5",
            categoryId: "teams",
            difficulty: 0,
            type: "sandbox",
            appModule: "teams",
            topic: "Úprava vlastní zprávy",
            taskDescription: "Při psaní zprávy v chatu jsi udělal(a) překlep. Oprav svou vlastní zprávu.",
            hint: "Na své zprávě otevři menu (••• nebo pravé tlačítko) a zvol Upravit.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "edit_message"
            }
        },
        {
            id: "m_teams_6",
            categoryId: "teams",
            difficulty: 0,
            type: "sandbox",
            appModule: "teams",
            topic: "Vyhledávání v Teams",
            taskDescription: "Potřebuješ najít starší zprávu o rozpočtu. Použij vyhledávání v Teams a zadej výraz 'rozpočet'.",
            hint: "Vyhledávací pole je nahoře uprostřed. Zadej výraz a potvrď Enterem.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "search_teams"
            }
        },
        {
            id: "m_teams_7",
            categoryId: "teams",
            difficulty: 0,
            type: "sandbox",
            appModule: "teams",
            topic: "Připnutí zprávy v kanálu",
            taskDescription: "V kanálu Projekty je důležitá zpráva s termínem odevzdání. Připni ji, aby byla vidět všem členům kanálu.",
            hint: "Otevři menu zprávy (••• nebo pravé tlačítko) a zvol Připnout.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "pin_message"
            }
        },
        {
            id: "m_teams_8",
            categoryId: "teams",
            difficulty: 0,
            type: "sandbox",
            appModule: "teams",
            topic: "Sdílení souboru v Teams",
            taskDescription: "Kolegyni potřebuješ poslat soubor z OneDrivu přímo do chatu. Sdílej dokument z OneDrivu.",
            hint: "V poli pro psaní zprávy klikni na ikonu přílohy a zvol soubor z OneDrivu.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "share_file",
                source: "onedrive"
            }
        },
        {
            id: "m_teams_9",
            categoryId: "teams",
            difficulty: 0,
            type: "sandbox",
            appModule: "teams",
            topic: "Správa oznámení týmu",
            taskDescription: "Chceš upravit, jaká oznámení ti tým posílá. Otevři nastavení oznámení týmu.",
            hint: "Klikni pravým tlačítkem na tým v levém panelu nebo použij ikonu zvonku u týmu.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "notif_settings"
            }
        },
        {
            id: "m_teams_10",
            categoryId: "teams",
            difficulty: 0,
            type: "sandbox",
            appModule: "teams",
            topic: "Online schůzka – kamera, mikrofon a ruka",
            taskDescription: "Jsi v online schůzce a chceš položit dotaz. Zvedni ruku, aby tě moderátor viděl.",
            hint: "Ovládací prvky schůzky jsou dole uprostřed. Hledej ikonu ruky.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "raise_hand",
                value: true
            }
        },
        {
            id: "m_teams_11",
            categoryId: "teams",
            difficulty: 1,
            type: "sandbox",
            appModule: "teams",
            topic: "Záznam schůzky",
            taskDescription: "Pro kolegy, kteří se schůzky nemohli zúčastnit, spusť záznam probíhající schůzky.",
            hint: "V horní liště schůzky otevři menu '••• Více' a zvol Záznam.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "meeting_feature",
                feature: "record"
            }
        },
        {
            id: "m_teams_12",
            categoryId: "teams",
            difficulty: 2,
            type: "sandbox",
            appModule: "teams",
            topic: "Skupinové místnosti (Breakout rooms)",
            taskDescription: "Workshop se má rozdělit do menších skupin. Otevři Breakout rooms a přiřaď účastníky do místností.",
            hint: "Ve schůzce otevři panel Breakout rooms (ikona místností) a použij možnost přiřadit účastníky.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "meeting_feature",
                feature: "breakout_assign"
            }
        },
        {
            id: "m_teams_13",
            categoryId: "teams",
            difficulty: 2,
            type: "sandbox",
            appModule: "teams",
            topic: "Správa členů týmu",
            taskDescription: "Novému kolegovi v týmu Projekty potřebuješ změnit roli na vlastníka. Otevři správu členů a změň roli.",
            hint: "Klikni na tým pravým tlačítkem (nebo •••) a zvol Spravovat členy, pak změň roli člena.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "manage_members",
                change: "role"
            }
        },
        {
            id: "m_teams_14",
            categoryId: "teams",
            difficulty: 3,
            type: "sandbox",
            appModule: "teams",
            topic: "Moderování velké schůzky – seznam účastníků",
            taskDescription: "Moderuješ velkou schůzku. Otevři panel Účastníci, abys měl(a) přehled, kdo je připojen.",
            hint: "V horní liště schůzky klikni na ikonu Účastníci (lidé).",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "meeting_feature",
                feature: "people"
            }
        },

        // ================= SLOUPEC 4: MICROSOFT OUTLOOK =================
        {
            id: "m_outlook_0",
            categoryId: "outlook",
            difficulty: 0,
            type: "sandbox",
            appModule: "outlook",
            topic: "Automatické odpovědi (Mimo kancelář) a nastavení podpisu",
            taskDescription: "Chystáš se na týdenní dovolenou a potřebuješ, aby každý, kdo ti v té době napíše e-mail, automaticky obdržel informaci o tvé nepřítomnosti a kontakt na zástup (případně potřebuješ upravit svůj e-mailový podpis).",
            hint: "Klikni vlevo nahoře na kartu 'Soubor' a zvol 'Automatické odpovědi'. V okně přepni na 'Posílat automatické odpovědi' a potvrď tlačítkem 'OK' (nebo jdi přes Soubor -> Možnosti -> Pošta -> Podpisy).",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: {
                action: "outlook_autoreply",
                status: "enabled"
            }
        },
        {
            id: "m_outlook_1",
            categoryId: "outlook",
            difficulty: 1,
            type: "sandbox",
            appModule: "outlook",
            topic: "Pravidla (Rules), kategorie a vytvoření úkolu z e-mailu",
            taskDescription: "Do doručené pošty ti chodí desítky zpráv denně. Aby ti nic důležitého neuteklo, potřebuješ na vybraný e-mail nastavit automatické pravidlo třídění, přiřadit mu barevnou kategorii nebo z něj rovnou vytvořit sledovaný úkol s připomenutím.",
            hint: "V seznamu doručené pošty klikni pravým tlačítkem myši na libovolný e-mail, najeď na 'Pravidla ›' a zvol 'Vytvořit pravidlo...' (nebo zvol 'Zpracovat ›' -> 'Přidat připomenutí...' či 'Zařadit do kategorií').",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: {
                action: "outlook_context_workflow",
                completed: true
            }
        },
        {
            id: "m_outlook_2",
            categoryId: "outlook",
            difficulty: 0,
            type: "sandbox",
            appModule: "outlook",
            topic: "Kalendář, rezervace místností a Pomocník pro plánování",
            taskDescription: "Potřebuješ svolat osobní schůzku v budově firmy. Musíš k ní najít a zarezervovat volnou zasedací místnost a zároveň zkontrolovat kalendáře všech účastníků, aby se schůzka nekryla s jinou událostí.",
            hint: "V levé svislé liště přepni na ikonu '📅 Kalendář' a klikni na 'Nová schůzka'. V pravém panelu Místnosti vyber volnou místnost, přepni nahoře na 'Pomocník pro plánování', vyber volný čas '11:35 – 12:00' a klikni na 'Poslat'.",
            timeLimitSeconds: 65,
            maxMistakes: 3,
            targetCondition: {
                action: "outlook_schedule_meeting",
                completed: true
            }
        },
        {
            id: "m_outlook_3",
            categoryId: "outlook",
            difficulty: 1,
            type: "sandbox",
            appModule: "outlook",
            topic: "Vytváření šablon e-mailů (*.oft) a správa schránky",
            taskDescription: "Pravidelně posíláš kolegům či klientům e-mail se stejnou strukturou a textem. Aby ses nemusel(a) pokaždé zdržovat kopírováním starých zpráv, vytvoř si nový e-mail a ulož ho jako oficiální šablonu Outlooku (*.oft).",
            hint: "Klikni v pásu karet na 'Nový e-mail'. V otevřeném okně zprávy klikni vlevo nahoře na 'Soubor' -> 'Uložit jako', v rozbalovacím seznamu 'Uložit jako typ' vyber 'Šablona Outlook (*.oft)' a klikni na 'Uložit'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: {
                action: "outlook_save_template",
                fileType: "oft"
            }
        },
        {
            id: "m_outlook_4",
            categoryId: "outlook",
            difficulty: 0,
            type: "sandbox",
            appModule: "outlook",
            topic: "Přikládání souborů",
            taskDescription: "Kolegyni máš poslat zprávu s přehledem rozpočtu, který máš uložený na OneDrive. Vytvoř e-mail a připoj k němu soubor.",
            hint: "V okně nového e-mailu klikni v pásu karet na '📎 Připojit soubor ▾'.",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "outlook_attach_file", completed: true }
        },
        {
            id: "m_outlook_5",
            categoryId: "outlook",
            difficulty: 0,
            type: "sandbox",
            appModule: "outlook",
            topic: "Třídění a kategorizace e-mailů",
            taskDescription: "V doručené poště se ti hromadí zprávy k různým projektům. Přiřaď jednomu e-mailu barevnou kategorii, aby ses v něm rychleji orientoval(a).",
            hint: "Klikni pravým tlačítkem myši na e-mail v seznamu, najeď na 'Zařadit do kategorií' a vyber barvu.",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "outlook_categorize", completed: true }
        },
        {
            id: "m_outlook_6",
            categoryId: "outlook",
            difficulty: 0,
            type: "sandbox",
            appModule: "outlook",
            topic: "Vyhledávání e-mailů",
            taskDescription: "Potřebuješ najít e-mail s fakturou od dodavatele, ale nepamatuješ si přesný předmět. Vyhledej ho v poště pomocí klíčového slova.",
            hint: "Klikni do pole 'Hledat' nahoře, napiš klíčové slovo (např. 'faktura') a stiskni Enter.",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "outlook_search", completed: true }
        },
        {
            id: "m_outlook_7",
            categoryId: "outlook",
            difficulty: 0,
            type: "sandbox",
            appModule: "outlook",
            topic: "Šablony e-mailů",
            taskDescription: "Pro opakované odpovědi na dotazy máš připravené e-mailové šablony. Otevři nastavení šablon e-mailů.",
            hint: "V nastavení Outlooku (Soubor -> Možnosti -> Pošta) klikni na 'Šablony e-mailů a písma...'.",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "outlook_email_templates", completed: true }
        },
        {
            id: "m_outlook_8",
            categoryId: "outlook",
            difficulty: 0,
            type: "sandbox",
            appModule: "outlook",
            topic: "Orientace v kalendáři",
            taskDescription: "Chceš zjistit, jaké schůzky máš dnes v kalendáři. Otevři kalendář a podívej se na detail jedné z událostí.",
            hint: "V levé svislé liště přepni na '📅 Kalendář' a klikni na libovolnou kartu události v agendě.",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "outlook_calendar_open", completed: true }
        },
        {
            id: "m_outlook_9",
            categoryId: "outlook",
            difficulty: 1,
            type: "sandbox",
            appModule: "outlook",
            topic: "My Templates",
            taskDescription: "Chceš si otevřít své šablony zpráv, abys je mohl(a) rychle využít v běžné práci.",
            hint: "Klikni na 'Nový e-mail', potom na 'Všechny aplikace' a vyber 'My Templates'.",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "outlook_view_templates", completed: true }
        },
        {
            id: "m_outlook_10",
            categoryId: "outlook",
            difficulty: 1,
            type: "sandbox",
            appModule: "outlook",
            topic: "Diktování e-mailu",
            taskDescription: "Máš rozepsaný dlouhý e-mail a místo psaní na klávesnici ho chceš nadiktovat. Aktivuj funkci diktování.",
            hint: "V okně nového e-mailu přepni na kartu 'Domů' a klikni na tlačítko 'Diktovat' 🎙️.",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "outlook_dictate", completed: true }
        },
        {
            id: "m_outlook_11",
            categoryId: "outlook",
            difficulty: 2,
            type: "sandbox",
            appModule: "outlook",
            topic: "Pokročilé filtry",
            taskDescription: "V poště hledáš nepřečtené zprávy s přílohou od konkrétního odesílatele. Nastav pokročilé filtrování e-mailů.",
            hint: "Nad seznamem e-mailů klikni na '⏳ Filtrovat e-maily ▾' a vyber pokročilý filtr.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "outlook_advanced_filter", completed: true }
        },
        {
            id: "m_outlook_12",
            categoryId: "outlook",
            difficulty: 2,
            type: "sandbox",
            appModule: "outlook",
            topic: "Rezervace schůzek přes Bookings",
            taskDescription: "Kolegové potřebují domluvit konzultaci s tebou bez zdlouhavého hledání volného termínu. Otevři stránku Bookings, kde si mohou termín sami rezervovat.",
            hint: "Na kartě 'Domů' v pásu karet klikni na 'Bookings stránka ▾'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "outlook_bookings", completed: true }
        },
        {
            id: "m_outlook_13",
            categoryId: "outlook",
            difficulty: 2,
            type: "sandbox",
            appModule: "outlook",
            topic: "Správa kalendářů a jejich sdílení",
            taskDescription: "Tým potřebuje vidět tvoji dostupnost. Nasdílej svůj kalendář kolegům.",
            hint: "V kalendáři klikni na 'Sdílet kalendář ▾' v pásu karet.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "outlook_share_calendar", completed: true }
        },
        {
            id: "m_outlook_14",
            categoryId: "outlook",
            difficulty: 3,
            type: "sandbox",
            appModule: "outlook",
            topic: "Archivace a údržba schránky",
            taskDescription: "Schránka je plná starých e-mailů z minulého roku. Přesuň vybranou zprávu do archivu, aby se v doručené poště neztrácela.",
            hint: "Vyber e-mail v seznamu a v pásu karet klikni na tlačítko pro archivaci.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "outlook_archive", completed: true }
        },

        // ================= SLOUPEC 5: MICROSOFT ONENOTE =================
        {
            id: "m_onenote_0",
            categoryId: "onenote",
            difficulty: 0,
            type: "sandbox",
            appModule: "onenote",
            topic: "Zápisky, stránky a sekce",
            taskDescription: "Rozjíždíš nový projekt a potřebuješ si v digitálním poznámkovém bloku OneNote přehledně rozčlenit zápisky. Založ v bloku nový tematický oddíl (sekci) nebo přidej novou prázdnou stránku.",
            hint: "V aplikaci OneNote klikni v levém sloupci na 'Vlastní oddíl+' pro vytvoření oddílu, nebo na tlačítko '+ Přidat stránku'.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "create_section_page", completed: true }
        },
        {
            id: "m_onenote_1",
            categoryId: "onenote",
            difficulty: 0,
            type: "sandbox",
            appModule: "onenote",
            topic: "Přidávání textu, obrázků a odkazů",
            taskDescription: "Při psaní poznámek z workshopu potřebuješ v textu zvýraznit klíčové pojmy a připojit přímý odkaz na související dokumentaci. Použij nástroje pro formátování textu a vložení odkazu přímo na plátně poznámky.",
            hint: "Na bílém plátně poznámky klikni pravým tlačítkem myši a z kontextového menu zvol 'Formátovat text / Odkaz'.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "format_text_link", completed: true }
        },
        {
            id: "m_onenote_2",
            categoryId: "onenote",
            difficulty: 0,
            type: "sandbox",
            appModule: "onenote",
            topic: "Vyhledávání poznámek",
            taskDescription: "Víš, že sis před několika týdny zapsal(a) důležitý postup, ale už si nepamatuješ, ve kterém oddílu nebo poznámkovém bloku to bylo. Použij vyhledávání v OneNotu a nastav rozsah hledání napříč bloky.",
            hint: "Klikni nahoře ve fialové liště do vyhledávacího pole a v rozbalovací nabídce vyber 'Všechny poznámkové bloky' nebo 'Tento oddíl'.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "search_notebooks", completed: true }
        },
        {
            id: "m_onenote_3",
            categoryId: "onenote",
            difficulty: 1,
            type: "sandbox",
            appModule: "onenote",
            topic: "Struktura bloků",
            taskDescription: "Potřebuješ zkontrolovat, kde je tvůj poznámkový blok uložen, zda se správně synchronizuje do cloudu a jaké má nastavení. Otevři přehled informací o poznámkovém bloku.",
            hint: "Otevři Backstage menu kliknutím vlevo nahoře na fialové tlačítko 'Soubor' a prohlédni si Informace o poznámkovém bloku.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "notebook_structure", completed: true }
        },
        {
            id: "m_onenote_4",
            categoryId: "onenote",
            difficulty: 1,
            type: "sandbox",
            appModule: "onenote",
            topic: "Organizace úkolů",
            taskDescription: "Z porady sis vypsal(a) seznam úkolů, které musíš během týdne splnit. Aby sis mohl(a) hotové položky odškrtávat, označ text v poznámce interaktivní značkou úkolu.",
            hint: "V pásu karet na kartě 'Domů' klikni na značku '☑️ Úkol (Ctrl+1)' a vlož do poznámek zaškrtávací seznam úkolů.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "organize_tasks", completed: true }
        },
        {
            id: "m_onenote_5",
            categoryId: "onenote",
            difficulty: 1,
            type: "sandbox",
            appModule: "onenote",
            topic: "Vkládání souborů",
            taskDescription: "Chceš mít k poznámkám z projektu přiloženou i původní prezentaci či PDF příručku, abys do ní mohl(a) rovnou psát své poznámky. Vlož do stránky OneNotu externí soubor nebo jeho výtisk.",
            hint: "Přepni nahoře na kartu 'Vložení' v pásu karet a klikni na '📎 Soubor jako příloha' nebo '📄 Výtisk souboru'.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "insert_files", completed: true }
        },
        {
            id: "m_onenote_6",
            categoryId: "onenote",
            difficulty: 2,
            type: "sandbox",
            appModule: "onenote",
            topic: "Sdílení poznámkového bloku",
            taskDescription: "Tvůj poznámkový blok potřebují číst a upravovat i ostatní členové projektového týmu. Nasdílej jim poznámkový blok tak, abyste v něm mohli pracovat společně v reálném čase.",
            hint: "Klikni vpravo nahoře na tlačítko 'Sdílet ▾' (nebo v menu Soubor -> Sdílet) a nasdílej poznámkový blok kolegům.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "share_notebook", completed: true }
        },
        {
            id: "m_onenote_7",
            categoryId: "onenote",
            difficulty: 2,
            type: "sandbox",
            appModule: "onenote",
            topic: "Šablony meetingů",
            taskDescription: "Chystáš se zapisovat poznámky z plánované porady. Místo ručního vypisování data, místa a seznamu účastníků si tyto detaily schůzky vlož do stránky automaticky ze svého kalendáře.",
            hint: "V pásu karet na kartě 'Domů' klikni na tlačítko '📅 Podrobnosti schůzky' a vlož zápis z kalendáře Outlooku.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "meeting_templates", completed: true }
        },
        {
            id: "m_onenote_8",
            categoryId: "onenote",
            difficulty: 2,
            type: "sandbox",
            appModule: "onenote",
            topic: "Propojení s Outlookem",
            taskDescription: "Schůzka právě skončila a ty potřebuješ svůj hotový zápis z OneNotu okamžitě rozeslat všem účastníkům e-mailem, aniž bys musel(a) text ručně kopírovat do Outlooku.",
            hint: "V pásu karet na kartě 'Domů' klikni na tlačítko '✉️ Odeslat stránku'.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "outlook_integration", completed: true }
        },
        {
            id: "m_onenote_9",
            categoryId: "onenote",
            difficulty: 3,
            type: "sandbox",
            appModule: "onenote",
            topic: "Propojení s Teams a Copilotem",
            taskDescription: "Chceš, aby byl tento sdílený poznámkový blok trvale po ruce všem kolegům přímo v komunikačním kanálu vašeho týmu v Microsoft Teams. Propoj svůj OneNote s kanálem v Teams.",
            hint: "Klikni v pásu karet na tlačítko '🟣 Sdílet do Teams' (nebo v menu Soubor -> Sdílet do kanálu Teams) a propoj blok s týmem.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "teams_integration", completed: true }
        },
        {
            id: "m_onenote_10",
            categoryId: "onenote",
            difficulty: 2,
            type: "sandbox",
            appModule: "onenote",
            topic: "OneNote 2-Specialista: Propojení s Teams",
            taskDescription: "Poznámkový blok 'Projekt Testovací firma' chce tým Jany Testovací připnout do kanálu MS Teams. Propoj blok s Teams přes Soubor → Informace.",
            hint: "Otevři Soubor → Informace a klikni na 'Připnout do kanálu MS Teams'.",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "teams_channel_pin", completed: true }
        },

        // ================= SLOUPEC 6: MICROSOFT FORMS =================
        {
            id: "m_forms_0",
            categoryId: "forms",
            difficulty: 0,
            type: "sandbox",
            appModule: "forms",
            topic: "Vyplnění dotazníku, který dostanu",
            taskDescription: "Kolega ti poslal odkaz na krátký znalostní kvíz v Microsoft Forms. Projdi si jednotlivé otázky v dotazníku, zkontroluj nebo vyber odpovědi a formulář na konci odešli ke zpracování.",
            hint: "V pohledu respondenta (kvíz 'Víte víc? Otestujte se v kvízu') sroluj pod třetí otázku a klikni na tlačítko 'Odeslat'.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "forms_fill_quiz", completed: true }
        },
        {
            id: "m_forms_1",
            categoryId: "forms",
            difficulty: 2,
            type: "sandbox",
            appModule: "forms",
            topic: "Vytvoření otázek v dotazníku",
            taskDescription: "Potřebuješ rychle sesbírat zpětnou vazbu od kolegů po skončení školení. Založ v aplikaci Microsoft Forms nový dotazník a vlož do něj první otázku (např. výběr z možností, textovou odpověď nebo hodnocení).",
            hint: "Na domovské obrazovce klikni na '➕ Nový formulář' (nebo přepni nahoře do Editoru dotazníku) a v sekci '+ Rychlý start s' klikni na libovolný typ otázky (např. 'Volba' nebo 'Text').",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "forms_create_questions", completed: true }
        },
        {
            id: "m_forms_2",
            categoryId: "forms",
            difficulty: 2,
            type: "sandbox",
            appModule: "forms",
            topic: "Administrace dotazníku – nastavení",
            taskDescription: "Než svůj dotazník rozešleš respondentům, potřebuješ upravit jeho pravidla – například nastavit, zda je anonymní nebo zaznamenává jména v organizaci, omezit počet odpovědí na jednu osobu nebo zapnout indikátor průběhu.",
            hint: "V editoru dotazníku klikni v horní liště na '⚙️ Nastavení' a v pravém panelu změň nebo zaškrtni libovolnou volbu.",
            timeLimitSeconds: 50,
            maxMistakes: 3,
            targetCondition: { action: "forms_admin_settings", completed: true }
        },
        {
            id: "m_forms_3",
            categoryId: "forms",
            difficulty: 3,
            type: "sandbox",
            appModule: "forms",
            topic: "Automatizace aktivity spojených s dotazníkem – rozesílka, sběr, vyhodnocení",
            taskDescription: "Tvůj dotazník je připravený. Nyní potřebuješ zautomatizovat proces sběru a vyhodnocení: rozešli pozvánku k vyplnění přes Outlook/Teams (nebo vygeneruj zkrácený odkaz), zapni automatické e-mailové notifikace o každé nové odpovědi nebo otevři výsledky.",
            hint: "V editoru dotazníku klikni vpravo nahoře na zelené tlačítko '▷ Shromažďovat odpovědi' a klikni na 'Poslat' či 'Kopírovat odkaz' (případně klikni na '📊 Zobrazit odpovědi' nebo v '⚙️ Nastavení' zaškrtni 'Dostávat e-mailové oznámení o každé odpovědi').",
            timeLimitSeconds: 45,
            maxMistakes: 3,
            targetCondition: { action: "forms_automation", completed: true }
        },

        // ================= MICROSOFT TO DO =================
        {
            id: "m_todo_0",
            categoryId: "todo",
            difficulty: 0,
            type: "sandbox",
            appModule: "todo",
            topic: "Vytvoření úkolu",
            taskDescription: "Vytvoř nový úkol s názvem „Připravit podklady na poradu“ v seznamu Úkoly.",
            hint: "Úkoly → pole „Přidat úkol“ → napiš název → Přidat",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "todo_create_task", minLength: 3 }
        },
        {
            id: "m_todo_1",
            categoryId: "todo",
            difficulty: 1,
            type: "sandbox",
            appModule: "todo",
            topic: "Termín a připomenutí",
            taskDescription: "Nastav úkolu „Objednat catering na firemní akci“ termín na zítra.",
            hint: "Vyber úkol → Připomenout mě / Termín → Zítra",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "todo_set_due_date", value: "tomorrow" }
        },
        {
            id: "m_todo_2",
            categoryId: "todo",
            difficulty: 1,
            type: "sandbox",
            appModule: "todo",
            topic: "Můj den",
            taskDescription: "Přidej úkol „Zkontrolovat zálohy“ do seznamu Můj den.",
            hint: "Vyber úkol → Přidat do Můj den (ikona sluníčka)",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "todo_add_to_my_day", value: true }
        },
        {
            id: "m_todo_3",
            categoryId: "todo",
            difficulty: 2,
            type: "sandbox",
            appModule: "todo",
            topic: "Sdílené seznamy",
            taskDescription: "Sdílej seznam Úkoly s kolegou Janem Novákem (jan.novak@example.com).",
            hint: "Levý panel → Úkoly → Sdílet → zadej adresu → Sdílet",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "todo_share_list", minLength: 5 }
        },
        {
            id: "m_todo_4",
            categoryId: "todo",
            difficulty: 2,
            type: "sandbox",
            appModule: "todo",
            topic: "Příloha k úkolu",
            taskDescription: "Přidej k úkolu „Připravit osnovu workshopu o OneNote“ soubor Osnova_workshop.docx.",
            hint: "Vyber úkol → Přidat soubor → vyber soubor",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "todo_attach_file" }
        },
        {
            id: "m_todo_5",
            categoryId: "todo",
            difficulty: 2,
            type: "sandbox",
            appModule: "todo",
            topic: "Úkol z Teams zprávy",
            taskDescription: "Vytvoř úkol z Teams zprávy v panelu Návrhy (tlačítko +).",
            hint: "Můj den → pravý panel Návrhy → u zprávy klikni na +",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "todo_from_teams", source: "teams" }
        },
        {
            id: "m_todo_6",
            categoryId: "todo",
            difficulty: 2,
            type: "sandbox",
            appModule: "todo",
            topic: "Úkol z e-mailu",
            taskDescription: "Z e-mailu přidej úkol do seznamu Úkoly.",
            hint: "Pravé tlačítko na e-mail → Přidat do Úkolů",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "todo_from_teams", source: "email" }
        },
        {
            id: "m_todo_7",
            categoryId: "todo",
            difficulty: 1,
            type: "sandbox",
            appModule: "todo",
            topic: "Důležité úkoly",
            taskDescription: "Označ úkol „Zkontrolovat rozpočet projektu“ jako důležitý.",
            hint: "Vyber úkol → hvězdička u úkolu (nebo Označit jako důležité)",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "todo_toggle_important", value: true }
        },

        // ================= MICROSOFT PLANNER =================
        {
            id: "m_planner_0",
            categoryId: "planner",
            difficulty: 2,
            type: "sandbox",
            appModule: "planner",
            topic: "Vytvoření úkolu",
            taskDescription: "V plánu „Tým – jarní kampaň“ vytvoř nový úkol s názvem alespoň 3 znaky a přidej ho do libovolného segmentu.",
            hint: "Ve sloupci segmentu klikni na „+ Přidat úkol“, vyplň název a potvrď tlačítkem „Přidat“.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "planner_create_task", minLength: 3 }
        },
        {
            id: "m_planner_1",
            categoryId: "planner",
            difficulty: 2,
            type: "sandbox",
            appModule: "planner",
            topic: "Připomenutí úkolu",
            taskDescription: "U úkolu „Příprava rozpočtu na Q1“ nastav připomenutí „1 den předem“ nebo „V 9:00 v den termínu“.",
            hint: "Otevři úkol kliknutím na kartu, v detailu v poli „Připomenutí“ vyber požadovanou možnost.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "planner_set_reminder", reminder: ["1 den předem", "V 9:00 v den termínu"] }
        },
        {
            id: "m_planner_2",
            categoryId: "planner",
            difficulty: 2,
            type: "sandbox",
            appModule: "planner",
            topic: "Opakované úkoly",
            taskDescription: "U úkolu „Předávání informací na týmovkách“ nastav opakování na „Týdně“ nebo „Měsíčně“.",
            hint: "Otevři úkol kliknutím na kartu a v detailu v poli „Opakování“ vyber „Týdně“ nebo „Měsíčně“.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "planner_set_recurrence", recurrence: ["Týdně", "Měsíčně"] }
        },
        {
            id: "m_planner_3",
            categoryId: "planner",
            difficulty: 3,
            type: "sandbox",
            appModule: "planner",
            topic: "Sdílení plánu",
            taskDescription: "Přidej do plánu „Tým – jarní kampaň“ nového člena s e-mailem alespoň 3 znaky (nebo jménem).",
            hint: "V horní části plánu klikni na tlačítko „Sdílet“, do pole zadej jméno a potvrď tlačítkem „Přidat“.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "planner_share_plan", minLength: 3 }
        },
        {
            id: "m_planner_4",
            categoryId: "planner",
            difficulty: 2,
            type: "sandbox",
            appModule: "planner",
            topic: "Příloha k úkolu",
            taskDescription: "K úkolu „Vyhodnocení průzkumu spokojenosti“ přidej soubor „Zápis_z_porady_07.docx“ nebo „Harmonogram_HRS_2026.xlsx“.",
            hint: "Otevři úkol, v detailu klikni na „Přidat soubor“ a vyber požadovaný soubor ze seznamu.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "planner_attach_file", file: ["Zápis_z_porady_07.docx", "Harmonogram_HRS_2026.xlsx"] }
        },
        {
            id: "m_planner_5",
            categoryId: "planner",
            difficulty: 2,
            type: "sandbox",
            appModule: "planner",
            topic: "Úkol z Teams zprávy",
            taskDescription: "Vytvoř úkol z Teams zprávy „Připravit podklady pro jarní setkání“ pomocí tlačítka „Z Teams zprávy“.",
            hint: "Na panelu plánu klikni na „Z Teams zprávy“ a v okně vyber zprávu „Připravit podklady pro jarní setkání“.",
            timeLimitSeconds: 60,
            maxMistakes: 3,
            targetCondition: { action: "planner_from_teams", source: "teams" }
        }
    ]
};