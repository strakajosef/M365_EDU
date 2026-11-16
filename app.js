/**
 * ============================================================================
 * HLAVNÍ HERNÍ ENGINE (app.js)
 * ============================================================================
 */

window.SandboxModules = window.SandboxModules || {};

const TIME_LIMIT_BY_LEVEL = {
    0: 90, // Základní: 90 vteřin
    1: 50, // Pokročilý: 50 vteřin
    2: 40, // Specialista: 40 vteřin
    3: 30  // Expert: 30 vteřin
};

const MODERATOR_IMAGES = {
    norm: "moderator_norm.png",
    ques: "moderator_ques.png",
    happy: "moderator_happy.png",
    no: "moderator_no.png"
};

const appState = {
    userLevel: 0, // 0 = Základní, 1 = Pokročilý, 2 = Specialista, 3 = Expert
    score: 0,
    answeredCount: 0,
    maxQuestions: GAME_DATA.config.maxQuestionsPerGame,
    wrongTopics: [],
    playedMissionIds: [],

    activeMission: null,
    activePoints: 0,
    activeCellBtn: null,
    isFreeTraining: false,
    activeGame: null,

    timerInterval: null,
    remainingSeconds: 0,
    remainingLives: 3,
    loadedCssLink: null,

    moderatorResetTimer: null,
    lastModeratorQuote: ""
};

const dom = {
    screenMenu: document.getElementById("screen-menu"),
    screenSandbox: document.getElementById("screen-sandbox"),
    btnBackToMenu: document.getElementById("btn-back-to-menu"),
    userLevelSelector: document.getElementById("user-level-selector"),

    scoreDisplay: document.getElementById("score-display"),
    moderatorText: document.getElementById("moderator-text"),
    speechBubble: document.getElementById("speech-bubble"),
    moderatorImg: document.getElementById("moderator-img"),
    moderatorAvatarBox: document.getElementById("moderator-avatar-box") || document.querySelector(".moderator-avatar"),
    moderatorHoverBadge: document.getElementById("moderator-hover-badge"),

    // Sandbox UI
    sandboxAppBadge: document.getElementById("sandbox-app-badge"),
    sandboxTopicTitle: document.getElementById("sandbox-topic-title"),
    sandboxTaskDesc: document.getElementById("sandbox-task-desc"),
    sandboxHintWrapper: document.getElementById("sandbox-hint-wrapper"),
    sandboxHintPopup: document.getElementById("sandbox-hint-popup"),
    sandboxTimerBox: document.getElementById("sandbox-timer-box"),
    sandboxTimer: document.getElementById("sandbox-timer"),
    sandboxLivesBox: document.getElementById("sandbox-lives-box"),
    sandboxLives: document.getElementById("sandbox-lives"),
    sandboxViewport: document.getElementById("sandbox-viewport"),
    btnAbortSandbox: document.getElementById("btn-abort-sandbox"),
    sandboxToast: document.getElementById("sandbox-toast"),

    // Vyhodnocení Modal
    modalEvaluation: document.getElementById("modal-evaluation"),
    finalScoreValue: document.getElementById("final-score-value"),
    evaluationRecommendations: document.getElementById("evaluation-recommendations"),
    btnPlayAgain: document.getElementById("btn-play-again"),
    btnEvalToMenu: document.getElementById("btn-eval-to-menu")
};

const COPILOT_AGENT_URL = "https://m365.cloud.microsoft/chat/?titleId=T_b041f691-d215-6f56-804c-cbe3b8d78c33&source=embedded-builder";
const COPILOT_AGENT_TEXT = "Ještě mě nemáš jako Copilot agenta? Tak na mě klikni!";
let menuHoverPrevText = null;

function isMainMenuVisible() {
    return dom.screenMenu && !dom.screenMenu.classList.contains("hidden");
}

document.addEventListener("DOMContentLoaded", () => {
    // 0. Vytvoření obrazovek a karet všech zaregistrovaných her
    GameRegistry.mountAll({
        grid: document.querySelector(".games-grid"),
        screenAnchor: dom.screenSandbox
    });

    // 1. Přepínání úrovně uživatele v Hlavním menu
    document.querySelectorAll("#user-level-selector .level-btn").forEach(btn => {
        btn.addEventListener("click", () => {
            document.querySelectorAll("#user-level-selector .level-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            appState.userLevel = parseInt(btn.getAttribute("data-level"), 10);
        });
    });

    // 2. Spuštění Volného tréninku
    document.querySelectorAll("[data-free-app]").forEach(btn => {
        btn.addEventListener("click", (e) => {
            e.stopPropagation();
            const appName = btn.getAttribute("data-free-app");
            startFreeSandbox(appName);
        });
    });

    dom.btnBackToMenu.addEventListener("click", () => {
        stopSandboxTimer();
        appState.isFreeTraining = false;
        showScreen("menu");
        setModeratorQuote("menu");
        setModeratorAvatar("norm");
    });

    dom.btnAbortSandbox.addEventListener("click", () => {
        if (appState.isFreeTraining) {
            stopSandboxTimer();
            appState.isFreeTraining = false;
            showScreen("menu");
            setModeratorQuote("menu");
            setModeratorAvatar("norm");
        } else {
            finishSandboxMission(false, "Mise byla předčasně ukončena.");
        }
    });

    dom.btnPlayAgain.addEventListener("click", () => {
        dom.modalEvaluation.classList.add("hidden");
        appState.activeGame?.start();
    });

    dom.btnEvalToMenu.addEventListener("click", () => {
        dom.modalEvaluation.classList.add("hidden");
        appState.isFreeTraining = false;
        showScreen("menu");
        setModeratorQuote("menu");
        setModeratorAvatar("norm");
    });

    // 4. Hover události na postavě moderátora pro zobrazení návodu ve Volném tréninku
    if (dom.moderatorAvatarBox) {
        dom.moderatorAvatarBox.addEventListener("mouseenter", () => {
            if (isMainMenuVisible()) {
                menuHoverPrevText = dom.moderatorText.textContent;
                dom.moderatorAvatarBox.classList.add("is-copilot-link");
                dom.moderatorText.textContent = COPILOT_AGENT_TEXT;
                return;
            }
            if (appState.isFreeTraining && appState.activeMission?.hint) {
                dom.speechBubble.classList.add("hint-active");
                dom.moderatorText.innerHTML = `<strong>💡 Návod jak na to:</strong><br>${appState.activeMission.hint}`;
            }
        });

        dom.moderatorAvatarBox.addEventListener("mouseleave", () => {
            if (menuHoverPrevText !== null) {
                dom.moderatorAvatarBox.classList.remove("is-copilot-link");
                dom.moderatorText.textContent = menuHoverPrevText;
                menuHoverPrevText = null;
                return;
            }
            if (appState.isFreeTraining) {
                dom.speechBubble.classList.remove("hint-active");
                dom.moderatorText.textContent = appState.lastModeratorQuote || "Trénuj vlastním tempem! Když si nebudeš vědět rady, najeď na mě myší.";
            }
        });

        dom.moderatorAvatarBox.addEventListener("click", () => {
            if (isMainMenuVisible()) {
                window.open(COPILOT_AGENT_URL, "_blank", "noopener");
            }
        });
    }

    setModeratorQuote("menu");
    setModeratorAvatar("norm");
});

/**
 * Nastaví obrázek moderátora.
 * Pokud je zadán durationMs (např. 5000 pro happy/no), po uplynutí doby vrátí výchozí postavu
 * (moderator_ques ve volném tréninku, jinak moderator_norm).
 */
function setModeratorAvatar(type, durationMs = 0) {
    if (!dom.moderatorImg) return;

    if (appState.moderatorResetTimer) {
        clearTimeout(appState.moderatorResetTimer);
        appState.moderatorResetTimer = null;
    }

    const src = MODERATOR_IMAGES[type] || MODERATOR_IMAGES.norm;
    dom.moderatorImg.src = src;

    // Vizuální indikace, že moderátor nabízí nápovědu po najetí myší (pouze ve volném tréninku)
    const showHintIndicator = appState.isFreeTraining && !dom.screenSandbox.classList.contains("hidden");
    if (dom.moderatorAvatarBox) {
        dom.moderatorAvatarBox.classList.toggle("has-hint", showHintIndicator);
    }
    if (dom.moderatorHoverBadge) {
        dom.moderatorHoverBadge.classList.toggle("hidden", !showHintIndicator);
    }

    if (durationMs > 0) {
        appState.moderatorResetTimer = setTimeout(() => {
            const defaultType = (appState.isFreeTraining && !dom.screenSandbox.classList.contains("hidden"))
                ? "ques"
                : "norm";
            dom.moderatorImg.src = MODERATOR_IMAGES[defaultType];
            appState.moderatorResetTimer = null;
        }, durationMs);
    }
}

/**
 * Zobrazí obrazovku "screen-<name>" (menu, sandbox nebo obrazovka libovolné hry).
 * Návrat do menu zruší časovače aktivní hry.
 */
function showScreen(screenName) {
    if (screenName === "menu") appState.activeGame?.cancel?.();
    document.querySelectorAll(".screen").forEach(s => {
        s.classList.toggle("hidden", s.id !== `screen-${screenName}`);
    });
    dom.btnBackToMenu.classList.toggle("hidden", screenName === "menu");
    dom.speechBubble?.classList.remove("hint-active");
}

function loadModuleStylesheet(appModuleName) {
    if (appState.loadedCssLink) {
        appState.loadedCssLink.remove();
        appState.loadedCssLink = null;
    }
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = `${appModuleName}.css`;
    link.id = `dynamic-css-${appModuleName}`;
    document.head.appendChild(link);
    appState.loadedCssLink = link;
}

// ============================================================================
// SOUTĚŽ RISKUJ! (NÁHODNÝ VÝBĚR OTÁZKY DLE ÚROVNĚ PŘI KLIKNUTÍ NA POLE)
// ============================================================================
// ============================================================================
// JÁDRO PRO HRY (EduaCore) – sdílí je Riskuj, Kolo štěstí a případné další hry
// ============================================================================

/** Resetuje běh hry a nastaví ji jako aktivní. */
function beginGame(game) {
    appState.activeGame = game;
    appState.score = 0;
    appState.answeredCount = 0;
    appState.wrongTopics = [];
    appState.playedMissionIds = [];
    appState.isFreeTraining = false;

    const levelInfo = GAME_DATA.config.difficulties[appState.userLevel];
    const timeLimit = TIME_LIMIT_BY_LEVEL[appState.userLevel] || 60;
    const badge = document.querySelector(`#screen-${game.id} [data-level-badge]`);
    if (badge) {
        badge.innerHTML = `
            Zvolená úroveň: <b>${appState.userLevel} – ${levelInfo.name}</b> (losují se úkoly úrovně 0 až ${appState.userLevel} | čas na úkol: <b>${timeLimit} s</b>)
        `;
    }

    updateScoreUI(0);
    updateCounterUI();
}

/** Nevybrané úkoly dané oblasti odpovídající úrovni hráče. */
function getEligibleMissions(categoryId) {
    return GAME_DATA.missions.filter(
        m => m.categoryId === categoryId
            && m.difficulty <= appState.userLevel
            && !appState.playedMissionIds.includes(m.id)
    );
}

/** Spustí sandbox s daným úkolem v soutěžním režimu. Výsledek zpracuje activeGame.onMissionFinished. */
function launchMission(mission, points) {
    appState.playedMissionIds.push(mission.id);
    appState.activeMission = mission;
    appState.activePoints = points;
    appState.isFreeTraining = false;
    openSandboxMission(mission);
}

// Přednost: oblasti, ze kterých ještě nic nevybral a nejsou na aktuální ploše;
// poté oblasti, ze kterých ještě nic nevybral, ale jsou na ploše;
// nakonec (když už žádné nezbydou) náhodné oblasti, ze kterých již něco vybral.

// ============================================================================
// OTEVŘENÝ SANDBOX (Soutěžní mise vs. Volný trénink s moderator_ques)
// ============================================================================
function openSandboxMission(mission) {
    setModeratorQuote("beforeSandbox");
    loadModuleStylesheet(mission.appModule);
    showScreen("sandbox");
    setModeratorAvatar("norm");

    const diffName = GAME_DATA.config.difficulties[mission.difficulty]?.name || "Úkol";
    dom.sandboxAppBadge.textContent = `${mission.appModule.toUpperCase()} (${mission.difficulty}-${diffName})`;
    dom.sandboxTopicTitle.textContent = `${mission.topic} (${appState.activePoints} bodů)`;
    dom.sandboxTaskDesc.textContent = mission.taskDescription;
    dom.btnAbortSandbox.textContent = "Vzdat misi (-body)";

    // Horní štítek nápovědy zůstává trvale skrytý (nápovědu drží moderátor ve volném tréninku)
    dom.sandboxHintWrapper?.classList.add("hidden");

    appState.remainingLives = mission.maxMistakes || 3;
    appState.remainingSeconds = TIME_LIMIT_BY_LEVEL[appState.userLevel] || TIME_LIMIT_BY_LEVEL[mission.difficulty] || 90;

    dom.sandboxTimerBox.classList.remove("hidden", "danger-pulse");
    dom.sandboxLivesBox.classList.remove("hidden");
    updateSandboxMetricsUI();

    startSandboxTimer();
    mountSandboxModule(mission.appModule, mission);
}

function startFreeSandbox(appModuleName, selectedMissionId = null) {
    const appMissions = GAME_DATA.missions.filter(
        m => m.appModule === appModuleName
    );
    if (!appMissions.length) return;

    const activeMission = appMissions.find(m => m.id === selectedMissionId) || appMissions[0];

    appState.isFreeTraining = true;
    appState.activeGame = null;
    appState.activeMission = activeMission;
    loadModuleStylesheet(appModuleName);
    showScreen("sandbox");

    // Ve volném tréninku se zobrazí postavička moderator_ques a po najetí myší ukáže návod
    setModeratorQuote("freeTraining");
    setModeratorAvatar("ques");

    dom.sandboxAppBadge.textContent = `${appModuleName.toUpperCase()} (VOLNÝ TRÉNINK)`;

    if (appMissions.length > 1) {
        dom.sandboxTopicTitle.innerHTML = `
            Tréninkový úkol: 
            <select id="free-mission-switcher" style="padding:3px 8px; border-radius:6px; border:1px solid #cbd5e1; font-weight:600; font-family:inherit;">
                ${appMissions.map((m, idx) => `
                    <option value="${m.id}" ${m.id === activeMission.id ? "selected" : ""}>
                        ${idx + 1}. [${GAME_DATA.config.difficulties[m.difficulty]?.name}] ${m.topic}
                    </option>
                `).join("")}
            </select>
        `;
        document.getElementById("free-mission-switcher")?.addEventListener("change", (e) => {
            startFreeSandbox(appModuleName, e.target.value);
        });
    } else {
        dom.sandboxTopicTitle.textContent = activeMission.topic;
    }

    dom.sandboxTaskDesc.textContent = activeMission.taskDescription;
    dom.btnAbortSandbox.textContent = "Ukončit trénink";

    // Původní horní bublinu nápovědy skryjeme, protože návod je nyní u postavičky moderátora
    dom.sandboxHintWrapper?.classList.add("hidden");

    dom.sandboxTimerBox.classList.add("hidden");
    dom.sandboxLivesBox.classList.add("hidden");

    mountSandboxModule(appModuleName, activeMission);
}

function mountSandboxModule(appModuleName, mission) {
    const moduleObj = window.SandboxModules[appModuleName];
    if (moduleObj && typeof moduleObj.init === "function") {
        moduleObj.init(dom.sandboxViewport, mission, {
            onAction: (actionPayload) => verifySandboxAction(actionPayload),
            onMistake: (reasonText) => {
                if (appState.isFreeTraining) {
                    setModeratorAvatar("no", 5000);
                    showSandboxToast(`⚠️ Pozor: ${reasonText}`, "#F59E0B");
                } else {
                    handleSandboxMistake(reasonText);
                }
            },
            onInfo: (infoText) => showSandboxToast(infoText, "#0078D4")
        });
    } else {
        console.error("Modul nenalezen:", appModuleName);
    }
}

function verifySandboxAction(payload) {
    const cond = appState.activeMission?.targetCondition;
    if (!cond) return;

    let isMatch = true;
    for (const key of Object.keys(cond)) {
        if (key === "minLength") {
            if (!payload.value || payload.value.length < cond.minLength) isMatch = false;
        } else if (Array.isArray(cond[key])) {
            if (!cond[key].includes(payload[key])) isMatch = false;
        } else if (payload[key] !== cond[key]) {
            isMatch = false;
        }
    }

    if (isMatch) {
        if (appState.isFreeTraining) {
            setModeratorQuote("correctAnswer");
            setModeratorAvatar("happy", 5000);
            showSandboxToast("🎉 Skvělá práce! Úkol v tréninkovém režimu jsi splnil(a). Můžeš si nahoře přepnout další úkol nebo aplikaci dál zkoumat.", "#10B981");
        } else {
            finishSandboxMission(true, `🎉 Mise splněna v čase! Získáváš +${appState.activePoints} bodů.`);
        }
    }
}

function handleSandboxMistake(reasonText) {
    appState.remainingLives--;
    updateSandboxMetricsUI();
    setModeratorQuote("sandboxMistake");
    setModeratorAvatar("no", 5000);
    showSandboxToast(`⚠️ Chybný krok (-1 pokus): ${reasonText}`, "#EF4444");

    if (appState.remainingLives <= 0 && !appState.isFreeTraining) {
        finishSandboxMission(false, "Vyčerpal(a) jsi všechny povolené pokusy v této misi!");
    }
}

function startSandboxTimer() {
    stopSandboxTimer();
    appState.timerInterval = setInterval(() => {
        appState.remainingSeconds--;
        updateSandboxMetricsUI();

        if (appState.remainingSeconds <= 10) {
            dom.sandboxTimerBox.classList.add("danger-pulse");
        }
        if (appState.remainingSeconds <= 0) {
            finishSandboxMission(false, "⏱️ Časový limit pro splnění praktického úkolu vypršel!");
        }
    }, 1000);
}

function stopSandboxTimer() {
    if (appState.timerInterval) {
        clearInterval(appState.timerInterval);
        appState.timerInterval = null;
    }
}

function updateSandboxMetricsUI() {
    const mins = String(Math.floor(appState.remainingSeconds / 60)).padStart(2, "0");
    const secs = String(appState.remainingSeconds % 60).padStart(2, "0");
    dom.sandboxTimer.textContent = `${mins}:${secs}`;
    dom.sandboxLives.textContent = `${appState.remainingLives} / ${appState.activeMission?.maxMistakes || 3}`;
}

function finishSandboxMission(isSuccess, message) {
    stopSandboxTimer();
    showSandboxToast(message, isSuccess ? "#10B981" : "#EF4444");

    if (isSuccess) {
        updateScoreUI(appState.activePoints);
        setModeratorQuote("correctAnswer");
        setModeratorAvatar("happy", 5000);
    } else {
        updateScoreUI(-appState.activePoints);
        setModeratorQuote("wrongAnswer");
        setModeratorAvatar("no", 5000);
        recordMistakeTopic(appState.activeMission);
    }

    appState.answeredCount++;
    updateCounterUI();

    setTimeout(() => {
        // Udržet happy/no postavičku po zbytek 5vteřinového intervalu
        if (dom.moderatorHoverBadge) dom.moderatorHoverBadge.classList.add("hidden");
        if (dom.moderatorAvatarBox) dom.moderatorAvatarBox.classList.remove("has-hint");
        appState.activeGame?.onMissionFinished?.();
        checkGameEnd();
    }, 1800);
}

function recordMistakeTopic(mission) {
    const cat = GAME_DATA.categories.find(c => c.id === mission.categoryId);
    appState.wrongTopics.push({
        categoryName: cat ? cat.name : mission.categoryId,
        topic: mission.topic
    });
}

/** Ukončí hru při vyčerpání 5 úkolů nebo když aktivní hra nemá další úkoly (force = vynucené ukončení). */
function checkGameEnd(force = false) {
    const noMoreTasks = !!appState.activeGame?.canContinue && !appState.activeGame.canContinue();
    if (force || noMoreTasks || appState.answeredCount >= appState.maxQuestions) {
        setModeratorQuote("gameOver");
        dom.finalScoreValue.textContent = appState.score;

        const errorCount = appState.wrongTopics.length;
        let verdict;
        let verdictColor;
        if (appState.score < 0 || errorCount >= 3) {
            verdict = "Žádná sláva 😒";
            verdictColor = "#d13438";
        } else if (errorCount === 0) {
            verdict = "Skvělý výkon!";
            verdictColor = "#107c10";
        } else if (errorCount === 1) {
            verdict = "Dobrý výkon 😊";
            verdictColor = "#107c10";
        } else {
            verdict = "To by ušlo. ";
            verdictColor = "#d83b01";
        }

        const finalVerdict = document.getElementById("final-verdict");
        if (finalVerdict) {
            finalVerdict.textContent = `${verdict} Zde je tvé vyhodnocení:`;
            finalVerdict.style.color = verdictColor;
        }

        if (errorCount === 0) {
            dom.evaluationRecommendations.innerHTML = "<p>Žádné chyby, žádná doporučení.</p>";
        } else {
            const uniqueItems = [];
            appState.wrongTopics.forEach(item => {
                const str = `Kategorie <strong>${item.categoryName}</strong> – zaměř se na téma: <em>${item.topic}</em>`;
                if (!uniqueItems.includes(str)) uniqueItems.push(str);
            });
            dom.evaluationRecommendations.innerHTML = `<ul>${uniqueItems.map(u => `<li>${u}</li>`).join("")}</ul>`;
        }

        dom.modalEvaluation.classList.remove("hidden");
    }
}

function updateScoreUI(delta) {
    appState.score += delta;
    dom.scoreDisplay.textContent = appState.score;
    dom.scoreDisplay.classList.remove("score-up", "score-down");
    if (delta > 0) dom.scoreDisplay.classList.add("score-up");
    if (delta < 0) dom.scoreDisplay.classList.add("score-down");
    setTimeout(() => dom.scoreDisplay.classList.remove("score-up", "score-down"), 450);
}

function updateCounterUI() {
    document.querySelectorAll("[data-game-counter]").forEach(el => {
        el.textContent = `${appState.answeredCount} / ${appState.maxQuestions}`;
    });
}

function setModeratorQuote(key) {
    const list = GAME_DATA.moderatorQuotes[key];
    if (!list || !list.length) return;
    const quote = list[Math.floor(Math.random() * list.length)];
    appState.lastModeratorQuote = quote;
    dom.speechBubble?.classList.remove("hint-active");
    dom.moderatorText.textContent = quote;
}

let toastTimer = null;
function showSandboxToast(text, borderColor) {
    dom.sandboxToast.textContent = text;
    dom.sandboxToast.style.borderLeftColor = borderColor;
    dom.sandboxToast.classList.remove("hidden");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => dom.sandboxToast.classList.add("hidden"), 4200);
}

/** Veřejné rozhraní jádra pro moduly her (riskuj.js, kolo_stesti.js, ...). */
window.EduaCore = {
    get state() { return appState; },
    showScreen,
    beginGame,
    getEligibleMissions,
    launchMission,
    checkGameEnd,
    setModeratorQuote,
    setModeratorAvatar,
    setModeratorText(text) {
        appState.lastModeratorQuote = text;
        dom.speechBubble?.classList.remove("hint-active");
        dom.moderatorText.textContent = text;
    }
};