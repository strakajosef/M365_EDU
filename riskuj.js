/**
 * ============================================================================
 * HRA RISKUJ! (riskuj.js)
 * ============================================================================
 * Herní plán se 3 náhodnými oblastmi. Po každém odehraném poli se doplní nové
 * náhodné oblasti, dokud je k dispozici něco nevybraného.
 * Jádro (sandbox, skóre, vyhodnocení) je v app.js přes window.EduaCore.
 */
(function () {
    const BOARD_CATEGORY_COUNT = 3;
    const ROW_POINTS = [100, 200, 300, 400];

    const core = () => window.EduaCore;
    const state = { pickedCategoryIds: [], boardCategoryIds: [], activeCellBtn: null };
    let boardEl = null;

    const game = {
        id: "riskuj",
        card: {
            title: "Riskuj",
            tag: "Hlavní soutěž na body",
            tagClass: "tag-active",
            description: "Vyber si políčko na herním plánu. Vylosuje se ti náhodný praktický úkol dle tvé úrovně. Splň ho přímo v simulátoru aplikace bez nápovědy v časovém limitu!",
            buttonText: "Spustit soutěž Riskuj!",
            imageText: "IKONA HRY: M365 RISKUJ!<br>",
            imageStyle: "background: linear-gradient(135deg, #0078D4, #464EB8);",
			imageUrl: "riskuj.png"
        },
        mountScreen,
        start,
        canContinue,
        onMissionFinished
    };

    GameRegistry.register(game);

    function mountScreen(parent, anchor) {
        const section = document.createElement("section");
        section.id = `screen-${game.id}`;
        section.className = "screen hidden jeopardy-bg-placeholder";
        section.innerHTML = `
            <div class="jeopardy-header-info">
                <span>Odehráno polí: <strong data-game-counter>0 / 5</strong></span>
                <span data-level-badge class="difficulty-legend">
                    Zvolená úroveň hráče: <b>0 - Základní</b>
                </span>
            </div>
            <div id="jeopardy-board" class="jeopardy-board"></div>
        `;
        parent.insertBefore(section, anchor);
        boardEl = section.querySelector("#jeopardy-board");
    }

    function start() {
        state.pickedCategoryIds = [];
        state.boardCategoryIds = [];
        state.activeCellBtn = null;

        core().beginGame(game);
        renderBoard();
        core().showScreen(game.id);
        core().setModeratorQuote("gameStart");
        core().setModeratorAvatar("norm");
    }

    function canContinue() {
        return GAME_DATA.categories.some(c => core().getEligibleMissions(c.id).length > 0);
    }

    // Přednost: oblasti, ze kterých ještě nic nevybral a nejsou na aktuální ploše;
    // poté oblasti, ze kterých ještě nic nevybral, ale jsou na ploše;
    // nakonec (když už žádné nezbydou) náhodné oblasti, ze kterých již něco vybral.
    function pickBoardCategories() {
        const available = GAME_DATA.categories.filter(c => core().getEligibleMissions(c.id).length > 0);
        const shuffle = arr => arr.sort(() => Math.random() - 0.5);
        const previous = state.boardCategoryIds;
        const picked = state.pickedCategoryIds;

        const fresh = shuffle(available.filter(c => !picked.includes(c.id) && !previous.includes(c.id)));
        const stillOnBoard = shuffle(available.filter(c => !picked.includes(c.id) && previous.includes(c.id)));
        const touched = shuffle(available.filter(c => picked.includes(c.id)));

        return [...fresh, ...stillOnBoard, ...touched]
            .slice(0, BOARD_CATEGORY_COUNT)
            .map(c => c.id);
    }

    function renderBoard() {
        boardEl.innerHTML = "";
        state.boardCategoryIds = pickBoardCategories();

        state.boardCategoryIds.forEach(catId => {
            const cat = GAME_DATA.categories.find(c => c.id === catId);
            const eligiblePool = core().getEligibleMissions(catId);

            if (!cat || eligiblePool.length === 0) return;

            const col = document.createElement("div");
            col.className = "jeopardy-column";

            col.innerHTML = `
                <div class="category-header" style="border-top-color: ${cat.color}">
                    <img src="${cat.iconPlaceholder}" alt="${cat.name}">
                    <h3>${cat.name}</h3>
                </div>
            `;

            // Pole nad počet dostupných úkolů zůstanou prázdná a nerozkliknutelná
            ROW_POINTS.forEach((pts, index) => {
                const btn = document.createElement("button");
                btn.className = "board-cell";

                if (index < eligiblePool.length) {
                    btn.innerHTML = `
                        <span>${pts}</span>
                        <small>🛠️ Praktický úkol</small>
                    `;
                    btn.addEventListener("click", () => handleCellClick(cat, pts, btn));
                } else {
                    btn.innerHTML = "";
                    btn.disabled = true;
                    btn.classList.add("board-cell-empty");
                }
                col.appendChild(btn);
            });

            boardEl.appendChild(col);
        });
    }

    function handleCellClick(category, points, cellBtn) {
        const pool = core().getEligibleMissions(category.id);
        if (pool.length === 0) return;

        const mission = pool[Math.floor(Math.random() * pool.length)];
        if (!state.pickedCategoryIds.includes(category.id)) {
            state.pickedCategoryIds.push(category.id);
        }
        state.activeCellBtn = cellBtn;
        core().launchMission(mission, points);
    }

    function onMissionFinished() {
        if (state.activeCellBtn) {
            state.activeCellBtn.disabled = true;
            state.activeCellBtn.classList.add("played");
            state.activeCellBtn = null;
        }
        renderBoard();
        core().showScreen(game.id);
    }
})();
