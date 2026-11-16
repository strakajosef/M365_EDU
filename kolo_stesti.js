/**
 * ============================================================================
 * HRA KOLO ŠTĚSTÍ (kolo_stesti.js)
 * ============================================================================
 * Každý segment kola je jeden nevybraný úkol pro úroveň hráče. Barva segmentu
 * odpovídá oblasti úkolu, uvnitř je ikona aplikace. Segmenty se rozmísťují náhodně.
 * Sílu roztočení udává doba držení tlačítka Roztoč (ukazatel nad tlačítkem).
 */
(function () {
    const MAX_HOLD_MS = 2000;
    const MODERATOR_DELAY_MS = 5000;

    const core = () => window.EduaCore;
    const state = {
        rotation: 0,
        spinning: false,
        holding: false,
        holdStart: 0,
        power: 0,
        rafId: null,
        segments: [],
        spinTimer: null,
        launchTimer: null
    };
    let els = {};

    const game = {
        id: "kolo-stesti",
        card: {
            title: "Kolo štěstí",
            tag: "Náhodný výběr oblasti",
            tagClass: "tag-active",
            description: "Roztoč kolo štěstí. Oblast, na které se zastaví, ti vybere náhodný úkol dle tvé úrovně. Sílu roztočení nastavíš tím, jak dlouho držíš tlačítko Roztoč.",
            buttonText: "Roztoč kolo štěstí",
            imageText: "IKONA HRY: KOLO ŠTĚSTÍ<br><small>(Náhodný úkol z kola)</small>",
            imageStyle: "background: linear-gradient(135deg, #f59e0b, #ef4444);",
			imageUrl: "kolo_stesti.png"
        },
        cssFile: "kolo_stesti.css",
        mountScreen,
        start,
        canContinue: () => getSegments().length > 0,
        onMissionFinished,
        cancel
    };

    GameRegistry.register(game);

    function mountScreen(parent, anchor) {
        const section = document.createElement("section");
        section.id = `screen-${game.id}`;
        section.className = "screen hidden kolo-screen";
        section.innerHTML = `
            <div class="jeopardy-header-info">
                <span>Odehráno úkolů: <strong data-game-counter>0 / 5</strong></span>
                <span data-level-badge class="difficulty-legend"></span>
            </div>
			<div class="kolo-banner">Kolo štěstí</div>
            <div class="kolo-coins"></div>
            <div class="kolo-wheel-box">
                <div class="kolo-pointer"></div>
                <div class="kolo-wheel"></div>
            </div>
            <div class="kolo-power">
                <div class="kolo-power-track"><div class="kolo-power-fill"></div></div>
                <span class="kolo-power-text">Síla: 0 %</span>
            </div>
            <button class="btn btn-primary kolo-spin-btn">Roztoč</button>
            <p class="kolo-hint">Drž tlačítko – čím déle, tím silnější roztočení.</p>
        `;
        parent.insertBefore(section, anchor);

        const coinsEl = section.querySelector(".kolo-coins");
        for (let i = 0; i < 14; i++) {
            const coin = document.createElement("span");
            coin.textContent = "🪙";
            coin.style.left = `${Math.random() * 100}%`;
            coin.style.top = `${Math.random() * 100}%`;
            coin.style.transform = `rotate(${Math.random() * 360}deg) scale(${0.6 + Math.random() * 0.8})`;
            coinsEl.appendChild(coin);
        }

        els = {
            wheel: section.querySelector(".kolo-wheel"),
            spinBtn: section.querySelector(".kolo-spin-btn"),
            powerFill: section.querySelector(".kolo-power-fill"),
            powerText: section.querySelector(".kolo-power-text")
        };

        els.spinBtn.addEventListener("pointerdown", onHoldStart);
        els.spinBtn.addEventListener("pointerup", onHoldEnd);
        els.spinBtn.addEventListener("pointercancel", onHoldEnd);
        els.spinBtn.addEventListener("keydown", e => {
            if ((e.code === "Space" || e.code === "Enter") && !e.repeat) {
                e.preventDefault();
                onHoldStart();
            }
        });
        els.spinBtn.addEventListener("keyup", e => {
            if (e.code === "Space" || e.code === "Enter") onHoldEnd();
        });
    }

    function start() {
        cancel();
        core().beginGame(game);
        renderWheel();
        updatePowerUI(0);
        core().showScreen(game.id);
        core().setModeratorText("Vítej u Kola štěstí! Roztoč kolo a nech štěstí rozhodnout, jaký úkol tě čeká.");
        core().setModeratorAvatar("norm");
    }

    function cancel() {
        clearTimeout(state.spinTimer);
        clearTimeout(state.launchTimer);
        cancelAnimationFrame(state.rafId);
        state.spinTimer = null;
        state.launchTimer = null;
        state.rafId = null;
        state.holding = false;
        state.spinning = false;
    }

    // Segmenty = nevybrané úkoly pro úroveň hráče, každý s oblastí (barva, ikona)
    function getSegments() {
        return GAME_DATA.categories.flatMap(c =>
            core().getEligibleMissions(c.id).map(mission => ({ mission, category: c }))
        );
    }

    function shuffle(arr) {
        for (let i = arr.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [arr[i], arr[j]] = [arr[j], arr[i]];
        }
        return arr;
    }

    function renderWheel() {
        state.segments = shuffle(getSegments());
        const n = state.segments.length;
        const span = n ? 360 / n : 360;

        els.wheel.style.background = n
            ? `conic-gradient(${state.segments.map((s, i) => `${s.category.color} ${i * span}deg ${(i + 1) * span}deg`).join(", ")})`
            : "#e2e8f0";
        els.wheel.innerHTML = state.segments.map((s, i) => `
            <div class="kolo-label-wrap" style="transform: rotate(${(i + 0.5) * span}deg)">
                <img class="kolo-icon" src="${s.category.iconPlaceholder}" alt="${s.category.name}">
            </div>
        `).join("");
        els.spinBtn.disabled = n === 0 || state.spinning;
    }

    function updatePowerUI(power) {
        els.powerFill.style.width = `${Math.round(power * 100)}%`;
        els.powerText.textContent = `Síla: ${Math.round(power * 100)} %`;
    }

    function onHoldStart() {
        if (state.spinning || state.holding || getSegments().length === 0) return;
        state.holding = true;
        state.holdStart = performance.now();
        trackPower();
    }

    function trackPower() {
        if (!state.holding) return;
        state.power = Math.min((performance.now() - state.holdStart) / MAX_HOLD_MS, 1);
        updatePowerUI(state.power);
        state.rafId = requestAnimationFrame(trackPower);
    }

    function onHoldEnd() {
        if (!state.holding) return;
        state.holding = false;
        cancelAnimationFrame(state.rafId);
        startSpin(Math.max(state.power, 0.15));
    }

    // Index segmentu, který je pod ukazatelem (nahoře) při dané rotaci kola
    function segmentIndexAt(rotation, n) {
        const span = 360 / n;
        const angle = ((360 - (rotation % 360)) % 360 + 360) % 360;
        return Math.floor(angle / span) % n;
    }

    function startSpin(power) {
        const segments = state.segments;
        const n = segments.length;
        if (n === 0 || state.spinning) return;

        state.spinning = true;
        els.spinBtn.disabled = true;
        core().setModeratorText("Roztáčím kolo… Uvidíme, kam padne!");

        // Síla určuje počet přeskočených výsečí (slabé stisknutí = jen pár polí) a délku točení
        const span = 360 / n;
        const maxSteps = Math.max(1, Math.round(power * n));
        const steps = 1 + Math.floor(Math.random() * maxSteps);
        const extraTurns = Math.floor(power * 2);
        const jitter = (Math.random() - 0.5) * span * 0.6;

        const current = state.rotation;
        const target = current + extraTurns * 360 + steps * span + jitter;
        const duration = 1.2 + 2.3 * power;
        state.rotation = target;
        els.wheel.style.transition = `transform ${duration}s cubic-bezier(0.15, 0.6, 0.1, 1)`;
        els.wheel.style.transform = `rotate(${target}deg)`;

        const winner = segments[segmentIndexAt(target, n)];
        state.spinTimer = setTimeout(() => onWheelStopped(winner), duration * 1000 + 50);
    }

    function onWheelStopped(segment) {
        state.spinTimer = null;
        core().setModeratorText(`Vytočil(a) jsi úkol z aplikace ${segment.category.name}: ${segment.mission.topic}. Hned se do toho dáme.`);
        core().setModeratorAvatar("ques");
        state.launchTimer = setTimeout(() => launchSegment(segment), MODERATOR_DELAY_MS);
    }

    function launchSegment(segment) {
        state.launchTimer = null;
        const points = GAME_DATA.config.difficulties[segment.mission.difficulty]?.points || 100;
        core().launchMission(segment.mission, points);
    }

    function onMissionFinished() {
        state.spinning = false;
        state.power = 0;
        updatePowerUI(0);
        renderWheel();
        core().showScreen(game.id);
    }
})();
