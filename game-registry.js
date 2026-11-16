/**
 * ============================================================================
 * REGISTR HER (game-registry.js)
 * ============================================================================
 * Hra se do hlavního menu připojí voláním GameRegistry.register({...}) ve svém souboru.
 * Odebrání hry = smazání jejího <script> tagu v index.html (případně jejího souboru).
 *
 * Definice hry:
 *   id                  – jedinečné id; obrazovka hry má id "screen-<id>"
 *   card                – { title, tag, tagClass, description, buttonText, imageText, imageStyle,
 *                          imageUrl (obrázek místo textu), comingSoon (tlačítko bez akce) }
 *   cssFile             – (volitelné) stylesheet hry, připojí se do <head>
 *   mountScreen(parent, anchor) – vytvoří obrazovku hry a vloží ji před anchor
 *   start()             – spustí hru (volá se po kliknutí na kartu v menu)
 *   canContinue()       – (volitelné) false = nejsou k dispozici další úkoly, hra skončí
 *   onMissionFinished() – (volitelné) volá se po splnění nebo nesplnění úkolu
 *   cancel()            – (volitelné) zruší časovače hry při návratu do menu
 */
window.GameRegistry = (function () {
    const games = [];

    function register(game) {
        if (!game || !game.id || typeof game.start !== "function") {
            console.error("GameRegistry: neplatná definice hry", game);
            return;
        }
        if (game.cssFile && !document.getElementById(`game-css-${game.id}`)) {
            const link = document.createElement("link");
            link.rel = "stylesheet";
            link.href = game.cssFile;
            link.id = `game-css-${game.id}`;
            document.head.appendChild(link);
        }
        games.push(game);
    }

    function createCard(game) {
        const c = game.card || {};
        const article = document.createElement("article");
        article.className = "game-card";
        article.id = `card-${game.id}`;
        const imageStyle = c.imageUrl
            ? `background: #fff url('${c.imageUrl}') center / contain no-repeat;`
            : (c.imageStyle || "");
        article.innerHTML = `
            <div class="game-card-image placeholder-box" style="${imageStyle}">
                <span>${c.imageUrl ? "" : (c.imageText || "")}</span>
            </div>
            <div class="game-card-body">
                <span class="tag ${c.tagClass || "tag-active"}">${c.tag || ""}</span>
                <h3>${c.title || game.id}</h3>
                <p>${c.description || ""}</p>
                <button class="btn btn-primary btn-full">${c.buttonText || "Spustit"}</button>
            </div>
        `;
        const button = article.querySelector("button");
        if (c.comingSoon) {
            button.disabled = true;
        } else {
            button.addEventListener("click", () => game.start());
        }
        return article;
    }

    /**
     * Vytvoří obrazovky všech registrovaných her a karty v menu.
     * grid – element .games-grid; screenAnchor – obrazovka, před kterou se hry vloží.
     */
    function mountAll({ grid, screenAnchor }) {
        const parent = screenAnchor.parentNode;
        const freeCard = grid.querySelector("#card-start-freesandbox");

        games.forEach(game => {
            game.mountScreen?.(parent, screenAnchor);
            grid.insertBefore(createCard(game), freeCard);
        });
    }

    return { register, mountAll };
})();
