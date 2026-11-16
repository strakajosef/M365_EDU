/**
 * ============================================================================
 * HRA JEDNORUKÝ BANDITA (jednoruky_bandita.js)
 * ============================================================================
 * Zástupná karta hry – zatím bez odkazu a bez herní logiky.
 * Odebrání hry = smazání jejího <script> tagu v index.html a tohoto souboru.
 */
(function () {
    const game = {
        id: "jednoruky-bandita",
        card: {
            title: "Jednoruký bandita",
            tag: "Připravujeme",
            tagClass: "tag-active",
            description: "Zatáhni za páku a uvidíš, co ti padne. Hra je ve výstavbě.",
            buttonText: "Ve výstavbě",
            imageUrl: "ve_vystavbe.png",
            comingSoon: true
        },
        // Hra nemá zatím žádnou obrazovku ani logiku; start se nikdy nevolá (tlačítko je neaktivní)
        start() {}
    };

    GameRegistry.register(game);
})();
