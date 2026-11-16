/**
 * ============================================================================
 * HRA UTEČ Z PRÁCE (utec_z_prace.js)
 * ============================================================================
 * Zástupná karta hry – zatím bez odkazu a bez herní logiky.
 * Odebrání hry = smazání jejího <script> tagu v index.html a tohoto souboru.
 */
(function () {
    const game = {
        id: "utec-z-prace",
        card: {
            title: "Uteč z práce",
            tag: "Připravujeme",
            tagClass: "tag-active",
            description: "Najdi cestu z kanceláře dřív, než tě chytí šéf. Hra je ve výstavbě.",
            buttonText: "Ve výstavbě",
            imageUrl: "ve_vystavbe.png",
            comingSoon: true
        },
        // Hra nemá zatím žádnou obrazovku ani logiku; start se nikdy nevolá (tlačítko je neaktivní)
        start() {}
    };

    GameRegistry.register(game);
})();
