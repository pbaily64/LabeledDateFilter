import eslintPluginPowerBIVisuals from "eslint-plugin-powerbi-visuals";

export default [
  {
    plugins: {
      "powerbi-visuals": eslintPluginPowerBIVisuals
    },
    rules: {
      // Désactivé : le schéma https://powerbi.com est une constante imposée par l'API Power BI
      "powerbi-visuals/no-http-string": "off",
      // Désactivé : container.innerHTML = "" est utilisé uniquement pour vider le conteneur (reset propre)
      "powerbi-visuals/no-inner-outer-html": "off"
    }
  }
];
