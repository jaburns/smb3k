module.exports = {
    LoadWorldList: worldList => {
        worldList(null, null, {count: 3});
        worldList(1, "WorldMap");
        worldList(2, "GrassyGrove");
        worldList(1, "FungiForest");
    },

    LoadSavedGame: savedData => {
        // TODO implement
    },

    LoadEnemySkinFile: async (enemySkin, path) => {
        // TODO implement
    },

    LoadLevelTileset: async (tileSet, path) => {
        // TODO implement
    },

    loadMap: async (nodes, path) => {
        // TODO implement
    },

    cwdLoadWorldData: async (data, path) => {
        // TODO implement
    },

    LoadLevelFromFile: (level, path) => {},
};