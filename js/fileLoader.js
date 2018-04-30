module.exports = {
    LoadWorldList: worldList => {
        worldList(null, null, {count: 3});
        worldList(1, "WorldMap");
        worldList(2, "GrassyGrove");
        worldList(1, "FungiForest");
    },

    LoadSavedGame: savedData => {},
    LoadEnemySkinFile: (enemySkin, path) => {},
    LoadLevelTileset: (tileSet, path) => {},
    loadMap: (nodes, path) => {},
    cwdLoadWorldData: (data, path) => {},
    LoadLevelFromFile: (level, path) => {},
};