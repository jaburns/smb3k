import { pathToURL } from './utils';

const blobToByteArray = async blob => new Promise(resolve => {
    const reader = new FileReader();
    reader.addEventListener("loadend", () => 
        resolve(Array.from(new Uint8Array(reader.result))));
    reader.readAsArrayBuffer(blob);
});

const loadBinaryFile = async url =>
    fetch(url)
        .then(response => response.blob())
        .then(blobToByteArray);

const chompByte = bytes => bytes.shift();

const chompShort = bytes => {
    const a = bytes.shift(), b = bytes.shift();
    return a + 0x100*b;
};

const chompLong = bytes => {
    const a = bytes.shift(), b = bytes.shift();
    const c = bytes.shift(), d = bytes.shift();
    return a + 0x100*b + 0x10000*c + 0x1000000*d;
};

const chompString = bytes => {
    let len = chompShort(bytes);
    const result = [];
    while (len > 0) {
        result.push(bytes.shift());
        len--;
    }
    return result.map(x => String.fromCharCode(x)).join("");
};

module.exports = {
    LoadWorldList: worldList => {
        worldList(null, null, {ubound: 2});
        worldList(0, "WorldMap");
        worldList(1, "GrassyGrove");
        worldList(2, "FungiForest");
    },

    LoadSavedGame: savedData => {
        console.log("FileLoader::LoadSavedGame");
    },

    cwdLoadWorldData: async (data, path) => {
        const url = pathToURL(path);
        console.log("FileLoader::cwdLoadWorldData", url);
        const bytes = await loadBinaryFile(url);

        data.WorldName = chompString(bytes);
        chompShort(bytes);

        const levelCount = chompLong(bytes);
        data.LevelData(null, null, {ubound: levelCount - 1});
        chompLong(bytes);

        for (let i = 0; i < levelCount; ++i) {
            data.LevelData(i).LevelName = chompString(bytes);
            data.LevelData(i).Filename = chompString(bytes);
            data.LevelData(i).Background = chompString(bytes);
            data.LevelData(i).EnemySkin = chompString(bytes);
            data.LevelData(i).MusicFile = chompString(bytes);
            data.LevelData(i).TileFile = chompString(bytes);

            data.LevelData(i).TimeGiven = chompShort(bytes);
            data.LevelData(i).iScrollStyle = chompByte(bytes);
            data.LevelData(i).iScrollSpeed = chompShort(bytes);

            data.LevelData(i).dfCoin.xSrc = chompShort(bytes);
            data.LevelData(i).dfCoin.ySrc = chompShort(bytes);
            data.LevelData(i).dfBrick.xSrc = chompShort(bytes);
            data.LevelData(i).dfBrick.ySrc = chompShort(bytes);
            data.LevelData(i).dfVine.xSrc = chompShort(bytes);
            data.LevelData(i).dfVine.ySrc = chompShort(bytes);

            for (let j = 1; j <= 16; ++j) {
                data.LevelData(i).PipeDest(j).destLevel = chompShort(bytes);
                data.LevelData(i).PipeDest(j).destTag = chompByte(bytes);
                data.LevelData(i).PipeDest(j).destDir = chompByte(bytes);
            }
        }
    },

    loadMap: async (nodes, path) => {
        const url = pathToURL(path);
        console.log("FileLoader::loadMap", url);
        const bytes = await loadBinaryFile(url);

        const nodeUbound = chompByte(bytes);
        nodes(null, null, {ubound: nodeUbound});
        for (let i = 1; i <= nodeUbound; i++) {
            nodes(i).zxPos = chompShort(bytes);
            nodes(i).zyPos = chompShort(bytes);
            nodes(i).zupNode = chompByte(bytes);
            nodes(i).zdownNode = chompByte(bytes);
            nodes(i).zleftNode = chompByte(bytes);
            nodes(i).zrightNode = chompByte(bytes);
            let x = chompByte(bytes);
            if (x >= 128) {
                x -= 128;
                nodes(i).zpassThrough = true;
            }
            if (x > 0) {
                x -= 1;
                nodes(i).znodeTag = x;
            }
            nodes(i).znodeImage = chompByte(bytes);
            nodes(i).zentryPoint = chompByte(bytes);
            nodes(i).zexitWorld = chompByte(bytes);
        }
    },

    LevelLoadFromFile: async (level, path) => {
        const url = pathToURL(path);
        console.log("FileLoader::LevelLoadFromFile", url);
        const bytes = await loadBinaryFile(url);

        level.bSideWarp = chompByte(bytes) == 0xFF;
        level.iStartX = chompShort(bytes);
        level.iStartY = chompShort(bytes);
        level.iWidth = chompShort(bytes);
        level.iHeight = chompShort(bytes);
        level.iTime = chompShort(bytes);

        level.Column(null, null, {ubound: level.iWidth});
        for (let xx = 0; xx <= level.iWidth; xx++) {
            level.Column(xx).Row(null, null, {preserve: true, ubound: level.iHeight});
            for (let yy = 0; yy <= level.iHeight; yy++) {
                level.Column(xx).Row(yy).notEmpty = chompByte(bytes) == 0xFF;
                level.Column(xx).Row(yy).xSrc = chompByte(bytes);
                level.Column(xx).Row(yy).ySrc = chompByte(bytes);
                level.Column(xx).Row(yy).tag = chompByte(bytes);
                level.Column(xx).Row(yy).TileEnemy = chompByte(bytes);
            }
        }

        console.log(level);
    },

    LoadEnemySkinFile: async (enemySkin, path) => {
        const url = pathToURL(path);
        console.log("FileLoader::LoadEnemySkinFile", url);
        const bytes = await loadBinaryFile(url);

        enemySkin.Goomba = chompString(bytes);
        enemySkin.PiranaPlants = chompString(bytes);
        enemySkin.BuzzyBeetle = chompString(bytes);
        enemySkin.DumbKoopa = chompString(bytes);
        enemySkin.SmartKoopa = chompString(bytes);
        enemySkin.BumptyPenguin = chompString(bytes);
        enemySkin.Spiney = chompString(bytes);
        enemySkin.LittleBoo = chompString(bytes);
        enemySkin.RotoDisc = chompString(bytes);
        enemySkin.LavaBubble = chompString(bytes);
        enemySkin.Thwomp = chompString(bytes);
        enemySkin.DryBones = chompString(bytes);
        enemySkin.MovingPlatform = chompString(bytes);
        enemySkin.PowButton = chompString(bytes);
        enemySkin.FreeCheepCheep = chompString(bytes);
        enemySkin.BlockCheepCheep = chompString(bytes);
        enemySkin.Layer3Lava = chompString(bytes);
        enemySkin.TallCactus = chompString(bytes);
        enemySkin.EnemyWings = chompString(bytes);
        enemySkin.Bouncer = chompString(bytes);
        enemySkin.BulletBill = chompString(bytes);
        enemySkin.Bobomb = chompString(bytes);
        enemySkin.BOSS_Goomboss = chompString(bytes);
        enemySkin.Wiggler = chompString(bytes);
        enemySkin.PowerBlock = chompString(bytes);
        enemySkin.SavePoint = chompString(bytes);
    },

    LoadLevelTileset: async (tileSet, path) => {
        const url = pathToURL(path);
        console.log("FileLoader::LoadLevelTileset", url);
        const bytes = await loadBinaryFile(url);
    
        tileSet.bWidth = chompByte(bytes);
        tileSet.bHeight = chompByte(bytes);

        tileSet.Column(null, null, {ubound: tileSet.bWidth});
        for (let xx = 0; xx <= tileSet.bWidth; xx++) {
            tileSet.Column(xx).Row(null, null, {preserve: true, ubound: tileSet.bHeight});
            for (let yy = 0; yy <= tileSet.bHeight; yy++) {
                tileSet.Column(xx).Row(yy, chompByte(bytes));
            }
        }
    },
};