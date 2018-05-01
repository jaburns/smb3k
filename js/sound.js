module.exports = () => {
    return {
        Initialize: hWnd => {/*nop*/},

        LoadSound: (path, channel, bits) => {
            console.log("Sound::LoadSound", path);
            // TODO implement
        },

        PlaySound: (id, loop) => {
            // TODO implement
        },

        StopSound: (id, pause) => {
            // TODO implement
        },

        StillPlaying: id => {
            // TODO implement
        },
    };
};