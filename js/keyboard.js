module.exports = () => {
    let _keys = {};

    return {
        Key_J: -1,//74,
        Key_F9: -1,//57,
        Key_F10: -1,//48,
        Key_L: -1,//76,
        Key_Q: -1, //81,
        Key_Z: 90,
        Key_LSHIFT: 16,
        Key_SPACE: 32,
        Key_X: 88,
        Key_BACKSPACE: -1,//8,
        Key_P: 80,
        Key_LEFT: 37,
        Key_RIGHT: 39,
        Key_DOWN: 40,
        Key_UP: 38,
        Key_ESCAPE: -1,//27,

        Initialize: () => {
            document.onkeydown = e => _keys[e.keyCode] = true;
            document.onkeyup   = e => delete _keys[e.keyCode];
        },

        IsDown: keyCode => _keys[keyCode] === true,
    };
};