module.exports = () => {
    let _keys = {};

    return {
        Key_J: 74,
        Key_F9: 57,
        Key_F10: 48,
        Key_L: 76,
        Key_Q: 81,
        Key_Z: 90,
        Key_LSHIFT: 16,
        Key_SPACE: 32,
        Key_X: 88,
        Key_BACKSPACE: 8,
        Key_P: 80,
        Key_LEFT: 37,
        Key_RIGHT: 39,
        Key_DOWN: 40,
        Key_UP: 38,
        Key_ESCAPE: 27,

        Initialize: () => {
            document.onkeydown = e => _keys[e.keyCode] = true;
            document.onkeyup   = e => delete _keys[e.keyCode];
        },

        IsDown: keyCode => _keys[keyCode] === true,
    };
};