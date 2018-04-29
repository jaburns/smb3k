window.False = false;
window.True = true;
window.vbGreen = "#0f0";
window.vbBlue = "#00f";

window.new_DXMouse = require('./dxmouse');
window.new_DXGraphics = require('./dxgraphics');
window.new_DXKeyboard = require('./dxkeyboard');
window.new_DXJoystick = require('./dxjoystick');
window.new_DXSound = require('./dxsound');
window.new_DXMusic = require('./dxmusic');

window.App = { Path: "" };

window.__fileLoader = require('./fileLoader');

window.__intDiv = (a, b) => Math.floor(a / b) >> 0;

window.__makeArray = (lowerIndex, count, makeElem) => {
    let arr = [];
    for (let i = 0; i < count; i++) {
        arr.push(makeElem());
    }

    return (lookup, write, redim) => {
        if (typeof redim === 'object') {
            const new_arr = [];

            for (let i = 0; i < redim.count; i++) {
                if (redim.preserve && i < arr.length) {
                    new_arr.push(arr[i])
                } else {
                    new_arr.push(makeElem());
                }
            }

            arr = new_arr;
        } 
        else if (typeof write !== 'undefined') {
            arr[lookup - lowerIndex] = write;
        } 
        else {
            return arr[lookup - lowerIndex];
        }
    };
};