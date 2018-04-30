window.False = false;
window.True = true;
window.vbGreen = "#0f0";
window.vbBlue = "#00f";

window.new_DXGraphics = require('./engine/graphics');
window.new_DXKeyboard = require('./engine/keyboard');
window.new_DXJoystick = require('./engine/joystick');
window.new_DXSound = require('./engine/sound');
window.new_DXMusic = require('./engine/music');
window.__fileLoader = require('./fileLoader');

window.App = { Path: "" };
window.CSng = x => x;
window.CLng = x => x >> 0;
window.CStr = x => x.toString();
window.Trim$ = x => x.toString().trim();
window.Abs = Math.abs;
window.Sin = Math.sin;
window.Cos = Math.cos;
window.Strings = { 
    Left$: (str, count) => str.substr(0, count),
    Right$: (str, count) => str.substr(str.length - count, count)
};
window.UBound = arr => arr(null, null, null, true);
window.DoEvents = async () => new Promise(resolve => setTimeout(resolve, 10));

window.__intDiv = (a, b) => Math.floor(a / b) >> 0;

window.__makeArray = (lowerIndex, count, makeElem) => {
    let arr = [];
    for (let i = 0; i < count; i++) {
        arr.push(makeElem());
    }

    return (lookup, write, redim, getUbound) => {
        if (getUbound === true) {
            return lowerIndex + arr.length - 1;
        }
        else if (typeof redim === 'object') {
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