window.new_DXGraphics = require('./graphics');
window.new_DXKeyboard = require('./keyboard');
window.new_DXJoystick = require('./joystick');
window.new_DXSound = require('./sound');
window.new_DXMusic = require('./music');

window.__fileLoader = require('./fileLoader');
window.False = false;
window.True = true;
window.vbRed = "#f00";
window.vbGreen = "#0f0";
window.vbBlue = "#00f";
window.vbWhite = "#fff";

window.App = { Path: "" };
window.Int = x => Math.floor(x) >> 0;
window.CSng = x => x;
window.CByte = x => Math.floor(x) >> 0;
window.CInt = x => Math.floor(x) >> 0;
window.CLng = x => Math.floor(x) >> 0;
window.CStr = x => x.toString();
window.Trim$ = x => x.toString().trim();
window.IIf = (a, b, c) => (a ? b : c);
window.Abs = Math.abs;
window.Sin = Math.sin;
window.Cos = Math.cos;
Math.Sqr = Math.sqrt;
window.Rnd = Math.random;
window.Sgn = Math.sign;
window.Mid$ = (s, a, b) => s.substr(a - 1, b);
window.Strings = { 
    Left$: (str, count) => str.substr(0, count),
    Right$: (str, count) => str.substr(str.length - count, count)
};
window.Len = x => x.toString().length;
window.UBound = arr => arr(null, null, null, true);
window.Randomize = () => {};
window.DoEvents = async () => new Promise(resolve => setTimeout(resolve, 50));
window.ConsoleLog = console.log.bind(console);

window.__intDiv = (a, b) => Math.floor(a / b) >> 0;

window.__makeArray = (lbound, ubound, makeElem) => {
    let arr = [];
    for (let i = lbound; i <= ubound; i++) {
        arr.push(makeElem());
    }

    return (lookup, write, redim, getUbound) => {
        if (getUbound === true) {
            return lbound + arr.length - 1;
        }
        else if (typeof redim === 'object') {
            const new_arr = [];

            for (let i = 0; i <= redim.ubound; i++) {
                new_arr.push(redim.preserve && i < arr.length 
                    ? arr[i]
                    : makeElem());
            }

            arr = new_arr;
        } 
        else if (typeof write !== 'undefined') {
            arr[lookup - lbound] = write;
        } 
        else {
            return arr[lookup - lbound];
        }
    };
};