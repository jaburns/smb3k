
const new_DXMouse = () => {
    return {};
};

const new_DXGraphics = () => {
    return {};
};

const new_DXKeyboard = () => {
    return {};
};

const new_DXJoystick = () => {
    return {};
};

const new_DXSound = () => {
    return {};
};

const new_DXMusic = () => {
    return {};
};

const __makeArray = (lowerIndex, count, makeElem) => {
    let arr = [];
    for (let i = 0; i < count; i++) {
        arr.push(makeElem());
    }

    return (lookup, write, redim) => {
        if (typeof redim === 'object') {
            // TODO implement
        } else if (typeof write !== 'undefined') {
            arr[lookup - lowerIndex] = write;
        } else {
            return arr[lookup - lowerIndex];
        }
    };
};

console.log("Hello world");

module.exports = __makeArray;

/*
let arr = __makeArray(1, 5, () => (0));
// "arr(2)" in expression
arr(2);  
// "arr(2) = 10"
arr(2, 10);
// "ReDim Preserve arr(LocalVar + 1)"
arr(null, null, {preserve: true, count: (LocalVar + 1)});
*/

const False = false;
const True = true;