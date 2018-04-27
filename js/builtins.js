window.new_DXMouse = () => {
    return {};
};

window.new_DXGraphics = () => {
    return {};
};

window.new_DXKeyboard = () => {
    return {};
};

window.new_DXJoystick = () => {
    return {};
};

window.new_DXSound = () => {
    return {};
};

window.new_DXMusic = () => {
    return {};
};

window.False = false;
window.True = true;
window.vbGreen = "#0f0";
window.vbBlue = "#00f";

window.__makeArray = (lowerIndex, count, makeElem) => {
    let arr = [];
    for (let i = 0; i < count; i++) {
        arr.push(makeElem());
    }

    return (lookup, write, redim) => {
        if (typeof redim === 'object') {
            // "ReDim Preserve arr(LocalVar + 1)"
            //arr(null, null, {preserve: true, count: (LocalVar + 1)});
        } else if (typeof write !== 'undefined') {
            arr[lookup - lowerIndex] = write;
        } else {
            return arr[lookup - lowerIndex];
        }
    };
};