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