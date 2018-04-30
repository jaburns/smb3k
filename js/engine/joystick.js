module.exports = () => {
    return {
        Initialize: hWnd => {/*nop*/},

        get JoystickPresent() { return false; },
        ButtonDown: (button) => false,
        get XAxis() { return 0; },
        get YAxis() { return 0; },

        Update: () => {/*nop*/},
    };
};