// ===== Built-ins ===========================

const DoEvents = () => {};
const Console_Log = console.log.bind(console);

// ===== Classes =============================

const new_SomeClass = () => {
    let localField;

    const Class_Initialize = () => {
        localField = 123;
    };

    Class_Initialize();

    return {
        get ReadValue() {
            return localField;
        },

        SetValueTo: (newValue) => {
            localField = newValue;
        },

        CallMethod: (a, b) => {
            Console_Log(a, b, SOME_CONST, localField);
        },
    };
};

// ===== Modules =============================

// ----- main.bas -----

let someGlobal;

const module_main = (() => {
    let someNumber;

    return {
        __main__: () => {
            someNumber = 2;
            CallOtherSub();
            someNumber = someNumber + 10;
            someGlobal = someNumber;
            CallOtherSub();
        }
    };
})();

const __main__ = module_main.__main__;

// ----- OtherModule.bas -----

const SOME_CONST = 12;

const module_OtherModule = (() => {
    let classInstance = new_SomeClass();

    return {
        CallOtherSub: () => {
            classInstance.CallMethod(1, someGlobal);
            classInstance.SetValueTo(classInstance.ReadValue + 10000);
            classInstance.CallMethod(0, 0);
        }
    };
})();

const CallOtherSub = module_OtherModule.CallOtherSub;

// ===========================================

__main__();