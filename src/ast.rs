#[derive(Debug)]
pub enum AccessLevel {
    Public,
    Private,
}

#[derive(Debug)]
pub struct FunctionParam {}

#[derive(Debug)]
pub enum FunctionKind {
    Sub,
    Function,
    PropertyGet,
    PropertySet,
}

#[derive(Debug)]
pub enum VarKind {
    Standard,
    DynamicArray,
    RangeArray(i32, i32),
    AutoInstantiate,
}

#[derive(Debug)]
pub struct VarDeclaration {
    pub name: String,
    pub type_name: String,
    pub kind: VarKind,
}

#[derive(Debug)]
pub enum TopLevelBlock {
    ClassMarker,

    Attribute {
        name: String,
        value: String,
    },

    OptionExplicit,

    Field {
        access_level: AccessLevel,
        declaration: VarDeclaration,
    },

    Constant {
        access_level: AccessLevel,
        name: String,
        type_name: String,
        value: String,
    },

    Type {
        access_level: AccessLevel,
        name: String,
        fields: Vec<VarDeclaration>,
    },

    Enum {
        access_level: AccessLevel,
        name: String,
        values: Vec<(String, i32)>,
    },

    Function {
        access_level: AccessLevel,
        kind: FunctionKind,
        name: String,
        params: Vec<FunctionParam>,
        return_type: String,
        body: Vec<StatementBlock>,
    },

    Empty,
}

#[derive(Debug)]
pub enum Expression {
}

#[derive(Debug)]
pub struct Argument {}

#[derive(Debug)]
pub enum DoLoopKind {
    While,
    Until,
}

#[derive(Debug)]
pub enum StatementBlock {
    OnError,

    Label {
        name: String,
    },

    GoTo {
        label_name: String,
    },

    Dim {
        declaration: VarDeclaration,
    },

    ReDim {
        preserve: bool,
        target_name: String,
        new_size: Expression, // int
    },

    Assignment {
        to_name: String,
        value: Expression, // typeof(lhs)
    },

    CallSub {
        name: String,
        args: Vec<Argument>,
    },

    IfBlock {
        condition: Expression, // bool
        main_body: Vec<StatementBlock>,
        else_body: Vec<StatementBlock>,
    },

    WithBlock {
        target_name: String,
        body: Vec<StatementBlock>,
    },

    ForLoop {
        index: String,
        lower_bound: Expression, // int
        upper_bound: Expression, // int
        step: i32,
        body: Vec<StatementBlock>,
    },

    DoLoop {
        condition: Expression, // bool
        kind: DoLoopKind,
        eval_at_top: bool,
        body: Vec<StatementBlock>,
    },

    FileOperation,

    Unknown {
        source: String,
    },
}
