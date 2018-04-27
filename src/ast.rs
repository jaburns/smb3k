#[derive(PartialEq, Eq, Debug)]
pub enum AccessLevel {
    Public,
    Private,
}

#[derive(PartialEq, Eq, Debug)]
pub enum FunctionKind {
    Sub,
    Function,
    PropertyGet,
}

#[derive(PartialEq, Eq, Debug)]
pub enum VarKind {
    Standard,
    DynamicArray,
    RangeArray(i32, i32),
    AutoInstantiate,
}

#[derive(Debug)]
pub struct FunctionParam {
    pub name: String,
    pub type_name: String,
    pub default_value: Option<String>,
}

#[derive(Debug)]
pub struct VarDeclaration {
    pub name: String,
    pub type_name: String,
    pub kind: VarKind,
}

#[derive(Debug)]
pub struct TypeDeclaration {
    pub access_level: AccessLevel,
    pub name: String,
    pub fields: Vec<VarDeclaration>,
}

#[derive(Debug)]
pub struct Module {
    pub name: String,
    pub is_class: bool,
    pub contents: Vec<TopLevelBlock>,
}

#[derive(Debug)]
pub enum TopLevelBlock {
    ClassMarker,

    Attribute {
        name: String,
        value: String,
    },

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

    Type(TypeDeclaration),

    Enum {
        access_level: AccessLevel,
        name: String,
        values: Vec<(String, Option<i32>)>,
    },

    Function {
        access_level: AccessLevel,
        kind: FunctionKind,
        is_async: bool,
        name: String,
        params: Vec<FunctionParam>,
        return_type: String,
        body: Vec<StatementBlock>,
    },

    Empty,
}

#[derive(Debug)]
pub struct Expression {
    pub body: String,
}

#[allow(dead_code)]
#[derive(PartialEq, Eq, Debug)]
pub enum DoLoopKind {
    While,
    Until,
}

#[allow(dead_code)]
#[derive(Debug)]
pub enum StatementBlock {
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
        args: Vec<Expression>,
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

    Unknown(Expression),

    Empty,
}
