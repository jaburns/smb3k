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
        body: Vec<StatementLine>,
    },

    Empty,
}

#[derive(Debug)]
pub struct Expression {
    pub body: String,
}

#[derive(PartialEq, Eq, Debug)]
pub enum DoLoopKind {
    None,
    While,
    Until,
}

#[derive(Debug)]
pub enum StatementLine {
    Dim(VarDeclaration),

    Set {
        target_name: Expression,
        type_name: Option<String>,
    },

    ReDim {
        preserve: bool,
        target_name: Expression,
        new_ubound: Expression,
    },

    Assignment {
        to_name: Expression,
        value: Expression,
    },

    CallSub {
        name: Expression,
        args: Vec<Option<Expression>>,
    },

    SingleLineIf {
        condition: Expression,
        if_body: Box<StatementLine>,
        else_body: Box<StatementLine>,
    },

    BeginIf(Expression),
    ElseIf(Expression),
    Else,

    BeginWith(Expression),

    BeginFor {
        index: String,
        lower_bound: Expression,
        upper_bound: Expression,
        step: String,
    },

    DoLoop {
        kind: DoLoopKind,
        condition: Expression,
        is_end: bool,
    },

    BeginSelect(Expression),
    CaseLabel(Expression),

    EndBlock,

    ExitSub,
    ExitFunction,
    ExitLoop,

    Unknown(String),

    Empty,
}
