
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

    Attribute { name: String, value: String },

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
pub struct StatementBlock {
    pub contents: String,
}
