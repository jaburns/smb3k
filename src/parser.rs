use pom::DataInput;
use pom::parser::*;

#[derive(Debug)]
pub enum AccessLevel {
    Public, 
    Private,
}

#[derive(Debug)]
pub struct FunctionParam {
}

#[derive(Debug)]
pub struct TypeDeclField {
    pub name: String,
    pub type_name: String,
}

#[derive(Debug)]
pub enum FunctionKind {
    Sub, 
    Function, 
    PropertyGet, 
    PropertySet,
}

#[derive(Debug)]
pub enum TopLevelBlock {
    ClassMarker,

    Attribute { 
        name: String, 
        value: String,
    },

    OptionExplicit,

    TopField {
        access_level: AccessLevel,
        is_const: bool, /// TODO not looking for "Const"
        name: String,
        type_name: String,
        is_new: bool,
    },

    Type {
        access_level: AccessLevel,
        name: String,
        fields: Vec<TypeDeclField>
    },

    /// TODO match enum
    Enum {
        access_level: AccessLevel,
        name: String,
        values: Vec<String>
    },

    Function {
        access_level: AccessLevel,
        kind: FunctionKind,
        name: String,
        params: Vec<FunctionParam>,
        return_type: String,
        body: Vec<StatementBlock>,
    },

    DebugDump { rest: String },

    Empty,
}

#[derive(Debug)]
pub struct StatementBlock {
    pub contents: String,
}


fn space() -> Parser<'static, u8, ()> {
	one_of(b" \t\r\n").repeat(0..).discard()
}

fn word() -> Parser<'static, u8, String> {
    none_of(b" \t\r\n`()").repeat(0..).convert(String::from_utf8)
}

fn string() -> Parser<'static, u8, String> {
    let inner = sym(b'"') * none_of(b"\"").repeat(0..) - sym(b'"');
    inner.convert(String::from_utf8)
}

fn access_level() -> Parser<'static, u8, Option<AccessLevel>> {
    word().map(|w| match w.as_str() {
        "Public"  => Some(AccessLevel::Public),
        "Private" => Some(AccessLevel::Private),
        _         => None
    })
}

fn attribute_decl() -> Parser<'static, u8, TopLevelBlock> {
    let matches = seq(b"Attribute") * space() * word() - space() - sym(b'=') - space() + (string() | word());
    matches.map(|(n, v)| TopLevelBlock::Attribute { name: n, value: v })
}

fn option_decl() -> Parser<'static, u8, TopLevelBlock> {
    seq(b"Option Explicit").map(|_| TopLevelBlock::OptionExplicit)
}

fn type_decl_field() -> Parser<'static, u8, TypeDeclField> {
    let matches = word() - space() - seq(b"As") - space() + word();
    matches.map(|(n, t)| TypeDeclField { name: n, type_name: t })
}

fn type_decl() -> Parser<'static, u8, TopLevelBlock> {
    let begin = access_level() - space() - seq(b"Type") - space() + word() - sym(b'`');
    let members = list(type_decl_field(), sym(b'`'));
    let obj = begin + members - seq(b"`End Type");
    obj.map(|((a, n), f)| TopLevelBlock::Type { access_level: a.unwrap(), name: n, fields: f })
}

fn top_field_decl() -> Parser<'static, u8, TopLevelBlock> { 
    let access_and_name = access_level() - space() + word() - space() - seq(b"As") - space();
    let maybe_new_and_type = (seq(b"New") - space()).opt() + word();

    (access_and_name + maybe_new_and_type).map(|((a,n),(m,t))| {
         TopLevelBlock::TopField { access_level: a.unwrap(), name: n, type_name: t, is_new: m.is_some(), is_const: false }
    })
}

fn function_kind_keyword() -> Parser<'static, u8, Option<FunctionKind>> {
    ((seq(b"Property") - space()).opt() * word()).map(|w| match w.as_str() {
        "Function" => Some(FunctionKind::Function),
        "Sub"      => Some(FunctionKind::Sub),
        "Get"      => Some(FunctionKind::PropertyGet),
        "Set"      => Some(FunctionKind::PropertySet),
        _          => None
    })
}

fn function_param_list() -> Parser<'static, u8, Vec<FunctionParam>> {
    let param = none_of(b",)").repeat(0..);
    let params = list(param, sym(b',') * space());
    let inner = sym(b'(') * space() * params.opt() - sym(b')');
    inner.map(|_| Vec::new())
}

fn function_decl() -> Parser<'static, u8, TopLevelBlock> { 
    let access_and_kind = access_level() - space() + function_kind_keyword() - space();
    let name_and_params = word() - space() + function_param_list() - space();
    let maybe_return_type = ((seq(b"As") - space()) * word() - space()).opt();

    let header = (access_and_kind + name_and_params) + maybe_return_type - sym(b'`');
    let body_lines = list(statement(), sym(b'`'));
    let obj = header + body_lines - sym(b'`') - end_function();

    obj.map(|((((a,k),(n,p)),r),b)| {
        TopLevelBlock::Function { access_level: a.unwrap(), kind: k.unwrap(), name: n, params: p, return_type: r.unwrap_or(String::new()), body: b }
    })
}

fn end_function() -> Parser<'static, u8, ()> {
    (seq(b"End Sub") | seq(b"End Function") | seq(b"End Property")).discard()
}

fn ignored_header_decl() -> Parser<'static, u8, TopLevelBlock> { 
    (seq(b"BEGIN") * none_of(b"E").repeat(0..) * seq(b"END")).map(|_| TopLevelBlock::Empty)
}

fn class_marker_decl() -> Parser<'static, u8, TopLevelBlock> { 
    seq(b"VERSION 1.0 CLASS").map(|_| TopLevelBlock::ClassMarker)
}

fn debug_consume_rest() -> Parser<'static, u8, TopLevelBlock> { 
    none_of(b"").repeat(0..).convert(String::from_utf8).map(|x| TopLevelBlock::DebugDump { rest: x })
}

fn statement() -> Parser<'static, u8, StatementBlock> { 
    !end_function() * none_of(b"`").repeat(0..).convert(String::from_utf8).map(|x| StatementBlock { contents: x })
}

fn top_level_block() -> Parser<'static, u8, TopLevelBlock> {
    option_decl() | attribute_decl() | type_decl() | top_field_decl() | function_decl() 
        | ignored_header_decl() | class_marker_decl() | debug_consume_rest()
}

fn do_parse() -> Parser<'static, u8, Vec<TopLevelBlock>> {
    list(top_level_block(), sym(b'`'))
}

fn without_comments(line: &str) -> &str {
    match line.find("'") {
        Some(pos) => &line[..pos],
        None => line
    }
}

pub fn parse_module(contents: &str) -> Vec<TopLevelBlock> {
    let trimmed_lines: Vec<String> = contents.lines()
        .map(|x| { without_comments(x).trim() })
        .filter_map(|x| { if x.len() > 0 { Some(String::from(x)) } else { None }})
        .collect();

    let tick_seperated_contents = trimmed_lines.join("`");
    let mut input = DataInput::new(tick_seperated_contents.as_bytes());

    do_parse().parse(&mut input).unwrap()
}