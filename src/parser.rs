use ast::*;
use pom::parser::*;
use pom::DataInput;

enum BlockParseResult {
    Block(TopLevelBlock),
    ParseFail(String),
    EOF,
}

fn space() -> Parser<'static, u8, ()> {
    one_of(b" \t\r\n").repeat(0..).discard()
}

fn word() -> Parser<'static, u8, String> {
    none_of(b" \t\r\n`()")
        .repeat(1..)
        .convert(String::from_utf8)
}

fn string() -> Parser<'static, u8, String> {
    let inner = sym(b'"') * none_of(b"\"").repeat(0..) - sym(b'"');
    inner.convert(String::from_utf8)
}

fn integer() -> Parser<'static, u8, i32> {
    let integer = one_of(b"123456789") - one_of(b"0123456789").repeat(0..) | sym(b'0');
    let number = sym(b'-').opt() + integer;
    number
        .collect()
        .convert(String::from_utf8)
        .map(|s| s.parse::<i32>().unwrap())
}

fn access_level() -> Parser<'static, u8, Option<AccessLevel>> {
    word().map(|w| match w.as_str() {
        "Public" => Some(AccessLevel::Public),
        "Private" => Some(AccessLevel::Private),
        _ => None,
    })
}

fn attribute_decl() -> Parser<'static, u8, TopLevelBlock> {
    let matched =
        seq(b"Attribute") * space() * word() - space() - sym(b'=') - space() + (string() | word());
    matched.map(|(n, v)| TopLevelBlock::Attribute { name: n, value: v })
}

fn array_inner_range() -> Parser<'static, u8, VarKind> {
    let inner = integer() - space() - seq(b"To") - space() + integer() - space();
    inner.map(|(a, b)| VarKind::RangeArray(a, b))
}

fn array_inner_single() -> Parser<'static, u8, VarKind> {
    let inner = integer() - space();
    inner.map(|a| VarKind::RangeArray(0, a))
}

fn maybe_array() -> Parser<'static, u8, Option<VarKind>> {
    let maybe_array_inner = (array_inner_range() | array_inner_single()).opt();
    let matched = (sym(b'(') * space() * maybe_array_inner - sym(b')')).opt() - space();

    matched.map(|outer| outer.map(|inner| inner.unwrap_or(VarKind::DynamicArray)))
}

fn var_decl() -> Parser<'static, u8, VarDeclaration> {
    let name_and_maybe_range = word() + maybe_array();
    let maybe_new_and_type_name = seq(b"As") * space() * (seq(b"New") - space()).opt() + word();
    let matched = name_and_maybe_range + maybe_new_and_type_name;

    matched.map(|((n, r), (new, t))| {
        let kind = match new {
            Some(_) => VarKind::AutoInstantiate,
            None => r.unwrap_or(VarKind::Standard),
        };

        VarDeclaration {
            name: n,
            type_name: t,
            kind: kind,
        }
    })
}

fn type_decl() -> Parser<'static, u8, TopLevelBlock> {
    let begin = access_level() - space() - seq(b"Type") - space() + word() - sym(b'`');
    let members = list(var_decl(), sym(b'`'));
    let matched = begin + members - seq(b"`End Type");

    matched.map(|((a, n), f)| {
        TopLevelBlock::Type(TypeDeclaration {
            access_level: a.unwrap(),
            name: n,
            fields: f,
        })
    })
}

fn enum_decl_member() -> Parser<'static, u8, (String, Option<i32>)> {
    let maybe_value = (sym(b'=') * space() * integer()).opt();
    !seq(b"End Enum") * word() - space() + maybe_value
}

fn enum_decl() -> Parser<'static, u8, TopLevelBlock> {
    let begin = access_level() - space() - seq(b"Enum") - space() + word() - sym(b'`');
    let members = list(enum_decl_member(), sym(b'`'));
    let matched = begin + members - seq(b"`End Enum");

    matched.map(|((a, n), v)| TopLevelBlock::Enum {
        access_level: a.unwrap(),
        name: n,
        values: v,
    })
}

fn const_decl() -> Parser<'static, u8, TopLevelBlock> {
    let access_and_name = access_level() - space() - seq(b"Const") - space() + word() - space();
    let type_and_value = seq(b"As") * space() * word() - space() - sym(b'=') - space() + word();
    let matched = access_and_name + type_and_value;

    matched.map(|((a, n), (t, v))| TopLevelBlock::Constant {
        access_level: a.unwrap(),
        name: n,
        type_name: t,
        value: v,
    })
}

fn top_field_decl() -> Parser<'static, u8, TopLevelBlock> {
    let matched = access_level() - space() + var_decl();
    matched.map(|(a, v)| TopLevelBlock::Field {
        access_level: a.unwrap(),
        declaration: v,
    })
}

fn function_kind_keyword() -> Parser<'static, u8, Option<FunctionKind>> {
    let matched = (seq(b"Property") - space()).opt() * word();
    matched.map(|w| match w.as_str() {
        "Function" => Some(FunctionKind::Function),
        "Sub" => Some(FunctionKind::Sub),
        "Get" => Some(FunctionKind::PropertyGet),
        _ => None,
    })
}

fn function_param() -> Parser<'static, u8, FunctionParam> {
    let name_and_type = (seq(b"Optional") - space()).opt() * (seq(b"ByVal") - space()).opt()
        * word() - space() - seq(b"As") - space() + word() - space();
    let default_value = (sym(b'=') * space() * word()).opt();
    let param = name_and_type + default_value - space();

    param.map(|((n, t), v)| FunctionParam {
        name: n,
        type_name: t,
        default_value: v,
    })
}

fn function_param_list() -> Parser<'static, u8, Vec<FunctionParam>> {
    let param = none_of(b",)").repeat(1..).convert(String::from_utf8);
    let params = list(param, sym(b',') * space());
    let inner = sym(b'(') * space() * params.opt() - sym(b')');
    inner.map(|x| match x {
        Some(params) => params
            .iter()
            .map(|param| {
                let mut input = DataInput::new(param.as_bytes());
                function_param()
                    .parse(&mut input)
                    .expect("Failed to parse function parameters!")
            })
            .collect(),
        None => Vec::new(),
    })
}

fn function_decl() -> Parser<'static, u8, TopLevelBlock> {
    let maybe_async = (seq(b"Async") - space()).opt();
    let access_and_kind = access_level() - space() + function_kind_keyword() - space();
    let name_and_params = word() - space() + function_param_list() - space();
    let maybe_return_type = ((seq(b"As") - space()) * word() - space()).opt();

    let header =
        ((maybe_async + access_and_kind) + name_and_params) + maybe_return_type - sym(b'`');
    let body_lines = list(statement(), sym(b'`'));
    let matched = header + body_lines - sym(b'`') - end_function();

    matched.map(
        |((((async, (a, k)), (n, p)), r), b)| TopLevelBlock::Function {
            access_level: a.unwrap(),
            kind: k.unwrap(),
            is_async: async.is_some(),
            name: n,
            params: p,
            return_type: r.unwrap_or(String::new()),
            body: b,
        },
    )
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

fn option_decl() -> Parser<'static, u8, TopLevelBlock> {
    seq(b"Option Explicit").map(|_| TopLevelBlock::Empty)
}

fn statement() -> Parser<'static, u8, StatementBlock> {
    !end_function()
        * none_of(b"`")
            .repeat(0..)
            .convert(String::from_utf8)
            .map(|x| StatementBlock::Unknown { source: x })
}

fn match_eof() -> Parser<'static, u8, BlockParseResult> {
    seq(b"__EOF__").map(|_| BlockParseResult::EOF)
}

fn top_level_block() -> Parser<'static, u8, BlockParseResult> {
    let block_matches = option_decl() | attribute_decl() | ignored_header_decl()
        | class_marker_decl() | function_decl() | type_decl() | enum_decl()
        | const_decl() | top_field_decl();

    block_matches.map(BlockParseResult::Block) | match_eof()
        | none_of(b"")
            .repeat(0..)
            .convert(String::from_utf8)
            .map(BlockParseResult::ParseFail)
}

fn do_parse() -> Parser<'static, u8, Vec<BlockParseResult>> {
    list(top_level_block(), sym(b'`'))
}

fn without_comments(line: &str) -> &str {
    match line.find("'") {
        Some(pos) => &line[..pos],
        None => line,
    }
}

fn is_module_class(blocks: &Vec<TopLevelBlock>) -> bool {
    blocks
        .iter()
        .find(|ref x| match x {
            TopLevelBlock::ClassMarker => true,
            _ => false,
        })
        .is_some()
}

fn find_module_name(blocks: &Vec<TopLevelBlock>) -> String {
    for item in blocks {
        match item {
            &TopLevelBlock::Attribute {
                ref name,
                ref value,
            } => {
                if name == "VB_Name" {
                    return value.clone();
                }
            }
            _ => {}
        };
    }

    panic!("Failed to parse! Module had no name.");
}

pub fn parse_module(contents: &str) -> Module {
    let mut trimmed_lines: Vec<String> = contents
        .lines()
        .map(|x| without_comments(x).trim())
        .filter_map(|x| {
            if x.len() > 0 {
                Some(String::from(x))
            } else {
                None
            }
        })
        .collect();

    trimmed_lines.push(String::from("__EOF__"));

    let tick_seperated_contents = trimmed_lines.join("`");
    let mut input = DataInput::new(tick_seperated_contents.as_bytes());

    let maybe_block_parse_results = do_parse().parse(&mut input);

    let mut results: Vec<TopLevelBlock> = Vec::new();
    let mut found_eof = false;
    for parse_result in maybe_block_parse_results.expect("Failed to parse! Parser result was Err") {
        match parse_result {
            BlockParseResult::Block(b) => {
                results.push(b);
            }
            BlockParseResult::ParseFail(s) => {
                panic!("Failed to parse! Remaining:\n\n{}\n\n", s);
            }
            BlockParseResult::EOF => {
                found_eof = true;
            }
        }
    }

    if !found_eof {
        panic!("Failed to parse! A parser was likely partially matched.\n\n");
    }

    Module {
        name: find_module_name(&results),
        is_class: is_module_class(&results),
        contents: results,
    }
}
