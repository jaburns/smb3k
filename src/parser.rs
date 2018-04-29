use ast::*;
use pom::parser::*;
use pom::DataInput;

enum BlockParseResult {
    Block(TopLevelBlock),
    ParseFail(String),
    EOF,
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

pub fn separate_array_access(expr: &str) -> (String, String) {
    let mut s = String::from(expr);
    s.pop();

    let mut paren_count = 0;
    let mut split_index = 0;
    for i in (0..s.len()).rev() {
        let c = s.chars().nth(i).unwrap();
        if c == ')' {
            paren_count += 1;
        }
        if c == '(' {
            if paren_count == 0 {
                split_index = i;
                break;
            } else {
                paren_count -= 1;
            }
        }
    }

    (
        String::from(&s[..split_index]),
        String::from(&s[(split_index + 1)..]),
    )
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

fn do_parse() -> Parser<'static, u8, Vec<BlockParseResult>> {
    list(top_level_block(), sym(b'`'))
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

fn ignored_header_decl() -> Parser<'static, u8, TopLevelBlock> {
    (seq(b"BEGIN") * none_of(b"E").repeat(0..) * seq(b"END")).map(|_| TopLevelBlock::Empty)
}

fn class_marker_decl() -> Parser<'static, u8, TopLevelBlock> {
    seq(b"VERSION 1.0 CLASS").map(|_| TopLevelBlock::ClassMarker)
}

fn option_decl() -> Parser<'static, u8, TopLevelBlock> {
    seq(b"Option Explicit").map(|_| TopLevelBlock::Empty)
}

fn match_eof() -> Parser<'static, u8, BlockParseResult> {
    seq(b"__EOF__").map(|_| BlockParseResult::EOF)
}

fn space() -> Parser<'static, u8, ()> {
    one_of(b" \t\r\n").repeat(0..).discard()
}

fn word() -> Parser<'static, u8, String> {
    none_of(b" \t\r\n`()")
        .repeat(1..)
        .convert(String::from_utf8)
}

fn not_space() -> Parser<'static, u8, String> {
    none_of(b" `").repeat(1..).convert(String::from_utf8)
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

fn statement() -> Parser<'static, u8, StatementLine> {
    let matched = on_error_statement() | dim_statement() | assignment_statement() | set_statement()
        | redim_statement() | label_statement() | single_line_if_statement()
        | begin_if_block() | else_if_line() | else_line() | begin_with_block()
        | end_block() | exit_sub_statement() | exit_function_statement()
        | exit_loop_statement() | begin_do_block() | end_do_block()
        | begin_for_block() | for_next_statement() | begin_select_block()
        | case_label_line() | file_operation() | call_sub_statement()
        | unknown_statement();

    !end_function() * matched
}

fn parse_single_statement(line: &str) -> StatementLine {
    let mut input = DataInput::new(line.as_bytes());
    statement()
        .parse(&mut input)
        .expect("Failed to parse statement!")
}

fn on_error_statement() -> Parser<'static, u8, StatementLine> {
    (seq(b"On Error") - none_of(b"`").repeat(0..)).map(|_| StatementLine::Empty)
}

fn redim_statement() -> Parser<'static, u8, StatementLine> {
    let matched = seq(b"ReDim") * space() * (seq(b"Preserve") - space()).opt() + rest_of_the_line();
    matched.map(|(p, ts)| {
        let (t, s) = separate_array_access(ts.as_str());

        StatementLine::ReDim {
            preserve: p.is_some(),
            target_name: Expression { body: t },
            new_size: Expression { body: s },
        }
    })
}

fn label_statement() -> Parser<'static, u8, StatementLine> {
    (word() * sym(b':')).map(|_| StatementLine::Empty)
}

fn dim_statement() -> Parser<'static, u8, StatementLine> {
    (seq(b"Dim") * space() * var_decl()).map(|v| StatementLine::Dim(v))
}

fn assignment_statement() -> Parser<'static, u8, StatementLine> {
    let matched = not_space() - space() - sym(b'=') - space() + rest_of_the_line();
    matched.map(|(to, exp)| StatementLine::Assignment {
        to_name: Expression { body: to },
        value: Expression { body: exp },
    })
}

fn set_statement_start() -> Parser<'static, u8, String> {
    seq(b"Set") * space() * not_space() - space() - sym(b'=') - space()
}

fn set_statement() -> Parser<'static, u8, StatementLine> {
    let set_new = set_statement_start() - seq(b"New") - space() + word() - space();
    let set_nothing = set_statement_start() - seq(b"Nothing") - space();

    set_new.map(|(v, t)| StatementLine::Set {
        target_name: Expression { body: v },
        type_name: Some(t),
    }) | set_nothing.map(|v| StatementLine::Set {
        target_name: Expression { body: v },
        type_name: None,
    })
}

fn join(with: &str, lines: &Vec<String>) -> String {
    let mut result = lines
        .iter()
        .fold(String::new(), |acc, line| acc + line + with);

    let len = result.len();
    result.truncate(len - with.len());
    result
}

fn if_then() -> Parser<'static, u8, String> {
    let spaced_expr = list(!seq(b"Then") * not_space(), sym(b' '));
    (seq(b"If") * space() * spaced_expr - space() - seq(b"Then")).map(|x| join(" ", &x))
}

fn single_line_if_statement() -> Parser<'static, u8, StatementLine> {
    let matched = if_then() - sym(b' ') + rest_of_the_line();

    matched.map(|(c, rest)| match rest.find(" Else ") {
        Some(pos) => StatementLine::SingleLineIf {
            condition: Expression { body: c },
            if_body: Box::new(parse_single_statement(&rest[..pos])),
            else_body: Box::new(parse_single_statement(&rest[(pos + 6)..])),
        },
        None => StatementLine::SingleLineIf {
            condition: Expression { body: c },
            if_body: Box::new(parse_single_statement(&rest)),
            else_body: Box::new(StatementLine::Empty),
        },
    })
}

fn begin_if_block() -> Parser<'static, u8, StatementLine> {
    if_then().map(|c| StatementLine::BeginIf(Expression { body: c }))
}

fn else_if_line() -> Parser<'static, u8, StatementLine> {
    (seq(b"Else") * if_then()).map(|c| StatementLine::ElseIf(Expression { body: c }))
}

fn else_line() -> Parser<'static, u8, StatementLine> {
    seq(b"Else").map(|_| StatementLine::Else)
}

fn begin_with_block() -> Parser<'static, u8, StatementLine> {
    let matched = seq(b"With") * space() * rest_of_the_line();
    matched.map(|x| StatementLine::BeginWith(Expression { body: x }))
}

fn for_separate_ubound_and_step(expr: &str) -> (String, String) {
    match expr.find(" Step ") {
        Some(i) => (String::from(&expr[..i]), String::from(&expr[(i + 6)..])),
        None => (String::from(expr), String::from("1")),
    }
}

fn begin_for_block() -> Parser<'static, u8, StatementLine> {
    let iter_name = seq(b"For") * space() * word() - space() - sym(b'=') - space();
    let lower = list(!seq(b"To") * not_space(), sym(b' ')).map(|x| join(" ", &x));
    let range = lower - space() - seq(b"To") - space() + rest_of_the_line();
    let matched = iter_name + range;

    matched.map(|(i, (l, us))| {
        let (u, s) = for_separate_ubound_and_step(us.as_str());
        StatementLine::BeginFor {
            index: i,
            lower_bound: Expression { body: l },
            upper_bound: Expression { body: u },
            step: s,
        }
    })
}

fn for_next_statement() -> Parser<'static, u8, StatementLine> {
    (seq(b"Next") - rest_of_the_line()).map(|_| StatementLine::EndBlock)
}

fn do_loop_condition(prefix: &'static [u8], end: bool) -> Parser<'static, u8, StatementLine> {
    let do_while = (seq(prefix) * seq(b" While") * space() * rest_of_the_line()).map(move |c| {
        StatementLine::DoLoop {
            kind: DoLoopKind::While,
            condition: Expression { body: c },
            is_end: end && true,
        }
    });
    let do_until = (seq(prefix) * seq(b" Until") * space() * rest_of_the_line()).map(move |c| {
        StatementLine::DoLoop {
            kind: DoLoopKind::Until,
            condition: Expression { body: c },
            is_end: end,
        }
    });
    let do_forever = (seq(prefix) * (!none_of(b"`"))).map(move |_| StatementLine::DoLoop {
        kind: DoLoopKind::None,
        condition: Expression {
            body: String::from(""),
        },
        is_end: end,
    });

    do_while | do_until | do_forever
}

fn begin_do_block() -> Parser<'static, u8, StatementLine> {
    do_loop_condition(b"Do", false)
}

fn end_do_block() -> Parser<'static, u8, StatementLine> {
    do_loop_condition(b"Loop", true)
}

fn begin_select_block() -> Parser<'static, u8, StatementLine> {
    let matched = seq(b"Select Case") * space() * rest_of_the_line();
    matched.map(|x| StatementLine::BeginSelect(Expression { body: x }))
}

fn case_label_line() -> Parser<'static, u8, StatementLine> {
    let matched = seq(b"Case") * space() * rest_of_the_line();
    matched.map(|x| StatementLine::CaseLabel(Expression { body: x }))
}

fn end_block() -> Parser<'static, u8, StatementLine> {
    (seq(b"End If") | seq(b"End With") | seq(b"End Select")).map(|_| StatementLine::EndBlock)
}

fn exit_sub_statement() -> Parser<'static, u8, StatementLine> {
    seq(b"Exit Sub").map(|_| StatementLine::ExitSub)
}

fn exit_function_statement() -> Parser<'static, u8, StatementLine> {
    (seq(b"Exit Function") | seq(b"Exit Property")).map(|_| StatementLine::ExitFunction)
}

fn exit_loop_statement() -> Parser<'static, u8, StatementLine> {
    (seq(b"Exit For") | seq(b"Exit Do") | seq(b"Exit While")).map(|_| StatementLine::ExitLoop)
}

fn call_sub_statement() -> Parser<'static, u8, StatementLine> {
    let arg = none_of(b"`,").repeat(0..).convert(String::from_utf8) - space();
    let args = list(arg, sym(b',') - space());
    let matched = not_space() - space() + args;

    matched.map(|(n, a)| StatementLine::CallSub {
        name: Expression { body: n },
        args: a.into_iter()
            .map(|x| {
                if x.len() < 1 {
                    None
                } else {
                    Some(Expression { body: x })
                }
            })
            .collect(),
    })
}

fn file_operation() -> Parser<'static, u8, StatementLine> {
    ((seq(b"Open ") | seq(b"Close ") | seq(b"Get ") | seq(b"Put ") | seq(b"Line ")) * rest_of_the_line())
        .map(|_| StatementLine::FileOperation)
}

fn rest_of_the_line() -> Parser<'static, u8, String> {
    none_of(b"`").repeat(0..).convert(String::from_utf8)
}

fn unknown_statement() -> Parser<'static, u8, StatementLine> {
    rest_of_the_line().map(|x| StatementLine::Unknown(x))
}
