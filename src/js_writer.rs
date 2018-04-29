use ast::*;
use parser::separate_array_access;

use std::collections::HashMap;

type TypeLookup<'a> = HashMap<String, &'a Vec<VarDeclaration>>;

fn write_let(decl: &VarDeclaration, type_lookup: &TypeLookup) -> String {
    format!(
        "let {} = {};",
        decl.name.as_str(),
        write_default_value(&decl.kind, decl.type_name.as_str(), type_lookup).as_str()
    )
}

fn write_default_value(decl_kind: &VarKind, type_name: &str, type_lookup: &TypeLookup) -> String {
    match decl_kind {
        VarKind::AutoInstantiate => format!("new_{}()", type_name),
        VarKind::RangeArray(lower, upper) => format!(
            "__makeArray({}, {}, ()=>({}))",
            lower,
            (upper - lower + 1),
            write_default_value(&VarKind::Standard, type_name, type_lookup)
        ),
        VarKind::DynamicArray => format!(
            "__makeArray(1, 0, ()=>({}))",
            write_default_value(&VarKind::Standard, type_name, type_lookup)
        ),
        VarKind::Standard => match type_name {
            "String" => String::from("\"\""),
            "Long" | "Single" | "Integer" | "Double" => String::from("0"),
            "Boolean" => String::from("False"),
            _ => write_default_object(type_name, type_lookup),
        },
    }
}

fn write_default_object(type_name: &str, type_lookup: &TypeLookup) -> String {
    if !type_lookup.contains_key(type_name) {
        return String::from("0");
    }

    let mut result = String::from("{");
    for member in type_lookup[type_name] {
        result.push_str(
            format!(
                "\"{}\":{},",
                member.name,
                write_default_value(&member.kind, member.type_name.as_str(), type_lookup).as_str()
            ).as_str(),
        );
    }
    result.pop();
    result.push_str("}");
    result
}

fn write_module_header(module: &Module) -> String {
    if module.is_class {
        format!("const new_{} = () => {{", module.name)
    } else {
        format!("const module_{} = (() => {{", module.name)
    }
}

fn write_module_footer(module: &Module) -> String {
    if module.is_class {
        String::from("};")
    } else {
        String::from("})();")
    }
}

fn join_lines(lines: &Vec<String>) -> String {
    lines
        .iter()
        .fold(String::new(), |acc, line| acc + "\n" + line) + "\n"
}

fn translate_expression(expression: &Expression) -> String {
    let mut ret = str::replace(expression.body.as_str(), " And ", " && ");
    ret = str::replace(ret.as_str(), " Or ", " || ");
    ret = str::replace(ret.as_str(), " = ", " == ");
    ret = str::replace(ret.as_str(), " <> ", " != ");
    ret = str::replace(ret.as_str(), "Not ", "! ");
    ret = str::replace(ret.as_str(), " .", " __with.");
    ret = str::replace(ret.as_str(), "(.", " (__with.");
    ret = str::replace(ret.as_str(), " Mod ", " % ");
    ret = str::replace(ret.as_str(), " & ", " + ");
    ret = str::replace(ret.as_str(), "\\", "/");

    if ret.chars().next().unwrap() == '.' {
        String::from("__with") + ret.as_str()
    } else {
        ret
    }
}

fn write_assignment(target: &Expression, value: &str) -> String {
    let mut result = String::new();
    let fixed_targ = translate_expression(target);

    if fixed_targ.ends_with(')') {
        let (array_targ, array_index) = separate_array_access(fixed_targ.as_str());
        result.push_str(array_targ.as_str());
        result.push_str("(");
        result.push_str(array_index.as_str());
        result.push_str(",");
        result.push_str(value);
        result.push_str(");");
    } else {
        result.push_str(fixed_targ.as_str());
        result.push_str(" = ");
        result.push_str(value);
        result.push_str(";");
    }

    result
}

fn write_statement_line(line: &StatementLine, type_lookup: &TypeLookup) -> String {
    let mut result = String::new();

    match line {
        StatementLine::Dim(declaration) => {
            result.push_str(write_let(declaration, type_lookup).as_str());
        }
        StatementLine::ReDim {
            preserve,
            target_name,
            new_size,
        } => {
            result.push_str(translate_expression(target_name).as_str());
            result.push_str("(null,null,{");
            if *preserve {
                result.push_str("preserve:true,");
            }
            result.push_str("count:(");
            result.push_str(translate_expression(new_size).as_str());
            result.push_str(")});");
        }
        StatementLine::Assignment { to_name, value } => {
            result
                .push_str(write_assignment(to_name, translate_expression(value).as_str()).as_str());
        }
        StatementLine::CallSub { name, args } => {
            result.push_str(translate_expression(name).as_str());
            result.push_str("(");
            for arg in args {
                match arg {
                    Some(a) => result.push_str(translate_expression(a).as_str()),
                    None => result.push_str("(void(0))"),
                };
                result.push_str(",");
            }
            result.pop();
            result.push_str(");");
        }
        StatementLine::SingleLineIf {
            condition,
            if_body,
            else_body,
        } => {
            result.push_str("if (");
            result.push_str(translate_expression(condition).as_str());
            result.push_str(") {");
            result.push_str(write_statement_line(if_body, type_lookup).as_str());
            result.push_str("} else {");
            result.push_str(write_statement_line(else_body, type_lookup).as_str());
            result.push_str("}");
        }
        StatementLine::BeginIf(condition) => {
            result.push_str("if (");
            result.push_str(translate_expression(condition).as_str());
            result.push_str(") {");
        }
        StatementLine::ElseIf(condition) => {
            result.push_str("} else if (");
            result.push_str(translate_expression(condition).as_str());
            result.push_str(") {");
        }
        StatementLine::Else => {
            result.push_str("} else {");
        }
        StatementLine::BeginWith(target) => {
            result.push_str("{const __with=");
            result.push_str(translate_expression(target).as_str());
            result.push_str(";");
        }
        StatementLine::BeginFor {
            index,
            lower_bound,
            upper_bound,
            step,
        } => {
            result.push_str("for (let ");
            result.push_str(index);
            result.push_str(" = (");
            result.push_str(translate_expression(lower_bound).as_str());
            result.push_str("); ");
            result.push_str(index);
            if step.chars().next().unwrap() == '-' {
                result.push_str(" >= (");
            } else {
                result.push_str(" <= (");
            }
            result.push_str(translate_expression(upper_bound).as_str());
            result.push_str("); ");
            result.push_str(index);
            result.push_str(" += (");
            result.push_str(step);
            result.push_str(")) {");
        }
        StatementLine::DoLoop {
            kind,
            condition,
            is_end,
        } => {
            if *is_end {
                match kind {
                    DoLoopKind::None => result.push_str("}"),
                    DoLoopKind::While => result.push_str(
                        format!("}} while ({});", translate_expression(condition)).as_str(),
                    ),
                    DoLoopKind::Until => result.push_str(
                        format!("}} while (!({}));", translate_expression(condition)).as_str(),
                    ),
                }
            } else {
                match kind {
                    DoLoopKind::None => result.push_str("do {"),
                    DoLoopKind::While => result.push_str(
                        format!("while ({}) {{", translate_expression(condition)).as_str(),
                    ),
                    DoLoopKind::Until => result.push_str(
                        format!("while (!({})) {{", translate_expression(condition)).as_str(),
                    ),
                }
            }
        }
        StatementLine::BeginSelect(expr) => {
            result.push_str("switch (");
            result.push_str(translate_expression(expr).as_str());
            result.push_str(") { case \"__NOP\":");
        }
        StatementLine::CaseLabel(expr) => {
            result.push_str("break; case ");
            result.push_str(translate_expression(expr).as_str());
            result.push_str(":");
        }
        StatementLine::Set {
            target_name,
            type_name,
        } => {
            let value = match type_name {
                Some(t) => format!("new_{}()", t),
                None => String::from("null"),
            };
            result.push_str(write_assignment(target_name, value.as_str()).as_str());
        }
        StatementLine::EndBlock => {
            result.push_str("}");
        }
        StatementLine::ExitSub => {
            result.push_str("return;");
        }
        StatementLine::ExitFunction => {
            result.push_str("return __retVal();");
        }
        StatementLine::ExitLoop => {
            result.push_str("break;");
        }
        StatementLine::Unknown(source) => {
            result.push_str("/*");
            result.push_str(source.as_str());
            result.push_str("*/");
        }
        _ => {}
    }

    result
}

fn write_function_body(body: &Vec<StatementLine>, type_lookup: &TypeLookup) -> String {
    let mut result = String::new();

    for s in body {
        result.push_str(write_statement_line(s, type_lookup).as_str());
        result.push_str("\n");
    }

    result
}

fn write_function_return_header(name: &str, return_type: &str, type_lookup: &TypeLookup) -> String {
    format!(
        "let {0} = {1}; const __retVal = () => {0};\n",
        name,
        write_default_value(&VarKind::Standard, return_type, type_lookup)
    )
}

fn write_function_return_footer() -> String {
    String::from("return __retVal();\n")
}

fn write_function(
    is_async: bool,
    name: &str,
    params: &Vec<FunctionParam>,
    body: &Vec<StatementLine>,
    return_type: Option<&str>,
    type_lookup: &TypeLookup,
) -> String {
    let mut result = String::new();

    if is_async {
        result.push_str("async ");
    }

    result.push_str("(");

    if params.len() > 0 {
        for p in params {
            result.push_str(format!("{},", p.name).as_str());
        }
        result.pop();
    }

    result.push_str(") => {\n");

    for p in params {
        match &p.default_value {
            Some(v) => {
                result.push_str(
                    format!("if (typeof {0} === 'undefined') {0} = {1};\n", p.name, v).as_str(),
                );
            }
            None => {}
        }
    }

    if let Some(ret) = return_type {
        result.push_str(write_function_return_header(name, ret, type_lookup).as_str());
    }

    result.push_str(write_function_body(body, type_lookup).as_str());

    if return_type.is_some() {
        result.push_str(write_function_return_footer().as_str());
    }

    result.push_str("}");

    result
}

fn write_module(module: &Module, type_lookup: &TypeLookup) -> String {
    let mut pre_header: Vec<String> = Vec::new();
    let mut post_header: Vec<String> = Vec::new();
    let mut return_block: Vec<String> = Vec::new();
    let mut post_footer: Vec<String> = Vec::new();

    for block in &module.contents {
        match block {
            TopLevelBlock::Constant {
                access_level,
                name,
                value,
                ..
            } => {
                let code = format!("const {} = {};", name.as_str(), value.as_str());
                match access_level {
                    AccessLevel::Public => pre_header.push(code),
                    AccessLevel::Private => post_header.push(code),
                };
            }

            TopLevelBlock::Field {
                access_level,
                declaration,
            } => {
                if module.is_class {
                    post_header.push(write_let(&declaration, type_lookup));
                    if *access_level == AccessLevel::Public {
                        return_block.push(format!(
                            "get {0}() {{ return {0}; }},",
                            declaration.name.as_str()
                        ));
                        return_block.push(format!(
                            "set {0}(x) {{ {0} = x; }},",
                            declaration.name.as_str()
                        ));
                    }
                } else {
                    match access_level {
                        AccessLevel::Public => {
                            pre_header.push(write_let(&declaration, type_lookup));
                        }
                        AccessLevel::Private => {
                            post_header.push(write_let(&declaration, type_lookup));
                        }
                    }
                }
            }

            TopLevelBlock::Enum {
                access_level,
                values,
                ..
            } => {
                let mut val: i32 = 0;

                for (name, custom_val) in values {
                    val = custom_val.unwrap_or(val);
                    let code = format!("const {} = {};", name.as_str(), val);
                    val += 1;

                    match access_level {
                        AccessLevel::Public => pre_header.push(code),
                        AccessLevel::Private => post_header.push(code),
                    };
                }
            }

            TopLevelBlock::Function {
                access_level,
                kind,
                is_async,
                name,
                params,
                body,
                return_type,
            } => {
                if *access_level == AccessLevel::Public && *kind == FunctionKind::PropertyGet
                    && params.len() == 0
                {
                    return_block.push(format!(
                        "get {}() {{ {} {} {} }},",
                        name.as_str(),
                        write_function_return_header(
                            name.as_str(),
                            return_type.as_str(),
                            type_lookup
                        ),
                        write_function_body(body, type_lookup).as_str(),
                        write_function_return_footer()
                    ));
                } else {
                    let ret_type =
                        if kind == &FunctionKind::Function || kind == &FunctionKind::PropertyGet {
                            Some(return_type.as_str())
                        } else {
                            None
                        };

                    post_header.push(format!(
                        "const {} = {};",
                        name.as_str(),
                        write_function(
                            *is_async,
                            name.as_str(),
                            params,
                            body,
                            ret_type,
                            type_lookup
                        ).as_str()
                    ));

                    if module.is_class && name == "Class_Initialize" {
                        post_header.push(String::from("Class_Initialize();"));
                    }

                    if *access_level == AccessLevel::Public {
                        return_block.push(format!("{},", name.as_str()));

                        if !module.is_class {
                            post_footer.push(format!(
                                "const {0} = module_{1}.{0};",
                                name.as_str(),
                                module.name.as_str()
                            ));
                        }
                    }
                }
            }

            _ => {}
        }
    }

    let mut result = String::new();

    result.push_str(format!("// Module: {} ", module.name.as_str()).as_str());
    result.push_str("=".repeat(40).as_str());
    result.push_str(join_lines(&pre_header).as_str());
    result.push_str(write_module_header(module).as_str());
    result.push_str(join_lines(&post_header).as_str());
    result.push_str("\nreturn {");
    result.push_str(join_lines(&return_block).as_str());
    result.push_str("};\n");
    result.push_str(write_module_footer(module).as_str());
    result.push_str(join_lines(&post_footer).as_str());

    result
}

fn collect_types(program: &Vec<Module>) -> TypeLookup {
    let mut result = HashMap::new();

    for module in program {
        for block in &module.contents {
            if let TopLevelBlock::Type(type_decl) = block {
                result.insert(type_decl.name.clone(), &type_decl.fields);
            }
        }
    }

    result
}

pub fn write_program(program: &Vec<Module>) -> String {
    let mut result = String::new();

    let type_lookup = collect_types(program);

    for module in program.iter().filter(|x| x.is_class) {
        result.push_str(write_module(module, &type_lookup).as_str());
        result.push_str("\n");
    }

    for module in program.iter().filter(|x| !x.is_class) {
        result.push_str(write_module(module, &type_lookup).as_str());
        result.push_str("\n");
    }

    result.push_str("module.exports = __main__;\n");

    result
}
