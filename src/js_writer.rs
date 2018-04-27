use ast::*;

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

fn write_function_body(body: &Vec<StatementBlock>) -> String {
    let mut result = String::new();

    for s in body {
        match s {
            StatementBlock::Unknown { source } => {
                result.push_str("// ");
                result.push_str(source.as_str());
                result.push_str("\n");
            }
            _ => {}
        }
    }

    result
}

fn write_function(
    is_async: bool,
    params: &Vec<FunctionParam>,
    body: &Vec<StatementBlock>,
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

    result.push_str(write_function_body(body).as_str());

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
                ..
            } => {
                if *access_level == AccessLevel::Public && *kind == FunctionKind::PropertyGet
                    && params.len() == 0
                {
                    // This is inaccessible to the class it is defined in. Unsure if this is a problem yet.
                    return_block.push(format!(
                        "get {}() {{ {} }},",
                        name.as_str(),
                        write_function_body(body).as_str()
                    ));
                } else {
                    post_header.push(format!(
                        "const {} = {};",
                        name.as_str(),
                        write_function(*is_async, params, body).as_str()
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
