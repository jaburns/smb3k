use ast::*;

fn write_let(decl: &VarDeclaration) -> String {
    match decl.kind {
        VarKind::AutoInstantiate => format!(
            "let {} = new_{}();",
            decl.name.as_str(),
            decl.type_name.as_str()
        ),
        _ => format!("let {};", decl.name.as_str()),
    }
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

fn write_function(params: &Vec<FunctionParam>, body: &Vec<StatementBlock>) -> String {
    let mut result = String::new();
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
                result.push_str(format!("if (typeof {0} === 'undefined') {0} = {1};\n", p.name, v).as_str());
            }
            None => {}
        }
    }

 // for s in body {
 //     match s {
 //         StatementBlock::Unknown { source } => {
 //             result.push_str(source.as_str());
 //             result.push_str("\n");
 //         }
 //         _ => {}
 //     }
 // }

    result.push_str("}");

    result
}

fn write_module(module: &Module) -> String {
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
            } => match access_level {
                AccessLevel::Public => {
                    pre_header.push(format!("const {} = {};", name.as_str(), value.as_str()));
                }
                AccessLevel::Private => {
                    post_header.push(format!("const {} = {};", name.as_str(), value.as_str()));
                }
            },

            TopLevelBlock::Field {
                access_level,
                declaration,
            } => {
                if module.is_class {
                    post_header.push(write_let(&declaration));
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
                            pre_header.push(write_let(&declaration));
                        }
                        AccessLevel::Private => {
                            post_header.push(write_let(&declaration));
                        }
                    }
                }
            }

            TopLevelBlock::Function {
                access_level,
                kind,
                name,
                params,
                body,
                ..
            } => {
                match access_level {
                    AccessLevel::Public => {
                        if *kind == FunctionKind::PropertyGet && params.len() > 0 {
                            return_block.push(format!(
                                "get {}: {},",
                                name.as_str(),
                                write_function(params, body).as_str()
                            ));
                        } else {
                            return_block.push(format!(
                                "{}: {},",
                                name.as_str(),
                                write_function(params, body).as_str()
                            ));
                        }

                        if !module.is_class {
                            post_footer.push(format!(
                                "const {0} = module_{1}.{0};",
                                name.as_str(),
                                module.name.as_str()
                            ));
                        }
                    }
                    AccessLevel::Private => {
                        post_header.push(format!(
                            "const {} = {};",
                            name.as_str(),
                            write_function(params, body).as_str()
                        ));
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

pub fn write_program(program: &Vec<Module>) -> String {
    let mut result = String::new();

    result.push_str("(() => {\n");

    for module in program {
        result.push_str(write_module(module).as_str());
        result.push_str("\n");
    }

    result.push_str("__main__();\n})();");

    result
}
