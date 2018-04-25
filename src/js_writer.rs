use ast::*;

fn get_fields(blocks: &Vec<TopLevelBlock>, at_level: AccessLevel) -> Vec<&VarDeclaration> {
    blocks.iter().filter_map(|x| match x { 
            TopLevelBlock::Field { 
                access_level,
                declaration
            } => if *access_level == at_level { Some(declaration) } else { None },
            _ => None
        })
        .collect()
}

fn get_functions(blocks: &Vec<TopLevelBlock>, at_level: AccessLevel) -> Vec<&TopLevelBlock> {
    blocks.iter().filter_map(|x| match x { 
            TopLevelBlock::Function { 
                access_level,
                ..
            } => if *access_level == at_level { Some(x) } else { None },
            _ => None
        })
        .collect()
}

fn write_let(decl: &VarDeclaration) -> String {
    match decl.kind {
        VarKind::AutoInstantiate =>
            format!("let {} = new_{}();\n", decl.name.as_str(), decl.type_name.as_str()),
        _ =>
            format!("let {};\n", decl.name.as_str()),
    }
}

fn write_module_header(module: &Module) -> String {
    if module.is_class {
        format!("const new_{} = () => {{\n", module.name)
    } else {
        format!("const module_{} = (() => {{\n", module.name)
    }
}

fn write_module_footer(module: &Module) -> String {
    if module.is_class {
        String::from("};\n")
    } else {
        String::from("})();\n")
    }
}

fn join_lines(lines: &Vec<String>) -> String {
    lines.iter().fold(String::new(), |acc, line| acc + "\n" + line) + "\n"
}

fn write_module(module: &Module) -> String {
    let mut pre_header: Vec<String> = Vec::new();
    let mut post_header: Vec<String> = Vec::new();
    let mut return_block: Vec<String> = Vec::new();
    let mut post_footer: Vec<String> = Vec::new();

    let public_fields = get_fields(&module.contents, AccessLevel::Public);
    let private_fields = get_fields(&module.contents, AccessLevel::Private);
    let public_functions = get_functions(&module.contents, AccessLevel::Public);
    let private_functions = get_functions(&module.contents, AccessLevel::Private);

    for block in module.contents {
        match block {
            TopLevelBlock::Constant { access_level, name, value, .. } => {
                match access_level {
                    AccessLevel::Public => {
                        pre_header.push(format!("const {} = {};", name.as_str(), value.as_str()));
                    }
                    AccessLevel::Private => {
                        post_header.push(format!("const {} = {};", name.as_str(), value.as_str()));
                    }
                }
            },

            TopLevelBlock::Field { access_level, declaration } => {
                if module.is_class {
                    post_header.push(write_let(declaration));
                    if access_level == AccessLevel::Public {
                        return_block.push_str(&format!("get {0}() {{ return {0}; }},\n", declaration.name.as_str()));
                        return_block.push_str(&format!("set {0}(x) {{ {0} = x; }},\n", declaration.name.as_str()));
                    }
                } else {
                    match access_level {

                            pre_header.push(format!("const {} = {};", name.as_str(), value.as_str()));
                        }
                        AccessLevel::Private => {
                            post_header.push(format!("const {} = {};", name.as_str(), value.as_str()));
                        }
                    }
                }

            },

            _ => {}
        }
    }

    if !module.is_class {
        for decl in &public_fields {
            result.push_str(write_let(0, decl).as_str());
        }
    }


    if module.is_class {
        for decl in &public_fields {
            result.push_str(write_let(0, decl).as_str());
        }
    }

    for decl in &private_fields {
        result.push_str(write_let(2, decl).as_str());
    }

    if !module.is_class {
        for block in &public_functions {
            if let TopLevelBlock::Function { name, .. } = block {
                result.push_str(format!("const {0} = module_{1}.{0};\n", name.as_str(), module.name.as_str()).as_str());
            }
        }
    }


    let mut result = String::new();

    result.push_str(format!("// Module: {}", module.name.as_str()));
    result.push_str(join_lines(&pre_header).as_str());
    result.push_str(write_module_header(module).as_str());
    result.push_str(join_lines(&post_header).as_str());
    result.push_str("\nreturn {\n");
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