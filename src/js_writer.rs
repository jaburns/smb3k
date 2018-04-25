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

fn write_let(indent: usize, decl: &VarDeclaration) -> String {
    match decl.kind {
        VarKind::AutoInstantiate =>
            format!("{}let {} = new_{}();\n", " ".repeat(indent), decl.name.as_str(), decl.type_name.as_str()),
        _ =>
            format!("{}let {};\n", " ".repeat(indent), decl.name.as_str()),
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
//  lines.iter().intersperse(",".to_string())
}

fn write_module(module: &Module) -> String {
    let mut pre_header: Vec<String> = Vec::new();
    let mut post_header: Vec<String> = Vec::new();
    let mut return_block: Vec<String> = Vec::new();
    let mut post_footer: Vec<String> = Vec::new();

/*
    let header = write_module_header(module)
    let footer = write_module_footer(module);
*/


    let public_fields = get_fields(&module.contents, AccessLevel::Public);
    let private_fields = get_fields(&module.contents, AccessLevel::Private);
    let public_functions = get_functions(&module.contents, AccessLevel::Public);
    let private_functions = get_functions(&module.contents, AccessLevel::Private);

    if !module.is_class {
        for decl in &public_fields {
            result.push_str(write_let(0, decl).as_str());
        }
    }

    result.push_str(write_module_header(module).as_str());

    if module.is_class {
        for decl in &public_fields {
            result.push_str(write_let(0, decl).as_str());
        }
    }

    for decl in &private_fields {
        result.push_str(write_let(2, decl).as_str());
    }

    // TODO write private functions

    result.push_str("\n  return {\n");

    // TODO write public functions

    result.push_str("  };\n");

    result.push_str(write_module_footer(module).as_str());

    if !module.is_class {
        for block in &public_functions {
            if let TopLevelBlock::Function { name, .. } = block {
                result.push_str(format!("const {0} = module_{1}.{0};\n", name.as_str(), module.name.as_str()).as_str());
            }
        }
    }

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