use ast::*;

fn get_fields(blocks: &Vec<TopLevelBlock>, at_level: AccessLevel) -> Vec<&VarDeclaration> {
    blocks.iter().filter_map(|ref x| match x { 
            TopLevelBlock::Field { 
                access_level,
                declaration
            } => if *access_level == at_level { Some(declaration) } else { None },
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

fn write_class(module: &Module) -> String {
    let mut result = String::new();

    let public_fields = get_fields(&module.contents, AccessLevel::Public);
    let private_fields = get_fields(&module.contents, AccessLevel::Private);

    result.push_str(format!("const new_{} = () => {{\n", module.name).as_str());

    for decl in &private_fields {
        result.push_str(write_let(2, decl).as_str());
    }
    for decl in &public_fields {
        result.push_str(write_let(2, decl).as_str());
    }

    // TODO write private functions and class Class_Initialize if it exists.

    result.push_str("\n  return {\n");

    // TODO write public functions

    for decl in &public_fields {
        result.push_str(&format!("    get {}() {{ return {}; }},\n", decl.name.as_str(), decl.name.as_str()));
        result.push_str(&format!("    set {}(x) {{ {} = x; }},\n", decl.name.as_str(), decl.name.as_str()));
    }

    result.push_str("  };\n};\n");

    result
}

fn write_module(module: &Module) -> String {
    let mut result = String::new();

    let public_fields = get_fields(&module.contents, AccessLevel::Public);
    let private_fields = get_fields(&module.contents, AccessLevel::Private);

    for decl in &public_fields {
        result.push_str(write_let(0, decl).as_str());
    }

    result.push_str(format!("const module_{} = () => {{\n", module.name).as_str());

    for decl in &private_fields {
        result.push_str(write_let(2, decl).as_str());
    }

    // TODO write private functions

    result.push_str("\n  return {\n");

    // TODO write public functions

    result.push_str("  };\n};\n");

    result
}

pub fn write_program(program: &Vec<Module>) -> String {
    let mut result = String::new();

    result.push_str("(() => {\n");

    for module in program {
        if module.is_class {
            result.push_str(write_class(module).as_str());
        } else {
            result.push_str(write_module(module).as_str());
        }
        result.push_str("\n");
    }

    result.push_str("__main__();\n})();");

    result
}