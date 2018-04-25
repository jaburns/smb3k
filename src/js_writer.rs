use ast::*;

fn get_public_fields(blocks: &Vec<TopLevelBlock>) -> Vec<String> {
    blocks.iter().filter_map(|ref x| match x { 
            TopLevelBlock::Field { 
                access_level,
                declaration
            } => if *access_level == AccessLevel::Public { Some(declaration) } else { None },
            _ => None
        })
        .map(|VarDeclaration { name, .. }| name.clone())
        .collect()
}

fn write_class(module: &Module) -> String {
    let mut result = String::new();

    let pub_fields = get_public_fields(&module.contents);

    for decl in pub_fields {
        result.push_str(&format!("let {};\n", decl.as_str()));
    }

    result.push_str(format!("const new_{} = () => {{\n", module.name).as_str());
    result.push_str("}\n");

    result
}

pub fn write_program(program: &Vec<Module>) -> String {
    let mut result = String::new();

    for module in program {
        if module.is_class {
            result.push_str(write_class(module).as_str());
        } else {
            result.push_str("module");
        }
        result.push_str("\n");
    }

    result
}