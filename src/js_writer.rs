use ast::*;

use std::collections::HashMap;
use std::fmt;

fn is_module_class(module: &Vec<TopLevelBlock>) -> bool {
    true
}

fn write_class(name: &str, blocks: &Vec<TopLevelBlock>) -> String {
    let mut result = String::new();

    result.push_str(format!("const new_{} = () => {{\n", name).as_str());

 // blocks.iter().filter(|&x| match x { 
 //     &TopLevelBlock::Field { 
 //         access_level,
 //         declaration
 //     } => access_level == AccessLevel::Public,
 //     _ => false
 // });

    result.push_str("}\n");

    result
}

pub fn write_program(program: &HashMap<String, Vec<TopLevelBlock>>) -> String {
    let mut result = String::new();

    for (module, blocks) in program {
        result.push_str(write_class(module, blocks).as_str());
        result.push_str("\n");
    }

    result
}