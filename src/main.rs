extern crate pom;

mod ast;
mod parser;

use std::fs::*;
use std::io::prelude::*;
use std::collections::HashMap;

use ast::*;
use parser::*;

type Program = HashMap<String, Vec<TopLevelBlock>>;

fn find_module_name(module: &Vec<TopLevelBlock>) -> Option<String> {
    for item in module {
        match item { 
            &TopLevelBlock::Attribute {
                 ref name,
                 ref value,
             } => {
                if name == "VB_Name" {
                    println!("Parsed module name: {}", value);
                    return Some(value.clone());
                }
            }
            _ => {}
        };
    }

    None
}

fn read_file(path: &str) -> (String, Vec<TopLevelBlock>) {
    let mut file = File::open(path).unwrap();
    let mut contents = String::new();
    file.read_to_string(&mut contents).unwrap();

    let result = parse_module(contents.as_str());

    (find_module_name(&result).expect("Module '{}' did not have attribute VB_Name!"), result)
}

fn load_program() -> Program {
    let mut result = HashMap::new();

    for maybe_path in read_dir("vb6/Source").unwrap() {
        let path = maybe_path.unwrap().path();
        let path_str = path.to_str().unwrap();

        if path_str.ends_with(".bas") || path_str.ends_with(".cls") {
            println!("Loading and parsing module at {}", path_str);
            let (name, parsed) = read_file(path_str);
            result.insert(name, parsed);
        }
    }

    result
}

fn main() {
    load_program();

    println!("\nSUCCESS\n");
}
