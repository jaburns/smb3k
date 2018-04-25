extern crate pom;

mod ast;
mod parser;
mod js_writer;

use std::collections::HashMap;
use std::fs::*;
use std::io::prelude::*;

use ast::*;
use parser::*;
use js_writer::*;

fn find_module_name(module: &Vec<TopLevelBlock>) -> Option<String> {
    for item in module {
        match item {
            &TopLevelBlock::Attribute {
                ref name,
                ref value,
            } => {
                if name == "VB_Name" {
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

    (
        find_module_name(&result).expect("Module '{}' did not have attribute VB_Name!"),
        result,
    )
}

fn load_program() -> Program {
    let mut result = HashMap::new();

    for maybe_path in read_dir("test_program").unwrap() {
        let path = maybe_path.unwrap().path();
        let path_str = path.to_str().unwrap();

        if path_str.ends_with(".bas") || path_str.ends_with(".cls") {
            println!("\nLoading and parsing module at {}", path_str);
            let (name, parsed) = read_file(path_str);

            for block in &parsed {
                println!("  :: {:?}", block);
            }

            result.insert(name, parsed);
        }
    }

    result
}

fn main() {
    let program = load_program();

    println!("\nDone parsing, generating JS...");

    let js = write_program(&program);

    println!("\n{}", js);

    println!("\nDone!\n");
}