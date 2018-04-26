extern crate pom;

mod ast;
mod js_writer;
mod parser;

use std::fs::*;
use std::io::prelude::*;

use ast::*;
use js_writer::*;
use parser::*;

fn read_file(path: &str) -> Module {
    let mut file = File::open(path).unwrap();
    let mut contents = String::new();
    file.read_to_string(&mut contents).unwrap();

    parse_module(contents.as_str())
}

fn load_program() -> Vec<Module> {
    let mut result = Vec::new();

    for maybe_path in read_dir("vb6/Source").unwrap() {
        let path = maybe_path.unwrap().path();
        let path_str = path.to_str().unwrap();

        if path_str.ends_with(".bas") || path_str.ends_with(".cls") {
            println!("\n//Loading and parsing module at {}", path_str);
            let module = read_file(path_str);

            for block in &module.contents {
                println!("//  :: {:?}", block);
            }

            result.push(module);
        }
    }

    result
}

fn main() {
    let program = load_program();

    println!("\n//Done parsing, generating JS...");

    let js = write_program(&program);

    println!("\n{}\n//Done!\n", js);
}
