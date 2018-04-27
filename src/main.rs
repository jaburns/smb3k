extern crate pom;

mod ast;
mod js_writer;
mod parser;

use std::fs::*;
use std::io::prelude::*;
use std::io::{BufWriter, Write};

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
            let module = read_file(path_str);
            result.push(module);
        }
    }

    result
}

fn main() {
    println!("\nLoading and parsing VB6 program...");
    let program = load_program();

    println!("Done parsing, generating JS...");
    let js = write_program(&program);


    let file = OpenOptions::new()
        .write(true)
        .truncate(true)
        .create(true)
        .open("js/game.js")
        .expect("Failed to open output file for writing!");

    let mut f = BufWriter::new(file);
    f.write_all(js.as_bytes()).expect("Failed to write to output file!");

    println!("Done!\n");
}
