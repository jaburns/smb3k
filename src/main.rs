extern crate pom;

mod ast;
mod parser;

use std::fs::File;
use std::io::prelude::*;
use parser::*;

fn main() {
    let mut file = File::open("vb6/Source/mGlobal.bas").unwrap();

    let mut contents = String::new();
    file.read_to_string(&mut contents).unwrap();

    let parsed = parse_module(contents.as_str());

    for p in parsed {
        println!(":: {:?}", p);
    }

    println!("\nSUCCESS\n");
}