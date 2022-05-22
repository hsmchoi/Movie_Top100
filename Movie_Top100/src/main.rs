use office::{Excel, Range, DataType};

let mut workbook = Excel::open(excel/movie_list_by_BBC.xlsx).unwrap();


fn main() {
    println!("Hello, world!");
    println!(workbook);

}
