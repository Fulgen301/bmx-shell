#![deny(unfulfilled_lint_expectations)]
#![deny(unused)]
#![deny(unused_imports)]
//#![deny(clippy::missing_safety_doc)]
//#![deny(clippy::undocumented_unsafe_blocks)]

pub mod bmx;
pub mod com;
pub mod export;
pub mod registry;
mod util;

pub fn add(left: u64, right: u64) -> u64 {
    left + right
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let result = add(2, 2);
        assert_eq!(result, 4);
    }
}
