use windows::{
    core::PCWSTR,
    Win32::{
        Foundation::HMODULE,
        System::LibraryLoader::{
            GetModuleFileNameW, GetModuleHandleExW, GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
            GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        },
    },
};

pub mod guid {
    use core::str;
    use std::{
        fmt::Display,
        io::{Cursor, Write},
    };

    use windows::core::GUID;

    struct GuidWrapper<'a>(&'a GUID);

    impl<'a> Display for GuidWrapper<'a> {
        fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            write!(
                f,
                "{{{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}}}",
                self.0.data1,
                self.0.data2,
                self.0.data3,
                self.0.data4[0],
                self.0.data4[1],
                self.0.data4[2],
                self.0.data4[3],
                self.0.data4[4],
                self.0.data4[5],
                self.0.data4[6],
                self.0.data4[7]
            )
        }
    }

    pub const fn from_str(value: &str) -> GUID {
        if !value.is_ascii() {
            panic!("Invalid GUID characters");
        }

        let value = value.as_bytes();

        const fn slice_to_str(slice: &[u8]) -> &str {
            unsafe { str::from_utf8_unchecked(slice) }
        }

        const fn to_hex(value1: u8, value2: u8) -> u8 {
            match u8::from_str_radix(slice_to_str(&[value1, value2]), 16) {
                Ok(data4) => data4,
                Err(_) => panic!("Invalid GUID data4"),
            }
        }

        match value {
            [a0, a1, a2, a3, a4, a5, a6, a7, b'-', b0, b1, b2, b3, b'-', c0, c1, c2, c3, b'-', d0, d1, d2, d3, b'-', e0, e1, e2, e3, e4, e5, e6, e7, e8, e9, e10, e11]
            | [b'{', a0, a1, a2, a3, a4, a5, a6, a7, b'-', b0, b1, b2, b3, b'-', c0, c1, c2, c3, b'-', d0, d1, d2, d3, b'-', e0, e1, e2, e3, e4, e5, e6, e7, e8, e9, e10, e11, b'}'] => {
                GUID {
                    data1: match u32::from_str_radix(
                        slice_to_str(&[*a0, *a1, *a2, *a3, *a4, *a5, *a6, *a7]),
                        16,
                    ) {
                        Ok(data1) => data1,
                        Err(_) => panic!("Invalid GUID data1"),
                    },
                    data2: match u16::from_str_radix(slice_to_str(&[*b0, *b1, *b2, *b3]), 16) {
                        Ok(data2) => data2,
                        Err(_) => panic!("Invalid GUID data2"),
                    },
                    data3: match u16::from_str_radix(slice_to_str(&[*c0, *c1, *c2, *c3]), 16) {
                        Ok(data3) => data3,
                        Err(_) => panic!("Invalid GUID data3"),
                    },
                    data4: [
                        to_hex(*d0, *d1),
                        to_hex(*d2, *d3),
                        to_hex(*e0, *e1),
                        to_hex(*e2, *e3),
                        to_hex(*e4, *e5),
                        to_hex(*e6, *e7),
                        to_hex(*e8, *e9),
                        to_hex(*e10, *e11),
                    ],
                }
            }
            _ => panic!("Invalid GUID"),
        }
    }

    pub trait GuidExt {
        fn to_ascii_with_nul(&self) -> [u8; 39];
        fn to_wide(&self) -> [u16; 39] {
            self.to_ascii_with_nul().map(|value| value as u16)
        }
    }

    impl GuidExt for GUID {
        fn to_ascii_with_nul(&self) -> [u8; 39] {
            let mut cursor = Cursor::new([0u8; 39]);
            write!(cursor, "{}", GuidWrapper(self)).unwrap();
            assert!(cursor.position() == 38);
            cursor.into_inner()
        }
    }
}

#[inline(never)]
pub unsafe fn get_this_module_handle() -> windows::core::Result<HMODULE> {
    let mut module = HMODULE::default();

    unsafe {
        GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            PCWSTR::from_raw(get_this_module_path as *const () as *const _),
            &raw mut module,
        )?;
    }

    Ok(module)
}

pub fn get_module_path(module: HMODULE) -> windows::core::Result<Vec<u16>> {
    let mut path = vec![0; 1024];

    loop {
        let size = unsafe { GetModuleFileNameW(module, &mut path) };

        if size == 0 {
            return Err(windows::core::Error::from_win32());
        } else if size as usize != path.len() {
            path.truncate((size as usize) + 1);
            return Ok(path);
        } else {
            path.resize(path.len() * 2, 0);
        }
    }
}

pub unsafe fn get_this_module_path() -> windows::core::Result<Vec<u16>> {
    get_module_path(get_this_module_handle()?)
}
