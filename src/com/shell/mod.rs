use std::ops::{Deref, DerefMut};

use windows::Win32::System::Com::CoTaskMemFree;
use windows_core::PWSTR;

pub mod command;
pub mod property_store;

pub struct CoTaskMemPWSTR(PWSTR);

impl CoTaskMemPWSTR {
    pub const fn new(value: PWSTR) -> Self {
        Self(value)
    }

    pub const fn null() -> Self {
        Self(PWSTR::null())
    }
}

impl Deref for CoTaskMemPWSTR {
    type Target = PWSTR;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl DerefMut for CoTaskMemPWSTR {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.0
    }
}

impl Drop for CoTaskMemPWSTR {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.0.as_ptr().cast()));
            };
        }
    }
}
