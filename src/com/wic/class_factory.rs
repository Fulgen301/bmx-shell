use std::ffi::c_void;
use std::sync::atomic::{AtomicUsize, Ordering};

use windows::Win32::Foundation::{CLASS_E_NOAGGREGATION, E_POINTER};
use windows::{
    core::{implement, GUID},
    Win32::{
        Foundation::BOOL,
        System::Com::{IClassFactory, IClassFactory_Impl},
    },
};
use windows_core::HRESULT;

static LOCK_COUNT: AtomicUsize = AtomicUsize::new(0);

#[implement(IClassFactory)]
pub struct ClassFactory {
    constructor: fn(*const GUID, *mut *mut c_void) -> HRESULT,
}

impl ClassFactory {
    pub fn new(constructor: fn(*const GUID, *mut *mut c_void) -> HRESULT) -> Self {
        Self { constructor }
    }
}

impl IClassFactory_Impl for ClassFactory_Impl {
    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn CreateInstance(
        &self,
        outer: Option<&windows::core::IUnknown>,
        iid: *const GUID,
        ppv: *mut *mut core::ffi::c_void,
    ) -> windows::core::Result<()> {
        if outer.is_some() {
            return Err(CLASS_E_NOAGGREGATION.into());
        }

        if iid.is_null() {
            return Err(E_POINTER.into());
        }

        if ppv.is_null() {
            return Err(E_POINTER.into());
        }

        (self.constructor)(iid, ppv).ok()
    }

    fn LockServer(&self, flock: BOOL) -> windows::core::Result<()> {
        if flock.as_bool() {
            LOCK_COUNT.fetch_add(1, Ordering::AcqRel);
        } else {
            LOCK_COUNT.fetch_sub(1, Ordering::AcqRel);
        }

        Ok(())
    }
}

pub fn can_unload_now() -> bool {
    LOCK_COUNT.load(Ordering::Acquire) == 0
}
