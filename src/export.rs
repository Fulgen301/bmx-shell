use std::os::raw::c_void;

use windows::{
    core::{w, HRESULT},
    Win32::{
        Foundation::{CLASS_E_CLASSNOTAVAILABLE, E_POINTER, S_OK},
        System::Registry::HKEY_CLASSES_ROOT,
    },
};
use windows_core::{ComObject, IUnknown, Interface, GUID};

use crate::{
    com::{
        shell::{command::transcode::Transcode, property_store::PropertyStore},
        wic::{class_factory::ClassFactory, decoder::BitmapDecoder, encoder::BitmapEncoder},
        CoClass,
    },
    registry::{
        register_server,
        transaction::{Key, Transaction},
        unregister_server,
    },
    util::get_this_module_path,
};

#[allow(non_snake_case)]
#[no_mangle]
unsafe extern "system" fn DllRegisterServer() -> HRESULT {
    fn do_register() -> windows::core::Result<()> {
        let transaction = Transaction::new(true)?;

        let classes_root = Key::predefined(&transaction, HKEY_CLASSES_ROOT, w!(""))?;
        /*let classes_root = Key::predefined(
            &transaction,
            HKEY_CURRENT_USER,
            w!("Software\\X16BMX\\BMX\\DryRun"),
        )?;*/
        register_server(&transaction, &classes_root, unsafe {
            &get_this_module_path()?
        })
    }

    match do_register() {
        Ok(()) => S_OK,
        Err(err) => err.into(),
    }
}

#[allow(non_snake_case)]
#[no_mangle]
unsafe extern "system" fn DllUnregisterServer() -> HRESULT {
    fn do_unregister() -> windows::core::Result<()> {
        let transaction = Transaction::new(true)?;

        let classes_root = Key::predefined(&transaction, HKEY_CLASSES_ROOT, w!(""))?;

        unregister_server(&transaction, &classes_root)
    }

    match do_unregister() {
        Ok(()) => S_OK,
        Err(err) => err.into(),
    }
}

#[allow(non_snake_case)]
#[no_mangle]
unsafe extern "system" fn DllGetClassObject(
    clsid: *const GUID,
    iid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    if clsid.is_null() {
        return E_POINTER;
    }

    if iid.is_null() {
        return E_POINTER;
    }

    if ppv.is_null() {
        return E_POINTER;
    }

    let class_factory = match unsafe { *clsid } {
        BitmapDecoder::CLSID => ClassFactory::new(|iid, ppv| unsafe {
            ComObject::new(BitmapDecoder::new())
                .as_interface::<IUnknown>()
                .query(iid, ppv)
        }),
        BitmapEncoder::CLSID => ClassFactory::new(|iid, ppv| unsafe {
            ComObject::new(BitmapEncoder::new())
                .as_interface::<IUnknown>()
                .query(iid, ppv)
        }),
        PropertyStore::CLSID => ClassFactory::new(|iid, ppv| unsafe {
            ComObject::new(PropertyStore::new())
                .as_interface::<IUnknown>()
                .query(iid, ppv)
        }),
        Transcode::CLSID => ClassFactory::new(|iid, ppv| unsafe {
            ComObject::new(Transcode::new())
                .as_interface::<IUnknown>()
                .query(iid, ppv)
        }),
        _ => return CLASS_E_CLASSNOTAVAILABLE,
    };

    ComObject::new(class_factory)
        .as_interface::<IUnknown>()
        .query(iid, ppv)
}
