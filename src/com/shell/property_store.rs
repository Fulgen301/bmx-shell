use std::sync::RwLock;

use windows::core::PROPVARIANT;
use windows::Win32::Foundation::{E_OUTOFMEMORY, S_FALSE};
use windows::Win32::Storage::EnhancedStorage::{PKEY_Image_Compression, PKEY_MIMEType};
use windows::Win32::System::Com::CoTaskMemAlloc;
use windows::Win32::System::Variant::VT_LPWSTR;
use windows::{
    core::{implement, w, Interface, HRESULT, PCWSTR},
    Win32::{
        Foundation::{ERROR_ALREADY_INITIALIZED, E_INVALIDARG, E_UNEXPECTED, STG_E_ACCESSDENIED},
        Storage::EnhancedStorage::{
            PKEY_Image_BitDepth, PKEY_Image_CompressionText, PKEY_Image_Dimensions,
            PKEY_Image_HorizontalSize, PKEY_Image_VerticalSize,
        },
        System::Com::{IStream, STGM_READ, STGM_WRITE},
        UI::Shell::PropertiesSystem::{
            IInitializeWithStream, IInitializeWithStream_Impl, IPropertyStore, IPropertyStoreCache,
            IPropertyStoreCapabilities, IPropertyStoreCapabilities_Impl, IPropertyStore_Impl,
            PSCreateMemoryPropertyStore, PROPERTYKEY, PSC_READONLY,
        },
    },
};
use windows_core::{GUID, HSTRING};

use crate::com::wic::com::MIME_TYPE;
use crate::com::CoClass;
use crate::util::guid;
use crate::{bmx::FileHeader, com::FileHeaderExt};

fn propvariant_init_lpwstr(string: PCWSTR) -> windows::core::Result<PROPVARIANT> {
    if string.is_null() {
        return Err(E_INVALIDARG.into());
    }

    let size = unsafe { string.len() } + 1;
    let buffer = unsafe { CoTaskMemAlloc(size * std::mem::size_of::<usize>()) };

    if buffer.is_null() {
        return Err(E_OUTOFMEMORY.into());
    }

    unsafe {
        string.as_ptr().copy_to_nonoverlapping(buffer.cast(), size);
    }

    unsafe fn propvariant_impl_helper_cast<'a, T: Sized>(
        propvariant: &'a PROPVARIANT,
        _helper: &'a T,
    ) -> T {
        std::mem::transmute_copy::<PROPVARIANT, T>(propvariant)
    }

    let propvar_impl = PROPVARIANT::new();
    let mut propvar_impl =
        unsafe { propvariant_impl_helper_cast(&propvar_impl, propvar_impl.as_raw()) };

    propvar_impl.Anonymous.Anonymous.vt = VT_LPWSTR.0;
    propvar_impl.Anonymous.Anonymous.Anonymous.pwszVal = buffer.cast();

    Ok(unsafe { PROPVARIANT::from_raw(propvar_impl) })
}

fn propvariant_init_string<T: AsRef<str>>(string: T) -> windows::core::Result<PROPVARIANT> {
    propvariant_init_lpwstr(PCWSTR::from_raw(HSTRING::from(string.as_ref()).as_ptr()))
}

struct PropertyStoreData {
    properties: IPropertyStoreCache,
}

#[derive(Default)]
#[implement(IPropertyStore, IPropertyStoreCapabilities, IInitializeWithStream)]
pub struct PropertyStore {
    inner: RwLock<Option<PropertyStoreData>>,
}

impl PropertyStore {
    pub const PREVIEW_DETAILS: PCWSTR =
        w!("prop:System.Image.Dimensions;System.Image.BitDepth;System.Image.Compression");

    pub fn new() -> Self {
        Self {
            inner: RwLock::new(None),
        }
    }

    fn initialize_from_header(
        &self,
        header: FileHeader,
    ) -> windows::core::Result<IPropertyStoreCache> {
        let properties = unsafe {
            let mut property_store = std::ptr::null_mut();
            PSCreateMemoryPropertyStore(&IPropertyStoreCache::IID, &raw mut property_store)?;

            IPropertyStoreCache::from_raw(property_store)
        };

        macro_rules! set_property {
            ($key:ident = $value:expr) => {
                unsafe {
                    properties.SetValueAndState(&$key, &PROPVARIANT::from($value), PSC_READONLY)?
                }
            };
        }

        macro_rules! set_properties {
            ($key: ident = $value:expr) => {
                set_property!($key = $value);
            };

            ($key: ident = $value:expr, $($rest:tt)*) => {
                set_property!($key = $value);
                set_properties!($($rest)*);
            };
        }

        set_properties!(
            PKEY_MIMEType = propvariant_init_lpwstr(MIME_TYPE)?,
            PKEY_Image_BitDepth = header.bit_depth as u32,
            PKEY_Image_Dimensions =
                propvariant_init_string(format!("{}x{}", header.width, header.height))?,
            PKEY_Image_HorizontalSize = header.width as u32,
            PKEY_Image_VerticalSize = header.height as u32
        );

        match header.compressed {
            0 => {
                set_properties!(PKEY_Image_Compression = 1u16);
            }
            1 => {
                set_properties!(
                    PKEY_Image_Compression = u16::MAX - 1,
                    PKEY_Image_CompressionText = propvariant_init_lpwstr(w!("LZSA"))?
                );
            }
            _ => {
                set_properties!(
                    PKEY_Image_Compression = u16::MAX,
                    PKEY_Image_CompressionText = propvariant_init_lpwstr(w!("Unknown"))?
                );
            }
        }

        Ok(properties)
    }

    fn with_property_store<F, R>(&self, op: F) -> windows::core::Result<R>
    where
        F: FnOnce(&IPropertyStoreCache) -> windows::core::Result<R>,
    {
        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        op(&inner.properties)
    }
}

impl CoClass for PropertyStore {
    const CLSID: GUID = guid::from_str("04f579e2-ace3-481c-81ee-f153ffd42551");
    const PROG_ID: PCWSTR = w!("X16BMX.PropertyStore.1");
    const VERSION_INDEPENDENT_PROG_ID: PCWSTR = w!("X16BMX.PropertyStore");
}

impl IPropertyStore_Impl for PropertyStore_Impl {
    fn GetCount(&self) -> windows::core::Result<u32> {
        self.with_property_store(|properties| unsafe { properties.GetCount() })
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn GetAt(&self, index: u32, key: *mut PROPERTYKEY) -> windows::core::Result<()> {
        self.with_property_store(
            #[allow(clippy::not_unsafe_ptr_arg_deref)]
            |properties| unsafe { properties.GetAt(index, key) },
        )
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn GetValue(
        &self,
        key: *const PROPERTYKEY,
    ) -> windows::core::Result<windows::core::PROPVARIANT> {
        self.with_property_store(|properties| unsafe { properties.GetValue(key) })
    }

    fn SetValue(
        &self,
        _key: *const PROPERTYKEY,
        _value: *const PROPVARIANT,
    ) -> windows::core::Result<()> {
        Err(STG_E_ACCESSDENIED.into())
    }

    fn Commit(&self) -> windows::core::Result<()> {
        Err(STG_E_ACCESSDENIED.into())
    }
}

impl IPropertyStoreCapabilities_Impl for PropertyStore_Impl {
    fn IsPropertyWritable(&self, _key: *const PROPERTYKEY) -> windows::core::Result<()> {
        Err(windows::core::Error::new(S_FALSE, ""))
    }
}

impl IInitializeWithStream_Impl for PropertyStore_Impl {
    fn Initialize(&self, stream: Option<&IStream>, grfmode: u32) -> windows::core::Result<()> {
        if grfmode & (STGM_READ.0 | STGM_WRITE.0) != 0 {
            return Err(STG_E_ACCESSDENIED.into());
        }

        let stream = stream.ok_or(E_INVALIDARG)?;

        let mut inner = self.inner.write().unwrap();

        if inner.is_some() {
            return Err(HRESULT::from_win32(ERROR_ALREADY_INITIALIZED.0).into());
        }

        let header = FileHeader::from_stream(stream)?;
        let properties = self.initialize_from_header(header)?;

        inner.replace(PropertyStoreData { properties });

        Ok(())
    }
}
