use std::ops::Deref;

use transaction::{Key, Transaction};
use windows::Win32::{
    Foundation::E_BLUETOOTH_ATT_ATTRIBUTE_NOT_FOUND,
    Graphics::Imaging::{
        CATID_WICBitmapDecoders, CATID_WICBitmapEncoders, GUID_WICPixelFormat1bppIndexed,
        GUID_WICPixelFormat2bppIndexed, GUID_WICPixelFormat4bppIndexed,
        GUID_WICPixelFormat8bppIndexed,
    },
    System::Registry::HKEY_LOCAL_MACHINE,
    UI::Shell::{IThumbnailProvider, SHChangeNotify, SHCNE_ASSOCCHANGED, SHCNF_FLAGS},
};
use windows_core::{w, Interface, PCWSTR};

use crate::{
    com::{
        shell::{command::transcode::Transcode, property_store::PropertyStore},
        wic::{
            com::{CONTAINER_FORMAT, EXTENSION, MIME_TYPE, PREVIEW_DETAILS, PROG_ID, VENDOR},
            decoder::BitmapDecoder,
            encoder::BitmapEncoder,
        },
        CoClass,
    },
    util::guid::GuidExt,
};

pub mod transaction {
    use std::cell::Cell;

    use crate::util::guid::GuidExt;

    use windows::{
        core::{w, Owned, GUID, PCWSTR},
        Win32::{
            Foundation::{
                ERROR_FILE_NOT_FOUND, ERROR_SUCCESS, E_ILLEGAL_STATE_CHANGE, HANDLE, WIN32_ERROR,
            },
            Storage::FileSystem::{CommitTransaction, CreateTransaction, RollbackTransaction},
            System::{
                Registry::{
                    RegCreateKeyTransactedW, RegDeleteTreeW, RegDeleteValueW,
                    RegOpenKeyTransactedW, HKEY, KEY_READ, KEY_WRITE, REG_BINARY, REG_DWORD,
                    REG_EXPAND_SZ, REG_OPEN_CREATE_OPTIONS, REG_OPTION_NON_VOLATILE,
                    REG_OPTION_VOLATILE, REG_QWORD, REG_SZ, REG_VALUE_TYPE,
                },
                Threading::INFINITE,
            },
        },
    };

    pub struct Transaction {
        handle: Owned<HANDLE>,
        key_options: REG_OPEN_CREATE_OPTIONS,
        committed: Cell<bool>,
    }

    impl Transaction {
        pub fn new(volatile: bool) -> windows::core::Result<Self> {
            Ok(Self {
                handle: unsafe {
                    Owned::new(CreateTransaction(
                        std::ptr::null_mut(),
                        std::ptr::null_mut(),
                        0,
                        0,
                        0,
                        INFINITE,
                        w!("bmx-shell"),
                    )?)
                },
                key_options: if volatile {
                    REG_OPTION_VOLATILE
                } else {
                    REG_OPTION_NON_VOLATILE
                },

                committed: Cell::new(false),
            })
        }

        pub fn commit(&self) -> windows::core::Result<()> {
            if self.committed.get() {
                return Err(E_ILLEGAL_STATE_CHANGE.into());
            }

            unsafe {
                CommitTransaction(*self.handle)?;
            }

            self.committed.replace(true);
            Ok(())
        }
    }

    impl Drop for Transaction {
        fn drop(&mut self) {
            if !self.committed.get() {
                unsafe {
                    let _ = RollbackTransaction(*self.handle);
                }
            }
        }
    }

    unsafe fn reg_create_key_transacted(
        key: HKEY,
        sub_key: PCWSTR,
        options: REG_OPEN_CREATE_OPTIONS,
        transaction: HANDLE,
    ) -> windows::core::Result<HKEY> {
        let mut result = HKEY::default();

        unsafe {
            RegCreateKeyTransactedW(
                key,
                sub_key,
                0,
                None,
                options,
                KEY_READ | KEY_WRITE,
                None,
                &raw mut result,
                None,
                transaction,
                None,
            )
            .ok()?;
        }

        Ok(result)
    }

    #[allow(unused)]
    unsafe fn open_key_transacted(
        key: HKEY,
        sub_key: PCWSTR,
        transaction: HANDLE,
    ) -> windows::core::Result<HKEY> {
        let mut result = HKEY::default();

        unsafe {
            RegOpenKeyTransactedW(
                key,
                sub_key,
                0,
                KEY_READ | KEY_WRITE,
                &raw mut result,
                transaction,
                None,
            )
            .ok()?;
        }

        Ok(result)
    }

    pub struct Key<'a> {
        transaction: &'a Transaction,
        key: Owned<HKEY>,
    }

    impl<'a> Key<'a> {
        pub fn predefined(
            transaction: &'a Transaction,
            key: HKEY,
            sub_key: PCWSTR,
        ) -> windows::core::Result<Self> {
            let mut result = HKEY::default();

            unsafe {
                RegCreateKeyTransactedW(
                    key,
                    sub_key,
                    0,
                    None,
                    transaction.key_options,
                    KEY_READ | KEY_WRITE,
                    None,
                    &raw mut result,
                    None,
                    *transaction.handle,
                    None,
                )
                .ok()?;
            }

            Ok(Self {
                transaction,
                key: unsafe {
                    Owned::new(reg_create_key_transacted(
                        key,
                        sub_key,
                        transaction.key_options,
                        *transaction.handle,
                    )?)
                },
            })
        }

        pub fn create_subkey(&self, sub_key: PCWSTR) -> windows::core::Result<Key<'a>> {
            Ok(Self {
                transaction: self.transaction,
                key: unsafe {
                    Owned::new(reg_create_key_transacted(
                        *self.key,
                        sub_key,
                        self.transaction.key_options,
                        *self.transaction.handle,
                    )?)
                },
            })
        }

        #[allow(unused)]
        pub fn open_subkey(&self, sub_key: PCWSTR) -> windows::core::Result<Key<'a>> {
            Ok(Self {
                transaction: self.transaction,
                key: unsafe {
                    Owned::new(open_key_transacted(
                        *self.key,
                        sub_key,
                        *self.transaction.handle,
                    )?)
                },
            })
        }

        pub fn delete_subkey(&self, subkey: PCWSTR) -> windows::core::Result<()> {
            self.delete_tree_internal(subkey)
        }

        pub fn delete_tree(&self) -> windows::core::Result<()> {
            self.delete_tree_internal(PCWSTR::null())
        }

        fn delete_tree_internal(&self, subkey: PCWSTR) -> windows::core::Result<()> {
            match unsafe { RegDeleteTreeW(*self.key, subkey) } {
                ERROR_SUCCESS | ERROR_FILE_NOT_FOUND => Ok(()),
                e => e.ok(),
            }
        }

        pub fn set_u32(&self, name: PCWSTR, value: u32) -> windows::core::Result<()> {
            self.set_value(name, Some(&value.to_le_bytes()), REG_DWORD)
        }

        #[allow(unused)]
        pub fn set_u64(&self, name: PCWSTR, value: u64) -> windows::core::Result<()> {
            self.set_value(name, Some(&value.to_le_bytes()), REG_QWORD)
        }

        pub fn set_binary(&self, name: PCWSTR, value: &[u8]) -> windows::core::Result<()> {
            self.set_value(name, Some(value), REG_BINARY)
        }

        #[allow(unused)]
        pub fn set_str(&self, name: PCWSTR, value: &str) -> windows::core::Result<()> {
            self.set_value(
                name,
                Some(&value.encode_utf16().collect::<Vec<_>>()),
                REG_SZ,
            )
        }

        #[allow(unused)]
        pub fn set_str_expand(&self, name: PCWSTR, value: &str) -> windows::core::Result<()> {
            self.set_value(
                name,
                Some(&value.encode_utf16().collect::<Vec<_>>()),
                REG_EXPAND_SZ,
            )
        }

        pub fn set_pcwstr(&self, name: PCWSTR, value: PCWSTR) -> windows::core::Result<()> {
            self.set_value(
                name,
                if value.is_null() {
                    None
                } else {
                    Some(unsafe { value.as_wide() })
                },
                REG_SZ,
            )
        }

        pub fn set_pcwstr_expand(&self, name: PCWSTR, value: PCWSTR) -> windows::core::Result<()> {
            self.set_value(
                name,
                if value.is_null() {
                    None
                } else {
                    Some(unsafe { value.as_wide() })
                },
                REG_EXPAND_SZ,
            )
        }

        pub fn set_guid(&self, name: PCWSTR, value: &GUID) -> windows::core::Result<()> {
            self.set_value(name, Some(&value.to_wide()), REG_SZ)
        }

        fn set_value<T>(
            &self,
            name: PCWSTR,
            value: Option<&[T]>,
            value_type: REG_VALUE_TYPE,
        ) -> windows::core::Result<()> {
            unsafe extern "system" {
                #[allow(unused)]
                fn RegSetValueExW(
                    hkey: HKEY,
                    lpvaluename: PCWSTR,
                    reserved: u32,
                    dwtype: REG_VALUE_TYPE,
                    lpdata: *const u8,
                    cbdata: u32,
                ) -> WIN32_ERROR;
            }

            unsafe {
                RegSetValueExW(
                    *self.key,
                    name,
                    0,
                    value_type,
                    value.map_or(std::ptr::null(), |v| v.as_ptr().cast()),
                    (value.map_or(0, |v| v.len()) * std::mem::size_of::<T>()) as u32,
                )
                .ok()
            }
        }

        pub fn delete_value(&self, name: PCWSTR) -> windows::core::Result<()> {
            match unsafe { RegDeleteValueW(*self.key, name) } {
                ERROR_SUCCESS | ERROR_FILE_NOT_FOUND => Ok(()),
                e => e.ok(),
            }
        }
    }
}

#[derive(Clone, Copy)]
struct NullTerminatedSlice<'a>(&'a [u16]);

impl<'a> NullTerminatedSlice<'a> {
    pub fn new(slice: &'a [u16]) -> Result<Self, ()> {
        if slice.last() != Some(&0u16) {
            Err(())
        } else {
            Ok(Self(slice))
        }
    }
}

impl Deref for NullTerminatedSlice<'_> {
    type Target = [u16];

    fn deref(&self) -> &Self::Target {
        self.0
    }
}

fn register_com_extension<'a, T: CoClass>(
    classes: &'a Key,
    module_path: NullTerminatedSlice,
    description: PCWSTR,
    apartment_type: PCWSTR,
) -> windows::core::Result<Key<'a>> {
    let clsid_string = T::CLSID.to_wide();
    let com_object = classes
        .create_subkey(w!("CLSID"))?
        .create_subkey(PCWSTR::from_raw(clsid_string.as_ptr()))?;

    com_object.set_pcwstr(PCWSTR::null(), description)?;

    com_object
        .create_subkey(w!("ProgId"))?
        .set_pcwstr(PCWSTR::null(), T::PROG_ID)?;

    com_object
        .create_subkey(w!("VersionIndependentProgId"))?
        .set_pcwstr(PCWSTR::null(), T::VERSION_INDEPENDENT_PROG_ID)?;

    let inproc = com_object.create_subkey(w!("InprocServer32"))?;
    inproc.set_pcwstr(PCWSTR::null(), PCWSTR::from_raw(module_path.as_ptr()))?;
    inproc.set_pcwstr(w!("ThreadingModel"), apartment_type)?;

    classes
        .create_subkey(T::PROG_ID)?
        .create_subkey(w!("CLSID"))?
        .set_guid(PCWSTR::null(), &T::CLSID)?;

    classes
        .create_subkey(T::VERSION_INDEPENDENT_PROG_ID)?
        .create_subkey(w!("CLSID"))?
        .set_guid(PCWSTR::null(), &T::CLSID)?;

    Ok(com_object)
}

fn unregister_com_extension<T: CoClass>(classes: &Key) -> windows::core::Result<()> {
    let mut buffer = [0u16; 39 + 6];
    unsafe {
        buffer[..6]
            .as_mut_ptr()
            .copy_from_nonoverlapping(w!("CLSID\\").as_ptr(), 6);
    }

    let clsid_string = T::CLSID.to_wide();
    buffer[6..].copy_from_slice(&clsid_string);
    classes.delete_subkey(PCWSTR::from_raw(buffer.as_ptr()))?;

    classes.delete_subkey(T::PROG_ID)?;
    classes.delete_subkey(T::VERSION_INDEPENDENT_PROG_ID)?;
    Ok(())
}

pub fn register_server<'a>(
    transaction: &'a Transaction,
    classes_root: &'a Key,
    module_path: &[u16],
) -> windows::core::Result<()> {
    let module_path = NullTerminatedSlice::new(module_path)
        .map_err(|_| windows::core::Error::from(E_BLUETOOTH_ATT_ATTRIBUTE_NOT_FOUND))?;

    {
        let prog_id = classes_root.create_subkey(PROG_ID)?;
        prog_id.set_pcwstr(PCWSTR::null(), w!("BMX File"))?;

        let drop_target = prog_id.create_subkey(w!("DropTarget"))?;
        drop_target.set_pcwstr(PCWSTR::null(), w!("{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}"))?;

        let shell = prog_id.create_subkey(w!("shell"))?;

        {
            let open = shell.create_subkey(w!("open"))?;
            open.set_pcwstr_expand(
                w!("MuiVerb"),
                w!("@%PROGRAMFILES%\\Windows Photo Viewer\\photoviewer.dll,-3043"),
            )?;

            let command = open.create_subkey(w!("command"))?;
            command.set_pcwstr_expand(PCWSTR::null(),  w!("%SystemRoot%\\System32\\rundll32.exe \"%ProgramFiles%\\Windows Photo Viewer\\PhotoViewer.dll\", ImageView_Fullscreen %1"))?;
        }

        {
            let printto = shell.create_subkey(w!("printto"))?;
            let command = printto.create_subkey(w!("command"))?;
            command.set_pcwstr_expand(w!("Name"), w!("%SystemRoot%\\System32\\rundll32.exe \"%SystemRoot%\\System32\\shimgvw.dll\", ImageView_PrintTo /pt \"%1\" \"%2\" \"%3\" \"%4\""))?;
        }

        let shellex = prog_id.create_subkey(w!("ShellEx"))?;
        let thumbnail_provider =
            shellex.create_subkey(PCWSTR::from_raw(IThumbnailProvider::IID.to_wide().as_ptr()))?;
        thumbnail_provider
            .set_pcwstr(PCWSTR::null(), w!("{C7657C4A-9F68-40fa-A4DF-96BC08EB3551}"))?;
    }

    {
        let bmx_decoder = register_com_extension::<BitmapDecoder>(
            classes_root,
            module_path,
            w!("BMX File"),
            w!("Both"),
        )?;

        bmx_decoder.set_pcwstr(w!("Author"), w!("Fulgen"))?;
        bmx_decoder.set_guid(w!("ContainerFormat"), &CONTAINER_FORMAT)?;
        bmx_decoder.set_pcwstr(w!("Description"), w!("BMX Decoder"))?;
        bmx_decoder.set_pcwstr(w!("FileExtensions"), EXTENSION)?;
        bmx_decoder.set_pcwstr(w!("FriendlyName"), w!("BMX Decoder"))?;
        bmx_decoder.set_pcwstr(w!("MimeTypes"), MIME_TYPE)?;
        bmx_decoder.set_u32(w!("SupportLossless"), 1)?;
        bmx_decoder.set_guid(w!("VendorGUID"), &VENDOR)?;

        let formats = bmx_decoder.create_subkey(w!("Formats"))?;
        _ = formats.create_subkey(PCWSTR::from_raw(
            GUID_WICPixelFormat1bppIndexed.to_wide().as_ptr(),
        ))?;
        _ = formats.create_subkey(PCWSTR::from_raw(
            GUID_WICPixelFormat2bppIndexed.to_wide().as_ptr(),
        ))?;
        _ = formats.create_subkey(PCWSTR::from_raw(
            GUID_WICPixelFormat4bppIndexed.to_wide().as_ptr(),
        ))?;
        _ = formats.create_subkey(PCWSTR::from_raw(
            GUID_WICPixelFormat8bppIndexed.to_wide().as_ptr(),
        ))?;

        let patterns = bmx_decoder.create_subkey(w!("Patterns"))?;
        let first_pattern = patterns.create_subkey(w!("0"))?;
        first_pattern.set_u32(w!("Position"), 0)?;

        first_pattern.set_binary(w!("Pattern"), b"BMX\x01")?;
        first_pattern.set_binary(w!("Mask"), &[0xFF, 0xFF, 0xFF, 0xFF])?;
        first_pattern.set_u32(w!("Length"), 4)?;
    }

    {
        let category = classes_root
            .create_subkey(w!("CLSID"))?
            .create_subkey(PCWSTR::from_raw(CATID_WICBitmapDecoders.to_wide().as_ptr()))?;

        let instance = category.create_subkey(w!("Instance"))?;

        let bmx_decoder =
            instance.create_subkey(PCWSTR::from_raw(BitmapDecoder::CLSID.to_wide().as_ptr()))?;
        bmx_decoder.set_guid(w!("CLSID"), &BitmapDecoder::CLSID)?;
        bmx_decoder.set_pcwstr(w!("FriendlyName"), w!("BMX Decoder"))?;
    }

    {
        let bmx_encoder = register_com_extension::<BitmapEncoder>(
            classes_root,
            module_path,
            w!("BMX File"),
            w!("Both"),
        )?;

        bmx_encoder.set_pcwstr(w!("Author"), w!("Fulgen"))?;
        bmx_encoder.set_guid(w!("ContainerFormat"), &CONTAINER_FORMAT)?;
        bmx_encoder.set_pcwstr(w!("Description"), w!("BMX Encoder"))?;
        bmx_encoder.set_pcwstr(w!("FileExtensions"), EXTENSION)?;
        bmx_encoder.set_pcwstr(w!("FriendlyName"), w!("BMX Encoder"))?;
        bmx_encoder.set_pcwstr(w!("MimeTypes"), MIME_TYPE)?;
        bmx_encoder.set_u32(w!("SupportLossless"), 1)?;
        bmx_encoder.set_guid(w!("VendorGUID"), &VENDOR)?;

        let formats = bmx_encoder.create_subkey(w!("Formats"))?;
        _ = formats.create_subkey(PCWSTR::from_raw(
            GUID_WICPixelFormat1bppIndexed.to_wide().as_ptr(),
        ))?;
        _ = formats.create_subkey(PCWSTR::from_raw(
            GUID_WICPixelFormat2bppIndexed.to_wide().as_ptr(),
        ))?;
        _ = formats.create_subkey(PCWSTR::from_raw(
            GUID_WICPixelFormat4bppIndexed.to_wide().as_ptr(),
        ))?;
        _ = formats.create_subkey(PCWSTR::from_raw(
            GUID_WICPixelFormat8bppIndexed.to_wide().as_ptr(),
        ))?;
    }

    {
        let category = classes_root
            .create_subkey(w!("CLSID"))?
            .create_subkey(PCWSTR::from_raw(CATID_WICBitmapEncoders.to_wide().as_ptr()))?;

        let instance = category.create_subkey(w!("Instance"))?;

        let bmx_encoder =
            instance.create_subkey(PCWSTR::from_raw(BitmapEncoder::CLSID.to_wide().as_ptr()))?;
        bmx_encoder.set_guid(w!("CLSID"), &BitmapEncoder::CLSID)?;
        bmx_encoder.set_pcwstr(w!("FriendlyName"), w!("BMX Encoder"))?;
    }

    {
        let bmx = classes_root.create_subkey(EXTENSION)?;
        bmx.set_pcwstr(PCWSTR::null(), PROG_ID)?;
        bmx.set_pcwstr(w!("Content Type"), MIME_TYPE)?;
        bmx.set_pcwstr(w!("PerceivedType"), w!("image"))?;

        let open_with_list = bmx.create_subkey(w!("OpenWithList"))?;
        _ = open_with_list.create_subkey(w!("PhotoViewer.dll"))?;
    }

    {
        let systems_file_associations = classes_root.create_subkey(w!("SystemFileAssociations"))?;
        let bmx = systems_file_associations.create_subkey(EXTENSION)?;
        bmx.set_pcwstr(w!("PreviewDetails"), PREVIEW_DETAILS)?;

        let open_with_list = bmx.create_subkey(w!("OpenWithList"))?;
        _ = open_with_list.create_subkey(w!("PhotoViewer.dll"))?;

        let shellex = bmx.create_subkey(w!("ShellEx"))?;
        let thumbnail_provider =
            shellex.create_subkey(PCWSTR::from_raw(IThumbnailProvider::IID.to_wide().as_ptr()))?;
        thumbnail_provider
            .set_pcwstr(PCWSTR::null(), w!("{C7657C4A-9F68-40fa-A4DF-96BC08EB3551}"))?;

        let context_menu_handlers = bmx.create_subkey(w!("ContextMenuHandlers"))?;
        let shell_image_preview = context_menu_handlers.create_subkey(w!("ShellImagePreview"))?;
        shell_image_preview
            .set_pcwstr(PCWSTR::null(), w!("{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}"))?;
    }

    {
        let current_version = Key::predefined(
            transaction,
            HKEY_LOCAL_MACHINE,
            w!("Software\\Microsoft\\Windows\\CurrentVersion"),
        )?;

        let kind_map = current_version.create_subkey(w!("Explorer\\KindMap"))?;
        kind_map.set_pcwstr(EXTENSION, w!("Picture"))?;
    }

    {
        let _property_store = register_com_extension::<PropertyStore>(
            classes_root,
            module_path,
            w!("BMXPropertyStore"),
            w!("Both"),
        );

        let property_handlers = Key::predefined(
            transaction,
            HKEY_LOCAL_MACHINE,
            w!("Software\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers"),
        )?;

        let bmx = property_handlers.create_subkey(EXTENSION)?;
        bmx.set_guid(PCWSTR::null(), &PropertyStore::CLSID)?;
    }

    {
        let _transcode = register_com_extension::<Transcode>(
            classes_root,
            module_path,
            w!("Transcode"),
            w!("Both"),
        );

        let file = classes_root.create_subkey(w!("*"))?;
        let shell = file.create_subkey(w!("shell"))?;
        let transcode = shell.create_subkey(w!("Transcode"))?;
        //transcode.set_pcwstr(w!("AppliesTo"), w!("System.Kind:picture"))?;
        transcode.set_guid(w!("ExplorerCommandHandler"), &Transcode::CLSID)?;
    }

    transaction
        .commit()
        .map(|_| unsafe { SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_FLAGS(0), None, None) })?;

    Ok(())
}

pub fn unregister_server<'a>(
    transaction: &'a Transaction,
    classes_root: &'a Key,
) -> windows::core::Result<()> {
    classes_root.delete_subkey(PROG_ID)?;

    unregister_com_extension::<BitmapDecoder>(classes_root)?;
    unregister_com_extension::<BitmapEncoder>(classes_root)?;
    unregister_com_extension::<PropertyStore>(classes_root)?;

    let clsid = classes_root.open_subkey(w!("CLSID"))?;

    clsid
        .open_subkey(PCWSTR::from_raw(CATID_WICBitmapDecoders.to_wide().as_ptr()))?
        .open_subkey(w!("Instance"))?
        .delete_subkey(PCWSTR::from_raw(BitmapDecoder::CLSID.to_wide().as_ptr()))?;

    clsid
        .open_subkey(PCWSTR::from_raw(CATID_WICBitmapEncoders.to_wide().as_ptr()))?
        .open_subkey(w!("Instance"))?
        .delete_subkey(PCWSTR::from_raw(BitmapEncoder::CLSID.to_wide().as_ptr()))?;

    classes_root.delete_subkey(EXTENSION)?;

    classes_root
        .open_subkey(w!("SystemFileAssociations"))?
        .delete_subkey(EXTENSION)?;

    Key::predefined(
        transaction,
        HKEY_LOCAL_MACHINE,
        w!("Software\\Microsoft\\Windows\\CurrentVersion\\KindMap"),
    )?
    .delete_value(EXTENSION)?;

    Key::predefined(
        transaction,
        HKEY_LOCAL_MACHINE,
        w!("Software\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers"),
    )?
    .delete_subkey(EXTENSION)?;

    transaction
        .commit()
        .map(|_| unsafe { SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_FLAGS(0), None, None) })?;

    Ok(())
}
