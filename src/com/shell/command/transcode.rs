use std::ffi::c_void;
use std::fmt::Display;
use std::mem::MaybeUninit;
use std::sync::{Mutex, RwLock};

#[allow(unused)]
use windows::core::{implement, ComObject, IUnknownImpl, Interface, GUID, HRESULT};
use windows::core::{w, Array, IUnknown, HSTRING, PCWSTR, PROPVARIANT, PWSTR};
use windows::Win32::Foundation::{
    BOOL, ERROR_ALREADY_INITIALIZED, ERROR_NO_MORE_ITEMS, E_FAIL, E_INVALIDARG, E_NOTIMPL,
    E_POINTER, E_UNEXPECTED, HWND, S_FALSE, S_OK, WINCODEC_ERR_UNSUPPORTEDOPERATION,
};
use windows::Win32::Graphics::Imaging::{
    IWICBitmapCodecInfo, IWICImagingFactory, WICBitmapEncoderNoCache, WICComponentEnumerateDefault,
    WICConvertBitmapSource, WICDecodeMetadataCacheOnDemand, WICDecoder, WICEncoder,
};
use windows::Win32::Storage::EnhancedStorage::{PKEY_Kind, PKEY_MIMEType};
use windows::Win32::Storage::FileSystem::FILE_ATTRIBUTE_NORMAL;
use windows::Win32::System::Com::StructuredStorage::IPropertyBag;
use windows::Win32::System::Com::Urlmon::E_PENDING;
use windows::Win32::System::Com::{
    CoCreateInstance, CreateBindCtx, IBindCtx, IEnumUnknown, IStream, BIND_OPTS,
    CLSCTX_INPROC_SERVER, STGM_WRITE,
};
use windows::Win32::System::Diagnostics::Debug::OutputDebugStringW;
use windows::Win32::System::Ole::{IObjectWithSite, IObjectWithSite_Impl};
use windows::Win32::System::Variant::{VT_LPWSTR, VT_VECTOR};
use windows::Win32::UI::Shell::Common::COMDLG_FILTERSPEC;
use windows::Win32::UI::Shell::PropertiesSystem::IPropertyStore;
use windows::Win32::UI::Shell::{
    BHID_PropertyStore, BHID_Stream, FileOpenDialog, FileOperation, FileSaveDialog,
    IEnumExplorerCommand, IEnumExplorerCommand_Impl, IExplorerCommand, IExplorerCommand_Impl,
    IFileDialog, IFileDialogControlEvents, IFileDialogControlEvents_Impl, IFileDialogCustomize,
    IFileDialogEvents, IFileDialogEvents_Impl, IFileOperation, IFileOperationProgressSink,
    IFileOperationProgressSink_Impl, IInitializeCommand, IInitializeCommand_Impl, IShellItem,
    IShellItemArray, IUnknown_GetWindow, SHGetFileInfoW, SHStrDupW, ECF_DEFAULT,
    ECF_HASSUBCOMMANDS, ECF_ISDROPDOWN, ECS_ENABLED, ECS_HIDDEN, FDE_OVERWRITE_RESPONSE,
    FDE_SHAREVIOLATION_RESPONSE, FOS_PICKFOLDERS, FOS_STRICTFILETYPES, SHFILEINFOW, SHGFI_TYPENAME,
    SHGFI_USEFILEATTRIBUTES, SIGDN_PARENTRELATIVEPARSING,
};
use windows::Win32::UI::WindowsAndMessaging::{MessageBoxW, MB_ICONERROR};

use crate::com::shell::CoTaskMemPWSTR;
use crate::com::wic::{
    codec_mime_types, create_imaging_factory, get_component_iterator, pixel_format_friendly_name,
    pixel_format_is_known,
};
use crate::com::CoClass;
use crate::get_with_buffer;

fn pcwstr_is_equal_to_slice_no_case(first: PCWSTR, second: &[u16]) -> bool {
    unsafe extern "C" {
        fn _wcsnicmp(a: *const u16, b: *const u16, count: usize) -> i32;
    }

    unsafe { _wcsnicmp(first.as_ptr(), second.as_ptr(), second.len()) == 0 }
}

fn pcwstr_is_equal_to_pcwstr_no_case(first: PCWSTR, second: PCWSTR) -> bool {
    unsafe extern "C" {
        fn _wcsicmp(a: *const u16, b: *const u16) -> i32;
    }

    unsafe { _wcsicmp(first.as_ptr(), second.as_ptr()) == 0 }
}

fn propvariant_to_lpwstr(variant: &PROPVARIANT) -> Option<PWSTR> {
    unsafe {
        let variant = variant.as_raw();
        if variant.Anonymous.Anonymous.vt == VT_LPWSTR.0 {
            Some(PWSTR::from_raw(
                variant.Anonymous.Anonymous.Anonymous.pwszVal,
            ))
        } else {
            None
        }
    }
}

fn propvariant_to_lpwstr_slice<'a>(variant: &PROPVARIANT) -> Option<&'a [PWSTR]> {
    unsafe {
        let variant = variant.as_raw();
        if variant.Anonymous.Anonymous.vt == (VT_VECTOR | VT_LPWSTR).0 {
            Some(std::slice::from_raw_parts(
                variant.Anonymous.Anonymous.Anonymous.caub.pElems as *const PWSTR,
                variant.Anonymous.Anonymous.Anonymous.caub.cElems as _,
            ))
        } else {
            None
        }
    }
}

fn debug_output<S: AsRef<str>>(s: S) {
    let mut string = s.as_ref().to_owned();
    string.push('\n');
    unsafe {
        OutputDebugStringW(PCWSTR::from_raw(HSTRING::from(string).as_ptr()));
    }
}

fn item_array_has_matching_decoders(
    items: &IShellItemArray,
    imaging_factory: &IWICImagingFactory,
) -> windows::core::Result<bool> {
    let count = unsafe { items.GetCount()? };

    for i in 0..count {
        let item = unsafe { items.GetItemAt(i)? };

        let properties: IPropertyStore = unsafe { item.BindToHandler(None, &BHID_PropertyStore)? };

        let variant = unsafe { properties.GetValue(&PKEY_Kind)? };

        let Some(kind) = propvariant_to_lpwstr_slice(&variant) else {
            return Ok(false);
        };

        if !kind.iter().any(|kind| {
            pcwstr_is_equal_to_pcwstr_no_case(PCWSTR::from_raw(kind.as_ptr()), w!("picture"))
        }) {
            debug_output("no picture");
            return Ok(false);
        }

        let variant = unsafe { properties.GetValue(&PKEY_MIMEType)? };

        let Some(item_mime_type) = propvariant_to_lpwstr(&variant) else {
            debug_output("no mime type");
            return Ok(false);
        };

        let item_mime_type = PCWSTR::from_raw(item_mime_type.as_ptr());

        if get_component_iterator::<IWICBitmapCodecInfo>(
            imaging_factory,
            WICDecoder,
            WICComponentEnumerateDefault,
        )?
        .filter_map(|result| result.ok())
        .any(|decoder| {
            let Ok(pixel_formats) = get_with_buffer!(&decoder, GetPixelFormats) else {
                debug_output("no pixel formats for decoder");
                return false;
            };

            if !pixel_formats.iter().any(pixel_format_is_known) {
                debug_output("no known pixel formats for decoder");
                return false;
            }

            let Ok(mime_types) = codec_mime_types(&decoder) else {
                debug_output("no mime types for decoder");
                return false;
            };
            mime_types
                .split(|c| *c == b',' as u16)
                .any(|wic_mime_type| {
                    pcwstr_is_equal_to_slice_no_case(item_mime_type, wic_mime_type)
                })
        }) {
            debug_output("found decoder");
            return Ok(true);
        }
    }

    Ok(false)
}

struct TranscodeData {
    #[allow(unused)]
    command_name: String,
    #[allow(unused)]
    properties: IPropertyBag,
    imaging_factory: IWICImagingFactory,
    site: Option<IUnknown>,
}

#[derive(Default)]
#[implement(IExplorerCommand, IInitializeCommand, IObjectWithSite)]
pub struct Transcode {
    inner: RwLock<Option<TranscodeData>>,
}

impl Transcode {
    pub fn new() -> Self {
        Self {
            inner: RwLock::new(None),
        }
    }
}

impl CoClass for Transcode {
    const CLSID: GUID = GUID::from_u128(0xbe8b5162_693a_4d66_9efb_01ea923c1f4du128);
    const PROG_ID: PCWSTR = w!("X16BMX.Transcode.1");
    const VERSION_INDEPENDENT_PROG_ID: PCWSTR = w!("X16BMX.Transcode");
}

impl IExplorerCommand_Impl for Transcode_Impl {
    fn GetTitle(&self, _items: Option<&IShellItemArray>) -> windows::core::Result<PWSTR> {
        unsafe { SHStrDupW(w!("Transcode")) }
    }

    fn GetIcon(&self, _items: Option<&IShellItemArray>) -> windows::core::Result<PWSTR> {
        Err(E_NOTIMPL.into())
    }

    fn GetToolTip(&self, _items: Option<&IShellItemArray>) -> windows::core::Result<PWSTR> {
        unsafe { SHStrDupW(w!("Transcode an image format into another")) }
    }

    fn GetCanonicalName(&self) -> windows::core::Result<GUID> {
        Ok(Transcode::CLSID)
    }

    fn GetState(
        &self,
        items: Option<&IShellItemArray>,
        ok_to_be_slow: BOOL,
    ) -> windows::core::Result<u32> {
        let items = items.ok_or(E_POINTER)?;

        if !ok_to_be_slow.as_bool() {
            return Err(E_PENDING.into());
        }

        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        if item_array_has_matching_decoders(items, &inner.imaging_factory)? {
            Ok(ECS_ENABLED.0 as _)
        } else {
            Ok(ECS_HIDDEN.0 as _)
        }
    }

    fn Invoke(
        &self,
        _psiitemarray: Option<&IShellItemArray>,
        _pbc: Option<&IBindCtx>,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn GetFlags(&self) -> windows::core::Result<u32> {
        Ok((ECF_DEFAULT.0 | ECF_HASSUBCOMMANDS.0 | ECF_ISDROPDOWN.0) as _)
    }

    fn EnumSubCommands(&self) -> windows::core::Result<IEnumExplorerCommand> {
        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        Ok(ComObject::new(TranscodeEnumSubcommands::new(&inner.imaging_factory)?).to_interface())
    }
}

impl IInitializeCommand_Impl for Transcode_Impl {
    fn Initialize(
        &self,
        command_name: &windows::core::PCWSTR,
        property_bag: Option<&IPropertyBag>,
    ) -> windows::core::Result<()> {
        let mut inner = self.inner.write().unwrap();

        if inner.is_some() {
            return Err(HRESULT::from_win32(ERROR_ALREADY_INITIALIZED.0).into());
        }

        inner.replace(TranscodeData {
            command_name: unsafe { command_name.to_string().map_err(|_| E_INVALIDARG)? },
            properties: property_bag.ok_or(E_POINTER)?.clone(),
            imaging_factory: create_imaging_factory()?,
            site: None,
        });

        Ok(())
    }
}

impl IObjectWithSite_Impl for Transcode_Impl {
    fn SetSite(&self, site: Option<&IUnknown>) -> windows::core::Result<()> {
        let mut inner = self.inner.write().unwrap();
        let inner = inner.as_mut().ok_or(E_UNEXPECTED)?;
        inner.site = site.cloned();
        Ok(())
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn GetSite(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        if ppv.is_null() {
            return Err(E_POINTER.into());
        }

        if riid.is_null() {
            unsafe {
                ppv.write(std::ptr::null_mut());
            }

            return Err(E_POINTER.into());
        }

        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        match inner.site {
            Some(ref site) => unsafe { site.query(riid, ppv).ok() },
            None => {
                unsafe {
                    ppv.write(std::ptr::null_mut());
                }
                Err(E_FAIL.into())
            }
        }
    }
}

struct TranscodeEnumSubcommandsData {
    imaging_factory: IWICImagingFactory,
    enumerator: IEnumUnknown,
}

#[implement(IEnumExplorerCommand)]
struct TranscodeEnumSubcommands {
    inner: Mutex<TranscodeEnumSubcommandsData>,
}

impl TranscodeEnumSubcommands {
    pub fn new(imaging_factory: &IWICImagingFactory) -> windows::core::Result<Self> {
        let enumerator = unsafe {
            imaging_factory
                .CreateComponentEnumerator(WICEncoder.0 as _, WICComponentEnumerateDefault.0 as _)?
        };

        Ok(Self {
            inner: Mutex::new(TranscodeEnumSubcommandsData {
                imaging_factory: imaging_factory.clone(),
                enumerator,
            }),
        })
    }
}

impl IEnumExplorerCommand_Impl for TranscodeEnumSubcommands_Impl {
    fn Clone(&self) -> windows::core::Result<IEnumExplorerCommand> {
        let inner = self.inner.lock().unwrap();
        Ok(ComObject::new(TranscodeEnumSubcommands {
            inner: Mutex::new(TranscodeEnumSubcommandsData {
                imaging_factory: inner.imaging_factory.clone(),
                enumerator: unsafe { inner.enumerator.Clone()? },
            }),
        })
        .to_interface())
    }

    fn Next(
        &self,
        mut count: u32,
        mut commands: *mut Option<IExplorerCommand>,
        fetched: *mut u32,
    ) -> windows::core::HRESULT {
        if count == 0 {
            if !fetched.is_null() {
                unsafe {
                    fetched.write(0);
                }
            }

            return S_OK;
        }

        if commands.is_null() {
            return E_POINTER;
        }

        let inner = self.inner.lock().unwrap();

        let mut total_count = 0;

        let result = loop {
            let mut buffer = [const { None }; 20];
            let mut inner_fetched = 0;

            let result = unsafe {
                inner.enumerator.Next(
                    &mut buffer[..count.min(20) as _],
                    Some(&raw mut inner_fetched),
                )
            };

            if result.is_err() {
                break result;
            }

            let element_count = inner_fetched as usize;

            for command in &buffer[..element_count as _] {
                let Some(command) = command else {
                    inner_fetched -= 1;
                    continue;
                };

                let Ok(codec_info) = command.cast::<IWICBitmapCodecInfo>() else {
                    inner_fetched -= 1;
                    continue;
                };

                let command = ComObject::new(TranscodeSubcommand::new(
                    &inner.imaging_factory,
                    &codec_info,
                ))
                .to_interface();

                unsafe {
                    commands.write(Some(command));
                    commands = commands.add(1);
                }
            }

            total_count += inner_fetched;
            count -= inner_fetched;

            if result == S_FALSE || count == 0 {
                break result;
            }
        };

        if !fetched.is_null() {
            unsafe {
                fetched.write(total_count);
            }
        }

        result
    }

    fn Reset(&self) -> windows::core::Result<()> {
        let inner = self.inner.lock().unwrap();
        unsafe { inner.enumerator.Reset() }
    }

    fn Skip(&self, count: u32) -> windows::core::Result<()> {
        let inner = self.inner.lock().unwrap();
        unsafe { inner.enumerator.Skip(count) }
    }
}

#[allow(dead_code)]
struct TranscodeSubcommandData {
    properties: Option<IPropertyBag>,
    imaging_factory: IWICImagingFactory,
    codec_info: IWICBitmapCodecInfo,
    site: Option<IUnknown>,
}

#[derive(Default)]
#[implement(IExplorerCommand, IInitializeCommand, IObjectWithSite)]
struct TranscodeSubcommand {
    inner: RwLock<Option<TranscodeSubcommandData>>,
}

impl TranscodeSubcommand {
    pub fn new(imaging_factory: &IWICImagingFactory, codec_info: &IWICBitmapCodecInfo) -> Self {
        Self {
            inner: RwLock::new(Some(TranscodeSubcommandData {
                properties: None,
                imaging_factory: imaging_factory.clone(),
                codec_info: codec_info.clone(),
                site: None,
            })),
        }
    }

    fn item_name_without_extension(item: &IShellItem) -> windows::core::Result<CoTaskMemPWSTR> {
        unsafe {
            let file_name = CoTaskMemPWSTR::new(item.GetDisplayName(SIGDN_PARENTRELATIVEPARSING)?);

            for i in (0..file_name.len()).rev() {
                if *file_name.as_ptr().add(i) == b'.' as u16 {
                    *file_name.as_ptr().add(i) = 0;
                    break;
                }
            }
            Ok(file_name)
        }
    }

    fn transcode_items(
        imaging_factory: &IWICImagingFactory,
        items: &IShellItemArray,
        result: SaveDialogResult,
        container_format: &GUID,
        codec_info: &IWICBitmapCodecInfo,
        owner_window: HWND,
    ) -> windows::core::Result<()> {
        let operation: IFileOperation =
            unsafe { CoCreateInstance(&FileOperation, None, CLSCTX_INPROC_SERVER)? };

        unsafe {
            operation.SetOwnerWindow(owner_window)?;
        }

        for i in 0..unsafe { items.GetCount()? } {
            let item = unsafe { items.GetItemAt(i)? };

            let operation_sink = ComObject::new(TranscodeOperation::new(
                imaging_factory,
                &item,
                container_format,
                &result.pixel_format,
            ));

            let extensions = get_with_buffer!(codec_info, GetFileExtensions)?;

            let extension = extensions
                .split(|c| *c == b',' as u16)
                .next()
                .ok_or(E_UNEXPECTED)?;

            let new_filename = [
                unsafe { TranscodeSubcommand::item_name_without_extension(&item)?.as_wide() },
                extension,
                std::slice::from_ref(&0u16),
            ]
            .concat();

            unsafe {
                operation.NewItem(
                    &result.item,
                    FILE_ATTRIBUTE_NORMAL.0,
                    PCWSTR::from_raw(new_filename.as_ptr()),
                    None,
                    Some(&operation_sink.to_interface()),
                )?;
            }
        }
        unsafe { operation.PerformOperations()? };
        Ok(())
    }

    fn transcode_item(
        imaging_factory: &IWICImagingFactory,
        item: &IShellItem,
        result: SaveDialogResult,
        container_format: &GUID,
        owner_window: HWND,
    ) -> windows::core::Result<()> {
        let operation: IFileOperation =
            unsafe { CoCreateInstance(&FileOperation, None, CLSCTX_INPROC_SERVER)? };

        unsafe {
            operation.SetOwnerWindow(owner_window)?;
        }

        let operation_sink = ComObject::new(TranscodeOperation::new(
            imaging_factory,
            item,
            container_format,
            &result.pixel_format,
        ));

        enum Filename {
            WithoutExtension(CoTaskMemPWSTR),
            WithExtension(Vec<u16>),
        }

        let filename = CoTaskMemPWSTR::new(unsafe {
            result.item.GetDisplayName(SIGDN_PARENTRELATIVEPARSING)?
        });

        let filename = match result.extension {
            Some(extension) => Filename::WithExtension(
                [
                    unsafe { filename.as_wide() },
                    &extension,
                    std::slice::from_ref(&0u16),
                ]
                .concat(),
            ),
            None => Filename::WithoutExtension(filename),
        };

        unsafe {
            operation.NewItem(
                &result.item.GetParent()?,
                FILE_ATTRIBUTE_NORMAL.0,
                PCWSTR::from_raw(match filename {
                    Filename::WithoutExtension(filename) => filename.as_ptr(),
                    Filename::WithExtension(filename) => filename.as_ptr(),
                }),
                None,
                Some(&operation_sink.to_interface()),
            )?;
        }

        unsafe { operation.PerformOperations() }.inspect_err(|err| unsafe {
            let message = operation_sink
                .error_message()
                .unwrap_or_else(|| err.message());

            MessageBoxW(
                owner_window,
                PCWSTR::from_raw(HSTRING::from(message).as_ptr()),
                w!("Transcoding Error"),
                MB_ICONERROR,
            );
        })?;

        Ok(())
    }
}

impl CoClass for TranscodeSubcommand {
    const CLSID: GUID = GUID::from_u128(0xa30460cf_027e_4157_ba2e_e49840b5e851u128);
    const PROG_ID: PCWSTR = w!("X16BMX.TranscodeSubcommand.1");
    const VERSION_INDEPENDENT_PROG_ID: PCWSTR = w!("X16BMX.TranscodeSubcommand");
}

impl IExplorerCommand_Impl for TranscodeSubcommand_Impl {
    fn GetTitle(&self, _items: Option<&IShellItemArray>) -> windows::core::Result<PWSTR> {
        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        unsafe {
            let mut actual = 0;
            _ = inner.codec_info.GetFriendlyName(&mut [], &raw mut actual);

            let mut buffer: Array<u16> = Array::with_len(actual as _);

            inner
                .codec_info
                .GetFriendlyName(&mut buffer, &raw mut actual)?;

            Ok(PWSTR::from_raw(buffer.into_abi().0))
        }
    }

    fn GetIcon(&self, _items: Option<&IShellItemArray>) -> windows::core::Result<PWSTR> {
        Err(E_NOTIMPL.into())
    }

    fn GetToolTip(&self, items: Option<&IShellItemArray>) -> windows::core::Result<PWSTR> {
        self.GetTitle(items)
    }

    fn GetCanonicalName(&self) -> windows::core::Result<GUID> {
        Ok(TranscodeSubcommand::CLSID)
    }

    fn GetState(
        &self,
        items: Option<&IShellItemArray>,
        ok_to_be_slow: BOOL,
    ) -> windows::core::Result<u32> {
        let items = items.ok_or(E_POINTER)?;

        if !ok_to_be_slow.as_bool() {
            return Err(E_PENDING.into());
        }

        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        if item_array_has_matching_decoders(items, &inner.imaging_factory)? {
            Ok(ECS_ENABLED.0 as _)
        } else {
            Ok(ECS_HIDDEN.0 as _)
        }
    }

    fn Invoke(
        &self,
        items: Option<&IShellItemArray>,
        _pbc: Option<&IBindCtx>,
    ) -> windows::core::Result<()> {
        let items = items.ok_or(E_POINTER)?;

        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        let one_item = unsafe { items.GetCount()? } == 1;

        let mode = if one_item {
            SaveDialogMode::File
        } else {
            SaveDialogMode::Folder
        };

        let file_name = if one_item {
            TranscodeSubcommand::item_name_without_extension(&unsafe { items.GetItemAt(0)? })?
        } else {
            CoTaskMemPWSTR::null()
        };

        let default_folder = unsafe { items.GetItemAt(0)?.GetParent()? };

        let file_extensions = get_with_buffer!(&inner.codec_info, GetFileExtensions)?;

        let known_pixel_formats = get_with_buffer!(&inner.codec_info, GetPixelFormats)?
            .into_iter()
            .filter(pixel_format_is_known)
            .collect::<Vec<_>>();

        let dialog = ComObject::new(SaveDialog::new());

        let result = dialog.show(
            PCWSTR::from_raw(file_name.as_ptr()),
            mode,
            Some(default_folder),
            file_extensions,
            known_pixel_formats,
        )?;

        let container_format = unsafe { inner.codec_info.GetContainerFormat()? };

        let owner_window = match inner.site {
            Some(ref site) => unsafe { IUnknown_GetWindow(site).unwrap_or(HWND::default()) },
            None => HWND::default(),
        };

        match mode {
            SaveDialogMode::Folder => TranscodeSubcommand::transcode_items(
                &inner.imaging_factory,
                items,
                result,
                &container_format,
                &inner.codec_info,
                owner_window,
            )?,
            SaveDialogMode::File => TranscodeSubcommand::transcode_item(
                &inner.imaging_factory,
                unsafe { &items.GetItemAt(0)? },
                result,
                &container_format,
                owner_window,
            )?,
        }

        Ok(())
    }

    fn GetFlags(&self) -> windows::core::Result<u32> {
        Ok((ECF_DEFAULT.0) as _)
    }

    fn EnumSubCommands(&self) -> windows::core::Result<IEnumExplorerCommand> {
        Err(E_NOTIMPL.into())
    }
}

impl IInitializeCommand_Impl for TranscodeSubcommand_Impl {
    fn Initialize(
        &self,
        _command_name: &windows::core::PCWSTR,
        _property_bag: Option<&IPropertyBag>,
    ) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }
}

impl IObjectWithSite_Impl for TranscodeSubcommand_Impl {
    fn SetSite(&self, site: Option<&IUnknown>) -> windows::core::Result<()> {
        let mut inner = self.inner.write().unwrap();
        let inner = inner.as_mut().ok_or(E_UNEXPECTED)?;
        inner.site = site.cloned();
        Ok(())
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn GetSite(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        if ppv.is_null() {
            return Err(E_POINTER.into());
        }

        if riid.is_null() {
            unsafe {
                ppv.write(std::ptr::null_mut());
            }

            return Err(E_POINTER.into());
        }

        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        match inner.site {
            Some(ref site) => unsafe { site.query(riid, ppv).ok() },
            None => {
                unsafe {
                    ppv.write(std::ptr::null_mut());
                }
                Err(E_FAIL.into())
            }
        }
    }
}

#[derive(Clone, Copy)]
enum SaveDialogMode {
    Folder,
    File,
}

#[derive(Clone)]
struct SaveDialogResult {
    pub pixel_format: GUID,
    pub item: IShellItem,
    pub extension: Option<Vec<u16>>,
}

#[expect(unused)]
struct SaveDialogData {
    mode: SaveDialogMode,
    extensions: Option<Vec<Vec<u16>>>,
    pixel_formats: Vec<GUID>,
    selected_item: u32,
}

#[implement(IFileDialogEvents, IFileDialogControlEvents)]
struct SaveDialog {
    inner: Mutex<Option<SaveDialogData>>,
}

impl SaveDialog {
    const COMBO_BOX_GROUP_CONTROL_ID: u32 = u32::from_le_bytes(*b"BMX\0");
    const COMBO_BOX_CONTROL_ID: u32 = SaveDialog::COMBO_BOX_GROUP_CONTROL_ID + 1;

    pub fn new() -> Self {
        Self {
            inner: Mutex::new(None),
        }
    }

    fn do_show(&self, dialog: &IFileDialog) -> windows::core::Result<SaveDialogResult> {
        unsafe { dialog.Show(None)? };

        let inner = self.inner.lock().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        /*let pixel_format = if inner.selected_item == 0 {
            GUID::zeroed()
        } else {
            inner.pixel_formats[inner.selected_item as usize - 1]
        };*/

        let pixel_format = inner.pixel_formats[inner.selected_item as usize];

        let extension = match inner.extensions {
            Some(ref extensions) => extensions
                .get(unsafe { dialog.GetFileTypeIndex()? } as usize)
                .cloned(),
            None => None,
        };

        Ok(SaveDialogResult {
            pixel_format,
            item: unsafe { dialog.GetResult()? },
            extension,
        })
    }
}

impl SaveDialog_Impl {
    pub fn show(
        &self,
        filename: PCWSTR,
        mode: SaveDialogMode,
        default_folder: Option<IShellItem>,
        file_extensions: Vec<u16>,
        pixel_formats: Vec<GUID>,
    ) -> windows::core::Result<SaveDialogResult> {
        let mut inner = self.inner.lock().unwrap();
        if inner.is_some() {
            return Err(HRESULT::from_win32(ERROR_ALREADY_INITIALIZED.0).into());
        }

        let clsid = match mode {
            SaveDialogMode::Folder => &FileOpenDialog,
            SaveDialogMode::File => &FileSaveDialog,
        };

        let dialog: IFileDialog = unsafe { CoCreateInstance(clsid, None, CLSCTX_INPROC_SERVER)? };

        let extensions = match mode {
            SaveDialogMode::File => {
                unsafe {
                    dialog.SetFileName(filename)?;
                    dialog.SetOptions(dialog.GetOptions()? | FOS_STRICTFILETYPES)?;
                    dialog.SetTitle(w!("Select Output File"))?;
                }

                let extensions = file_extensions
                    .split(|c| *c == b',' as u16)
                    .map(|ext| ext.to_vec())
                    .collect::<Vec<_>>();

                let extension_type_names = extensions
                    .iter()
                    .filter_map(|ext| unsafe {
                        let mut ext_buffer = vec![0u16; ext.len() + 2];
                        ext_buffer[0] = b'*' as u16;
                        ext_buffer[1..ext.len() + 1].copy_from_slice(ext);

                        let mut file_info = MaybeUninit::uninit();
                        if SHGetFileInfoW(
                            PCWSTR::from_raw(ext_buffer[1..].as_ptr()),
                            FILE_ATTRIBUTE_NORMAL,
                            Some(file_info.as_mut_ptr()),
                            std::mem::size_of::<SHFILEINFOW>() as _,
                            SHGFI_TYPENAME | SHGFI_USEFILEATTRIBUTES,
                        ) != 0
                        {
                            Some((ext_buffer, file_info.assume_init().szTypeName))
                        } else {
                            None
                        }
                    })
                    .collect::<Vec<_>>();

                let mut filter_spec = extension_type_names
                    .iter()
                    .map(|(ext, type_name)| COMDLG_FILTERSPEC {
                        pszName: PCWSTR::from_raw(type_name.as_ptr()),
                        pszSpec: PCWSTR::from_raw(ext.as_ptr()),
                    })
                    .collect::<Vec<_>>();

                let mut all_formats_buf = vec![];

                for extension in extensions.iter() {
                    all_formats_buf.push(b'*' as _);
                    all_formats_buf.extend_from_slice(extension);
                    all_formats_buf.push(b',' as _);
                }

                {
                    let len = all_formats_buf.len();
                    all_formats_buf[len - 1] = 0;
                }

                filter_spec.extend_from_slice(&[
                    COMDLG_FILTERSPEC {
                        pszName: w!("All Image Files"),
                        pszSpec: PCWSTR::from_raw(all_formats_buf.as_ptr()),
                    },
                    COMDLG_FILTERSPEC {
                        pszName: w!("All Files"),
                        pszSpec: w!("*.*"),
                    },
                ]);

                unsafe {
                    dialog.SetFileTypes(&filter_spec)?;
                    dialog.SetDefaultExtension(PCWSTR::from_raw(extensions[0].as_ptr().add(1)))?;
                }

                Some(extensions)
            }
            SaveDialogMode::Folder => unsafe {
                dialog.SetOptions(dialog.GetOptions()? | FOS_PICKFOLDERS)?;
                dialog.SetTitle(w!("Select Output Folder"))?;
                None
            },
        };

        if let Some(default_folder) = default_folder {
            unsafe { dialog.SetDefaultFolder(&default_folder)? };
        }

        let customize: IFileDialogCustomize = dialog.cast()?;

        unsafe {
            customize.StartVisualGroup(SaveDialog::COMBO_BOX_GROUP_CONTROL_ID, w!("Format:"))?;
            customize.AddComboBox(SaveDialog::COMBO_BOX_CONTROL_ID)?;
            customize.EndVisualGroup()?;
            customize.MakeProminent(SaveDialog::COMBO_BOX_GROUP_CONTROL_ID)?;

            //customize.AddControlItem(SaveDialog::COMBO_BOX_CONTROL_ID, 0, w!("From Source"))?;
        }

        for (i, pixel_format) in pixel_formats.iter().enumerate() {
            let name = pixel_format_friendly_name(pixel_format);
            if name.is_null() {
                continue;
            }

            unsafe {
                customize.AddControlItem(
                    SaveDialog::COMBO_BOX_CONTROL_ID,
                    /*(i + 1)*/ i as _,
                    name,
                )?;
            }
        }

        unsafe { customize.SetSelectedControlItem(SaveDialog::COMBO_BOX_CONTROL_ID, 0)? };

        let cookie = unsafe { dialog.Advise(&self.to_interface::<IFileDialogEvents>())? };

        inner.replace(SaveDialogData {
            mode,
            extensions,
            pixel_formats,
            selected_item: 0,
        });

        std::mem::drop(inner);

        let result = self.do_show(&dialog);

        unsafe {
            dialog.Unadvise(cookie)?;
        }

        result
    }
}

#[expect(unused_variables)]
impl IFileDialogEvents_Impl for SaveDialog_Impl {
    fn OnFileOk(&self, dialog: Option<&IFileDialog>) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }

    fn OnFolderChange(&self, pfd: Option<&IFileDialog>) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }

    fn OnFolderChanging(
        &self,
        pfd: Option<&IFileDialog>,
        psifolder: Option<&IShellItem>,
    ) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }

    fn OnOverwrite(
        &self,
        pfd: Option<&IFileDialog>,
        psi: Option<&IShellItem>,
    ) -> windows::core::Result<FDE_OVERWRITE_RESPONSE> {
        Err(E_NOTIMPL.into())
    }

    fn OnSelectionChange(&self, pfd: Option<&IFileDialog>) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }

    fn OnShareViolation(
        &self,
        pfd: Option<&IFileDialog>,
        psi: Option<&IShellItem>,
    ) -> windows::core::Result<FDE_SHAREVIOLATION_RESPONSE> {
        Err(E_NOTIMPL.into())
    }

    fn OnTypeChange(&self, pfd: Option<&IFileDialog>) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }
}

#[expect(unused_variables)]
impl IFileDialogControlEvents_Impl for SaveDialog_Impl {
    fn OnButtonClicked(
        &self,
        _pfdc: Option<&IFileDialogCustomize>,
        _dwidctl: u32,
    ) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }

    fn OnCheckButtonToggled(
        &self,
        _pfdc: Option<&IFileDialogCustomize>,
        _dwidctl: u32,
        _bchecked: BOOL,
    ) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }

    fn OnControlActivating(
        &self,
        _pfdc: Option<&IFileDialogCustomize>,
        _dwidctl: u32,
    ) -> windows::core::Result<()> {
        Err(E_NOTIMPL.into())
    }

    fn OnItemSelected(
        &self,
        pfdc: Option<&IFileDialogCustomize>,
        control_id: u32,
        item_id: u32,
    ) -> windows::core::Result<()> {
        if control_id == SaveDialog::COMBO_BOX_CONTROL_ID {
            let mut inner = self.inner.lock().unwrap();
            let inner = inner.as_mut().ok_or(E_UNEXPECTED)?;

            if
            /*item_id == 0 || (item_id - 1)*/
            item_id < inner.pixel_formats.len() as u32 {
                inner.selected_item = item_id;
                Ok(())
            } else {
                Err(E_INVALIDARG.into())
            }
        } else {
            Err(E_NOTIMPL.into())
        }
    }
}

struct TranscodeOperationData {
    imaging_factory: IWICImagingFactory,
    source: IShellItem,
    container_format: GUID,
    pixel_format: GUID,
    error_message: Option<String>,
}

#[implement(IFileOperationProgressSink)]
struct TranscodeOperation {
    inner: Mutex<TranscodeOperationData>,
}

impl TranscodeOperation {
    pub fn new(
        imaging_factory: &IWICImagingFactory,
        source: &IShellItem,
        container_format: &GUID,
        pixel_format: &GUID,
    ) -> Self {
        Self {
            inner: Mutex::new(TranscodeOperationData {
                imaging_factory: imaging_factory.clone(),
                source: source.clone(),
                container_format: *container_format,
                pixel_format: *pixel_format,
                error_message: None,
            }),
        }
    }

    pub fn error_message(&self) -> Option<String> {
        self.inner.lock().unwrap().error_message.clone()
    }
}

impl IFileOperationProgressSink_Impl for TranscodeOperation_Impl {
    fn FinishOperations(&self, _hrresult: windows::core::HRESULT) -> windows::core::Result<()> {
        Ok(())
    }

    fn PauseTimer(&self) -> windows::core::Result<()> {
        Ok(())
    }

    fn PreCopyItem(
        &self,
        _dwflags: u32,
        _psiitem: Option<&IShellItem>,
        _psidestinationfolder: Option<&IShellItem>,
        _psznewname: &windows::core::PCWSTR,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn PostCopyItem(
        &self,
        _dwflags: u32,
        _psiitem: Option<&IShellItem>,
        _psidestinationfolder: Option<&IShellItem>,
        _psznewname: &windows::core::PCWSTR,
        _hrcopy: windows::core::HRESULT,
        _psinewlycreated: Option<&IShellItem>,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn PreDeleteItem(
        &self,
        _dwflags: u32,
        _psiitem: Option<&IShellItem>,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn PostDeleteItem(
        &self,
        _dwflags: u32,
        _psiitem: Option<&IShellItem>,
        _hrdelete: windows::core::HRESULT,
        _psinewlycreated: Option<&IShellItem>,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn PreMoveItem(
        &self,
        _dwflags: u32,
        _psiitem: Option<&IShellItem>,
        _psidestinationfolder: Option<&IShellItem>,
        _psznewname: &windows::core::PCWSTR,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn PostMoveItem(
        &self,
        _dwflags: u32,
        _psiitem: Option<&IShellItem>,
        _psidestinationfolder: Option<&IShellItem>,
        _psznewname: &windows::core::PCWSTR,
        _hrmove: windows::core::HRESULT,
        _psinewlycreated: Option<&IShellItem>,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn PreNewItem(
        &self,
        _dwflags: u32,
        _psidestinationfolder: Option<&IShellItem>,
        _psznewname: &windows::core::PCWSTR,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn PostNewItem(
        &self,
        _dwflags: u32,
        _psidestinationfolder: Option<&IShellItem>,
        _psznewname: &windows::core::PCWSTR,
        _psztemplatename: &windows::core::PCWSTR,
        _dwfileattributes: u32,
        hrnew: windows::core::HRESULT,
        new_item: Option<&IShellItem>,
    ) -> windows::core::Result<()> {
        hrnew.ok()?;
        let new_item = new_item.ok_or(E_POINTER)?;

        let mut inner = self.inner.lock().unwrap();

        transcode(
            &inner.imaging_factory,
            &inner.source,
            new_item,
            &inner.container_format,
            &inner.pixel_format,
        )
        .inspect_err(|err| match err {
            TranscodeError::Win(_) => {}
            err => {
                inner.error_message = Some(err.to_string());
            }
        })
        .map_err(|err| err.into())
    }

    fn PreRenameItem(
        &self,
        _dwflags: u32,
        _psiitem: Option<&IShellItem>,
        _psznewname: &windows::core::PCWSTR,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn PostRenameItem(
        &self,
        _dwflags: u32,
        _psiitem: Option<&IShellItem>,
        _psznewname: &windows::core::PCWSTR,
        _hrrename: windows::core::HRESULT,
        _psinewlycreated: Option<&IShellItem>,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    fn ResetTimer(&self) -> windows::core::Result<()> {
        Ok(())
    }

    fn ResumeTimer(&self) -> windows::core::Result<()> {
        Ok(())
    }

    fn StartOperations(&self) -> windows::core::Result<()> {
        Ok(())
    }

    fn UpdateProgress(&self, _iworktotal: u32, _iworksofar: u32) -> windows::core::Result<()> {
        Ok(())
    }
}

enum TranscodeError {
    Win(windows::core::Error),
    NoFrames,
    DoesNotSupportMultiframe,
}

impl Display for TranscodeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Win(err) => write!(f, "{}", err),
            Self::NoFrames => write!(f, "No frames in source"),
            Self::DoesNotSupportMultiframe => {
                write!(
                    f,
                    "Source has multiple frames, which the encoder does not support."
                )
            }
        }
    }
}

impl From<windows::core::Error> for TranscodeError {
    fn from(err: windows::core::Error) -> Self {
        Self::Win(err)
    }
}

impl From<HRESULT> for TranscodeError {
    fn from(hr: HRESULT) -> Self {
        Self::Win(hr.into())
    }
}

impl From<TranscodeError> for windows::core::Error {
    fn from(err: TranscodeError) -> Self {
        match err {
            TranscodeError::Win(err) => err,
            err => windows::core::Error::new(
                match err {
                    TranscodeError::NoFrames => HRESULT::from_win32(ERROR_NO_MORE_ITEMS.0),
                    TranscodeError::DoesNotSupportMultiframe => WINCODEC_ERR_UNSUPPORTEDOPERATION,
                    _ => unreachable!(),
                },
                err.to_string(),
            ),
        }
    }
}

fn transcode(
    imaging_factory: &IWICImagingFactory,
    source: &IShellItem,
    target: &IShellItem,
    container_format: &GUID,
    pixel_format: &GUID,
) -> Result<(), TranscodeError> {
    let source_stream: IStream = unsafe { source.BindToHandler(None, &BHID_Stream)? };
    let bind_ctx = unsafe { CreateBindCtx(0)? };

    let mut bind_options = BIND_OPTS {
        cbStruct: std::mem::size_of::<BIND_OPTS>() as _,
        ..Default::default()
    };

    unsafe { bind_ctx.GetBindOptions(&raw mut bind_options)? };

    bind_options.grfMode = STGM_WRITE.0;
    unsafe { bind_ctx.SetBindOptions(&raw const bind_options)? };

    let target_stream: IStream = unsafe { target.BindToHandler(Some(&bind_ctx), &BHID_Stream)? };

    let decoder = unsafe {
        imaging_factory.CreateDecoderFromStream(
            &source_stream,
            std::ptr::null(),
            WICDecodeMetadataCacheOnDemand,
        )?
    };

    let frame_count = unsafe { decoder.GetFrameCount()? };
    if frame_count < 1 {
        return Err(TranscodeError::NoFrames);
    }

    let encoder = unsafe { imaging_factory.CreateEncoder(container_format, std::ptr::null())? };

    if frame_count > 1 {
        let encoder_info = unsafe { encoder.GetEncoderInfo()? };

        if unsafe { !encoder_info.DoesSupportMultiframe()?.as_bool() } {
            return Err(TranscodeError::DoesNotSupportMultiframe);
        }
    }

    unsafe {
        encoder.Initialize(&target_stream, WICBitmapEncoderNoCache)?;
    }

    for i in 0..frame_count {
        let frame = {
            let frame = unsafe { decoder.GetFrame(i)? }.cast()?;
            if *pixel_format != GUID::zeroed() {
                unsafe { WICConvertBitmapSource(pixel_format, &frame)? }
            } else {
                frame
            }
        };

        let mut property_bag = None;

        let frame_encode = unsafe {
            let mut frame_encode = None;
            encoder.CreateNewFrame(&raw mut frame_encode, &raw mut property_bag)?;
            frame_encode.ok_or(E_FAIL)?
        };

        unsafe {
            (Interface::vtable(&frame_encode).Initialize)(
                Interface::as_raw(&frame_encode),
                property_bag
                    .as_ref()
                    .map_or(std::ptr::null_mut(), Interface::as_raw),
            )
            .ok()?;
            frame_encode.WriteSource(&frame, std::ptr::null())?;
            frame_encode.Commit()?;
        }
    }

    unsafe {
        encoder.Commit()?;
    }

    Ok(())
}
