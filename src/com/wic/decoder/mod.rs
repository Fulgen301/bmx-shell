use std::sync::RwLock;

use windows::Win32::Foundation::{E_NOTIMPL, E_UNEXPECTED, WINCODEC_ERR_INSUFFICIENTBUFFER};
use windows::Win32::Graphics::Imaging::{
    IWICMetadataBlockReader_Impl, IWICMetadataReader, IWICStream, WICRect,
};
use windows::Win32::System::Com::IEnumUnknown;
use windows::{
    core::{implement, ComObject, IUnknownImpl, Interface, GUID, HRESULT},
    Win32::{
        Foundation::{ERROR_ALREADY_INITIALIZED, E_INVALIDARG},
        Graphics::Imaging::{
            CLSID_WICImagingFactory, IWICBitmapDecoder, IWICBitmapDecoderInfo,
            IWICBitmapDecoder_Impl, IWICBitmapFrameDecode, IWICBitmapFrameDecode_Impl,
            IWICBitmapSource, IWICBitmapSource_Impl, IWICColorContext, IWICComponentInfo,
            IWICImagingFactory, IWICMetadataBlockReader, IWICMetadataQueryReader, IWICPalette,
            WICBitmapDecoderCapabilityCanDecodeAllImages,
            WICBitmapDecoderCapabilityCanDecodeSomeImages, WICDecodeOptions,
        },
        System::Com::{CoCreateInstance, IStream, CLSCTX_INPROC_SERVER, STREAM_SEEK_SET},
    },
};
use windows_core::{w, PCWSTR};

use super::super::wic::util::bytes_per_line;
use super::super::wic::util::StreamPositionPreserver;
use crate::bmx::{FileHeader, PaletteEntry};
use crate::com::{stream_read_exact, stream_read_exact_items, stream_tell, FileHeaderExt};
use crate::util::guid;

use super::super::CoClass;
use super::com::CONTAINER_FORMAT;
use super::util::bit_depth_to_pixel_format;

struct BitmapDecoderData {
    imaging_factory: IWICImagingFactory,
    stream: IWICStream,
    header: FileHeader,
    palette: IWICPalette,
}

#[derive(Default)]
#[implement(IWICBitmapDecoder)]
pub struct BitmapDecoder {
    inner: RwLock<Option<BitmapDecoderData>>,
}

impl BitmapDecoder {
    pub fn new() -> Self {
        Default::default()
    }
}

impl CoClass for BitmapDecoder {
    const CLSID: GUID = guid::from_str("5c8a66da-1c32-4d8e-8ead-c579214a6522");
    const PROG_ID: PCWSTR = w!("X16BMX.BMXDecoder.1");
    const VERSION_INDEPENDENT_PROG_ID: PCWSTR = w!("X16BMX.BMXDecoder");
}

impl IWICBitmapDecoder_Impl for BitmapDecoder_Impl {
    fn QueryCapability(&self, stream: Option<&IStream>) -> windows::core::Result<u32> {
        let stream = stream.ok_or(E_INVALIDARG)?;

        let _position_preserver = StreamPositionPreserver::new(stream.clone())?;
        let header = FileHeader::from_stream(stream)?;

        if header.compressed == 0 {
            Ok(WICBitmapDecoderCapabilityCanDecodeAllImages.0 as u32
                | WICBitmapDecoderCapabilityCanDecodeSomeImages.0 as u32)
        } else {
            Ok(0)
        }
    }

    fn Initialize(
        &self,
        stream: Option<&IStream>,
        _cacheoptions: WICDecodeOptions,
    ) -> windows::core::Result<()> {
        let stream = stream.ok_or(E_INVALIDARG)?;

        let mut inner = self.inner.write().unwrap();
        if inner.is_some() {
            return Err(HRESULT::from_win32(ERROR_ALREADY_INITIALIZED.0).into());
        }

        let stream_position_preserver = StreamPositionPreserver::new(stream.clone())?;

        let begin_position = stream_tell(stream)?;

        let header = FileHeader::from_stream(stream)?;

        let imaging_factory: IWICImagingFactory =
            unsafe { CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER)? };

        let image_size = header.data_start as u64
            + bytes_per_line(header.width, header.bit_depth) as u64 * header.height as u64;

        let stream = {
            let wic_stream = unsafe { imaging_factory.CreateStream()? };

            unsafe {
                wic_stream.InitializeFromIStreamRegion(
                    stream,
                    stream_position_preserver.position,
                    image_size,
                )?
            };

            wic_stream
        };

        unsafe {
            stream.Seek(std::mem::size_of_val(&header) as _, STREAM_SEEK_SET, None)?;
        }

        let palette = unsafe { imaging_factory.CreatePalette()? };

        let palette_entry_count = header.palette_entry_count();
        let mut palette_entries: [PaletteEntry; 256] = [Default::default(); 256];
        let palette_entries = &mut palette_entries[..palette_entry_count];

        stream_read_exact_items(&stream, palette_entries)?;

        let mut wic_colors = [0u32; 256];

        for i in 0..palette_entry_count {
            wic_colors[i] = palette_entries[i].to_wic();
        }

        unsafe {
            palette.InitializeCustom(&wic_colors[..palette_entry_count])?;
        }

        unsafe {
            stream.Seek(
                (begin_position + (header.data_start as u64)) as i64,
                STREAM_SEEK_SET,
                None,
            )?;
        }

        inner.replace(BitmapDecoderData {
            imaging_factory,
            stream,
            header,
            palette,
        });

        Ok(())
    }

    fn GetContainerFormat(&self) -> windows::core::Result<windows::core::GUID> {
        Ok(CONTAINER_FORMAT)
    }

    fn GetDecoderInfo(&self) -> windows::core::Result<IWICBitmapDecoderInfo> {
        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        let component_info: IWICComponentInfo = unsafe {
            inner
                .imaging_factory
                .CreateComponentInfo(&BitmapDecoder::CLSID)?
        };

        component_info.cast()
    }

    fn GetFrameCount(&self) -> windows::core::Result<u32> {
        Ok(1)
    }

    fn GetFrame(&self, index: u32) -> windows::core::Result<IWICBitmapFrameDecode> {
        if index > 0 {
            Err(E_INVALIDARG.into())
        } else {
            Ok(ComObject::new(FrameDecoder::new(self.to_object())).into_interface())
        }
    }

    fn GetPreview(&self) -> windows::core::Result<IWICBitmapSource> {
        self.GetFrame(0)?.cast()
    }

    fn GetThumbnail(&self) -> windows::core::Result<IWICBitmapSource> {
        self.GetFrame(0)?.cast()
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn GetColorContexts(
        &self,
        count: u32,
        color_contexts: *mut Option<IWICColorContext>,
        actual_count: *mut u32,
    ) -> windows::core::Result<()> {
        if color_contexts.is_null() {
            if count > 0 {
                Err(E_INVALIDARG.into())
            } else {
                unsafe {
                    *actual_count = 0;
                }
                Ok(())
            }
        } else {
            unsafe {
                *color_contexts = None;
            }
            Err(E_INVALIDARG.into())
        }
    }

    fn GetMetadataQueryReader(&self) -> windows::core::Result<IWICMetadataQueryReader> {
        Err(E_NOTIMPL.into())
    }

    fn CopyPalette(&self, palette: Option<&IWICPalette>) -> windows::core::Result<()> {
        let palette = palette.ok_or(E_INVALIDARG)?;

        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;

        let mut colors = [0u32; 256];
        let mut actual_colors = 0;

        unsafe {
            inner
                .palette
                .GetColors(&mut colors, &raw mut actual_colors)?;
            palette.InitializeCustom(&colors[..actual_colors as _])
        }
    }
}

struct FrameDecoderData {
    parent: ComObject<BitmapDecoder>,
}

#[implement(IWICBitmapFrameDecode, IWICMetadataBlockReader)]
pub struct FrameDecoder {
    inner: RwLock<FrameDecoderData>,
}

impl FrameDecoder {
    pub fn new(parent: ComObject<BitmapDecoder>) -> FrameDecoder {
        FrameDecoder {
            inner: RwLock::new(FrameDecoderData { parent }),
        }
    }
}

impl IWICBitmapSource_Impl for FrameDecoder_Impl {
    fn GetPixelFormat(&self) -> windows::core::Result<windows::core::GUID> {
        let inner = self.inner.read().unwrap();
        let parent_inner = inner.parent.inner.read().unwrap();
        let parent_inner = parent_inner.as_ref().ok_or(E_UNEXPECTED)?;

        bit_depth_to_pixel_format(parent_inner.header.bit_depth).ok_or(E_UNEXPECTED.into())
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn GetResolution(&self, x: *mut f64, y: *mut f64) -> windows::core::Result<()> {
        unsafe {
            *x = 96.0f64;
            *y = 96.0f64;
        }

        Ok(())
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn GetSize(&self, width: *mut u32, height: *mut u32) -> windows::core::Result<()> {
        let inner = self.inner.read().unwrap();
        let parent_inner = inner.parent.inner.read().unwrap();
        let parent_inner = parent_inner.as_ref().ok_or(E_UNEXPECTED)?;

        unsafe {
            *width = parent_inner.header.width as _;
            *height = parent_inner.header.height as _;
        }

        Ok(())
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn CopyPixels(
        &self,
        rect: *const WICRect,
        stride: u32,
        buffer_size: u32,
        buffer: *mut u8,
    ) -> windows::core::Result<()> {
        let inner = self.inner.read().unwrap();
        let parent_inner = inner.parent.inner.read().unwrap();
        let parent_inner = parent_inner.as_ref().ok_or(E_UNEXPECTED)?;

        if (stride as u16)
            < bytes_per_line(
                parent_inner.header.width as _,
                parent_inner.header.bit_depth as _,
            )
        {
            return Err(WINCODEC_ERR_INSUFFICIENTBUFFER.into());
        }

        let rect = if rect.is_null() {
            None
        } else {
            Some(unsafe { &*rect })
        };

        let min_buffer_size = match rect {
            Some(rect) => {
                bytes_per_line(rect.Width as u16, parent_inner.header.bit_depth) as u32
                    * rect.Height as u32
            }
            None => {
                bytes_per_line(parent_inner.header.width, parent_inner.header.bit_depth) as u32
                    * parent_inner.header.height as u32
            }
        };

        if buffer_size < min_buffer_size {
            return Err(WINCODEC_ERR_INSUFFICIENTBUFFER.into());
        }

        let stream = &parent_inner.stream;

        match rect {
            Some(rect) => {
                if rect.X < 0
                    || rect.Y < 0
                    || rect.X + rect.Width > parent_inner.header.width as i32
                    || rect.Y + rect.Height > parent_inner.header.height as i32
                {
                    return Err(E_INVALIDARG.into());
                }

                let offset =
                    bytes_per_line(parent_inner.header.width, parent_inner.header.bit_depth) as u32
                        * (rect.Y as u32)
                        + (rect.X as u32);

                unsafe {
                    stream.Seek(
                        parent_inner.header.data_start as i64 + offset as i64,
                        STREAM_SEEK_SET,
                        None,
                    )?;
                }

                let bytes_per_line =
                    bytes_per_line(rect.Width as u16, parent_inner.header.bit_depth);

                let mut buffer = buffer;

                for i in 0..rect.Height {
                    if i > 0 && rect.X > 0 {
                        // skip the first bytes
                        stream_read_exact(stream, unsafe {
                            std::slice::from_raw_parts_mut(buffer, rect.X as _)
                        })?;
                    }

                    stream_read_exact(stream, unsafe {
                        std::slice::from_raw_parts_mut(buffer, bytes_per_line as _)
                    })?;

                    unsafe {
                        buffer = buffer.add(stride as _);
                    }
                }
            }
            None => {
                let bytes_per_line =
                    bytes_per_line(parent_inner.header.width, parent_inner.header.bit_depth);

                let mut buffer = buffer;

                for _ in 0..parent_inner.header.height {
                    stream_read_exact(stream, unsafe {
                        std::slice::from_raw_parts_mut(buffer, bytes_per_line as _)
                    })?;

                    unsafe {
                        buffer = buffer.add(stride as _);
                    }
                }
            }
        }

        Ok(())
    }

    fn CopyPalette(&self, palette: Option<&IWICPalette>) -> windows::core::Result<()> {
        let palette = palette.ok_or(E_INVALIDARG)?;

        let inner = self.inner.read().unwrap();
        inner.parent.CopyPalette(Some(palette))
    }
}

impl IWICBitmapFrameDecode_Impl for FrameDecoder_Impl {
    fn GetThumbnail(&self) -> windows::core::Result<IWICBitmapSource> {
        unsafe { self.cast() }
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn GetColorContexts(
        &self,
        count: u32,
        color_contexts: *mut Option<IWICColorContext>,
        actual_count: *mut u32,
    ) -> windows::core::Result<()> {
        if color_contexts.is_null() {
            if count > 0 {
                Err(E_INVALIDARG.into())
            } else {
                unsafe {
                    *actual_count = 0;
                }
                Ok(())
            }
        } else {
            unsafe {
                *color_contexts = None;
            }
            Err(E_INVALIDARG.into())
        }
    }

    fn GetMetadataQueryReader(&self) -> windows::core::Result<IWICMetadataQueryReader> {
        Err(E_NOTIMPL.into())
    }
}

impl IWICMetadataBlockReader_Impl for FrameDecoder_Impl {
    fn GetContainerFormat(&self) -> windows::core::Result<GUID> {
        Ok(CONTAINER_FORMAT)
    }

    fn GetCount(&self) -> windows::core::Result<u32> {
        Ok(0)
    }

    fn GetEnumerator(&self) -> windows::core::Result<IEnumUnknown> {
        Err(E_NOTIMPL.into())
    }

    fn GetReaderByIndex(&self, _index: u32) -> windows::core::Result<IWICMetadataReader> {
        Err(E_INVALIDARG.into())
    }
}
