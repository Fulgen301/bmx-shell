use std::sync::RwLock;

use windows::Win32::Foundation::{
    E_ILLEGAL_STATE_CHANGE, E_NOTIMPL, E_POINTER, E_UNEXPECTED, WINCODEC_ERR_CODECTOOMANYSCANLINES,
    WINCODEC_ERR_SOURCERECTDOESNOTMATCHDIMENSIONS, WINCODEC_ERR_UNEXPECTEDSIZE,
    WINCODEC_ERR_UNSUPPORTEDOPERATION,
};
use windows::Win32::Graphics::Imaging::{
    GUID_WICPixelFormat1bppIndexed, GUID_WICPixelFormat2bppIndexed, GUID_WICPixelFormat4bppIndexed,
    GUID_WICPixelFormat8bppIndexed, IWICBitmapEncoderInfo, IWICBitmapFrameEncode,
    IWICBitmapFrameEncode_Impl, IWICMetadataQueryWriter, WICBitmapEncoderCacheOption,
    WICBitmapPaletteTypeFixedHalftone256, WICRect,
};
use windows::Win32::System::Com::StructuredStorage::IPropertyBag2;
use windows::{
    core::{implement, ComObject, IUnknownImpl, Interface, GUID, HRESULT},
    Win32::{
        Foundation::{ERROR_ALREADY_INITIALIZED, E_INVALIDARG},
        Graphics::Imaging::{
            CLSID_WICImagingFactory, IWICBitmapEncoder, IWICBitmapEncoder_Impl, IWICBitmapSource,
            IWICColorContext, IWICImagingFactory, IWICPalette,
        },
        System::Com::{CoCreateInstance, IStream, CLSCTX_INPROC_SERVER},
    },
};
use windows_core::{w, PCWSTR};

use super::util::{bytes_per_line, pixel_format_to_bit_depth};
use crate::bmx::{FileHeader, PaletteEntry};
use crate::com::stream_write_exact_items;
use crate::util::guid;

use super::super::CoClass;
use super::com::CONTAINER_FORMAT;

enum PaletteToUse {
    Frame(IWICPalette),
    BitmapSource(IWICPalette),
}

struct Chunk {
    data: Vec<u8>,
    stride: u16,
    lines: u16,
}

struct BitmapEncoderData {
    imaging_factory: IWICImagingFactory,
    stream: IStream,
    palette: Option<IWICPalette>,
    has_frame: bool,
}

#[derive(Default)]
#[implement(IWICBitmapEncoder)]
pub struct BitmapEncoder {
    inner: RwLock<Option<BitmapEncoderData>>,
}

impl BitmapEncoder {
    pub fn new() -> Self {
        Self::default()
    }
}

impl CoClass for BitmapEncoder {
    const CLSID: GUID = guid::from_str("9d718e6d-4c95-4dc9-abd1-156a50488ebd");
    const PROG_ID: PCWSTR = w!("X16BMX.BMXEncoder.1");
    const VERSION_INDEPENDENT_PROG_ID: PCWSTR = w!("X16BMX.BMXEncoder");
}

impl IWICBitmapEncoder_Impl for BitmapEncoder_Impl {
    fn Initialize(
        &self,
        stream: Option<&IStream>,
        _cache_option: WICBitmapEncoderCacheOption,
    ) -> windows::core::Result<()> {
        let stream = stream.ok_or(E_INVALIDARG)?;

        let mut inner = self.inner.write().unwrap();
        if inner.is_some() {
            return Err(HRESULT::from_win32(ERROR_ALREADY_INITIALIZED.0).into());
        }

        let imaging_factory: IWICImagingFactory =
            unsafe { CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER)? };

        inner.replace(BitmapEncoderData {
            imaging_factory,
            stream: stream.clone(),
            palette: None,
            has_frame: false,
        });

        Ok(())
    }

    fn GetContainerFormat(&self) -> windows::core::Result<GUID> {
        Ok(CONTAINER_FORMAT)
    }

    fn GetEncoderInfo(&self) -> windows::core::Result<IWICBitmapEncoderInfo> {
        let inner = self.inner.read().unwrap();
        let inner = inner.as_ref().ok_or(E_UNEXPECTED)?;
        let component_info = unsafe {
            inner
                .imaging_factory
                .CreateComponentInfo(&BitmapEncoder::CLSID)?
        };

        component_info.cast()
    }

    fn SetColorContexts(
        &self,
        _count: u32,
        _colorcontext: *const Option<IWICColorContext>,
    ) -> windows::core::Result<()> {
        Err(WINCODEC_ERR_UNSUPPORTEDOPERATION.into())
    }

    fn SetPalette(&self, palette: Option<&IWICPalette>) -> windows::core::Result<()> {
        let palette = palette.ok_or(E_POINTER)?;

        let mut inner = self.inner.write().unwrap();
        let inner = inner.as_mut().ok_or(E_UNEXPECTED)?;

        inner.palette = Some(palette.clone());

        Ok(())
    }

    fn SetThumbnail(&self, _thumbnail: Option<&IWICBitmapSource>) -> windows::core::Result<()> {
        Err(WINCODEC_ERR_UNSUPPORTEDOPERATION.into())
    }

    fn SetPreview(&self, _preview: Option<&IWICBitmapSource>) -> windows::core::Result<()> {
        Err(WINCODEC_ERR_UNSUPPORTEDOPERATION.into())
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn CreateNewFrame(
        &self,
        frame_encode: *mut Option<IWICBitmapFrameEncode>,
        encoder_options: *mut Option<IPropertyBag2>,
    ) -> windows::core::Result<()> {
        let mut inner = self.inner.write().unwrap();
        let inner = inner.as_mut().ok_or(E_UNEXPECTED)?;

        if inner.has_frame {
            if !frame_encode.is_null() {
                unsafe { frame_encode.write(None) };
            }

            if !encoder_options.is_null() {
                unsafe { encoder_options.write(None) };
            }

            Err(WINCODEC_ERR_UNSUPPORTEDOPERATION.into())
        } else {
            if !encoder_options.is_null() {
                unsafe { encoder_options.write(None) };
            }

            let frame_encoder: IWICBitmapFrameEncode =
                ComObject::new(FrameEncoder::new(self.to_object())).to_interface();

            unsafe { frame_encode.write(Some(frame_encoder)) };

            inner.has_frame = true;

            Ok(())
        }
    }

    fn Commit(&self) -> windows::core::Result<()> {
        Ok(())
    }

    fn GetMetadataQueryWriter(&self) -> windows::core::Result<IWICMetadataQueryWriter> {
        Err(E_NOTIMPL.into())
    }
}

struct FrameEncoderData {
    parent: ComObject<BitmapEncoder>,
    header: Option<FileHeader>,
    palette: Option<PaletteToUse>,
    image_data: Vec<Chunk>,
    accumulated_height: u16,
}

#[implement(IWICBitmapFrameEncode)]
struct FrameEncoder {
    inner: RwLock<FrameEncoderData>,
}

impl FrameEncoder {
    pub fn new(parent: ComObject<BitmapEncoder>) -> Self {
        Self {
            inner: RwLock::new(FrameEncoderData {
                parent,
                header: None,
                palette: None,
                image_data: Vec::new(),
                accumulated_height: 0,
            }),
        }
    }
}

impl IWICBitmapFrameEncode_Impl for FrameEncoder_Impl {
    fn Initialize(&self, _encoder_options: Option<&IPropertyBag2>) -> windows::core::Result<()> {
        let mut inner = self.inner.write().unwrap();
        if inner.header.is_some() {
            return Err(HRESULT::from_win32(ERROR_ALREADY_INITIALIZED.0).into());
        }

        inner.header.replace(FileHeader::default());
        Ok(())
    }

    fn SetSize(&self, width: u32, height: u32) -> windows::core::Result<()> {
        let width: u16 = width
            .try_into()
            .map_err(|e| windows::core::Error::new(E_INVALIDARG, format!("{}", e)))?;
        let height: u16 = height
            .try_into()
            .map_err(|e| windows::core::Error::new(E_INVALIDARG, format!("{}", e)))?;

        if width == 0 {
            return Err(windows::core::Error::new(
                E_INVALIDARG,
                "width must not be 0",
            ));
        }

        if height == 0 {
            return Err(windows::core::Error::new(
                E_INVALIDARG,
                "height must not be 0",
            ));
        }

        let mut inner = self.inner.write().unwrap();
        let header = inner.header.as_mut().ok_or(E_UNEXPECTED)?;

        if (header.width != 0 && header.width != width)
            || (header.height != 0 && header.height != height)
        {
            return Err(windows::core::Error::new(
                E_INVALIDARG,
                "Size has already been set",
            ));
        }

        header.width = width;
        header.height = height;

        Ok(())
    }

    fn SetResolution(&self, _x: f64, _y: f64) -> windows::core::Result<()> {
        Ok(())
    }

    fn SetPixelFormat(&self, pixelformat: *mut GUID) -> windows::core::Result<()> {
        if pixelformat.is_null() {
            return Err(E_POINTER.into());
        }

        let pixelformat = unsafe { &mut *pixelformat };

        let mut inner = self.inner.write().unwrap();
        let header = inner.header.as_mut().ok_or(E_UNEXPECTED)?;

        #[allow(non_upper_case_globals)]
        let bit_depth = match *pixelformat {
            GUID_WICPixelFormat1bppIndexed => 1,
            GUID_WICPixelFormat2bppIndexed => 2,
            GUID_WICPixelFormat4bppIndexed => 4,
            GUID_WICPixelFormat8bppIndexed => 8,
            _ => {
                *pixelformat = GUID_WICPixelFormat8bppIndexed;
                8
            }
        };

        header.bit_depth = bit_depth;

        Ok(())
    }

    fn SetColorContexts(
        &self,
        _count: u32,
        _color_contexts: *const Option<IWICColorContext>,
    ) -> windows::core::Result<()> {
        Err(WINCODEC_ERR_UNSUPPORTEDOPERATION.into())
    }

    fn SetPalette(&self, palette: Option<&IWICPalette>) -> windows::core::Result<()> {
        let palette = palette.ok_or(E_POINTER)?;

        let mut inner = self.inner.write().unwrap();
        inner.palette = Some(PaletteToUse::Frame(palette.clone()));

        Ok(())
    }

    fn SetThumbnail(&self, _thumbnail: Option<&IWICBitmapSource>) -> windows::core::Result<()> {
        Err(WINCODEC_ERR_UNSUPPORTEDOPERATION.into())
    }

    fn WritePixels(
        &self,
        line_count: u32,
        stride: u32,
        buffer_size: u32,
        pixels: *const u8,
    ) -> windows::core::Result<()> {
        if pixels.is_null() {
            return Err(E_POINTER.into());
        }

        if buffer_size < stride {
            return Err(windows::core::Error::new(
                E_INVALIDARG,
                "Buffer size must not be smaller than stride",
            ));
        }

        let line_count: u16 = line_count
            .try_into()
            .map_err(|_| windows::core::Error::new(E_INVALIDARG, "line count out of range"))?;

        let mut inner = self.inner.write().unwrap();
        let header = inner.header.as_ref().ok_or(E_UNEXPECTED)?;

        if header.bit_depth == 0 {
            return Err(windows::core::Error::new(
                E_ILLEGAL_STATE_CHANGE,
                "Pixel format must be set before writing pixels",
            ));
        }

        if header.width == 0 {
            return Err(windows::core::Error::new(
                E_ILLEGAL_STATE_CHANGE,
                "Size must be set before writing pixels",
            ));
        }

        if inner.accumulated_height + line_count > header.height {
            return Err(windows::core::Error::new(
                WINCODEC_ERR_CODECTOOMANYSCANLINES,
                "Too many scanlines",
            ));
        }

        let data = unsafe { std::slice::from_raw_parts(pixels, buffer_size as _) }.to_vec();
        inner.image_data.push(Chunk {
            data,
            stride: stride as _,
            lines: line_count as _,
        });

        inner.accumulated_height += line_count;

        Ok(())
    }

    fn WriteSource(
        &self,
        bitmap_source: Option<&IWICBitmapSource>,
        rect: *const WICRect,
    ) -> windows::core::Result<()> {
        trait WICRectExt {
            fn intersect(&self, other: &Self) -> Self;
        }

        impl WICRectExt for WICRect {
            fn intersect(&self, other: &Self) -> Self {
                let x = self.X.max(other.X);
                let y = self.Y.max(other.Y);
                let width = (self.X + self.Width).min(other.X + other.Width) - x;
                let height = (self.Y + self.Height).min(other.Y + other.Height) - y;

                if width < 0 || height < 0 {
                    Default::default()
                } else {
                    WICRect {
                        X: x,
                        Y: y,
                        Width: width,
                        Height: height,
                    }
                }
            }
        }

        let bitmap_source = bitmap_source.ok_or(E_POINTER)?;

        let rect = if rect.is_null() {
            None
        } else {
            Some(unsafe { &*rect })
        };

        if let Some(rect) = rect {
            if rect.X < 0 || rect.Y < 0 || rect.Width <= 0 || rect.Height <= 0 {
                return Err(windows::core::Error::new(E_INVALIDARG, "Invalid rect"));
            }

            if rect.Width > u16::MAX as _ || rect.Height > u16::MAX as _ {
                return Err(windows::core::Error::new(E_INVALIDARG, "Rect too large"));
            }
        }

        let (source_width, source_height) = unsafe {
            let mut source_width = 0;
            let mut source_height = 0;
            bitmap_source.GetSize(&raw mut source_width, &raw mut source_height)?;
            (source_width, source_height)
        };

        let (dpi_x, dpi_y) = unsafe {
            let mut dpi_x = 0.0;
            let mut dpi_y = 0.0;
            bitmap_source.GetResolution(&raw mut dpi_x, &raw mut dpi_y)?;
            (dpi_x, dpi_y)
        };

        if (96.0 - dpi_x).abs() > 0.5 || (96.0 - dpi_y).abs() > 0.5 {
            return Err(windows::core::Error::new(
                WINCODEC_ERR_UNSUPPORTEDOPERATION,
                format!("DPI must be 96, got {dpi_x}x{dpi_y}"),
            ));
        }

        let pixel_format = unsafe { bitmap_source.GetPixelFormat()? };
        let pixel_format_bit_depth = pixel_format_to_bit_depth(&pixel_format)
            .ok_or(windows::core::Error::new(
                WINCODEC_ERR_UNSUPPORTEDOPERATION,
                "Invalid pixel format",
            ))?
            .get();

        let mut inner = self.inner.write().unwrap();

        let inner_accumulated_height = inner.accumulated_height;

        let (effective_source_rect, header_width_zero) = {
            let header = inner.header.as_mut().ok_or(E_UNEXPECTED)?;
            let header_width_zero = header.width == 0;

            if header.bit_depth != 0 && header.bit_depth != pixel_format_bit_depth {
                return Err(windows::core::Error::new(
                    E_INVALIDARG,
                    format!(
                        "Mismatch between pixel format and bit depth (header: {}, pixel format: {}",
                        header.bit_depth, pixel_format_bit_depth
                    ),
                ));
            }

            let effective_source_rect = WICRect {
                X: 0,
                Y: 0,
                Width: source_width as _,
                Height: source_height as _,
            };

            let effective_source_rect = if let Some(rect) = rect {
                effective_source_rect.intersect(rect)
            } else {
                effective_source_rect
            };

            if !header_width_zero {
                if header.width != effective_source_rect.Width as _ {
                    return Err(windows::core::Error::new(
                        WINCODEC_ERR_SOURCERECTDOESNOTMATCHDIMENSIONS,
                        "Width mismatch between source and frame",
                    ));
                }
                if inner_accumulated_height + effective_source_rect.Height as u16 > header.height {
                    return Err(windows::core::Error::new(
                        WINCODEC_ERR_CODECTOOMANYSCANLINES,
                        "Too many scanlines",
                    ));
                }
            }

            (effective_source_rect, header_width_zero)
        };

        let source_palette = if inner.palette.is_none() {
            let parent = inner.parent.inner.read().unwrap();
            let parent = parent.as_ref().ok_or(E_UNEXPECTED)?;
            let palette = unsafe { parent.imaging_factory.CreatePalette()? };
            unsafe {
                bitmap_source.CopyPalette(&palette)?;
            }

            Some(palette)
        } else {
            None
        };

        let bytes_per_line = bytes_per_line(
            effective_source_rect.Width as _,
            pixel_format_bit_depth as _,
        );

        let stride = (bytes_per_line + 3) & !3;

        let mut data = vec![0; stride as usize * effective_source_rect.Height as usize];
        unsafe {
            bitmap_source.CopyPixels(
                rect.map_or(std::ptr::null(), |f| f),
                stride as _,
                &mut data,
            )?;
        }

        if header_width_zero {
            inner.image_data.clear();
            inner.accumulated_height = 0;
        }

        inner.image_data.push(Chunk {
            data,
            stride: stride as _,
            lines: effective_source_rect.Height as _,
        });

        if header_width_zero {
            let header = inner.header.as_mut().unwrap();
            header.width = effective_source_rect.Width as _;
            header.height = effective_source_rect.Height as _;
            header.bit_depth = pixel_format_bit_depth;
        }

        if inner.palette.is_none() {
            inner.palette = Some(PaletteToUse::BitmapSource(source_palette.unwrap()));
        }

        inner.accumulated_height += effective_source_rect.Height as u16;

        Ok(())
    }

    fn Commit(&self) -> windows::core::Result<()> {
        let mut inner = self.inner.write().unwrap();
        let (width, height, bit_depth) = {
            let header = inner.header.as_ref().ok_or(E_UNEXPECTED)?;
            (header.width, header.height, header.bit_depth)
        };

        if bit_depth == 0 {
            return Err(windows::core::Error::new(
                E_ILLEGAL_STATE_CHANGE,
                "Pixel format must be set before committing",
            ));
        }

        if width == 0 {
            return Err(windows::core::Error::new(
                WINCODEC_ERR_UNEXPECTEDSIZE,
                "Size must be set before committing",
            ));
        }

        if inner
            .image_data
            .iter()
            .map(|chunk| chunk.lines)
            .sum::<u16>()
            != height
        {
            return Err(windows::core::Error::new(
                WINCODEC_ERR_UNEXPECTEDSIZE,
                "Not enough scanlines written",
            ));
        }

        {
            let header = inner.header.as_mut().unwrap();
            header.vera_color_depth_register = match header.bit_depth {
                1 => 0,
                2 => 1,
                4 => 2,
                8 => 3,
                _ => unreachable!(),
            };
        }

        let (palette_to_use, stream) = {
            let parent = inner.parent.inner.read().unwrap();
            let parent = parent.as_ref().ok_or(E_UNEXPECTED)?;

            let stream = parent.stream.clone();

            let palette_to_use = match inner.palette {
                Some(PaletteToUse::Frame(ref palette)) => palette.clone(),
                Some(PaletteToUse::BitmapSource(ref palette)) => match parent.palette {
                    Some(ref parent_palette) => parent_palette.clone(),
                    None => palette.clone(),
                },
                None => match parent.palette {
                    Some(ref palette) => palette.clone(),
                    None => {
                        let palette = unsafe { parent.imaging_factory.CreatePalette()? };

                        unsafe {
                            palette.InitializePredefined(
                                WICBitmapPaletteTypeFixedHalftone256,
                                false,
                            )?;
                        }

                        palette
                    }
                },
            };

            (palette_to_use, stream)
        };

        let mut colors = [0u32; 256];
        let mut actual_colors = 0;
        unsafe {
            palette_to_use.GetColors(&mut colors, &raw mut actual_colors)?;
        }

        let actual_colors = actual_colors as usize;

        let mut bmx_palette = [PaletteEntry::default(); 256];
        for i in 0..actual_colors {
            bmx_palette[i] = PaletteEntry::from_wic(colors[i]);
        }

        let header = inner.header.as_mut().unwrap();

        header.pal_used = if actual_colors == 256 {
            0
        } else {
            actual_colors as u8
        };

        header.data_start = (std::mem::size_of_val(header)
            + std::mem::size_of_val(&bmx_palette[..actual_colors]))
            as _;

        assert!(header.validate().is_ok());

        stream_write_exact_items(&stream, std::slice::from_ref(header))?;
        stream_write_exact_items(&stream, &bmx_palette[..actual_colors])?;

        let bytes_per_line = bytes_per_line(header.width, header.bit_depth);

        for chunk in &inner.image_data {
            if chunk.stride == bytes_per_line {
                stream_write_exact_items(&stream, &chunk.data)?;
            } else {
                for line in chunk.data.chunks_exact(chunk.stride as _) {
                    stream_write_exact_items(&stream, &line[..bytes_per_line as _])?;
                }
            }
        }

        Ok(())
    }

    fn GetMetadataQueryWriter(&self) -> windows::core::Result<IWICMetadataQueryWriter> {
        Err(E_NOTIMPL.into())
    }
}
