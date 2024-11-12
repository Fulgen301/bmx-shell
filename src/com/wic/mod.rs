use std::iter::FusedIterator;

use imp::Sealed;
use windows::Win32::{
    Foundation::{E_POINTER, S_FALSE, S_OK},
    Graphics::Imaging::*,
    System::Com::{CoCreateInstance, IEnumUnknown, CLSCTX_INPROC_SERVER},
};
use windows_core::{w, IUnknown, Interface, GUID, PCWSTR, PWSTR};

pub mod class_factory;
pub mod com;
pub mod decoder;
pub mod encoder;
mod util;

pub fn create_imaging_factory() -> windows::core::Result<IWICImagingFactory> {
    unsafe { CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER) }
}

struct EnumUnknownIterator {
    iterator: Option<IEnumUnknown>,
}

impl EnumUnknownIterator {
    fn new(iterator: IEnumUnknown) -> Self {
        Self {
            iterator: Some(iterator),
        }
    }
}

impl Iterator for EnumUnknownIterator {
    type Item = windows::core::Result<IUnknown>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.iterator {
            Some(ref mut iterator) => {
                let mut component = None;

                match unsafe { iterator.Next(std::slice::from_mut(&mut component), None) } {
                    S_OK => Some(component.ok_or(E_POINTER.into())),
                    S_FALSE => {
                        self.iterator = None;
                        None
                    }
                    err => Some(Err(err.into())),
                }
            }
            None => None,
        }
    }
}

impl FusedIterator for EnumUnknownIterator {}

pub struct TypedComponentIterator<T: Interface> {
    iterator: EnumUnknownIterator,
    _marker: std::marker::PhantomData<T>,
}

impl<T: Interface> TypedComponentIterator<T> {
    fn new(iterator: IEnumUnknown) -> Self {
        Self {
            iterator: EnumUnknownIterator::new(iterator),
            _marker: std::marker::PhantomData,
        }
    }
}

impl<T: Interface> Iterator for TypedComponentIterator<T> {
    type Item = windows::core::Result<T>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.iterator.next() {
            Some(Ok(item)) => Some(item.cast()),
            Some(Err(err)) => Some(Err(err)),
            None => None,
        }
    }
}

pub fn get_component_iterator<T: Interface>(
    imaging_factory: &IWICImagingFactory,
    component_types: WICComponentType,
    enumerate_options: WICComponentEnumerateOptions,
) -> windows::core::Result<impl Iterator<Item = windows::core::Result<T>> + use<T>> {
    Ok(TypedComponentIterator::<T>::new(unsafe {
        imaging_factory
            .CreateComponentEnumerator(component_types.0 as _, enumerate_options.0 as _)?
    }))
}

/*pub fn get_with_buffer<T: Clone + Default>(
    op: impl Fn(&mut [T], *mut u32) -> windows::core::Result<()>,
) -> windows::core::Result<Vec<T>> {
    let mut actual = 0;

    _ = op(&mut [], &raw mut actual);

    let mut buffer = vec![Default::default(); actual as usize];
    op(&mut buffer, &raw mut actual)?;

    buffer.resize(actual as usize, Default::default());
    Ok(buffer)
}*/

mod imp {
    use windows_core::PWSTR;

    pub trait Sealed {}

    impl<T> Sealed for *const T {}
    impl<T> Sealed for *mut T {}
    impl Sealed for PWSTR {}
}

pub trait AsNullPtr: Copy + Sealed {
    fn null() -> Self;
}

impl<T> AsNullPtr for *const T {
    fn null() -> Self {
        std::ptr::null()
    }
}

impl<T> AsNullPtr for *mut T {
    fn null() -> Self {
        std::ptr::null_mut()
    }
}

impl AsNullPtr for PWSTR {
    fn null() -> Self {
        PWSTR::null()
    }
}

#[macro_export]
macro_rules! get_with_buffer {
    ($com_object:expr, $method:ident) => {{
        let vtable = Interface::vtable($com_object);

        let mut buffer = vec![];
        let mut actual = 0;

        fn __null_ptr<T: $crate::com::wic::AsNullPtr>() -> T {
            $crate::com::wic::AsNullPtr::null()
        }

        unsafe {
            _ = (vtable.$method)(
                Interface::as_raw($com_object),
                0,
                __null_ptr(),
                &raw mut actual,
            )
        };

        buffer.resize(actual as usize, Default::default());

        let result = unsafe { $com_object.$method(&mut buffer, &raw mut actual) };

        match result {
            Ok(()) => {
                buffer.resize(actual as usize, Default::default());
                Ok(buffer)
            }
            Err(err) => Err(err),
        }
    }};
}

pub fn codec_mime_types(codec: &IWICBitmapCodecInfo) -> windows::core::Result<Vec<u16>> {
    //get_with_buffer(codec, method_index!(IWICBitmapCodecInfo_Vtbl::GetMimeTypes))
    get_with_buffer!(codec, GetMimeTypes)
}

pub fn pixel_format_is_known(pixel_format: &GUID) -> bool {
    !pixel_format_friendly_name(pixel_format).is_null()
}

pub fn pixel_format_friendly_name(pixel_format: &GUID) -> PCWSTR {
    #[allow(non_upper_case_globals)]
    match *pixel_format {
        GUID_WICPixelFormatDontCare => w!("Don't Care"),
        GUID_WICPixelFormat1bppIndexed => w!("1-bit Indexed"),
        GUID_WICPixelFormat2bppIndexed => w!("2-bit Indexed"),
        GUID_WICPixelFormat4bppIndexed => w!("4-bit Indexed"),
        GUID_WICPixelFormat8bppIndexed => w!("8-bit Indexed"),
        GUID_WICPixelFormatBlackWhite => w!("Black and White"),
        GUID_WICPixelFormat2bppGray => w!("2-bit Grayscale"),
        GUID_WICPixelFormat4bppGray => w!("4-bit Grayscale"),
        GUID_WICPixelFormat8bppGray => w!("8-bit Grayscale"),
        GUID_WICPixelFormat8bppAlpha => w!("8-bit Alpha"),
        GUID_WICPixelFormat16bppBGR555 => w!("16-bit BGR555"),
        GUID_WICPixelFormat16bppBGR565 => w!("16-bit BGR565"),
        GUID_WICPixelFormat16bppBGRA5551 => w!("16-bit BGRA5551"),
        GUID_WICPixelFormat16bppGray => w!("16-bit Grayscale"),
        GUID_WICPixelFormat24bppBGR => w!("24-bit BGR"),
        GUID_WICPixelFormat24bppRGB => w!("24-bit RGB"),
        GUID_WICPixelFormat32bppBGR => w!("32-bit BGR"),
        GUID_WICPixelFormat32bppBGRA => w!("32-bit BGRA"),
        GUID_WICPixelFormat32bppPBGRA => w!("32-bit Premultiplied BGRA"),
        GUID_WICPixelFormat32bppGrayFloat => w!("32-bit Floating Point Grayscale"),
        GUID_WICPixelFormat32bppRGB => w!("32-bit RGB"),
        GUID_WICPixelFormat32bppRGBA => w!("32-bit RGBA"),
        GUID_WICPixelFormat32bppPRGBA => w!("32-bit Premultiplied RGBA"),
        GUID_WICPixelFormat48bppRGB => w!("48-bit RGB"),
        GUID_WICPixelFormat48bppBGR => w!("48-bit BGR"),
        GUID_WICPixelFormat64bppRGB => w!("64-bit RGB"),
        GUID_WICPixelFormat64bppRGBA => w!("64-bit RGBA"),
        GUID_WICPixelFormat64bppBGRA => w!("64-bit BGRA"),
        GUID_WICPixelFormat64bppPRGBA => w!("64-bit Premultiplied RGBA"),
        GUID_WICPixelFormat64bppPBGRA => w!("64-bit Premultiplied BGRA"),
        GUID_WICPixelFormat16bppGrayFixedPoint => w!("16-bit Fixed Point Grayscale"),
        GUID_WICPixelFormat32bppBGR101010 => w!("32-bit BGR101010"),
        GUID_WICPixelFormat48bppRGBFixedPoint => w!("48-bit Fixed Point RGB"),
        GUID_WICPixelFormat48bppBGRFixedPoint => w!("48-bit Fixed Point BGR"),
        GUID_WICPixelFormat96bppRGBFixedPoint => w!("96-bit Fixed Point RGB"),
        GUID_WICPixelFormat96bppRGBFloat => w!("96-bit Floating Point RGB"),
        GUID_WICPixelFormat128bppRGBAFloat => w!("128-bit Floating Point RGBA"),
        GUID_WICPixelFormat128bppPRGBAFloat => {
            w!("128-bit Premultiplied Floating Point RGBA")
        }
        GUID_WICPixelFormat128bppRGBFloat => w!("128-bit Floating Point RGB"),
        GUID_WICPixelFormat32bppCMYK => w!("32-bit CMYK"),
        GUID_WICPixelFormat64bppRGBAFixedPoint => w!("64-bit Fixed Point RGBA"),
        GUID_WICPixelFormat64bppBGRAFixedPoint => w!("64-bit Fixed Point BGRA"),
        GUID_WICPixelFormat64bppRGBFixedPoint => w!("64-bit Fixed Point RGB"),
        GUID_WICPixelFormat128bppRGBAFixedPoint => w!("128-bit Fixed Point RGBA"),
        GUID_WICPixelFormat128bppRGBFixedPoint => w!("128-bit Fixed Point RGB"),
        GUID_WICPixelFormat64bppRGBAHalf => w!("64-bit Half RGBA"),
        GUID_WICPixelFormat64bppPRGBAHalf => w!("64-bit Half Premultiplied RGBA"),
        GUID_WICPixelFormat64bppRGBHalf => w!("64-bit Half RGB"),
        GUID_WICPixelFormat48bppRGBHalf => w!("48-bit Half RGB"),
        GUID_WICPixelFormat32bppRGBE => w!("32-bit RGBE"),
        GUID_WICPixelFormat16bppGrayHalf => w!("16-bit Half Grayscale"),
        GUID_WICPixelFormat32bppGrayFixedPoint => w!("32-bit Fixed Point Grayscale"),
        GUID_WICPixelFormat32bppRGBA1010102 => w!("32-bit RGBA1010102"),
        GUID_WICPixelFormat32bppRGBA1010102XR => w!("32-bit RGBA1010102 Extended Range"),
        GUID_WICPixelFormat32bppR10G10B10A2 => w!("32-bit R10G10B10A2"),
        GUID_WICPixelFormat32bppR10G10B10A2HDR10 => w!("32-bit R10G10B10A2 HDR10"),
        GUID_WICPixelFormat64bppCMYK => w!("64-bit CMYK"),
        GUID_WICPixelFormat24bpp3Channels => w!("24-bit 3 Channels"),
        GUID_WICPixelFormat32bpp4Channels => w!("32-bit 4 Channels"),
        GUID_WICPixelFormat40bpp5Channels => w!("40-bit 5 Channels"),
        GUID_WICPixelFormat48bpp6Channels => w!("48-bit 6 Channels"),
        GUID_WICPixelFormat56bpp7Channels => w!("56-bit 7 Channels"),
        GUID_WICPixelFormat64bpp8Channels => w!("64-bit 8 Channels"),
        GUID_WICPixelFormat48bpp3Channels => w!("48-bit 3 Channels"),
        GUID_WICPixelFormat64bpp4Channels => w!("64-bit 4 Channels"),
        GUID_WICPixelFormat80bpp5Channels => w!("80-bit 5 Channels"),
        GUID_WICPixelFormat96bpp6Channels => w!("96-bit 6 Channels"),
        GUID_WICPixelFormat112bpp7Channels => w!("112-bit 7 Channels"),
        GUID_WICPixelFormat128bpp8Channels => w!("128-bit 8 Channels"),
        GUID_WICPixelFormat40bppCMYKAlpha => w!("40-bit CMYK Alpha"),
        GUID_WICPixelFormat80bppCMYKAlpha => w!("80-bit CMYK Alpha"),
        GUID_WICPixelFormat32bpp3ChannelsAlpha => w!("32-bit 3 Channels Alpha"),
        GUID_WICPixelFormat40bpp4ChannelsAlpha => w!("40-bit 4 Channels Alpha"),
        GUID_WICPixelFormat48bpp5ChannelsAlpha => w!("48-bit 5 Channels Alpha"),
        GUID_WICPixelFormat56bpp6ChannelsAlpha => w!("56-bit 6 Channels Alpha"),
        GUID_WICPixelFormat64bpp7ChannelsAlpha => w!("64-bit 7 Channels Alpha"),
        GUID_WICPixelFormat72bpp8ChannelsAlpha => w!("72-bit 8 Channels Alpha"),
        GUID_WICPixelFormat64bpp3ChannelsAlpha => w!("64-bit 3 Channels Alpha"),
        GUID_WICPixelFormat80bpp4ChannelsAlpha => w!("80-bit 4 Channels Alpha"),
        GUID_WICPixelFormat96bpp5ChannelsAlpha => w!("96-bit 5 Channels Alpha"),
        GUID_WICPixelFormat112bpp6ChannelsAlpha => w!("112-bit 6 Channels Alpha"),
        GUID_WICPixelFormat128bpp7ChannelsAlpha => w!("128-bit 7 Channels Alpha"),
        GUID_WICPixelFormat144bpp8ChannelsAlpha => w!("144-bit 8 Channels Alpha"),
        GUID_WICPixelFormat8bppY => w!("8-bit Y"),
        GUID_WICPixelFormat8bppCb => w!("8-bit Cb"),
        GUID_WICPixelFormat8bppCr => w!("8-bit Cr"),
        GUID_WICPixelFormat16bppCbCr => w!("16-bit CbCr"),
        GUID_WICPixelFormat16bppYQuantizedDctCoefficients => {
            w!("16-bit Y Quantized Dct Coefficients")
        }
        GUID_WICPixelFormat16bppCbQuantizedDctCoefficients => {
            w!("16-bit Cb Quantized Dct Coefficients")
        }
        GUID_WICPixelFormat16bppCrQuantizedDctCoefficients => {
            w!("16-bit Cr Quantized Dct Coefficients")
        }
        _ => PCWSTR::null(),
    }
}
