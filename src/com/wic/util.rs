use std::num::NonZeroU8;

use windows::Win32::{
    Graphics::Imaging::{
        GUID_WICPixelFormat1bppIndexed, GUID_WICPixelFormat2bppIndexed,
        GUID_WICPixelFormat4bppIndexed, GUID_WICPixelFormat8bppIndexed,
    },
    System::Com::{IStream, STGC_DEFAULT, STREAM_SEEK_CUR, STREAM_SEEK_SET},
};
use windows_core::GUID;

pub struct StreamPositionPreserver {
    stream: IStream,
    pub position: u64,
}

impl StreamPositionPreserver {
    pub fn new(stream: IStream) -> windows::core::Result<Self> {
        let mut position = 0;
        unsafe {
            stream.Seek(0, STREAM_SEEK_CUR, Some(&raw mut position))?;
        }

        Ok(Self { stream, position })
    }
}

impl Drop for StreamPositionPreserver {
    fn drop(&mut self) {
        unsafe {
            let _ = self
                .stream
                .Seek(self.position as i64, STREAM_SEEK_SET, None);
        }
    }
}

pub struct StreamReadWriteWrapper<'a> {
    stream: &'a IStream,
}

impl<'a> std::io::Read for StreamReadWriteWrapper<'a> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let mut read = 0;
        unsafe {
            self.stream
                .Read(
                    buf.as_mut_ptr().cast(),
                    buf.len().try_into().unwrap(),
                    Some(&raw mut read),
                )
                .ok()
                .map_or_else(
                    |err| Err(std::io::Error::from_raw_os_error(err.code().0)),
                    |_| Ok(read as _),
                )
        }
    }
}

impl<'a> std::io::Write for StreamReadWriteWrapper<'a> {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        let mut written = 0;
        unsafe {
            self.stream.Write(
                buf.as_ptr().cast(),
                buf.len().try_into().unwrap(),
                Some(&raw mut written),
            )
        }
        .ok()
        .map_or_else(
            |err| Err(std::io::Error::from_raw_os_error(err.code().0)),
            |_| Ok(written as _),
        )
    }

    fn flush(&mut self) -> std::io::Result<()> {
        unsafe { self.stream.Commit(STGC_DEFAULT) }
            .map_err(|err| std::io::Error::from_raw_os_error(err.code().0))
    }
}

pub fn bytes_per_line(width: u16, bit_depth: u8) -> u16 {
    ((width as u32 * (bit_depth as u32) + 7) / 8) as u16
}

pub fn bit_depth_to_pixel_format(bit_depth: u8) -> Option<GUID> {
    match bit_depth {
        1 => Some(GUID_WICPixelFormat1bppIndexed),
        2 => Some(GUID_WICPixelFormat2bppIndexed),
        4 => Some(GUID_WICPixelFormat4bppIndexed),
        8 => Some(GUID_WICPixelFormat8bppIndexed),
        _ => None,
    }
}

pub fn pixel_format_to_bit_depth(pixel_format: &GUID) -> Option<NonZeroU8> {
    #[allow(non_upper_case_globals)]
    match *pixel_format {
        GUID_WICPixelFormat1bppIndexed => NonZeroU8::new(1),
        GUID_WICPixelFormat2bppIndexed => NonZeroU8::new(2),
        GUID_WICPixelFormat4bppIndexed => NonZeroU8::new(4),
        GUID_WICPixelFormat8bppIndexed => NonZeroU8::new(8),
        _ => None,
    }
}
