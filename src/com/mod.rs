use windows::Win32::{
    Foundation::{E_UNEXPECTED, S_FALSE, S_OK, WINCODEC_ERR_BADHEADER},
    System::Com::{IStream, STREAM_SEEK_CUR},
};
use windows_core::{GUID, PCWSTR};

use crate::bmx::{FileHeader, FileHeaderError};

pub mod shell;
mod util;
pub mod wic;

pub trait CoClass {
    const CLSID: GUID;
    const PROG_ID: PCWSTR;
    const VERSION_INDEPENDENT_PROG_ID: PCWSTR;
}

pub fn stream_read_exact(stream: &IStream, buf: &mut [u8]) -> windows::core::Result<usize> {
    let mut read = 0;
    unsafe {
        match stream.Read(
            buf.as_mut_ptr().cast(),
            buf.len().try_into().unwrap(),
            Some(&raw mut read),
        ) {
            S_OK => Ok(read as _),
            S_FALSE => Err(E_UNEXPECTED.into()),
            err => Err(err.into()),
        }
    }
}

pub fn stream_read_exact_items<T>(stream: &IStream, buf: &mut [T]) -> windows::core::Result<usize> {
    let mut read = 0;
    unsafe {
        match stream.Read(
            buf.as_mut_ptr().cast(),
            std::mem::size_of_val(buf).try_into().unwrap(),
            Some(&raw mut read),
        ) {
            S_OK => Ok(read as _),
            S_FALSE => Err(E_UNEXPECTED.into()),
            err => Err(err.into()),
        }
    }
}

pub fn stream_write_exact_items<T>(stream: &IStream, buf: &[T]) -> windows::core::Result<usize> {
    let mut written = 0;
    unsafe {
        match stream.Write(
            buf.as_ptr().cast(),
            std::mem::size_of_val(buf).try_into().unwrap(),
            Some(&raw mut written),
        ) {
            S_OK => Ok(written as _),
            S_FALSE => Err(E_UNEXPECTED.into()),
            err => Err(err.into()),
        }
    }
}

pub fn stream_tell(stream: &IStream) -> windows::core::Result<u64> {
    let mut position = 0;
    unsafe {
        stream.Seek(0, STREAM_SEEK_CUR, Some(&raw mut position))?;
    }

    Ok(position)
}

pub trait FileHeaderExt: Sized {
    fn from_stream(stream: &IStream) -> windows::core::Result<Self>;
}

impl FileHeaderExt for FileHeader {
    fn from_stream(stream: &IStream) -> windows::core::Result<Self> {
        let mut header = [0u8; std::mem::size_of::<FileHeader>()];
        stream_read_exact(stream, &mut header)?;
        FileHeader::from_bytes(&header).map_err(FileHeaderErrorExt::to_win_error)
    }
}

pub trait FileHeaderErrorExt: Sized {
    fn to_win_error(self) -> windows::core::Error;
}

impl FileHeaderErrorExt for FileHeaderError {
    fn to_win_error(self) -> windows::core::Error {
        windows::core::Error::new(WINCODEC_ERR_BADHEADER, self.to_string())
    }
}
