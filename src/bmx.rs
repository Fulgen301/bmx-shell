use std::{fmt::Display, num::NonZeroU8};

#[repr(C)]
#[derive(Clone, Debug)]
pub struct FileHeader {
    pub file_id: [NonZeroU8; 3],
    pub version: u8,
    pub bit_depth: u8,
    pub vera_color_depth_register: u8,
    pub width: u16,
    pub height: u16,
    pub pal_used: u8,
    pub pal_start: u8,
    pub data_start: u16,
    pub compressed: i8,
    pub vera_border_color: u8,
    pub reserved: [u8; 16],
}

impl FileHeader {
    pub const fn from_bytes(bytes: &[u8]) -> Result<FileHeader, FileHeaderError> {
        if bytes.len() != 32 {
            return Err(FileHeaderError::InvalidHeaderSize);
        }

        let file_id = [
            match NonZeroU8::new(bytes[0]) {
                Some(byte) => byte,
                None => return Err(FileHeaderError::InvalidFileId),
            },
            match NonZeroU8::new(bytes[1]) {
                Some(byte) => byte,
                None => return Err(FileHeaderError::InvalidFileId),
            },
            match NonZeroU8::new(bytes[2]) {
                Some(byte) => byte,
                None => return Err(FileHeaderError::InvalidFileId),
            },
        ];

        let header = FileHeader {
            file_id,
            version: bytes[3],
            bit_depth: bytes[4],
            vera_color_depth_register: bytes[5],
            width: u16::from_le_bytes([bytes[6], bytes[7]]),
            height: u16::from_le_bytes([bytes[8], bytes[9]]),
            pal_used: bytes[10],
            pal_start: bytes[11],
            data_start: u16::from_le_bytes([bytes[12], bytes[13]]),
            compressed: bytes[14] as i8,
            vera_border_color: bytes[15],
            reserved: [
                bytes[16], bytes[17], bytes[18], bytes[19], bytes[20], bytes[21], bytes[22],
                bytes[23], bytes[24], bytes[25], bytes[26], bytes[27], bytes[28], bytes[29],
                bytes[30], bytes[31],
            ],
        };

        match header.validate() {
            Ok(()) => Ok(header),
            Err(err) => Err(err),
        }
    }

    pub const fn validate(&self) -> Result<(), FileHeaderError> {
        if self.file_id[0].get() != b'B'
            || self.file_id[1].get() != b'M'
            || self.file_id[2].get() != b'X'
        {
            return Err(FileHeaderError::InvalidFileId);
        }

        if self.version != 1 {
            return Err(FileHeaderError::InvalidVersion);
        }

        if !matches!(self.bit_depth, 1 | 2 | 4 | 8) {
            return Err(FileHeaderError::InvalidBitDepth);
        }

        if !matches!(self.vera_color_depth_register, 0..=3) {
            return Err(FileHeaderError::InvalidVeraColorDepthRegister);
        }

        if !matches!(
            (self.bit_depth, self.vera_color_depth_register),
            (1, 0) | (2, 1) | (4, 2) | (8, 3)
        ) {
            return Err(FileHeaderError::BitDepthMismatch);
        }

        if (self.data_start as usize)
            < std::mem::size_of::<FileHeader>()
                + std::mem::size_of::<PaletteEntry>() * self.palette_entry_count()
        {
            return Err(FileHeaderError::InvalidDataStart);
        }

        Ok(())
    }

    pub const fn to_bytes(&self) -> [u8; 32] {
        let width = self.width.to_le_bytes();
        let height = self.height.to_le_bytes();
        let data_start = self.data_start.to_be_bytes();

        [
            self.file_id[0].get(),
            self.file_id[1].get(),
            self.file_id[2].get(),
            self.version,
            self.bit_depth,
            self.vera_color_depth_register,
            width[0],
            width[1],
            height[0],
            height[1],
            self.pal_used,
            self.pal_start,
            data_start[0],
            data_start[1],
            self.compressed as u8,
            self.vera_border_color,
            self.reserved[0],
            self.reserved[1],
            self.reserved[2],
            self.reserved[3],
            self.reserved[4],
            self.reserved[5],
            self.reserved[6],
            self.reserved[7],
            self.reserved[8],
            self.reserved[9],
            self.reserved[10],
            self.reserved[11],
            self.reserved[12],
            self.reserved[13],
            self.reserved[14],
            self.reserved[15],
        ]
    }

    pub const fn palette_entry_count(&self) -> usize {
        if self.pal_used == 0 {
            256 as _
        } else {
            self.pal_used as _
        }
    }
}

impl Default for FileHeader {
    fn default() -> Self {
        Self {
            file_id: [
                unsafe { NonZeroU8::new_unchecked(b'B') },
                unsafe { NonZeroU8::new_unchecked(b'M') },
                unsafe { NonZeroU8::new_unchecked(b'X') },
            ],
            version: 1,
            bit_depth: 0,
            vera_color_depth_register: 0,
            width: 0,
            height: 0,
            pal_used: 0,
            pal_start: 0,
            data_start: 0,
            compressed: 0,
            vera_border_color: 0,
            reserved: [0; 16],
        }
    }
}

const _: () =
    assert!(std::mem::size_of::<FileHeader>() == std::mem::size_of::<Option<FileHeader>>());

#[derive(Clone, Copy, Debug)]
pub enum FileHeaderError {
    InvalidHeaderSize,
    InvalidFileId,
    InvalidVersion,
    InvalidBitDepth,
    InvalidVeraColorDepthRegister,
    BitDepthMismatch,
    InvalidDataStart,
    InvalidVeraBorderColor,
}

impl Display for FileHeaderError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match *self {
            FileHeaderError::InvalidHeaderSize => write!(f, "Invalid header size"),
            FileHeaderError::InvalidFileId => write!(f, "Invalid file ID"),
            FileHeaderError::InvalidVersion => write!(f, "Invalid version"),
            FileHeaderError::InvalidBitDepth => write!(f, "Invalid bit depth"),
            FileHeaderError::InvalidVeraColorDepthRegister => {
                write!(f, "Invalid VERA color depth register")
            }
            FileHeaderError::BitDepthMismatch => {
                write!(
                    f,
                    "Mismatch between bit depth and VERA color depth register"
                )
            }
            FileHeaderError::InvalidDataStart => write!(f, "Invalid data start"),
            FileHeaderError::InvalidVeraBorderColor => write!(f, "Invalid Vera border color"),
        }
    }
}

#[repr(C)]
#[derive(Clone, Copy, Debug, Default)]
pub struct PaletteEntry {
    pub gb: u8,
    pub r: u8,
}

impl PaletteEntry {
    pub const fn from_rgb(r: u8, g: u8, b: u8) -> Self {
        Self {
            gb: (g >> 4) << 4 | (b >> 4),
            r: r >> 4,
        }
    }

    pub const fn to_rgb(&self) -> (u8, u8, u8) {
        let r = self.r;
        let g = self.gb >> 4;
        let b = self.gb & 0x0F;

        (r << 4, g << 4, b << 4)
    }

    pub const fn from_wic(color: u32) -> Self {
        let r = (color >> 16) as u8;
        let g = (color >> 8) as u8;
        let b = color as u8;

        Self::from_rgb(r, g, b)
    }

    pub const fn to_wic(&self) -> u32 {
        let (r, g, b) = self.to_rgb();
        0xFF000000 | (r as u32) << 16 | (g as u32) << 8 | (b as u32)
    }
}
