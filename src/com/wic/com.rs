use windows_core::{w, GUID, PCWSTR};

pub const VENDOR: GUID = GUID::from_values(
    0x9cac5e90,
    0xf9e5,
    0x4870,
    [0xac, 0x97, 0x9f, 0xef, 0x9a, 0x3c, 0xee, 0x41],
);

pub const CONTAINER_FORMAT: GUID = GUID::from_values(
    0x858f0257,
    0x3fa7,
    0x4014,
    [0xbc, 0xfa, 0x3c, 0x03, 0x3a, 0x0c, 0xab, 0x52],
);

pub const FORMAT: GUID = GUID::from_values(
    0x57ba4938,
    0x16a5,
    0x417e,
    [0xa9, 0xaf, 0x01, 0x26, 0x16, 0x2d, 0x38, 0x39],
);

pub const MIME_TYPE: PCWSTR = w!("image/vnd.X16BMX.bmx");

pub const PROG_ID: PCWSTR = w!("bmxfile");
pub const EXTENSION: PCWSTR = w!(".bmx");
pub const PREVIEW_DETAILS: PCWSTR =
    w!("prop:System.Image.Dimensions;System.Image.BitDepth;System.Image.Compression");
