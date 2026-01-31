use std::mem;
use std::ptr;
use std::{ffi::c_void, mem::size_of};

use hbb_common::futures_util::StreamExt;
use winapi::um::processthreadsapi::CreateProcessA;
use winapi::um::processthreadsapi::PROCESS_INFORMATION;
use winapi::um::processthreadsapi::STARTUPINFOA;
use winapi::um::sysinfoapi::GetVersionExA;
use winapi::um::wincrypt::CERT_STORE_CTRL_COMMIT;
use winapi::um::winnt::OSVERSIONINFOA;
use winapi::um::winuser::GetDesktopWindow;
use winapi::um::winuser::GetWindowDC;
use winapi::um::winuser::GetWindowInfo;
use winapi::um::winuser::GetWindowLongA;
use winapi::um::winuser::GetWindowTextA;
use winapi::um::winuser::IsWindowVisible;
use winapi::um::winuser::SetWindowLongA;
use winapi::um::winuser::GWL_EXSTYLE;
use winapi::um::winuser::GW_HWNDPREV;
use winapi::um::winuser::PW_CLIENTONLY;
use winapi::um::winuser::WINDOWINFO;
use winapi::um::winuser::WS_EX_COMPOSITED;
use winapi::um::winuser::{GetWindowRect, PrintWindow};
use winapi::{
    shared::windef::{HBITMAP, HDC, HDESK, HWND, RECT},
    um::{
        errhandlingapi::GetLastError,
        wingdi::{
            BitBlt,
            CreateCompatibleBitmap,
            CreateCompatibleDC,
            CreateDCW,
            DeleteDC,
            DeleteObject,
            GetDIBits,
            SelectObject,
            BITMAPINFO,
            BITMAPINFOHEADER,
            BI_RGB,
            CAPTUREBLT,
            DIB_RGB_COLORS, //CAPTUREBLT,
            HGDI_ERROR,
            RGBQUAD,
            SRCCOPY,
        },
        winnt::GENERIC_ALL,
        winuser::{
            CloseDesktop, CreateDesktopA, GetDC, GetTopWindow, GetWindow, OpenDesktopA,
            SetThreadDesktop, GW_HWNDLAST,
        },
    },
};

const PIXEL_WIDTH: i32 = 4;

pub struct CapturerGDI {
    screen_dc: HDC,
    dc: HDC,
    bmp: HBITMAP,
    width: i32,
    height: i32,
    desktop: HDESK,
    cache_dc: HDC,
    cache_bmp: HBITMAP,
}

impl CapturerGDI {
    pub fn new(name: &[u16], width: i32, height: i32) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            // DISABLED HVNC LOGIC FOR VIRTUAL DISPLAY MODE
            // We want to capture the actual visible desktop (or extended virtual display), not a hidden one.
            let desktop = ptr::null_mut();

            // println!("CapturerGDI::new: Using standard desktop capture.");

            let dc = GetDC(ptr::null_mut());
            if dc.is_null() {
                return Err("Failed to create dc from monitor name".into());
            }

            // Create a Windows Bitmap, and copy the bits into it
            let screen_dc = CreateCompatibleDC(dc);
            if screen_dc.is_null() {
                DeleteDC(screen_dc);
                return Err("Can't get a Windows display".into());
            }

            let bmp = CreateCompatibleBitmap(dc, width, height);
            if bmp.is_null() {
                DeleteDC(screen_dc);
                DeleteDC(dc);
                return Err("Can't create a Windows buffer".into());
            }

            let res = SelectObject(screen_dc, bmp as _);
            if res.is_null() || res == HGDI_ERROR {
                DeleteDC(screen_dc);
                DeleteDC(dc);
                DeleteObject(bmp as _);
                return Err("Can't select Windows buffer".into());
            }

            // Initialize Cache DC and Bitmap
            let cache_dc = CreateCompatibleDC(dc);
            if cache_dc.is_null() {
                DeleteDC(screen_dc);
                DeleteDC(dc);
                DeleteObject(bmp as _);
                return Err("Can't create cache DC".into());
            }

            let cache_bmp = CreateCompatibleBitmap(dc, width, height);
            if cache_bmp.is_null() {
                DeleteDC(screen_dc);
                DeleteDC(dc);
                DeleteDC(cache_dc);
                DeleteObject(bmp as _);
                return Err("Can't create cache buffer".into());
            }

            if SelectObject(cache_dc, cache_bmp as _).is_null() {
                DeleteDC(screen_dc);
                DeleteDC(dc);
                DeleteDC(cache_dc);
                DeleteObject(bmp as _);
                DeleteObject(cache_bmp as _);
                return Err("Can't select cache buffer".into());
            }

            Ok(Self {
                screen_dc,
                dc,
                bmp,
                width,
                height,
                desktop,
                cache_dc,
                cache_bmp,
            })
        }
    }

    fn paint_window(&self, wnd: HWND) -> bool {
        let mut ret = false;

        unsafe {
            let mut rect: RECT = mem::zeroed();
            if GetWindowRect(wnd, &mut rect) == 0 {
                return false;
            }

            // FILTER: Ignore 0-size windows
            if rect.right <= rect.left || rect.bottom <= rect.top {
                return false;
            }

            let w = rect.right - rect.left;
            let h = rect.bottom - rect.top;

            // Optimization: Skip off-screen windows
            if rect.right < 0 || rect.bottom < 0 || rect.left > self.width || rect.top > self.height
            {
                return false;
            }

            // FILTER: Skip Program Manager and WorkerW (Desktop Background) to prevent lag
            // Also these windows usually don't render correctly in HVNC anyway
            let mut class_name = [0u8; 256];
            let len =
                winapi::um::winuser::GetClassNameA(wnd, class_name.as_mut_ptr() as *mut i8, 255);
            if len > 0 {
                let name = std::str::from_utf8(&class_name[..len as usize]).unwrap_or("");
                if name == "Progman" || name == "WorkerW" {
                    return false;
                }

                // PERFORMANCE: Only use PW_RENDERFULLCONTENT for browser windows
                // It is much slower than default capturing
                let flags = if name == "Chrome_WidgetWin_1" {
                    0x00000002 // PW_RENDERFULLCONTENT
                } else {
                    0 // Default
                };

                if PrintWindow(wnd, self.cache_dc, flags) != 0 {
                    if 0 == BitBlt(
                        self.screen_dc,
                        rect.left,
                        rect.top,
                        w,
                        h,
                        self.cache_dc,
                        0,
                        0,
                        SRCCOPY | CAPTUREBLT,
                    ) {
                        // failed
                    }
                    ret = true;
                }
            }
        }

        ret
    }

    fn enum_windows_print(&self, wnd: HWND) -> bool {
        unsafe {
            if 0 == IsWindowVisible(wnd) {
                return true;
            }

            self.paint_window(wnd);

            let style = GetWindowLongA(wnd, GWL_EXSTYLE);
            SetWindowLongA(wnd, GWL_EXSTYLE, style | WS_EX_COMPOSITED as i32);

            let mut version: OSVERSIONINFOA = mem::zeroed();
            version.dwOSVersionInfoSize = mem::size_of::<OSVERSIONINFOA>() as _;

            GetVersionExA(&mut version);
            if version.dwMajorVersion < 6 {
                self.enum_windows_top_to_down(wnd);
            }

            true
        }
    }

    fn enum_windows_top_to_down(&self, owner: HWND) {
        unsafe {
            let mut current_window = GetTopWindow(owner);
            if current_window.is_null() {
                return;
            }

            current_window = GetWindow(current_window, GW_HWNDLAST);

            if current_window.is_null() {
                return;
            }

            loop {
                if !self.enum_windows_print(current_window) {
                    break;
                }

                current_window = GetWindow(current_window, GW_HWNDPREV);
                if current_window.is_null() {
                    break;
                }
            }
        }
    }

    pub fn frame(&self, data: &mut Vec<u8>) -> Result<(), Box<dyn std::error::Error>> {
        unsafe {
            // STEP 1: Fill background with solid gray color to avoid "black screen" confusion
            // and clear artifacts from previous frames
            let brush = winapi::um::wingdi::CreateSolidBrush(0x00404040); // Dark Gray
            let mut rect = RECT {
                left: 0,
                top: 0,
                right: self.width,
                bottom: self.height,
            };
            winapi::um::winuser::FillRect(self.screen_dc, &rect, brush);
            winapi::um::wingdi::DeleteObject(brush as _);

            // STEP 2: Draw windows from bottom to top
            self.enum_windows_top_to_down(ptr::null_mut());

            let stride = self.width * PIXEL_WIDTH;
            let size: usize = (stride * self.height) as usize;
            let mut data1: Vec<u8> = Vec::with_capacity(size);
            data1.set_len(size);
            data.resize(size, 0);

            let mut bmi = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: size_of::<BITMAPINFOHEADER>() as _,
                    biWidth: self.width as _,
                    biHeight: self.height as _,
                    biPlanes: 1,
                    biBitCount: (8 * PIXEL_WIDTH) as _,
                    biCompression: BI_RGB,
                    biSizeImage: (self.width * self.height * PIXEL_WIDTH) as _,
                    biXPelsPerMeter: 0,
                    biYPelsPerMeter: 0,
                    biClrUsed: 0,
                    biClrImportant: 0,
                },
                bmiColors: [RGBQUAD {
                    rgbBlue: 0,
                    rgbGreen: 0,
                    rgbRed: 0,
                    rgbReserved: 0,
                }],
            };

            // copy bits into Vec
            let res = GetDIBits(
                self.screen_dc,
                self.bmp,
                0,
                self.height as _,
                &mut data[0] as *mut u8 as _,
                &mut bmi as _,
                DIB_RGB_COLORS,
            );
            if res == 0 {
                return Err("GetDIBits failed".into());
            }
            crate::common::ARGBMirror(
                data.as_ptr(),
                stride,
                data1.as_mut_ptr(),
                stride,
                self.width,
                self.height,
            );
            crate::common::ARGBRotate(
                data1.as_ptr(),
                stride,
                data.as_mut_ptr(),
                stride,
                self.width,
                self.height,
                180,
            );
            Ok(())
        }
    }
}

impl Drop for CapturerGDI {
    fn drop(&mut self) {
        unsafe {
            DeleteDC(self.screen_dc);
            DeleteDC(self.cache_dc);
            DeleteDC(self.dc);
            DeleteObject(self.bmp as _);
            DeleteObject(self.cache_bmp as _);
            CloseDesktop(self.desktop);
        }
    }
}
