use self::winapi::ctypes::c_int;
use self::winapi::shared::{basetsd::ULONG_PTR, minwindef::*, windef::*};
use self::winapi::um::winbase::*;
use self::winapi::um::winuser::*;
use hbb_common::lazy_static::lazy_static;
use hbb_common::tokio::sync::RwLock;
use winapi;

use crate::win::keycodes::*;
use crate::{Key, KeyboardControllable, MouseButton, MouseControllable};
use std::mem::*;
use std::ptr;

// 补充拖拽需要的 WinAPI 常量
const SWP_NOSIZE: UINT = 0x0001;
const SWP_NOZORDER: UINT = 0x0004;
const SWP_NOACTIVATE: UINT = 0x0010;

extern "system" {
    pub fn GetLastError() -> DWORD;
}

/// The main struct for handling the event emitting
#[derive(Default)]
pub struct Enigo;
static mut LAYOUT: HKL = std::ptr::null_mut();

/// The dwExtraInfo value in keyboard and mouse structure that used in
/// SendInput()
pub const ENIGO_INPUT_EXTRA_VALUE: ULONG_PTR = 100;

static mut MOUSE_DOWN: bool = false;
static mut X: i32 = 0;
static mut Y: i32 = 0;

static mut LAST_MOUSE_DOWN: u32 = 0;
static mut LAST_MOUSE_DOWN_X_Y: (i32, i32) = (0, 0);
static mut LAST_POINT: (i32, i32) = (0, 0);

const DBL_CLK_AREA_DELTA: i32 = 5;

static mut MOVE_WINDOW: HWND = 0 as _;
static mut MOVE_WINDOW_TYP: isize = 0;

static mut GLOBAL_WND: HWND = 0 as _;

///
pub fn get_gbl_wnd() -> HWND {
    unsafe { GLOBAL_WND }
}

fn print_wnd_name(wnd: HWND) -> String {
    unsafe {
        let mut x = vec![0_u8; 64];
        GetClassNameA(wnd, x.as_mut_ptr() as _, 64);

        return std::ffi::CStr::from_ptr(x.as_ptr() as _)
            .to_str()
            .unwrap_or("")
            .to_owned();
    }
}

fn wnd_cls_name(wnd: HWND) -> String {
    unsafe {
        let mut x = vec![0_u8; 32];
        GetClassNameA(wnd, x.as_mut_ptr() as _, 64);

        return std::ffi::CStr::from_ptr(x.as_ptr() as _)
            .to_str()
            .unwrap_or("")
            .to_owned();
    }
}

#[inline]
fn is_dbl_clk(evt: u32, x: i32, y: i32) -> bool {
    unsafe {
        evt == LAST_MOUSE_DOWN
            && (x - LAST_MOUSE_DOWN_X_Y.0).abs() < DBL_CLK_AREA_DELTA
            && (y - LAST_MOUSE_DOWN_X_Y.1).abs() < DBL_CLK_AREA_DELTA
    }
}

#[inline]
fn lparam_from_point(point: POINT) -> isize {
    point.x as isize | (point.y as isize) << 16
}

fn mouse_event(flags: u32, data: u32, dx: i32, dy: i32) -> DWORD {
    // Threshold for Virtual Screen (Primary Screen Width)
    let primary_width = unsafe { GetSystemMetrics(SM_CXSCREEN) };
    
    // RustDesk typically sends coordinates relative to the captured display.
    // We add primary_width to X to target the virtual screen.
    // This assumes RustDesk sends relative coordinates (0-1920).
    // For "Shadow Mode", we assume we are strictly operating on the 2nd screen.
    let mut virtual_x = dx + primary_width; 
    let mut virtual_y = dy;

    if (flags & MOUSEEVENTF_ABSOLUTE) == 0 {
        unsafe {
            virtual_x = X + dx;
            virtual_y = Y + dy;
        }
    }

    const MOVE: u32 = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;
    
    const HTCLIENT: isize = 1;
    const HTCAPTION: isize = 2;
    const HTLEFT: isize = 10;
    const HTRIGHT: isize = 11;
    const HTTOP: isize = 12;
    const HTTOPLEFT: isize = 13;
    const HTTOPRIGHT: isize = 14;
    const HTBOTTOM: isize = 15;
    const HTBOTTOMLEFT: isize = 16;
    const HTBOTTOMRIGHT: isize = 17;

    unsafe {
        match flags {
            MOVE => {
                let last_x = X;
                let last_y = Y;
                X = virtual_x;
                Y = virtual_y;

                if MOUSE_DOWN && MOVE_WINDOW != 0 as HWND {
                    let wnd = MOVE_WINDOW;
                    let delta_x = X - last_x;
                    let delta_y = Y - last_y;

                    let mut rect: RECT = zeroed();
                    GetWindowRect(wnd, &mut rect);
                    
                    let mut wx = rect.left;
                    let mut wy = rect.top;
                    let mut width = rect.right - rect.left;
                    let mut height = rect.bottom - rect.top;

                    match MOVE_WINDOW_TYP {
                        HTCAPTION => {
                            wx += delta_x;
                            wy += delta_y;
                        }
                        HTTOP => {
                            wy += delta_y;
                            height -= delta_y;
                        }
                        HTBOTTOM => {
                            height += delta_y;
                        }
                        HTLEFT => {
                            wx += delta_x;
                            width -= delta_x;
                        }
                        HTRIGHT => {
                            width += delta_x;
                        }
                        HTTOPLEFT => {
                            wy += delta_y;
                            height -= delta_y;
                            wx += delta_x;
                            width -= delta_x;
                        }
                        HTTOPRIGHT => {
                            wy += delta_y;
                            height -= delta_y;
                            width += delta_x;
                        }
                        HTBOTTOMLEFT => {
                            height += delta_y;
                            wx += delta_x;
                            width -= delta_x;
                        }
                        HTBOTTOMRIGHT => {
                            height += delta_y;
                            width += delta_x;
                        }
                        _ => {}
                    }

                    if MOVE_WINDOW_TYP != HTCLIENT {
                        SetWindowPos(wnd, 0 as _, wx, wy, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
                    }
                }
            }
            _ => {
                let mut point = POINT { x: X, y: Y };
                // Find the window under mouse
                let mut wnd = WindowFromPoint(point);
                let screen_lparam = lparam_from_point(point);

                // Traverse down to the deepest child window
                let mut curr_wnd;
                let mut client_wnd = wnd;
                loop {
                    curr_wnd = client_wnd;
                    ScreenToClient(curr_wnd, &mut point);
                    client_wnd = ChildWindowFromPoint(client_wnd, point);
                    if client_wnd == curr_wnd || client_wnd.is_null() {
                        break;
                    }
                }
                
                // Fallback: Use the window returned by WindowFromPoint if traversal failed weirdly
                if client_wnd.is_null() {
                    client_wnd = wnd;
                } else {
                    // Update wnd to the actual top-level or control found, 
                    // but for dragging we usually want the top-level parent of the control
                    // However, let's keep logic simple: wnd is Top Level, client_wnd is target
                }

                // If we found a child, we might want to ensure 'wnd' is the real top-level parent for dragging
                let mut root = wnd;
                loop {
                    let parent = GetParent(root);
                    if parent.is_null() { break; }
                    root = parent;
                }
                // Only update global wnd if it's different, to avoid messing up state
                // Actually, let's stick to using 'wnd' as the one we clicked
                
                GLOBAL_WND = wnd;
                let client_lparam = lparam_from_point(point);

                match flags {
                    MOUSEEVENTF_LEFTDOWN => {
                        MOUSE_DOWN = true;
                        
                        // Hit Test on the ROOT window (for dragging/resizing)
                        let hit_test = SendMessageA(root, WM_NCHITTEST, 0, screen_lparam);
                        
                        if hit_test == HTCLIENT {
                            // Normal Click on Client Area
                            MOVE_WINDOW = 0 as HWND; // Not dragging
                            
                            // 1. Focus the window (Important for CMD/Input)
                            SetForegroundWindow(root);
                            
                            // 2. Handle Start Button special case
                            let start_button = FindWindowA("Button\0".as_ptr() as _, ptr::null());
                            let mut rect: RECT = zeroed();
                            GetWindowRect(start_button, &mut rect);
                            if PtInRect(&rect, POINT { x: X, y: Y }) != 0 {
                                PostMessageA(start_button, BM_CLICK, 0, 0);
                                return 1;
                            }

                            // 3. Send Click
                            // Special handling for double clicks
                            if is_dbl_clk(MOUSEEVENTF_LEFTDOWN, X, Y) {
                                PostMessageA(client_wnd, WM_LBUTTONDBLCLK, MK_LBUTTON, client_lparam);
                                LAST_MOUSE_DOWN = 0;
                                LAST_MOUSE_DOWN_X_Y = (0, 0);
                            } else {
                                PostMessageA(client_wnd, WM_LBUTTONDOWN, MK_LBUTTON, client_lparam);
                                LAST_MOUSE_DOWN = MOUSEEVENTF_LEFTDOWN;
                                LAST_MOUSE_DOWN_X_Y = (X, Y);
                            }
                        } else {
                            // Non-Client Area (Title bar, borders) -> Start Dragging/Resizing
                            MOVE_WINDOW = root;
                            MOVE_WINDOW_TYP = hit_test;
                            // Also focus it
                            SetForegroundWindow(root);
                            // We don't send WM_LBUTTONDOWN to client_wnd here, 
                            // Windows handles NC clicks automatically usually, 
                            // but since we are blocking input, we handle drag manually in MOVE event.
                        }
                    }
                    MOUSEEVENTF_LEFTUP => {
                        MOUSE_DOWN = false;
                        MOVE_WINDOW = 0 as _; // Stop dragging

                        // Always send UP to the client window to finish any click
                        PostMessageA(client_wnd, WM_LBUTTONUP, 0, client_lparam);
                        
                        // Handle sys commands if we clicked buttons (Min/Max/Close)
                         if wnd_cls_name(wnd) != "SysTreeView32" {
                             let ret = SendMessageA(wnd, WM_NCHITTEST, 0, screen_lparam);
                             match ret {
                                 HTCLOSE => { PostMessageA(wnd, WM_CLOSE, 0, 0); }
                                 HTMINBUTTON => { PostMessageA(wnd, WM_SYSCOMMAND, SC_MINIMIZE, 0); }
                                 HTMAXBUTTON => {
                                     let mut placement: WINDOWPLACEMENT = zeroed();
                                     placement.length = size_of::<WINDOWPLACEMENT>() as _;
                                     GetWindowPlacement(wnd, &mut placement);
                                     if placement.flags as i32 & SW_SHOWMAXIMIZED != 0 {
                                         PostMessageA(wnd, WM_SYSCOMMAND, SC_RESTORE, 0);
                                     } else {
                                         PostMessageA(wnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0);
                                     }
                                 }
                                 _ => {}
                             }
                         }
                    }
                    MOUSEEVENTF_MIDDLEDOWN => {
                        PostMessageA(client_wnd, WM_MBUTTONDOWN, MK_MBUTTON, client_lparam);
                    }
                    MOUSEEVENTF_MIDDLEUP => {
                        PostMessageA(client_wnd, WM_MBUTTONUP, 0, client_lparam);
                    }
                    MOUSEEVENTF_RIGHTDOWN => {
                        SetForegroundWindow(root); // Focus on right click too
                        PostMessageA(client_wnd, WM_RBUTTONDOWN, MK_RBUTTON, client_lparam);
                    }
                    MOUSEEVENTF_RIGHTUP => {
                        PostMessageA(client_wnd, WM_RBUTTONUP, 0, client_lparam);
                    }
                    MOUSEEVENTF_XDOWN => {
                        PostMessageA(client_wnd, WM_MOUSEWHEEL, MK_RBUTTON, client_lparam);
                    }
                    MOUSEEVENTF_HWHEEL => {
                        PostMessageA(client_wnd, WM_MOUSEHWHEEL, (data << 16) as usize, screen_lparam);
                    }
                    MOUSEEVENTF_WHEEL => {
                        PostMessageA(client_wnd, WM_MOUSEWHEEL, (data << 16) as usize, screen_lparam);
                    }
                    _ => {}
                }
            }
        }
    }

    return 1;
}
                        HTTOP => {
                            wy += delta_y;
                            height -= delta_y;
                        }
                        HTBOTTOM => {
                            height += delta_y;
                        }
                        HTLEFT => {
                            wx += delta_x;
                            width -= delta_x;
                        }
                        HTRIGHT => {
                            width += delta_x;
                        }
                        HTTOPLEFT => {
                            wy += delta_y;
                            height -= delta_y;
                            wx += delta_x;
                            width -= delta_x;
                        }
                        HTTOPRIGHT => {
                            wy += delta_y;
                            height -= delta_y;
                            width += delta_x;
                        }
                        HTBOTTOMLEFT => {
                            height += delta_y;
                            wx += delta_x;
                            width -= delta_x;
                        }
                        HTBOTTOMRIGHT => {
                            height += delta_y;
                            width += delta_x;
                        }
                        _ => {}
                    }

                    if MOVE_WINDOW_TYP != HTCLIENT {
                        SetWindowPos(
                            wnd,
                            0 as _,
                            wx,
                            wy,
                            width,
                            height,
                            SWP_NOZORDER | SWP_NOACTIVATE,
                        );
                    }
                }
            }
            _ => {
                let mut point = POINT { x: X, y: Y };
                let mut wnd = WindowFromPoint(point);
                let screen_lparam = lparam_from_point(point);

                let mut curr_wnd;
                let mut client_wnd = wnd;
                loop {
                    curr_wnd = client_wnd;
                    ScreenToClient(curr_wnd, &mut point);
                    client_wnd = ChildWindowFromPoint(client_wnd, point);
                    if client_wnd == curr_wnd || client_wnd.is_null() {
                        break;
                    }
                }

                if client_wnd.is_null() {
                    client_wnd = wnd;
                }

                let mut root = wnd;
                loop {
                    let parent = GetParent(root);
                    if parent.is_null() {
                        break;
                    }
                    root = parent;
                }

                GLOBAL_WND = wnd;
                let client_lparam = lparam_from_point(point);

                match flags {
                    MOUSEEVENTF_LEFTDOWN => {
                        MOUSE_DOWN = true;

                        let hit_test = SendMessageA(root, WM_NCHITTEST, 0, screen_lparam);

                        if hit_test == HTCLIENT {
                            MOVE_WINDOW = 0 as HWND;

                            SetForegroundWindow(root);

                            let start_button = FindWindowA("Button\0".as_ptr() as _, ptr::null());
                            let mut rect: RECT = zeroed();
                            GetWindowRect(start_button, &mut rect);
                            if PtInRect(&rect, POINT { x: X, y: Y }) != 0 {
                                PostMessageA(start_button, BM_CLICK, 0, 0);
                                return 1;
                            }

                            if is_dbl_clk(MOUSEEVENTF_LEFTDOWN, X, Y) {
                                PostMessageA(
                                    client_wnd,
                                    WM_LBUTTONDBLCLK,
                                    MK_LBUTTON,
                                    client_lparam,
                                );
                                LAST_MOUSE_DOWN = 0;
                                LAST_MOUSE_DOWN_X_Y = (0, 0);
                            } else {
                                PostMessageA(client_wnd, WM_LBUTTONDOWN, MK_LBUTTON, client_lparam);
                                LAST_MOUSE_DOWN = MOUSEEVENTF_LEFTDOWN;
                                LAST_MOUSE_DOWN_X_Y = (X, Y);
                            }
                        } else {
                            MOVE_WINDOW = root;
                            MOVE_WINDOW_TYP = hit_test;
                            SetForegroundWindow(root);
                        }
                    }
                    MOUSEEVENTF_LEFTUP => {
                        MOUSE_DOWN = false;
                        MOVE_WINDOW = 0 as _;

                        PostMessageA(client_wnd, WM_LBUTTONUP, 0, client_lparam);

                        if wnd_cls_name(wnd) != "SysTreeView32" {
                            let ret = SendMessageA(wnd, WM_NCHITTEST, 0, screen_lparam);
                            match ret {
                                HTCLOSE => {
                                    PostMessageA(wnd, WM_CLOSE, 0, 0);
                                }
                                HTMINBUTTON => {
                                    PostMessageA(wnd, WM_SYSCOMMAND, SC_MINIMIZE, 0);
                                }
                                HTMAXBUTTON => {
                                    let mut placement: WINDOWPLACEMENT = zeroed();
                                    placement.length = size_of::<WINDOWPLACEMENT>() as _;
                                    GetWindowPlacement(wnd, &mut placement);
                                    if placement.flags as i32 & SW_SHOWMAXIMIZED != 0 {
                                        PostMessageA(wnd, WM_SYSCOMMAND, SC_RESTORE, 0);
                                    } else {
                                        PostMessageA(wnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0);
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    MOUSEEVENTF_MIDDLEDOWN => {
                        PostMessageA(client_wnd, WM_MBUTTONDOWN, MK_MBUTTON, client_lparam);
                    }
                    MOUSEEVENTF_MIDDLEUP => {
                        PostMessageA(client_wnd, WM_MBUTTONUP, 0, client_lparam);
                    }
                    MOUSEEVENTF_RIGHTDOWN => {
                        SetForegroundWindow(root);
                        PostMessageA(client_wnd, WM_RBUTTONDOWN, MK_RBUTTON, client_lparam);
                    }
                    MOUSEEVENTF_RIGHTUP => {
                        PostMessageA(client_wnd, WM_RBUTTONUP, 0, client_lparam);
                    }
                    MOUSEEVENTF_XDOWN => {
                        PostMessageA(client_wnd, WM_MOUSEWHEEL, MK_RBUTTON, client_lparam);
                    }
                    MOUSEEVENTF_HWHEEL => {
                        PostMessageA(
                            client_wnd,
                            WM_MOUSEHWHEEL,
                            (data << 16) as usize,
                            screen_lparam,
                        );
                    }
                    MOUSEEVENTF_WHEEL => {
                        PostMessageA(
                            client_wnd,
                            WM_MOUSEWHEEL,
                            (data << 16) as usize,
                            screen_lparam,
                        );
                    }
                    _ => {}
                }
            }
        }
    }

    return 1;
}

fn keybd_event(mut flags: u32, vk: u16, scan: u16) -> DWORD {
    // let mut scan = scan;
    // unsafe {
    //     // https://github.com/rustdesk/rustdesk/issues/366
    //     if scan == 0 {
    //         if LAYOUT.is_null() {
    //             let current_window_thread_id =
    //                 GetWindowThreadProcessId(GetForegroundWindow(),
    // std::ptr::null_mut());             LAYOUT =
    // GetKeyboardLayout(current_window_thread_id);         }
    //         scan = MapVirtualKeyExW(vk as _, 0, LAYOUT) as _;
    //     }
    // }

    // if flags & KEYEVENTF_UNICODE == 0 {
    //     if scan >> 8 == 0xE0 || scan >> 8 == 0xE1 {
    //         flags |= winapi::um::winuser::KEYEVENTF_EXTENDEDKEY;
    //     }
    // }
    // let mut union: INPUT_u = unsafe { std::mem::zeroed() };
    // unsafe {
    //     *union.ki_mut() = KEYBDINPUT {
    //         wVk: vk,
    //         wScan: scan,
    //         dwFlags: flags,
    //         time: 0,
    //         dwExtraInfo: ENIGO_INPUT_EXTRA_VALUE,
    //     };
    // }
    // let mut inputs = [INPUT {
    //     type_: INPUT_KEYBOARD,
    //     u: union,
    // }; 1];
    // unsafe {
    //     SendInput(
    //         inputs.len() as UINT,
    //         inputs.as_mut_ptr(),
    //         size_of::<INPUT>() as c_int,
    //     )
    // }
    // println!("{flags} {vk} {scan}");

    return 1;
}

fn get_error() -> String {
    unsafe {
        let buff_size = 256;
        let mut buff: Vec<u16> = Vec::with_capacity(buff_size);
        buff.resize(buff_size, 0);
        let errno = GetLastError();
        let chars_copied = FormatMessageW(
            FORMAT_MESSAGE_IGNORE_INSERTS
                | FORMAT_MESSAGE_FROM_SYSTEM
                | FORMAT_MESSAGE_ARGUMENT_ARRAY,
            std::ptr::null(),
            errno,
            0,
            buff.as_mut_ptr(),
            (buff_size + 1) as u32,
            std::ptr::null_mut(),
        );
        if chars_copied == 0 {
            return "".to_owned();
        }
        let mut curr_char: usize = chars_copied as usize;
        while curr_char > 0 {
            let ch = buff[curr_char];

            if ch >= ' ' as u16 {
                break;
            }
            curr_char -= 1;
        }
        let sl = std::slice::from_raw_parts(buff.as_ptr(), curr_char);
        let err_msg = String::from_utf16(sl);
        return err_msg.unwrap_or("".to_owned());
    }
}

impl MouseControllable for Enigo {
    fn as_any(&self) -> &dyn std::any::Any {
        self
    }

    fn as_mut_any(&mut self) -> &mut dyn std::any::Any {
        self
    }

    fn mouse_move_to(&mut self, x: i32, y: i32) {
        mouse_event(
            MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK,
            0,
            x,
            y,
            // (x - unsafe { GetSystemMetrics(SM_XVIRTUALSCREEN) }) * 65535
            //     / unsafe { GetSystemMetrics(SM_CXVIRTUALSCREEN) },
            // (y - unsafe { GetSystemMetrics(SM_YVIRTUALSCREEN) }) * 65535
            //     / unsafe { GetSystemMetrics(SM_CYVIRTUALSCREEN) },
        );
    }

    fn mouse_move_relative(&mut self, x: i32, y: i32) {
        mouse_event(MOUSEEVENTF_MOVE, 0, x, y);
    }

    fn mouse_down(&mut self, button: MouseButton) -> crate::ResultType {
        let res = mouse_event(
            match button {
                MouseButton::Left => MOUSEEVENTF_LEFTDOWN,
                MouseButton::Middle => MOUSEEVENTF_MIDDLEDOWN,
                MouseButton::Right => MOUSEEVENTF_RIGHTDOWN,
                MouseButton::Back => MOUSEEVENTF_XDOWN,
                MouseButton::Forward => MOUSEEVENTF_XDOWN,
                _ => {
                    log::info!("Unsupported button {:?}", button);
                    return Ok(());
                }
            },
            match button {
                MouseButton::Back => XBUTTON1 as u32 * WHEEL_DELTA as u32,
                MouseButton::Forward => XBUTTON2 as u32 * WHEEL_DELTA as u32,
                _ => 0,
            },
            0,
            0,
        );
        if res == 0 {
            let err = get_error();
            if !err.is_empty() {
                return Err(err.into());
            }
        }
        Ok(())
    }

    fn mouse_up(&mut self, button: MouseButton) {
        mouse_event(
            match button {
                MouseButton::Left => MOUSEEVENTF_LEFTUP,
                MouseButton::Middle => MOUSEEVENTF_MIDDLEUP,
                MouseButton::Right => MOUSEEVENTF_RIGHTUP,
                MouseButton::Back => MOUSEEVENTF_XUP,
                MouseButton::Forward => MOUSEEVENTF_XUP,
                _ => {
                    log::info!("Unsupported button {:?}", button);
                    return;
                }
            },
            match button {
                MouseButton::Back => XBUTTON1 as _,
                MouseButton::Forward => XBUTTON2 as _,
                _ => 0,
            },
            0,
            0,
        );
    }

    fn mouse_click(&mut self, button: MouseButton) {
        self.mouse_down(button).ok();
        self.mouse_up(button);
    }

    fn mouse_scroll_x(&mut self, length: i32) {
        mouse_event(MOUSEEVENTF_HWHEEL, length as _, 0, 0);
    }

    fn mouse_scroll_y(&mut self, length: i32) {
        mouse_event(MOUSEEVENTF_WHEEL, length as _, 0, 0);
    }
}

impl KeyboardControllable for Enigo {
    fn as_any(&self) -> &dyn std::any::Any {
        self
    }

    fn as_mut_any(&mut self) -> &mut dyn std::any::Any {
        self
    }

    fn key_sequence(&mut self, sequence: &str) {
        let mut buffer = [0; 2];

        for c in sequence.chars() {
            // Windows uses uft-16 encoding. We need to check
            // for variable length characters. As such some
            // characters can be 32 bit long and those are
            // encoded in such called hight and low surrogates
            // each 16 bit wide that needs to be send after
            // another to the SendInput function without
            // being interrupted by "keyup"
            let result = c.encode_utf16(&mut buffer);
            if result.len() == 1 {
                self.unicode_key_click(result[0]);
            } else {
                for utf16_surrogate in result {
                    self.unicode_key_down(utf16_surrogate.clone());
                }
                // do i need to produce a keyup?
                // self.unicode_key_up(0);
            }
        }
    }

    fn key_click(&mut self, key: Key) {
        let vk = self.key_to_keycode(key);
        keybd_event(0, vk, 0);
        keybd_event(KEYEVENTF_KEYUP, vk, 0);
    }

    fn key_down(&mut self, key: Key) -> crate::ResultType {
        match &key {
            Key::Layout(c) => {
                // to-do: dup code
                // https://github.com/rustdesk/rustdesk/blob/1bc0dd791ed8344997024dc46626bd2ca7df73d2/src/server/input_service.rs#L1348
                let code = self.get_layoutdependent_keycode(*c);
                if code as u16 != 0xFFFF {
                    let vk = code & 0x00FF;
                    let flag = code >> 8;
                    let modifiers = [Key::Shift, Key::Control, Key::Alt];
                    let mod_len = modifiers.len();
                    for pos in 0..mod_len {
                        if flag & (0x0001 << pos) != 0 {
                            self.key_down(modifiers[pos])?;
                        }
                    }

                    unsafe {
                        if GLOBAL_WND != 0 as _ {
                            PostMessageA(GLOBAL_WND, WM_KEYDOWN, vk as usize, 0);
                        } else {
                            let res = keybd_event(0, vk, 0);
                            let err = if res == 0 { get_error() } else { "".to_owned() };
                            if !err.is_empty() {
                                return Err(err.into());
                            }
                        }
                    }

                    for pos in 0..mod_len {
                        let rpos = mod_len - 1 - pos;
                        if flag & (0x0001 << rpos) != 0 {
                            self.key_up(modifiers[pos]);
                        }
                    }
                } else {
                    return Err(format!("Failed to get keycode of {}", c).into());
                }
            }
            _ => {
                let code = self.key_to_keycode(key);
                if code == 0 || code == 65535 {
                    return Err("".into());
                }
                unsafe {
                    if GLOBAL_WND != 0 as _ {
                        PostMessageA(GLOBAL_WND, WM_KEYDOWN, code as usize, 0);
                    } else {
                        let res = keybd_event(0, code, 0);
                        if res == 0 {
                            let err = get_error();
                            if !err.is_empty() {
                                return Err(err.into());
                            }
                        }
                    }
                }
            }
        }
        Ok(())
    }

    fn key_up(&mut self, key: Key) {
        let code = self.key_to_keycode(key);
        unsafe {
            if GLOBAL_WND != 0 as _ {
                PostMessageA(GLOBAL_WND, WM_KEYUP, code as usize, 0);
            } else {
                keybd_event(KEYEVENTF_KEYUP, code, 0);
            }
        }
    }

    fn get_key_state(&mut self, key: Key) -> bool {
        let keycode = self.key_to_keycode(key);
        let x = unsafe { GetKeyState(keycode as _) };
        if key == Key::CapsLock || key == Key::NumLock || key == Key::Scroll {
            return (x & 0x1) == 0x1;
        }
        return (x as u16 & 0x8000) == 0x8000;
    }
}

impl Enigo {
    /// Gets the (width, height) of the main display in screen coordinates
    /// (pixels).
    ///
    /// # Example
    ///
    /// ```no_run
    /// use enigo::*;
    /// let mut size = Enigo::main_display_size();
    /// ```
    pub fn main_display_size() -> (usize, usize) {
        let w = unsafe { GetSystemMetrics(SM_CXSCREEN) as usize };
        let h = unsafe { GetSystemMetrics(SM_CYSCREEN) as usize };
        (w, h)
    }

    /// Gets the location of mouse in screen coordinates (pixels).
    ///
    /// # Example
    ///
    /// ```no_run
    /// use enigo::*;
    /// let mut location = Enigo::mouse_location();
    /// ```
    pub fn mouse_location() -> (i32, i32) {
        let mut point = POINT { x: 0, y: 0 };
        let result = unsafe { GetCursorPos(&mut point) };
        if result != 0 {
            (point.x, point.y)
        } else {
            (0, 0)
        }
    }

    fn unicode_key_click(&self, unicode_char: u16) {
        self.unicode_key_down(unicode_char);
        self.unicode_key_up(unicode_char);
    }

    fn unicode_key_down(&self, unicode_char: u16) {
        unsafe {
            if GLOBAL_WND != 0 as _ {
                PostMessageA(GLOBAL_WND, WM_CHAR, unicode_char as _, 0);
            } else {
                keybd_event(KEYEVENTF_UNICODE, 0, unicode_char);
            }
        }
    }

    fn unicode_key_up(&self, unicode_char: u16) {
        // WM_CHAR handles the character input, explicit KeyUp is often not needed for WM_CHAR semantics
        // in basic automation, but for completeness or if fallback is used:
        unsafe {
            if GLOBAL_WND == 0 as _ {
                keybd_event(KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, unicode_char);
            }
        }
    }

    fn key_to_keycode(&self, key: Key) -> u16 {
        // do not use the codes from crate winapi they're
        // wrongly typed with i32 instead of i16 use the
        // ones provided by win/keycodes.rs that are prefixed
        // with an 'E' infront of the original name
        #[allow(deprecated)]
        // I mean duh, we still need to support deprecated keys until they're removed
        match key {
            Key::Alt => EVK_MENU,
            Key::Backspace => EVK_BACK,
            Key::CapsLock => EVK_CAPITAL,
            Key::Control => EVK_LCONTROL,
            Key::Delete => EVK_DELETE,
            Key::DownArrow => EVK_DOWN,
            Key::End => EVK_END,
            Key::Escape => EVK_ESCAPE,
            Key::F1 => EVK_F1,
            Key::F10 => EVK_F10,
            Key::F11 => EVK_F11,
            Key::F12 => EVK_F12,
            Key::F2 => EVK_F2,
            Key::F3 => EVK_F3,
            Key::F4 => EVK_F4,
            Key::F5 => EVK_F5,
            Key::F6 => EVK_F6,
            Key::F7 => EVK_F7,
            Key::F8 => EVK_F8,
            Key::F9 => EVK_F9,
            Key::Home => EVK_HOME,
            Key::LeftArrow => EVK_LEFT,
            Key::Option => EVK_MENU,
            Key::PageDown => EVK_NEXT,
            Key::PageUp => EVK_PRIOR,
            Key::Return => EVK_RETURN,
            Key::RightArrow => EVK_RIGHT,
            Key::Shift => EVK_SHIFT,
            Key::Space => EVK_SPACE,
            Key::Tab => EVK_TAB,
            Key::UpArrow => EVK_UP,
            Key::Numpad0 => EVK_NUMPAD0,
            Key::Numpad1 => EVK_NUMPAD1,
            Key::Numpad2 => EVK_NUMPAD2,
            Key::Numpad3 => EVK_NUMPAD3,
            Key::Numpad4 => EVK_NUMPAD4,
            Key::Numpad5 => EVK_NUMPAD5,
            Key::Numpad6 => EVK_NUMPAD6,
            Key::Numpad7 => EVK_NUMPAD7,
            Key::Numpad8 => EVK_NUMPAD8,
            Key::Numpad9 => EVK_NUMPAD9,
            Key::Cancel => EVK_CANCEL,
            Key::Clear => EVK_CLEAR,
            Key::Pause => EVK_PAUSE,
            Key::Kana => EVK_KANA,
            Key::Hangul => EVK_HANGUL,
            Key::Junja => EVK_JUNJA,
            Key::Final => EVK_FINAL,
            Key::Hanja => EVK_HANJA,
            Key::Kanji => EVK_KANJI,
            Key::Convert => EVK_CONVERT,
            Key::Select => EVK_SELECT,
            Key::Print => EVK_PRINT,
            Key::Execute => EVK_EXECUTE,
            Key::Snapshot => EVK_SNAPSHOT,
            Key::Insert => EVK_INSERT,
            Key::Help => EVK_HELP,
            Key::Sleep => EVK_SLEEP,
            Key::Separator => EVK_SEPARATOR,
            Key::Mute => EVK_VOLUME_MUTE,
            Key::VolumeDown => EVK_VOLUME_DOWN,
            Key::VolumeUp => EVK_VOLUME_UP,
            Key::Scroll => EVK_SCROLL,
            Key::NumLock => EVK_NUMLOCK,
            Key::RWin => EVK_RWIN,
            Key::Apps => EVK_APPS,
            Key::Add => EVK_ADD,
            Key::Multiply => EVK_MULTIPLY,
            Key::Decimal => EVK_DECIMAL,
            Key::Subtract => EVK_SUBTRACT,
            Key::Divide => EVK_DIVIDE,
            Key::NumpadEnter => EVK_RETURN,
            Key::Equals => '=' as _,
            Key::RightShift => EVK_RSHIFT,
            Key::RightControl => EVK_RCONTROL,
            Key::RightAlt => EVK_RMENU,

            Key::Raw(raw_keycode) => raw_keycode,
            Key::Super | Key::Command | Key::Windows | Key::Meta => EVK_LWIN,
            Key::Layout(..) => {
                // unreachable
                0
            }
        }
    }

    fn get_layoutdependent_keycode(&self, chr: char) -> u16 {
        unsafe {
            LAYOUT = std::ptr::null_mut();
        }
        // NOTE VkKeyScanW uses the current keyboard LAYOUT
        // to specify a LAYOUT use VkKeyScanExW and GetKeyboardLayout
        // or load one with LoadKeyboardLayoutW
        let current_window_thread_id =
            unsafe { GetWindowThreadProcessId(GetForegroundWindow(), std::ptr::null_mut()) };
        unsafe { LAYOUT = GetKeyboardLayout(current_window_thread_id) };
        unsafe { VkKeyScanExW(chr as _, LAYOUT) as _ }
    }
}
