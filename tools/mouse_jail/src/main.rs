use std::{thread, time};
use winapi::um::winuser::{ClipCursor, GetSystemMetrics, SM_CXSCREEN, SM_CYSCREEN, RECT};

fn main() {
    println!("启动鼠标牢笼 (Mouse Jail)...");
    println!("物理鼠标将被限制在主屏幕范围内。");
    println!("按 Ctrl+C 退出。");

    let width = unsafe { GetSystemMetrics(SM_CXSCREEN) };
    let height = unsafe { GetSystemMetrics(SM_CYSCREEN) };

    println!("主屏幕分辨率: {}x{}", width, height);

    let rect = RECT {
        left: 0,
        top: 0,
        right: width,
        bottom: height,
    };

    loop {
        unsafe {
            ClipCursor(&rect);
        }
        thread::sleep(time::Duration::from_secs(1));
    }
}
        thread::sleep(time::Duration::from_secs(1));
    }
}
