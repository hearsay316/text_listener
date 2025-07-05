// 禁用在 Windows 上运行时弹出的控制台窗口
// #![windows_subsystem = "windows"]

use std::io;

// --- 方法一：轮询剪贴板 ---
// 这是最简单、最稳定的方法。
mod clipboard_poller {
    use arboard::Clipboard;
    use std::{thread, time::Duration};

    pub fn run() {
        println!("方法一：剪贴板轮询模式已启动。");
        println!("请在任何地方复制文本 (Ctrl+C)，这里会显示出来。按 Ctrl+C 退出此程序。");

        let mut clipboard = Clipboard::new().expect("无法初始化剪贴板");
        let mut previous_text = clipboard.get_text().unwrap_or_default();

        loop {
            let current_text = clipboard.get_text().unwrap_or_default();
            if !current_text.is_empty() && current_text != previous_text {
                println!("\n--- [剪贴板更新] ---");
                println!("{}", current_text);
                println!("--- [内容结束] ---\n");
                previous_text = current_text;
            }
            thread::sleep(Duration::from_millis(500)); // 每 500 毫秒检查一次
        }
    }
}

// --- 方法三：全局鼠标钩子 + 模拟按键 ---
// 这是一个"黑科技"方法，有侵入性，并且需要 unsafe 代码。
// 它会监听鼠标左键的抬起，然后模拟 Ctrl+C，再从剪贴板读取。
mod global_hook_simulator {
    use arboard::Clipboard;
    use std::{sync::atomic::{AtomicBool, Ordering}, thread, time::{Duration, Instant}};
    use windows::Win32::{
        Foundation::{LPARAM, LRESULT, WPARAM},
        System::Threading::GetCurrentThreadId,
        UI::{
            Input::KeyboardAndMouse::{
                SendInput, INPUT, INPUT_KEYBOARD, KEYBDINPUT, KEYEVENTF_KEYUP, VK_C, VK_LCONTROL,
                VK_ESCAPE,
            },
            WindowsAndMessaging::{
                CallNextHookEx, GetMessageW, SetWindowsHookExW, UnhookWindowsHookEx, HHOOK, MSG,
                WH_MOUSE_LL, WM_LBUTTONUP, PostThreadMessageW, WM_USER, TranslateMessage, DispatchMessageW,
                WM_QUIT, WM_KEYDOWN, WH_KEYBOARD_LL, KBDLLHOOKSTRUCT,
            },
        },
        System::Console::{SetConsoleCtrlHandler, CTRL_C_EVENT, CTRL_BREAK_EVENT, CTRL_CLOSE_EVENT, CTRL_LOGOFF_EVENT, CTRL_SHUTDOWN_EVENT},
    };

    // 全局变量来存储钩子句柄和状态
    static mut MOUSE_HOOK: Option<HHOOK> = None;
    static mut KEYBOARD_HOOK: Option<HHOOK> = None;
    static mut LAST_CLICK_TIME: Option<Instant> = None;
    static mut MAIN_THREAD_ID: u32 = 0;
    static SHOULD_EXIT: AtomicBool = AtomicBool::new(false);
    static IS_SIMULATING_CTRL_C: AtomicBool = AtomicBool::new(false);
    
    // 控制台信号处理函数
    unsafe extern "system" fn console_ctrl_handler(ctrl_type: u32) -> windows::Win32::Foundation::BOOL {
        match ctrl_type {
            CTRL_C_EVENT | CTRL_BREAK_EVENT => {
                // 只有在程序模拟 Ctrl+C 时才拦截信号，否则让用户正常操作通过
                if IS_SIMULATING_CTRL_C.load(Ordering::Relaxed) {
                    println!("[调试] 拦截了程序模拟的 Ctrl+C 信号，防止程序退出");
                    windows::Win32::Foundation::BOOL::from(true) // 返回 TRUE 表示已处理该信号
                } else {
                    println!("[事件] 检测到用户的 Ctrl+C 操作，正在优雅退出...");
                    SHOULD_EXIT.store(true, Ordering::Relaxed);
                    // 发送退出消息到主线程
                    let _ = PostThreadMessageW(MAIN_THREAD_ID, WM_QUIT, WPARAM(0), LPARAM(0));
                    windows::Win32::Foundation::BOOL::from(true) // 返回 TRUE 表示我们已经处理了这个信号
                }
            }
            CTRL_CLOSE_EVENT | CTRL_LOGOFF_EVENT | CTRL_SHUTDOWN_EVENT => {
                println!("[事件] 检测到系统关闭信号，正在清理资源...");
                SHOULD_EXIT.store(true, Ordering::Relaxed);
                // 给程序一点时间来清理资源
                thread::sleep(Duration::from_millis(100));
                windows::Win32::Foundation::BOOL::from(true)
            }
            _ => windows::Win32::Foundation::BOOL::from(false), // 其他信号交给默认处理器
        }
    }

    // 模拟按下和释放 Ctrl+C
    fn simulate_ctrl_c() {
        // 设置标志，表示程序正在模拟 Ctrl+C
        IS_SIMULATING_CTRL_C.store(true, Ordering::Relaxed);
        
        // 需要 unsafe 因为我们在调用系统 API
        unsafe {
            let inputs = &mut [
                // Press LCtrl
                INPUT {
                    r#type: INPUT_KEYBOARD,
                    Anonymous: windows::Win32::UI::Input::KeyboardAndMouse::INPUT_0 {
                        ki: KEYBDINPUT {
                            wVk: VK_LCONTROL,
                            ..Default::default()
                        },
                    },
                },
                // Press C
                INPUT {
                    r#type: INPUT_KEYBOARD,
                    Anonymous: windows::Win32::UI::Input::KeyboardAndMouse::INPUT_0 {
                        ki: KEYBDINPUT {
                            wVk: VK_C,
                            ..Default::default()
                        },
                    },
                },
                // Release C
                INPUT {
                    r#type: INPUT_KEYBOARD,
                    Anonymous: windows::Win32::UI::Input::KeyboardAndMouse::INPUT_0 {
                        ki: KEYBDINPUT {
                            wVk: VK_C,
                            dwFlags: KEYEVENTF_KEYUP,
                            ..Default::default()
                        },
                    },
                },
                // Release LCtrl
                INPUT {
                    r#type: INPUT_KEYBOARD,
                    Anonymous: windows::Win32::UI::Input::KeyboardAndMouse::INPUT_0 {
                        ki: KEYBDINPUT {
                            wVk: VK_LCONTROL,
                            dwFlags: KEYEVENTF_KEYUP,
                            ..Default::default()
                        },
                    },
                },
            ];
            SendInput(inputs, std::mem::size_of::<INPUT>() as i32);
        }
        
        // 短暂延迟后清除标志，确保信号处理器有时间处理
        thread::sleep(Duration::from_millis(50));
        IS_SIMULATING_CTRL_C.store(false, Ordering::Relaxed);
    }

    // 键盘钩子的回调函数
    unsafe extern "system" fn low_level_keyboard_proc(
        n_code: i32,
        w_param: WPARAM,
        l_param: LPARAM,
    ) -> LRESULT {
        if n_code >= 0 && w_param.0 as u32 == WM_KEYDOWN {
             let kbd_struct = *(l_param.0 as *const KBDLLHOOKSTRUCT);
             if kbd_struct.vkCode == VK_ESCAPE.0 as u32 {
                println!("[事件] 检测到 ESC 键，准备退出...");
                SHOULD_EXIT.store(true, Ordering::Relaxed);
                let _ = PostThreadMessageW(MAIN_THREAD_ID, WM_QUIT, WPARAM(0), LPARAM(0));
                return LRESULT(1); // 阻止 ESC 键传递给其他应用
            }
        }
        CallNextHookEx(KEYBOARD_HOOK.unwrap(), n_code, w_param, l_param)
    }

    // 鼠标钩子的回调函数
    // 这个函数会在每次鼠标事件发生时被 Windows 调用
    unsafe extern "system" fn low_level_mouse_proc(
        n_code: i32,
        w_param: WPARAM,
        l_param: LPARAM,
    ) -> LRESULT {
        if n_code >= 0 {
            // 当鼠标左键抬起时
            if w_param.0 as u32 == WM_LBUTTONUP {
                let now = Instant::now();
                
                // 防抖动：如果距离上次点击时间太短，则忽略
                if let Some(last_time) = LAST_CLICK_TIME {
                    if now.duration_since(last_time) < Duration::from_millis(300) {
                        return CallNextHookEx(MOUSE_HOOK.unwrap(), n_code, w_param, l_param);
                    }
                }
                LAST_CLICK_TIME = Some(now);
                
                println!("[事件] 检测到鼠标左键抬起。");
                
                // 不在钩子回调中执行耗时操作，而是发送消息到主线程处理
                let _ = PostThreadMessageW(MAIN_THREAD_ID, WM_USER + 1, WPARAM(0), LPARAM(0));
            }
        }
        // 把事件传递给下一个钩子，否则整个系统会卡住！
        CallNextHookEx(MOUSE_HOOK.unwrap(), n_code, w_param, l_param)
    }
    
    // 处理文本捕获的函数，在主线程中执行
    fn handle_text_capture() {
        match Clipboard::new() {
            Ok(mut clipboard) => {
                // 1. 保存用户当前的剪贴板内容
                let user_clipboard_backup = clipboard.get_text().ok();
                println!("[操作] 已备份用户剪贴板内容");
                
                // 2. 模拟 Ctrl+C
                println!("[操作] 正在模拟 Ctrl+C...");
                simulate_ctrl_c();

                // 3. 等待一小段时间，让目标应用有时间把文本放到剪贴板
                thread::sleep(Duration::from_millis(150));

                // 4. 从剪贴板读取捕获的内容
                match clipboard.get_text() {
                    Ok(captured_text) => {
                        if !captured_text.is_empty() && captured_text.trim().len() > 0 {
                            // 检查是否与用户备份的内容相同，避免显示用户自己的内容
                            let is_same_as_backup = user_clipboard_backup
                                .as_ref()
                                .map(|backup| backup == &captured_text)
                                .unwrap_or(false);
                            
                            if !is_same_as_backup {
                                println!("\n--- [自动捕获内容] ---");
                                println!("{}", captured_text);
                                println!("--- [内容结束] ---\n");
                            } else {
                                println!("[结果] 检测到的内容与用户剪贴板相同，可能没有新的选中文本。");
                            }
                        } else {
                            println!("[结果] 剪贴板为空或只包含空白字符，可能没有选中文本。");
                        }
                    }
                    Err(e) => println!("[错误] 读取剪贴板失败: {:?}", e),
                }
                
                // 5. 恢复用户的剪贴板内容
                if let Some(backup_content) = user_clipboard_backup {
                    if let Err(e) = clipboard.set_text(backup_content) {
                        println!("[警告] 恢复用户剪贴板内容失败: {:?}", e);
                    } else {
                        println!("[操作] 已恢复用户剪贴板内容");
                    }
                } else {
                    // 如果用户原本剪贴板为空，清空剪贴板
                    let _ = clipboard.set_text("".to_string());
                    println!("[操作] 已清空剪贴板（用户原本为空）");
                }
            }
            Err(e) => println!("[错误] 无法初始化剪贴板: {:?}", e),
        }
    }

    pub fn run() {
        println!("方法三：全局鼠标钩子模式已启动。");
        println!("请在任何地方用鼠标选中一段文本，然后松开左键。");
        println!("✅ 改进：程序会自动备份和恢复你的剪贴板内容，不影响正常使用");
        println!("提示：程序会自动过滤重复点击，300ms内的连续点击会被忽略。");
        println!("退出方式：按 ESC 键退出，或关闭此控制台窗口");
        
        // 重置退出标志
        SHOULD_EXIT.store(false, Ordering::Relaxed);

        // 需要 unsafe 因为我们在设置一个全局钩子
        unsafe {
            // 设置控制台信号处理器，防止模拟的 Ctrl+C 导致程序退出
            if let Err(e) = SetConsoleCtrlHandler(Some(console_ctrl_handler), true) {
                println!("[警告] 设置控制台信号处理器失败: {:?}", e);
            } else {
                println!("[状态] 控制台信号处理器已设置，程序不会因模拟 Ctrl+C 而退出。");
            }
            // 获取当前线程ID
            MAIN_THREAD_ID = GetCurrentThreadId();
            
            // 设置键盘钩子用于检测 ESC 键
            let keyboard_hook = match SetWindowsHookExW(
                WH_KEYBOARD_LL,
                Some(low_level_keyboard_proc),
                None,
                0,
            ) {
                Ok(h) => h,
                Err(e) => {
                    println!("[错误] 设置键盘钩子失败: {:?}", e);
                    return;
                }
            };
            KEYBOARD_HOOK = Some(keyboard_hook);
            
            // 设置一个低级鼠标钩子
            let mouse_hook = match SetWindowsHookExW(
                WH_MOUSE_LL,
                Some(low_level_mouse_proc),
                None, // hmod: None 表示钩子与任何特定模块无关
                0,    // dwThreadId: 0 表示这是一个全局钩子
            ) {
                Ok(h) => h,
                Err(e) => {
                    println!("[错误] 设置鼠标钩子失败: {:?}", e);
                    // 如果鼠标钩子失败，也要清理键盘钩子
                    let _ = UnhookWindowsHookEx(keyboard_hook);
                    return;
                }
            };
            MOUSE_HOOK = Some(mouse_hook);

            println!("[状态] 鼠标和键盘钩子已成功安装，开始监听...");

            // 运行一个消息循环，这是接收钩子事件所必需的
            println!("[状态] 钩子已激活，开始持续监听鼠标事件...");
            println!("[提示] 现在可以在任何地方选中文本并松开鼠标左键进行捕获。");
            
            let mut msg: MSG = Default::default();
            loop {
                // 检查是否需要退出
                if SHOULD_EXIT.load(Ordering::Relaxed) {
                    println!("[状态] 检测到退出信号，正在停止监听...");
                    break;
                }
                
                let result = GetMessageW(&mut msg, None, 0, 0);
                
                // 检查是否收到退出消息
                if !result.as_bool() || msg.message == WM_QUIT {
                    println!("[状态] 收到系统退出信号，正在停止监听...");
                    break;
                }
                
                // 检查是否是我们的自定义消息
                if msg.message == WM_USER + 1 {
                    handle_text_capture();
                } else {
                    // 处理其他消息
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
            }

            // 程序退出前，卸载钩子
            if let Err(e) = UnhookWindowsHookEx(mouse_hook) {
                println!("[警告] 卸载鼠标钩子时出错: {:?}", e);
            } else {
                println!("[状态] 鼠标钩子已成功卸载。");
            }
            
            if let Err(e) = UnhookWindowsHookEx(keyboard_hook) {
                println!("[警告] 卸载键盘钩子时出错: {:?}", e);
            } else {
                println!("[状态] 键盘钩子已成功卸载。");
            }
        }
    }
}

// --- 方法二：Windows UI Automation ---
// 这是最“正确”但也是最复杂的方法。
// 由于其极端复杂性，提供一个完整的、健壮的示例非常困难。
// 下面的代码是一个“概念验证”，展示了其基本思路，但省略了大量的错误处理和复杂的逻辑。
mod ui_automation_conceptual {
    use std::{thread, time::Duration};
    use windows::{
        core::ComInterface,
        Win32::{
            System::Com::{
                CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER,
                COINIT_MULTITHREADED,
            },
            UI::Accessibility::{
                CUIAutomation, IUIAutomation, IUIAutomationTextPattern, UIA_TextPatternId,
            },
        },
    };

    pub fn run() {
        println!("方法二：UI Automation 模式 (概念示例)。");
        println!("这个示例将尝试获取当前焦点窗口中选中的文本。");
        println!("注意：这非常复杂，且不保证对所有应用都有效。");

        unsafe {
            if let Err(e) = CoInitializeEx(None, COINIT_MULTITHREADED) {
                println!("COM 初始化失败: {:?}", e);
                return;
            }

            let automation: IUIAutomation =
                match CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER) {
                    Ok(inst) => inst,
                    Err(e) => {
                        println!("创建 UI Automation 实例失败: {:?}", e);
                        CoUninitialize();
                        return;
                    }
                };

            println!("请在 5 秒内切换到另一个窗口并选中文本...");
            thread::sleep(Duration::from_secs(5));

            match automation.GetFocusedElement() {
                Ok(focused_element) => {
                    println!("成功获取到焦点元素。正在尝试获取文本模式...");
                    
                    match focused_element.GetCurrentPattern(UIA_TextPatternId) {
                        Ok(pattern_unknown) => {
                            // 需要将获取到的 IUnknown 转换为 IUIAutomationTextPattern
                            match pattern_unknown.cast::<IUIAutomationTextPattern>() {
                                Ok(text_pattern) => {
                                    println!("成功获取文本模式。正在获取选区...");
                                    match text_pattern.GetSelection() {
                                        Ok(selection) => {
                                            let selection_len = selection.Length().unwrap_or(0);
                                            if selection_len > 0 {
                                                println!("找到 {} 个选区。", selection_len);
                                                for i in 0..selection_len {
                                                    if let Ok(range) = selection.GetElement(i) {
                                                        if let Ok(text) = range.GetText(-1) {
                                                            println!("--- [UIA 捕获内容] ---");
                                                            println!("{}", text.to_string());
                                                            println!("--- [内容结束] ---\n");
                                                        }
                                                    }
                                                }
                                            } else {
                                                println!("未找到任何选区。");
                                            }
                                        }
                                        Err(e) => println!("获取选区失败: {:?}", e),
                                    }
                                }
                                Err(e) => println!("无法将 Pattern 转换为 TextPattern: {:?}", e),
                            }
                        }
                        Err(_) => {
                            println!("此焦点元素不支持文本模式 (TextPattern)。");
                        }
                    }
                }
                Err(e) => println!("获取焦点元素失败: {:?}", e),
            }

            CoUninitialize();
        }
    }
}

fn main() {
    loop {
        println!("\n请选择要运行的 Demo 模式:");
        println!("1. 剪贴板轮询 (最稳定，推荐)");
        println!("2. UI Automation (最复杂，概念演示)");
        println!("3. 全局鼠标钩子 (有风险，侵入式)");
        println!("q. 退出");
        print!("请输入选项 (1, 2, 3, q): ");

        io::Write::flush(&mut io::stdout()).unwrap();

        let mut choice = String::new();
        io::stdin().read_line(&mut choice).unwrap();

        match choice.trim() {
            "1" => clipboard_poller::run(),
            "2" => ui_automation_conceptual::run(),
            "3" => global_hook_simulator::run(),
            "q" | "Q" => {
                println!("程序退出。");
                break;
            }
            _ => println!("无效选项，请重新输入。"),
        }
    }
}
