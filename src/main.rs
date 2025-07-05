// 禁用在 Windows 上运行时弹出的控制台窗口
// #![windows_subsystem = "windows"]

use std::io;
use windows::Win32::System::Threading::GetCurrentThreadId;

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
    use std::{sync::{Arc, Mutex}, thread, time::{Duration, Instant}};
    use windows::Win32::{
        Foundation::{LPARAM, LRESULT, WPARAM},
        UI::{
            Input::KeyboardAndMouse::{
                SendInput, INPUT, INPUT_KEYBOARD, KEYBDINPUT, KEYEVENTF_KEYUP, VK_C, VK_LCONTROL,
            },
            WindowsAndMessaging::{
                CallNextHookEx, GetMessageW, SetWindowsHookExW, UnhookWindowsHookEx, HHOOK, MSG,
                WH_MOUSE_LL, WM_LBUTTONUP, PostThreadMessageW, WM_USER,
            },
        },
    };

    // 全局变量来存储钩子句柄和状态
    static mut MOUSE_HOOK: Option<HHOOK> = None;
    static mut LAST_CLICK_TIME: Option<Instant> = None;
    static mut MAIN_THREAD_ID: u32 = 0;

    // 模拟按下和释放 Ctrl+C
    fn simulate_ctrl_c() {
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
                PostThreadMessageW(MAIN_THREAD_ID, WM_USER + 1, WPARAM(0), LPARAM(0));
            }
        }
        // 把事件传递给下一个钩子，否则整个系统会卡住！
        CallNextHookEx(MOUSE_HOOK.unwrap(), n_code, w_param, l_param)
    }
    
    // 处理文本捕获的函数，在主线程中执行
    fn handle_text_capture() {
        // 1. 模拟 Ctrl+C
        println!("[操作] 正在模拟 Ctrl+C...");
        simulate_ctrl_c();

        // 2. 等待一小段时间，让目标应用有时间把文本放到剪贴板
        thread::sleep(Duration::from_millis(150));

        // 3. 从剪贴板读取
        match Clipboard::new() {
            Ok(mut clipboard) => {
                match clipboard.get_text() {
                    Ok(text) => {
                        if !text.is_empty() && text.trim().len() > 0 {
                            println!("\n--- [自动捕获内容] ---");
                            println!("{}", text);
                            println!("--- [内容结束] ---\n");
                        } else {
                            println!("[结果] 剪贴板为空或只包含空白字符，可能没有选中文本。");
                        }
                    }
                    Err(e) => println!("[错误] 读取剪贴板失败: {:?}", e),
                }
            }
            Err(e) => println!("[错误] 无法初始化剪贴板: {:?}", e),
        }
    }

    pub fn run() {
        println!("方法三：全局鼠标钩子模式已启动。");
        println!("请在任何地方用鼠标选中一段文本，然后松开左键。");
        println!("警告：此模式会覆盖你的剪贴板。按 Ctrl+C 退出此程序。");
        println!("提示：程序会自动过滤重复点击，300ms内的连续点击会被忽略。");

        // 需要 unsafe 因为我们在设置一个全局钩子
        unsafe {
            // 获取当前线程ID
            MAIN_THREAD_ID = windows::Win32::System::Threading::GetCurrentThreadId();
            
            // 设置一个低级鼠标钩子
            let hook = match SetWindowsHookExW(
                WH_MOUSE_LL,
                Some(low_level_mouse_proc),
                None, // hmod: None 表示钩子与任何特定模块无关
                0,    // dwThreadId: 0 表示这是一个全局钩子
            ) {
                Ok(h) => h,
                Err(e) => {
                    println!("[错误] 设置鼠标钩子失败: {:?}", e);
                    return;
                }
            };
            MOUSE_HOOK = Some(hook);

            println!("[状态] 鼠标钩子已成功安装，开始监听...");

            // 运行一个消息循环，这是接收钩子事件所必需的
            let mut msg: MSG = Default::default();
            while GetMessageW(&mut msg, None, 0, 0).as_bool() {
                // 检查是否是我们的自定义消息
                if msg.message == WM_USER + 1 {
                    handle_text_capture();
                }
                // 处理其他消息
            }

            // 程序退出前，卸载钩子
            if let Err(e) = UnhookWindowsHookEx(hook) {
                println!("[警告] 卸载钩子时出错: {:?}", e);
            } else {
                println!("[状态] 鼠标钩子已成功卸载。");
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
