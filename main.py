import tkinter as tk
from tkinter import ttk, messagebox
import keyboard
import time
import threading
import random
import json
import os
import ctypes

# --- 全局常量 & 翻译字典 ---
CONFIG_FILE = "paste_config.json"
WINDOW_SIZE_MINI = "300x100"
WINDOW_SIZE_SETTINGS = "400x550"

# 国际化文本字典
TRANSLATIONS = {
    "app_title": {
        "zh": "强制粘贴 - 模拟键盘输入",
        "en": "ForcePaste - Auto Typer"
    },
    "btn_start": {
        "zh": "强制粘贴 (Start)",
        "en": "Force Paste (Start)"
    },
    "btn_settings": {
        "zh": "设置 (Settings)",
        "en": "Settings (设置)"
    },
    "btn_stop": {
        "zh": "正在中断...",
        "en": "Stopping..."
    },
    "btn_wait": {
        "zh": "点击中断",
        "en": "Click to Stop"
    },
    "btn_typing": {
        "zh": "正在输入...",
        "en": "Typing..."
    },
    "msg_empty": {
        "zh": "剪贴板为空!",
        "en": "Clipboard Empty!"
    },
    # 设置页
    "set_title": {
        "zh": "设置",
        "en": "Settings"
    },
    "grp_lang": {
        "zh": "语言 (Language)",
        "en": "Language (语言)"
    },
    "grp_delay": {
        "zh": "延迟设置 (ms)",
        "en": "Delay Settings (ms)"
    },
    "lbl_btn_delay": {
        "zh": "按钮启动延迟:",
        "en": "Button Start Delay:"
    },
    "lbl_hotkey_delay": {
        "zh": "快捷键启动延迟:",
        "en": "Hotkey Start Delay:"
    },
    "lbl_char_delay": {
        "zh": "字符输入延迟:",
        "en": "Char Input Delay:"
    },
    "lbl_jitter": {
        "zh": "随机波动范围:",
        "en": "Random Jitter:"
    },
    "grp_control": {
        "zh": "控制",
        "en": "Controls"
    },
    "lbl_hotkey": {
        "zh": "触发快捷键:",
        "en": "Trigger Hotkey:"
    },
    "grp_misc": {
        "zh": "功能开关",
        "en": "Features"
    },
    "chk_top": {
        "zh": "窗口始终置顶",
        "en": "Always on Top"
    },
    "chk_tab": {
        "zh": "Tab 转 4 空格",
        "en": "Convert Tab to 4 Spaces"
    },
    "chk_shift_enter": {
        "zh": "使用 Shift+Enter 换行",
        "en": "Use Shift+Enter for Newline"
    },
    "chk_staircase": {
        "zh": "启用防楼梯 (适用 Python 编辑器)",
        "en": "Anti-Staircase (For Python editor)"
    },
    "btn_save": {
        "zh": "保存 (Save)",
        "en": "Save (保存)"
    },
    "btn_cancel": {
        "zh": "取消 (Cancel)",
        "en": "Cancel (取消)"
    },
    "err_save": {
        "zh": "配置无效",
        "en": "Invalid Configuration"
    }
}


class ConfigManager:
    """配置管理类：加载、保存和获取用户设置"""

    def __init__(self, filename=CONFIG_FILE):
        self.filename = filename
        self.defaults = {
            "language": "zh",
            "btn_delay": 3000,
            "hotkey_delay": 100,
            "char_delay": 20,
            "random_jitter": 5,
            "hotkey": "ctrl+shift+y",
            "always_on_top": True,
            "tab_to_space": True,
            "anti_staircase": False,
            "shift_enter": False
        }
        self.config = self.load_config()

    def load_config(self):
        """加载配置，若文件缺失或损坏则返回默认值"""
        if not os.path.exists(self.filename):
            return self.defaults.copy()

        try:
            with open(self.filename, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
                # 确保新加入的字段也能被合并
                return {**self.defaults, **user_config}
        except Exception:
            return self.defaults.copy()

    def save_config(self):
        """保存当前配置到文件"""
        with open(self.filename, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4)

    def get(self, key):
        return self.config.get(key, self.defaults.get(key))

    def set(self, key, value):
        self.config[key] = value


class PasteEngine:
    """粘贴引擎核心逻辑"""

    def __init__(self, config_manager):
        self.cfg = config_manager
        self.stop_event = threading.Event()
        self.target_hwnd = None

    def get_clipboard_text(self, root):
        try:
            return root.clipboard_get()
        except tk.TclError:
            return ""

    def smart_sleep(self):
        base = self.cfg.get("char_delay")
        jitter = self.cfg.get("random_jitter")
        actual_delay = base + random.uniform(-jitter, jitter)
        time.sleep(max(0, actual_delay) / 1000.0)

    def process_text(self, text):
        if self.cfg.get("tab_to_space"):
            text = text.replace("\t", "    ")
        return text.replace("\r\n", "\n").replace("\r", "\n")

    def check_focus_safety(self):
        current_hwnd = ctypes.windll.user32.GetForegroundWindow()
        if self.target_hwnd and current_hwnd != self.target_hwnd:
            print("[Safety] Focus lost, stopping input.")
            self.stop_event.set()
            return False
        return True

    def wait_keys_release(self):
        start_time = time.time()
        while time.time() - start_time < 2.0:
            if not any(keyboard.is_pressed(k) for k in ['ctrl', 'alt', 'shift', 'win']):
                return
            time.sleep(0.05)
        print("[Warn] Key release timeout.")

    def execute_paste(self, text, is_hotkey=False):
        if self.stop_event.is_set():
            return

        if is_hotkey:
            delay = self.cfg.get("hotkey_delay")
            if delay > 0:
                time.sleep(delay / 1000.0)
            self.wait_keys_release()

        self.target_hwnd = ctypes.windll.user32.GetForegroundWindow()
        text = self.process_text(text)

        try:
            if self.cfg.get("anti_staircase"):
                self._paste_anti_staircase(text)
            else:
                self._paste_normal(text)
        except Exception as e:
            print(f"Paste Error: {e}")

    def abort(self):
        self.stop_event.set()

    def _perform_newline(self):
        if self.cfg.get("shift_enter"):
            keyboard.send('shift+enter')
        else:
            keyboard.send('enter')
        self.smart_sleep()

    def _paste_normal(self, text):
        for char in text:
            if self.stop_event.is_set() or not self.check_focus_safety():
                return
            if char == '\n':
                self._perform_newline()
            else:
                keyboard.write(char)
                self.smart_sleep()

    def _paste_anti_staircase(self, text):
        lines = text.split('\n')
        total_lines = len(lines)

        for idx, line in enumerate(lines):
            if self.stop_event.is_set() or not self.check_focus_safety(): return

            keyboard.send('home')
            time.sleep(0.01)
            keyboard.write('#')
            time.sleep(0.01)

            for char in line:
                if self.stop_event.is_set() or not self.check_focus_safety(): return
                keyboard.write(char)
                self.smart_sleep()

            if self.stop_event.is_set() or not self.check_focus_safety(): return

            keyboard.send('home')
            time.sleep(0.01)
            keyboard.send('right')
            time.sleep(0.01)
            keyboard.send('backspace')
            time.sleep(0.01)
            keyboard.send('end')
            time.sleep(0.01)

            if idx < total_lines - 1:
                self._perform_newline()


class AppUI:
    """图形界面逻辑"""

    def __init__(self, root):
        self.root = root
        self.cfg = ConfigManager()
        self.engine = PasteEngine(self.cfg)
        self.is_working = False
        self.hotkey_hook = None
        self.vars = {}

        # 初始化界面
        self.setup_window()
        self.update_hotkey()
        self.init_floating_mode()

    def T(self, key):
        lang = self.cfg.get("language")
        return TRANSLATIONS.get(key, {}).get(lang, key)

    def setup_window(self):
        self.root.title(self.T("app_title"))
        self.root.geometry(WINDOW_SIZE_MINI)
        self.root.resizable(False, False)
        self.root.attributes("-topmost", self.cfg.get("always_on_top"))

    def update_hotkey(self):
        if self.hotkey_hook:
            try:
                keyboard.remove_hotkey(self.hotkey_hook)
            except Exception:
                pass

        hotkey_str = self.cfg.get("hotkey")
        try:
            self.hotkey_hook = keyboard.add_hotkey(hotkey_str, self.on_hotkey_triggered)
        except Exception as e:
            print(f"Hotkey Error: {e}")

    def on_hotkey_triggered(self):
        if self.is_working:
            self.stop_sequence()
        else:
            text = self.engine.get_clipboard_text(self.root)
            if text:
                self.start_sequence(text, is_hotkey=True)

    # --- 界面模式切换 ---

    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def init_floating_mode(self):
        """迷你悬浮窗模式"""
        self.clear_window()
        self.root.geometry(WINDOW_SIZE_MINI)
        self.root.title(self.T("app_title"))  # 刷新标题语言

        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        self.btn_paste = tk.Button(frame, text=self.T("btn_start"),
                                   command=self.on_ui_paste_click,
                                   bg="#f0f0f0", font=("Arial", 10, "bold"), height=2)
        self.btn_paste.pack(fill=tk.X, pady=(0, 5))

        self.btn_settings = tk.Button(frame, text=self.T("btn_settings"),
                                      command=self.init_settings_mode,
                                      font=("Arial", 9))
        self.btn_settings.pack(fill=tk.X)

    def init_settings_mode(self):
        """详细设置模式"""
        self.clear_window()
        self.root.geometry(WINDOW_SIZE_SETTINGS)

        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text=self.T("set_title"), font=("Arial", 12, "bold")).pack(pady=(0, 10))

        # --- 语言设置 ---
        grp_lang = tk.LabelFrame(main_frame, text=self.T("grp_lang"))
        grp_lang.pack(fill=tk.X, pady=5)

        lang_var = tk.StringVar(value=self.cfg.get("language"))
        self.vars["language"] = lang_var

        f_lang = tk.Frame(grp_lang)
        f_lang.pack(fill=tk.X, padx=5, pady=5)

        # 语言单选按钮
        tk.Radiobutton(f_lang, text="中文", variable=lang_var, value="zh").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(f_lang, text="English", variable=lang_var, value="en").pack(side=tk.LEFT, padx=10)

        # 延迟设置组
        grp_delay = tk.LabelFrame(main_frame, text=self.T("grp_delay"))
        grp_delay.pack(fill=tk.X, pady=5)

        self._add_config_entry(grp_delay, self.T("lbl_btn_delay"), "btn_delay")
        self._add_config_entry(grp_delay, self.T("lbl_hotkey_delay"), "hotkey_delay")
        self._add_config_entry(grp_delay, self.T("lbl_char_delay"), "char_delay")
        self._add_config_entry(grp_delay, self.T("lbl_jitter"), "random_jitter")

        # 快捷键设置组
        grp_hotkey = tk.LabelFrame(main_frame, text=self.T("grp_control"))
        grp_hotkey.pack(fill=tk.X, pady=5)
        self._add_config_entry(grp_hotkey, self.T("lbl_hotkey"), "hotkey", is_str=True)

        # 杂项设置组
        grp_misc = tk.LabelFrame(main_frame, text=self.T("grp_misc"))
        grp_misc.pack(fill=tk.X, pady=5)

        self._add_check_btn(grp_misc, self.T("chk_top"), "always_on_top")
        self._add_check_btn(grp_misc, self.T("chk_tab"), "tab_to_space")
        self._add_check_btn(grp_misc, self.T("chk_shift_enter"), "shift_enter")

        # 重点功能高亮
        f_stair = tk.Frame(grp_misc)
        f_stair.pack(fill=tk.X, pady=2)
        stair_var = tk.BooleanVar(value=self.cfg.get("anti_staircase"))
        self.vars["anti_staircase"] = stair_var
        tk.Checkbutton(f_stair, text=self.T("chk_staircase"), variable=stair_var,
                       fg="blue", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)

        # 底部按钮
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=15)
        tk.Button(btn_frame, text=self.T("btn_save"), command=self.save_and_return,
                  bg="#4caf50", fg="white", width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text=self.T("btn_cancel"), command=self.init_floating_mode,
                  width=15).pack(side=tk.RIGHT, padx=5)

    # --- 辅助方法 ---

    def _add_config_entry(self, parent, label, key, is_str=False):
        f = tk.Frame(parent)
        f.pack(fill=tk.X, padx=5, pady=5)
        tk.Label(f, text=label, width=20, anchor='w').pack(side=tk.LEFT)  # 增加Label宽度适配英文
        val = self.cfg.get(key)
        var = tk.StringVar(value=val) if is_str else tk.IntVar(value=val)
        self.vars[key] = var
        tk.Entry(f, textvariable=var).pack(side=tk.RIGHT, expand=True, fill=tk.X)

    def _add_check_btn(self, parent, label, key):
        var = tk.BooleanVar(value=self.cfg.get(key))
        self.vars[key] = var
        tk.Checkbutton(parent, text=label, variable=var).pack(anchor='w', padx=5)

    def save_and_return(self):
        try:
            for key, var in self.vars.items():
                self.cfg.set(key, var.get())
            self.cfg.save_config()
            self.setup_window()  # 应用窗口设置（如置顶、标题变化）
            self.update_hotkey()
            self.init_floating_mode()  # 返回主界面，主界面会重新读取语言配置
        except Exception as e:
            messagebox.showerror(self.T("btn_save"), f"{self.T('err_save')}: {e}")

    # --- 主逻辑控制 ---

    def on_ui_paste_click(self):
        if self.is_working:
            self.stop_sequence()
        else:
            text = self.engine.get_clipboard_text(self.root)
            if not text:
                self._flash_message(self.T("msg_empty"), 1000)
                return
            self.start_sequence(text, is_hotkey=False)

    def start_sequence(self, text, is_hotkey):
        self.is_working = True
        self.engine.stop_event.clear()

        if not is_hotkey:
            # UI 触发：显示倒计时
            self.btn_paste.config(bg="#d9d9d9", relief=tk.SUNKEN)
            self.btn_settings.config(state="disabled")
            delay_ms = self.cfg.get("btn_delay")
            threading.Thread(target=self._thread_countdown_and_paste, args=(text, delay_ms)).start()
        else:
            # 快捷键触发
            threading.Thread(target=self._thread_paste_only, args=(text,)).start()

    def stop_sequence(self):
        self.engine.abort()
        if hasattr(self, 'btn_paste') and self.btn_paste.winfo_exists():
            self.btn_paste.config(text=self.T("btn_stop"))

    def reset_ui_state(self):
        self.is_working = False
        self.engine.stop_event.clear()
        self.engine.target_hwnd = None
        if hasattr(self, 'btn_paste') and self.btn_paste.winfo_exists():
            self.btn_paste.config(text=self.T("btn_start"), state="normal", bg="#f0f0f0", relief=tk.RAISED)
            self.btn_settings.config(state="normal")

    def _flash_message(self, msg, duration):
        orig_text = self.btn_paste.cget("text")
        self.btn_paste.config(text=msg)
        self.root.after(duration, lambda: self.btn_paste.config(text=orig_text))

    def _thread_countdown_and_paste(self, text, delay_ms):
        steps = int(delay_ms / 100)
        for i in range(steps, 0, -1):
            if self.engine.stop_event.is_set():
                break
            remaining = f"{i / 10:.1f}s"
            # 动态更新按钮文本
            msg = f"{self.T('btn_wait')} ({remaining})"
            self.root.after(0, lambda m=msg: self.btn_paste.config(text=m))
            time.sleep(0.1)

        if not self.engine.stop_event.is_set():
            self.root.after(0, lambda: self.btn_paste.config(text=self.T("btn_typing")))
            self.engine.execute_paste(text, is_hotkey=False)

        self.root.after(0, self.reset_ui_state)

    def _thread_paste_only(self, text):
        self.engine.execute_paste(text, is_hotkey=True)
        self.is_working = False


if __name__ == "__main__":
    # 管理员权限检查
    try:
        is_admin = os.getuid() == 0
    except AttributeError:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0

    if not is_admin:
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(
            "Permission Denied / 权限不足",
            "Run as Administrator / 请以管理员身份运行"
        )
        root.destroy()
    else:
        root = tk.Tk()
        app = AppUI(root)


        def on_closing():
            app.engine.abort()
            os._exit(0)


        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()