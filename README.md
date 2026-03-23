import os
import sys
import subprocess
import webbrowser
import tkinter as tk
from tkinter import messagebox, ttk
import ctypes
import threading
import time

# Windows 11風スタイル (要: pip install pywinstyles)
try:
    import pywinstyles
    OS_STYLE_AVAILABLE = True
except ImportError:
    OS_STYLE_AVAILABLE = False

# --- アプリ情報 ---
APP_TITLE_BASE = "OneDrive Toolkit"
APP_VERSION = "26.3.21"
APP_COPYRIGHT = "©2026 HEXs"

# --- 多言語データ ---
LANG_DATA = {
    "ja": {
        "status": "状態:",
        "installed": "導入済み",
        "not_installed": "未導入",
        "btn_un": "削除する",
        "btn_in": "導入する",
        "menu_lang": "言語",
        "menu_help": "ヘルプ",
        "msg_confirm": "アンインストールしますか？",
        "progress_run": "処理中...",
        "progress_done": "完了",
        "usage": "ボタンを押すとOneDriveの状態を切り替えます。"
    },
    "en": {
        "status": "Status:",
        "installed": "Installed",
        "not_installed": "Missing",
        "btn_un": "Uninstall",
        "btn_in": "Install",
        "menu_lang": "Language",
        "menu_help": "Help",
        "msg_confirm": "Uninstall now?",
        "progress_run": "Processing...",
        "progress_done": "Done",
        "usage": "Click the button to manage OneDrive status."
    }
}

class OneDriveApp:
    def __init__(self, root):
        self.root = root
        self.current_lang = "ja"
        
        # ウィンドウ設定
        self.root.geometry("320x240")
        self.root.resizable(False, False)
        
        # アイコン設定 (EXE埋め込み対応)
        self.set_app_icon()

        # Windows 11 スタイル適用
        if OS_STYLE_AVAILABLE:
            pywinstyles.apply_style(self.root, "mica")

        self.main_frame = ttk.Frame(root, padding=15)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.create_menu()
        self.create_widgets()
        self.refresh_status()

    def set_app_icon(self):
        """ウィンドウ左上のアイコンを設定"""
        try:
            if hasattr(sys, '_MEIPASS'):
                # EXE化された際の一時展開先パス
                icon_path = os.path.join(sys._MEIPASS, "icon.png")
            else:
                # スクリプト実行時のパス
                icon_path = "icon.png"
            
            if os.path.exists(icon_path):
                self.icon_img = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(False, self.icon_img)
        except Exception as e:
            print(f"Icon Load Error: {e}")

    def create_menu(self):
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)
        
        self.lang_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(menu=self.lang_menu)
        self.lang_menu.add_command(label="日本語", command=lambda: self.change_language("ja"))
        self.lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
        
        self.help_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(menu=self.help_menu)
        self.help_menu.add_command(label="Help / Version", command=self.show_about)

    def create_widgets(self):
        self.lbl_status = ttk.Label(self.main_frame, font=("Segoe UI", 9))
        self.lbl_status.pack(pady=(0, 5))
        
        self.lbl_main = ttk.Label(self.main_frame, font=("Segoe UI Bold", 11))
        self.lbl_main.pack(pady=(0, 15))

        self.btn_action = ttk.Button(self.main_frame, command=self.on_click)
        self.btn_action.pack(fill=tk.X, ipady=3)

        self.lbl_prog = ttk.Label(self.main_frame, font=("Segoe UI", 8), foreground="gray")
        self.lbl_prog.pack(pady=(15, 0), anchor=tk.W)
        
        self.pbar = ttk.Progressbar(self.main_frame, mode='determinate', length=100)
        self.pbar.pack(fill=tk.X, pady=(2, 10))

        ttk.Label(self.main_frame, text=f"{APP_COPYRIGHT}", font=("Segoe UI", 7), foreground="#999").pack(side=tk.BOTTOM)

    def get_path(self):
        paths = [
            os.path.expandvars(r"%ProgramFiles%\Microsoft OneDrive\OneDrive.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft OneDrive\OneDrive.exe"),
            os.path.expandvars(r"%LocalAppData%\Microsoft\OneDrive\OneDrive.exe")
        ]
        return next((p for p in paths if os.path.exists(p)), None)

    def update_ui(self):
        tr = LANG_DATA[self.current_lang]
        self.root.title(f"OD Toolkit v{APP_VERSION}")
        self.menubar.entryconfig(1, label=tr["menu_lang"])
        self.menubar.entryconfig(2, label=tr["menu_help"])
        self.lbl_status.config(text=tr["status"])
        
        if self.get_path():
            self.lbl_main.config(text=tr["installed"], foreground="#0078D4")
            self.btn_action.config(text=tr["btn_un"])
        else:
            self.lbl_main.config(text=tr["not_installed"], foreground="#777")
            self.btn_action.config(text=tr["btn_in"])

    def refresh_status(self):
        self.update_ui()

    def change_language(self, lang):
        self.current_lang = lang
        self.update_ui()

    def on_click(self):
        tr = LANG_DATA[self.current_lang]
        if self.get_path():
            if messagebox.askyesno("Confirm", tr["msg_confirm"]):
                threading.Thread(target=self.exec_un, daemon=True).start()
        else:
            webbrowser.open("https://www.microsoft.com/ja-jp/microsoft-365/onedrive/download")

    def exec_un(self):
        tr = LANG_DATA[self.current_lang]
        self.btn_action.config(state="disabled")
        self.lbl_prog.config(text=tr["progress_run"])
        
        # プロセス終了
        os.system("taskkill /f /im OneDrive.exe >nul 2>&1")
        self.pbar['value'] = 30
        
        # アンインストーラー特定
        un_path = os.path.expandvars(r"%SystemRoot%\SysWOW64\OneDriveSetup.exe")
        if not os.path.exists(un_path):
            un_path = os.path.expandvars(r"%SystemRoot%\System32\OneDriveSetup.exe")
            
        if os.path.exists(un_path):
            try:
                subprocess.run([un_path, "/uninstall"], shell=True)
                for i in range(31, 101, 10):
                    self.pbar['value'] = i
                    time.sleep(0.5)
            except Exception as e:
                messagebox.showerror("Error", str(e))
        
        self.lbl_prog.config(text=tr["progress_done"])
        time.sleep(1)
        self.pbar['value'] = 0
        self.lbl_prog.config(text="")
        self.btn_action.config(state="normal")
        self.root.after(0, self.refresh_status)

    def show_about(self):
        tr = LANG_DATA[self.current_lang]
        msg = f"{APP_TITLE_BASE}\nVer {APP_VERSION}\n{APP_COPYRIGHT}\n\n{tr['usage']}"
        messagebox.showinfo("About", msg)

if __name__ == "__main__":
    if not ctypes.windll.shell32.IsUserAnAdmin():
        # 管理者権限へ昇格
        script_path = os.path.abspath(sys.argv[0])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script_path}"', None, 1)
    else:
        root = tk.Tk()
        app = OneDriveApp(root)
        root.mainloop()
