import os
import sys
import subprocess
import webbrowser
import tkinter as tk
from tkinter import messagebox, ttk
import ctypes
import threading
import time
import shutil

# Windows 11風スタイル (要: pip install pywinstyles)
try:
    import pywinstyles
    OS_STYLE_AVAILABLE = True
except ImportError:
    OS_STYLE_AVAILABLE = False

# --- アプリ情報 ---
APP_TITLE_BASE = "OneDrive Toolkit"
APP_VERSION = "26.4.28" # アップデート
APP_COPYRIGHT = "©2026 HEXs"

# --- 多言語データ ---
LANG_DATA = {
    "ja": {
        "status": "状態:",
        "installed": "導入済み",
        "not_installed": "未導入 (クリーン)",
        "btn_un": "完全に削除する",
        "btn_in": "公式サイトから導入",
        "menu_lang": "言語",
        "menu_help": "ヘルプ",
        "msg_confirm": "OneDriveをシステムから完全に削除しますか？\n(レジストリや残存ファイルも対象です)",
        "progress_run": "クリーンアップ中...",
        "progress_done": "削除完了",
        "usage": "OneDriveを根本から削除、または再導入サイトを開きます。"
    },
    "en": {
        "status": "Status:",
        "installed": "Installed",
        "not_installed": "Not Installed",
        "btn_un": "Deep Uninstall",
        "btn_in": "Install (Web)",
        "menu_lang": "Language",
        "menu_help": "Help",
        "msg_confirm": "Deep uninstall OneDrive?\n(Includes registry and cache cleanup)",
        "progress_run": "Cleaning up...",
        "progress_done": "Finished",
        "usage": "Completely remove OneDrive or access the official install page."
    }
}

class OneDriveApp:
    def __init__(self, root):
        self.root = root
        self.current_lang = "ja"
        
        self.root.geometry("340x260")
        self.root.resizable(False, False)
        
        self.set_app_icon()

        if OS_STYLE_AVAILABLE:
            pywinstyles.apply_style(self.root, "mica")

        self.main_frame = ttk.Frame(root, padding=15)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.create_menu()
        self.create_widgets()
        self.refresh_status()

    def set_app_icon(self):
        try:
            icon_path = os.path.join(sys._MEIPASS, "icon.png") if hasattr(sys, '_MEIPASS') else "icon.png"
            if os.path.exists(icon_path):
                self.icon_img = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(False, self.icon_img)
        except Exception: pass

    def create_menu(self):
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)
        self.lang_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Language", menu=self.lang_menu)
        self.lang_menu.add_command(label="日本語", command=lambda: self.change_language("ja"))
        self.lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
        self.help_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Help", menu=self.help_menu)
        self.help_menu.add_command(label="About", command=self.show_about)

    def create_widgets(self):
        self.lbl_status = ttk.Label(self.main_frame, font=("Segoe UI", 9))
        self.lbl_status.pack(pady=(0, 5))
        
        self.lbl_main = ttk.Label(self.main_frame, font=("Segoe UI Bold", 11))
        self.lbl_main.pack(pady=(0, 15))

        self.btn_action = ttk.Button(self.main_frame, command=self.on_click)
        self.btn_action.pack(fill=tk.X, ipady=5)

        self.lbl_prog = ttk.Label(self.main_frame, font=("Segoe UI", 8), foreground="gray")
        self.lbl_prog.pack(pady=(15, 0), anchor=tk.W)
        
        self.pbar = ttk.Progressbar(self.main_frame, mode='determinate', length=100)
        self.pbar.pack(fill=tk.X, pady=(2, 10))

        ttk.Label(self.main_frame, text=f"{APP_COPYRIGHT}", font=("Segoe UI", 7), foreground="#999").pack(side=tk.BOTTOM)

    def get_path(self):
        """実行ファイルの存在確認"""
        paths = [
            os.path.expandvars(r"%ProgramFiles%\Microsoft OneDrive\OneDrive.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft OneDrive\OneDrive.exe"),
            os.path.expandvars(r"%LocalAppData%\Microsoft\OneDrive\OneDrive.exe")
        ]
        return next((p for p in paths if os.path.exists(p)), None)

    def update_ui(self):
        tr = LANG_DATA[self.current_lang]
        self.root.title(f"{APP_TITLE_BASE} v{APP_VERSION}")
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
                threading.Thread(target=self.deep_uninstall, daemon=True).start()
        else:
            # 再インストール用URL (直リンクではなくダウンロードトップ)
            webbrowser.open("https://www.microsoft.com/ja-jp/microsoft-365/onedrive/download")

    def deep_uninstall(self):
        """徹底的な削除処理"""
        tr = LANG_DATA[self.current_lang]
        self.btn_action.config(state="disabled")
        self.lbl_prog.config(text=tr["progress_run"])
        
        # 1. プロセス強制終了
        subprocess.run("taskkill /f /im OneDrive.exe", shell=True, capture_output=True)
        self.pbar['value'] = 20
        
        # 2. 標準アンインストーラー実行
        un_paths = [
            os.path.expandvars(r"%SystemRoot%\SysWOW64\OneDriveSetup.exe"),
            os.path.expandvars(r"%SystemRoot%\System32\OneDriveSetup.exe")
        ]
        for p in un_paths:
            if os.path.exists(p):
                subprocess.run([p, "/uninstall"], shell=True)
                break
        self.pbar['value'] = 50
        time.sleep(2) # アンインストールの完了を少し待つ

        # 3. 残存フォルダの削除
        folders = [
            os.path.expandvars(r"%LocalAppData%\Microsoft\OneDrive"),
            os.path.expandvars(r"%ProgramData%\Microsoft OneDrive"),
            os.path.expandvars(r"%UserProfile%\OneDrive") # 空の場合のみ推奨
        ]
        for folder in folders:
            if os.path.exists(folder):
                try:
                    shutil.rmtree(folder, ignore_errors=True)
                except: pass
        self.pbar['value'] = 75

        # 4. レジストリ（サイドバー項目）の削除
        # エクスプローラーのサイドバーからOneDriveを消すフラグ
        reg_cmd = 'REG ADD "HKCU\\Software\\Classes\\CLSID\\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /v "System.IsPinnedToNameSpaceTree" /t REG_DWORD /d 0 /f'
        subprocess.run(reg_cmd, shell=True, capture_output=True)
        
        self.pbar['value'] = 100
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
        script_path = os.path.abspath(sys.argv[0])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script_path}"', None, 1)
    else:
        root = tk.Tk()
        app = OneDriveApp(root)
        root.mainloop()