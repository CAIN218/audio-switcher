import os
import subprocess
import keyboard
import tkinter as tk
from tkinter import messagebox
import time
import sys
import warnings
import threading
import queue
import ctypes
import json
import urllib.request
import webbrowser
import winshell
from win32com.client import Dispatch

import comtypes
from pycaw.pycaw import AudioUtilities
from pycaw.constants import DEVICE_STATE, EDataFlow

# customtkinter（設定画面専用のUIライブラリ）はここでは読み込まない。
# バックグラウンド常駐時はホットキー監視しかしないため、
# 使わない重いUIアセットをメモリに乗せないよう open_setting_window() 内で遅延importする。

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "app_config.txt")
ICON_PATH = os.path.join(BASE_DIR, "app_icon.ico")

__version__ = "1.1.0"
GITHUB_REPO = "CAIN218/audio-switcher"

DEFAULT_OVERLAY_HOTKEY = "ctrl+alt+s"

last_switch_time = 0
_com_thread_local = threading.local()

# ----------------------------------------------------------------------
# デザイントークン — アナログ・ミキシングコンソール／パッチベイをモチーフにした
# グラファイト×カッパー（銅）配色。設定画面とトースト通知の両方で共有する。
# ----------------------------------------------------------------------
COLOR_BG          = "#16151a"  # アプリ全体の背景（暖色寄りの黒鉛色）
COLOR_PANEL       = "#1d1b21"  # カードパネル
COLOR_INPUT       = "#26232b"  # 入力欄・コンボボックス
COLOR_BORDER      = "#322e38"  # 通常の境界線
COLOR_ACCENT      = "#d68a4c"  # カッパー（VUメーターの針をイメージしたアクセント）
COLOR_ACCENT_HOVER = "#e6a05f"
COLOR_ACCENT_DIM  = "#5c4530"  # アクセントの控えめ版（非アクティブ表示用）
COLOR_TEXT_MAIN   = "#f2ece1"  # 暖色寄りのオフホワイト
COLOR_TEXT_SUB    = "#8f8878"  # 落ち着いたタウプグレー
COLOR_TEXT_FAINT  = "#5c5850"

FONT_DISPLAY = ("Segoe UI Semibold", 17)
FONT_LABEL   = ("Segoe UI", 10, "bold")
FONT_BODY    = ("Segoe UI", 10)
FONT_MONO    = ("Consolas", 11, "bold")


def draw_routing_indicator(canvas, active_side):
    """A/Bどちらのレイヤーが現在アクティブかをパッチベイ風に可視化するサイン要素"""
    canvas.delete("all")
    w = int(canvas["width"])
    h = int(canvas["height"])
    cy = h // 2
    ax, bx = 30, w - 30
    r = 10

    canvas.create_line(ax, cy, bx, cy, fill=COLOR_BORDER, width=2)

    for label, x in (("A", ax), ("B", bx)):
        is_active = (active_side == label)
        if is_active:
            canvas.create_oval(x - r - 5, cy - r - 5, x + r + 5, cy + r + 5,
                                outline=COLOR_ACCENT_DIM, width=2)
        fill = COLOR_ACCENT if is_active else COLOR_PANEL
        outline = COLOR_ACCENT if is_active else COLOR_TEXT_FAINT
        canvas.create_oval(x - r, cy - r, x + r, cy + r, fill=fill, outline=outline, width=2)
        text_color = COLOR_BG if is_active else COLOR_TEXT_SUB
        canvas.create_text(x, cy, text=label, fill=text_color, font=("Segoe UI", 9, "bold"))


def ensure_com_initialized():
    """pycaw(comtypes)はスレッド毎にCOM初期化が必要。二重初期化は無害に握りつぶす"""
    if not getattr(_com_thread_local, "initialized", False):
        try:
            comtypes.CoInitialize()
        except OSError:
            pass  # 既に初期化済みのスレッド
        _com_thread_local.initialized = True

def kill_previous_instances():
    """裏で既に動いている古いAudioSwitcherのプロセスをすべて安全に終了させる"""
    try:
        current_pid = os.getpid()
        target_process = "AudioSwitcher.exe"
        if getattr(sys, 'frozen', False):
            subprocess.run(
                f'taskkill /F /IM {target_process} /FI "PID ne {current_pid}"',
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
    except Exception:
        pass

def get_device_info():
    """pycaw(Windows Core Audio API)経由でデバイスリストと現在のデフォルトを取得（外部exe不要）"""
    try:
        ensure_com_initialized()

        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            devices = AudioUtilities.GetAllDevices(
                data_flow=EDataFlow.eRender.value,
                device_state=DEVICE_STATE.ACTIVE.value,
            )
            default_device = AudioUtilities.GetSpeakers()

        default_id = getattr(default_device, "id", None)

        device_dict = {}
        current_default_name = None

        for dev in devices:
            display_name = (dev.FriendlyName or dev.id or "").strip()
            if not display_name:
                continue

            # 同名デバイスが複数ある場合はIDの末尾を付けて衝突を回避
            if display_name in device_dict:
                display_name = f"{display_name} ({dev.id[-8:]})"

            device_dict[display_name] = dev.id

            if default_id and dev.id == default_id:
                current_default_name = display_name

        return device_dict, current_default_name
    except Exception as e:
        log_error("get_device_info", e)
        return {}, None

def load_config():
    """設定ファイルの読み込み。5,6行目(複数デバイス切り替えオーバーレイの
    ホットキーと有効/無効)は後から追加した項目なので、無い場合は
    デフォルト値にフォールバックし、古い4行だけの設定ファイルもそのまま
    読み込めるようにしている。"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
        dev_a = lines[0] if len(lines) > 0 else ""
        dev_b = lines[1] if len(lines) > 1 else ""
        key = lines[2] if len(lines) > 2 else "ctrl+alt+a"
        startup_val = lines[3] if len(lines) > 3 else "0"
        overlay_key = lines[4] if len(lines) > 4 else DEFAULT_OVERLAY_HOTKEY
        overlay_enabled = lines[5] if len(lines) > 5 else "0"
        return dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled
    return "", "", "ctrl+alt+a", "0", DEFAULT_OVERLAY_HOTKEY, "0"

def save_config(dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled):
    """設定ファイルの保存とスタートアップ制御"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(f"{dev_a}\n{dev_b}\n{key}\n{startup_val}\n{overlay_key}\n{overlay_enabled}")
    manage_startup(startup_val)

def manage_startup(startup_val):
    """Windowsのスタートアップフォルダへの自動登録・解除（生スクリプト/exeハイブリッド対応・エラー修正版）"""
    try:
        startup_dir = winshell.startup()
        shortcut_path = os.path.join(startup_dir, "AudioSwitcher.lnk")
        
        if startup_val == "1":
            if getattr(sys, 'frozen', False):
                target = sys.executable
                arguments = "/background"
                icon_target = target  # exe化されている時は自身のexeアイコンを使う
            else:
                # コンソール(黒い画面)が出ないよう、可能ならpythonw.exeを使う
                pythonw_path = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
                target = pythonw_path if os.path.exists(pythonw_path) else sys.executable
                arguments = f'"{os.path.abspath(__file__)}" /background'
                icon_target = None    # 生スクリプトの時はアイコン設定をパスする（エラー回避）
                
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = target
            shortcut.Arguments = arguments
            shortcut.WorkingDirectory = BASE_DIR
            
            # 生スクリプト実行時（icon_targetがNone）は設定をスキップしてクラッシュを防ぐ
            if icon_target:
                shortcut.IconLocation = icon_target
                
            shortcut.save()
        else:
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
    except Exception as e:
        tk.messagebox.showerror("Startup Error", f"Failed to update startup folder: {str(e)}")

class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class _MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", _RECT),
                ("rcWork", _RECT), ("dwFlags", ctypes.c_ulong)]


def _active_monitor_rect(use_work_area):
    """マウスカーソルがあるモニターの領域を返す。マルチモニター環境では
    プライマリモニター固定だと、実際に見ている画面とは違う場所に
    トースト/オーバーレイが出てしまうため、常にカーソルがある＝いま
    使っているモニターを基準にする。use_work_area=Trueならタスクバーを
    除いた作業領域、Falseならモニター全体(オーバーレイの背景を暗くする
    時は、タスクバーの領域も含めて画面全体を覆いたいのでこちらを使う)。"""
    pt = _POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    MONITOR_DEFAULTTONEAREST = 2
    hmonitor = ctypes.windll.user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
    info = _MONITORINFO()
    info.cbSize = ctypes.sizeof(_MONITORINFO)
    ctypes.windll.user32.GetMonitorInfoW(hmonitor, ctypes.byref(info))
    rc = info.rcWork if use_work_area else info.rcMonitor
    return rc.left, rc.top, rc.right, rc.bottom


def _active_monitor_work_area():
    return _active_monitor_rect(True)


def _active_monitor_full_rect():
    return _active_monitor_rect(False)


def _rounded_rect(cv, x1, y1, x2, y2, r, **kwargs):
    points = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
        x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return cv.create_polygon(points, smooth=True, **kwargs)


# 切り替え毎にトースト用のTk()を新規生成すると、Tclインタプリタが使い捨てられる
# たびにメモリが積み上がっていく（destroy()してもTcl側のオブジェクトプールは
# 完全には解放されない）。常駐中は1枚のToplevelを使い回し、フェード管理だけ
# やり直すことでインタプリタの再生成を避ける。
class _ToastManager:
    WIDTH, HEIGHT = 320, 68

    def __init__(self, root):
        self.root = root
        self.window = None
        self.canvas = None
        self.fade_job = None
        self.current_alpha = 0.0
        self.on_click = None
        self._hold_ms = 2600

    def _ensure_window(self):
        if self.window is not None:
            return
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.attributes("-alpha", 0.0)

        transparent_key = "#010203"  # 実際のUIでは使わない色をウィンドウの透過キーにする
        toast.configure(bg=transparent_key)
        toast.attributes("-transparentcolor", transparent_key)

        canvas = tk.Canvas(toast, width=self.WIDTH, height=self.HEIGHT,
                            bg=transparent_key, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        # クリック可能なトースト(アップデート通知など)用。show()側で都度差し替える。
        canvas.bind("<Button-1>", lambda e: self.on_click() if self.on_click else None)
        toast.withdraw()

        self.window = toast
        self.canvas = canvas

    def _place_bottom_right(self):
        # withdraw()後にdeiconify()すると、overrideredirectウィンドウの位置が
        # Windows側でリセットされることがあるため、表示するたびに座標を
        # 再設定して確実に固定する。また常にプライマリモニターではなく、
        # 現在マウスカーソルがあるモニターの作業領域（タスクバー除く）の
        # 右下を使うことで、マルチモニター環境でも実際に見ている画面に出す。
        try:
            left, top, right, bottom = _active_monitor_work_area()
        except Exception:
            left, top = 0, 0
            right = self.window.winfo_screenwidth()
            bottom = self.window.winfo_screenheight()
        x = right - self.WIDTH - 6
        y = bottom - self.HEIGHT - 6
        self.window.geometry(f"{self.WIDTH}x{self.HEIGHT}+{x}+{y}")

    def show(self, device_name):
        self._ensure_window()
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None

        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 16,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        # 左端のカッパーのアクセントバー
        canvas.create_rectangle(0, 12, 4, self.HEIGHT - 12, fill=COLOR_ACCENT, outline="")
        canvas.create_text(22, 22, anchor="w", text="AUDIO ROUTING SWITCHED",
                            fill=COLOR_ACCENT, font=("Segoe UI", 8, "bold"))
        display_text = device_name if len(device_name) <= 34 else device_name[:32] + "..."
        canvas.create_text(22, 44, anchor="w", text=display_text,
                            fill=COLOR_TEXT_MAIN, font=("Segoe UI", 10, "bold"))
        canvas.configure(cursor="")

        self.on_click = None
        self._hold_ms = 2600
        self._place_bottom_right()
        self.window.deiconify()
        self._place_bottom_right()
        self.current_alpha = 0.0
        self._fade_in()

    def show_update_available(self, latest_tag, release_url):
        """新しいバージョンが見つかった時の控えめな通知。クリックでリリース
        ページを開く。頻繁に出るものではないので、通常のトーストより長めに
        表示したままにする。"""
        self._ensure_window()
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None

        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 16,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        canvas.create_rectangle(0, 12, 4, self.HEIGHT - 12, fill=COLOR_ACCENT, outline="")
        canvas.create_text(22, 22, anchor="w", text="UPDATE AVAILABLE",
                            fill=COLOR_ACCENT, font=("Segoe UI", 8, "bold"))
        canvas.create_text(22, 44, anchor="w", text=f"{latest_tag} — click to view",
                            fill=COLOR_TEXT_MAIN, font=("Segoe UI", 10, "bold"))
        canvas.configure(cursor="hand2")

        self.on_click = lambda: webbrowser.open(release_url)
        self._hold_ms = 8000
        self._place_bottom_right()
        self.window.deiconify()
        self._place_bottom_right()
        self.current_alpha = 0.0
        self._fade_in()

    def _fade_in(self):
        if self.current_alpha < 0.96:
            self.current_alpha += 0.12
            self.window.attributes("-alpha", min(self.current_alpha, 0.96))
            self.fade_job = self.root.after(15, self._fade_in)
        else:
            self.fade_job = self.root.after(self._hold_ms, self._fade_out)

    def _fade_out(self):
        if self.current_alpha > 0.0:
            self.current_alpha -= 0.06
            self.window.attributes("-alpha", max(self.current_alpha, 0.0))
            self.fade_job = self.root.after(15, self._fade_out)
        else:
            self.window.withdraw()
            self.fade_job = None


# ----------------------------------------------------------------------
# 3台以上のデバイスから選ぶための全画面オーバーレイ。
#
# Tkinterの-alphaはウィンドウ単位でしか効かず、1枚のウィンドウの中で
# 「背景だけ半透明・中身だけ不透明」ということはできない。そのため、
# 画面全体を覆って暗くするだけの層(dim)と、実際の選択パネルを描く
# 不透明な層(card)を、別々のToplevelとして重ねて表示している。
# ----------------------------------------------------------------------
class _DeviceOverlay:
    WIDTH = 360
    ROW_HEIGHT = 46
    HEADER_H = 44
    PADDING = 16
    DIM_ALPHA = 0.55

    def __init__(self, root):
        self.root = root
        self.dim_window = None
        self.card_window = None
        self.canvas = None
        self.names = []
        self.device_dict = {}
        self.current_default = None
        self.selected_index = 0
        self.height = 0
        self.fade_job = None
        self.current_alpha = 0.0
        self.on_select = None  # callback(display_name, target_id)

    def is_open(self):
        return self.card_window is not None

    def open(self, device_dict, current_default):
        self.close()

        names = list(device_dict.keys())
        if not names:
            return

        self.names = names
        self.device_dict = device_dict
        self.current_default = current_default
        self.selected_index = names.index(current_default) if current_default in names else 0
        self.height = self.HEADER_H + len(names) * self.ROW_HEIGHT + self.PADDING

        try:
            left, top, right, bottom = _active_monitor_full_rect()
        except Exception:
            left, top = 0, 0
            right = self.root.winfo_screenwidth()
            bottom = self.root.winfo_screenheight()

        dim = tk.Toplevel(self.root)
        dim.overrideredirect(True)
        dim.attributes("-topmost", True)
        dim.attributes("-alpha", 0.0)
        dim.configure(bg="#000000")
        dim.geometry(f"{right - left}x{bottom - top}+{left}+{top}")
        dim.bind("<Button-1>", lambda e: self.close())
        dim.bind("<Escape>", lambda e: self.close())

        card = tk.Toplevel(self.root)
        card.overrideredirect(True)
        card.attributes("-topmost", True)
        card.attributes("-alpha", 0.0)
        card.configure(bg=COLOR_BG)

        canvas = tk.Canvas(card, width=self.WIDTH, height=self.height, bg=COLOR_BG, highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        cx = left + (right - left - self.WIDTH) // 2
        cy = top + (bottom - top - self.height) // 2
        card.geometry(f"{self.WIDTH}x{self.height}+{cx}+{cy}")

        self.dim_window = dim
        self.card_window = card
        self.canvas = canvas

        card.bind("<Up>", lambda e: self._move(-1))
        card.bind("<Down>", lambda e: self._move(1))
        card.bind("<Return>", lambda e: self._confirm())
        card.bind("<Escape>", lambda e: self.close())
        for i in range(min(9, len(names))):
            card.bind(str(i + 1), lambda e, idx=i: self._confirm(idx))
        card.bind("<FocusOut>", lambda e: self.root.after(150, self._close_if_unfocused))

        self._redraw()

        dim.lift()
        card.lift()
        card.focus_force()

        self.current_alpha = 0.0
        self._fade_in()

    def _redraw(self):
        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.height - 1, 16,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        canvas.create_text(self.PADDING, self.HEADER_H / 2, anchor="w",
                            text="SELECT OUTPUT DEVICE", fill=COLOR_TEXT_SUB,
                            font=("Segoe UI", 9, "bold"))
        canvas.create_line(self.PADDING * 0.5, self.HEADER_H, self.WIDTH - self.PADDING * 0.5, self.HEADER_H,
                            fill=COLOR_BORDER)

        for i, name in enumerate(self.names):
            y0 = self.HEADER_H + i * self.ROW_HEIGHT
            y1 = y0 + self.ROW_HEIGHT
            is_selected = (i == self.selected_index)
            is_current = (name == self.current_default)

            row_bg = _rounded_rect(
                canvas, self.PADDING * 0.5, y0 + 3, self.WIDTH - self.PADDING * 0.5, y1 - 3, 10,
                fill=(COLOR_ACCENT_DIM if is_selected else COLOR_PANEL), outline=""
            )

            display = name if len(name) <= 32 else name[:30] + "..."
            label = f"● {display}" if is_current else display
            text_color = COLOR_ACCENT if is_selected else (COLOR_TEXT_MAIN if is_current else COLOR_TEXT_SUB)
            text_id = canvas.create_text(
                self.PADDING * 1.2, (y0 + y1) / 2, anchor="w", text=label,
                fill=text_color, font=("Segoe UI", 10, "bold" if is_selected else "normal")
            )

            items = [row_bg, text_id]
            if i < 9:
                items.append(canvas.create_text(
                    self.WIDTH - self.PADDING, (y0 + y1) / 2, anchor="e",
                    text=str(i + 1), fill=COLOR_TEXT_FAINT, font=("Segoe UI", 9)
                ))

            for item in items:
                canvas.tag_bind(item, "<Button-1>", lambda e, idx=i: self._confirm(idx))
                canvas.tag_bind(item, "<Enter>", lambda e, idx=i: self._hover(idx))

    def _hover(self, idx):
        if idx == self.selected_index:
            return
        self.selected_index = idx
        self._redraw()

    def _move(self, delta):
        if not self.names:
            return
        self.selected_index = (self.selected_index + delta) % len(self.names)
        self._redraw()

    def _confirm(self, idx=None):
        if idx is not None:
            self.selected_index = idx
        if not (0 <= self.selected_index < len(self.names)):
            self.close()
            return
        name = self.names[self.selected_index]
        target_id = self.device_dict.get(name)
        callback = self.on_select
        self.close()
        if callback and target_id:
            callback(name, target_id)

    def _close_if_unfocused(self):
        if self.card_window is None:
            return
        try:
            focused = self.card_window.focus_get()
        except Exception:
            focused = None
        if focused is None:
            self.close()

    def _fade_in(self):
        if self.card_window is None:
            return
        if self.current_alpha < 0.98:
            self.current_alpha += 0.14
            alpha = min(self.current_alpha, 0.98)
            try:
                self.dim_window.attributes("-alpha", min(alpha, self.DIM_ALPHA))
                self.card_window.attributes("-alpha", alpha)
            except Exception:
                return
            self.fade_job = self.root.after(12, self._fade_in)
        else:
            try:
                self.dim_window.attributes("-alpha", self.DIM_ALPHA)
                self.card_window.attributes("-alpha", 1.0)
            except Exception:
                pass

    def close(self):
        if self.fade_job:
            try:
                self.root.after_cancel(self.fade_job)
            except Exception:
                pass
            self.fade_job = None
        for w in (self.card_window, self.dim_window):
            if w is not None:
                try:
                    w.destroy()
                except Exception:
                    pass
        self.card_window = None
        self.dim_window = None
        self.canvas = None


_toast_queue = queue.Queue()

def log_error(context, exc):
    """サイレントに握りつぶさず、原因調査用にエラーをログファイルへ記録する"""
    try:
        log_path = os.path.join(BASE_DIR, "error.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {context}: {repr(exc)}\n")
            if exc is not None and exc.__traceback__ is not None:
                import traceback
                f.write("".join(traceback.format_tb(exc.__traceback__)))
    except Exception:
        pass


def _log_thread_exception(args):
    """個々のスレッド内で例外が握りつぶされず消えるのを防ぐ（console=Falseのため標準出力は捨てられる）"""
    log_error(f"thread:{args.thread.name if args.thread else '?'}", args.exc_value)


threading.excepthook = _log_thread_exception


def _parse_version(tag):
    """'v1.2.3' -> (1, 2, 3)。数値以外が混ざっていても比較不能で落ちない
    ように、パースできない要素は0として扱う。"""
    parts = []
    for p in tag.lstrip("vV").split("."):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


_update_queue = queue.Queue()

def _update_check_loop():
    """1日1回、GitHubの最新リリースを確認する。バックグラウンドの常駐
    ツールなので、失敗(オフライン等)は完全に握りつぶし通常動作に一切
    影響させない。新しいバージョンが見つかった時だけ、トースト表示用の
    キューに積んでメインスレッド側(poll_toast_queue)に伝える。"""
    while True:
        try:
            req = urllib.request.Request(
                f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
                headers={"Accept": "application/vnd.github+json", "User-Agent": "AudioSwitcher-update-check"},
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            latest_tag = data.get("tag_name", "")
            release_url = data.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases/latest"
            if latest_tag and _parse_version(latest_tag) > _parse_version(__version__):
                _update_queue.put((latest_tag, release_url))
        except Exception as e:
            log_error("update_check", e)
        time.sleep(24 * 60 * 60)


_switch_request_queue = queue.Queue()

def switch_audio():
    """ホットキーのコールバック本体。低レベルキーボードフックのコールバック
    スレッドから直接呼ばれるため、ここでは絶対にブロックしてはいけない。
    Windowsは低レベルフックの応答が遅い(既定で数秒程度)と、そのフックを
    黙って無効化することがあり、以前はここでCOM/オーディオAPIを直接呼んで
    いたため、まれに発生する遅延がきっかけでホットキーが無反応になる不具合
    があった。実際の処理は全て別スレッド(_switch_worker)に投げ、ここでは
    デバウンス判定とキューへの投入だけを行う。"""
    global last_switch_time
    current_time = time.time()

    if current_time - last_switch_time < 0.6:
        return
    last_switch_time = current_time

    _switch_request_queue.put(None)

def _switch_worker():
    """switch_audio()から渡された切り替え要求を、フックコールバックの外側
    (専用スレッド)で実際に処理する。COM初期化やSetDefaultDeviceがここで
    多少時間を要しても、キーボードフックの応答性には影響しない。"""
    while True:
        _switch_request_queue.get()
        try:
            dev_a, dev_b, _, _, _, _ = load_config()
            device_dict, current_default = get_device_info()

            if not current_default:
                target_display_name = dev_a
            else:
                if current_default == dev_a:
                    target_display_name = dev_b
                else:
                    target_display_name = dev_a

            target_id = device_dict.get(target_display_name)
            if target_id:
                try:
                    ensure_com_initialized()
                    # roles=NoneでeConsole/eMultimedia/eCommunicationsの全ロールに設定（SoundVolumeViewの"all"相当）
                    AudioUtilities.SetDefaultDevice(target_id)
                except Exception as e:
                    log_error("SetDefaultDevice", e)
                    continue
                _toast_queue.put(target_display_name)
        except Exception as e:
            log_error("switch_worker", e)


_direct_switch_queue = queue.Queue()

def _direct_switch_worker():
    """複数デバイス切り替えオーバーレイで選んだ、任意の特定デバイスへの
    切り替えを処理する。A/Bトグル用の_switch_workerとは別経路にしている
    (こちらは対象デバイスを都度指定できる必要があるため)。"""
    while True:
        target_id, target_display_name = _direct_switch_queue.get()
        try:
            ensure_com_initialized()
            AudioUtilities.SetDefaultDevice(target_id)
            _toast_queue.put(target_display_name)
        except Exception as e:
            log_error("direct_switch_worker", e)


last_overlay_time = 0
_overlay_open_queue = queue.Queue()

def open_device_overlay():
    """複数デバイス切り替えオーバーレイ用ホットキーのコールバック。
    switch_audio()と同じ理由で、ここでは絶対にブロックしない。実際に
    オーバーレイを開く処理はメインスレッド側(poll_toast_queue)で行う。"""
    global last_overlay_time
    current_time = time.time()

    if current_time - last_overlay_time < 0.6:
        return
    last_overlay_time = current_time

    _overlay_open_queue.put(None)


def _rehook_watchdog(hotkeys, interval_sec=30):
    """UACの同意プロンプトやロック画面(Secure Desktop)に切り替わっている間、
    このプロセスの低レベルキーボードフックには何のイベントも届かない。
    そのタイミングでたまたまctrl/altを押していると、対応するKEY_UPイベントが
    永遠に届かず、keyboardライブラリ内部の「押されているキー」の記録
    (_pressed_events等)にゴースト状態が残り続ける。以降そのホットキーの
    組み合わせが二度と成立しなくなる(＝プロセスは動いているのにショート
    カットだけ無反応になる)のはこれが原因。
    unhook_all()はこの内部状態(_pressed_events等)には触れないため無意味で、
    ここで明示的にクリアしてゴーストを解消する。
    hotkeys: [(キーの組み合わせ, コールバック), ...] を全て張り直す。"""
    while True:
        time.sleep(interval_sec)
        try:
            with keyboard._pressed_events_lock:
                keyboard._pressed_events.clear()
            keyboard._logically_pressed_keys.clear()
            keyboard._listener.active_modifiers.clear()
            keyboard._listener.modifier_states.clear()

            keyboard.unhook_all()
            for key_combo, callback in hotkeys:
                keyboard.add_hotkey(key_combo, callback, trigger_on_release=False)
        except Exception as e:
            log_error("rehook_watchdog", e)


def start_backend():
    """バックグラウンドでのキーボード監視 + 常駐トースト用Tkループの開始
    Tclインタプリタを1つだけ常駐させ、切り替えのたびに使い回すことで
    インタプリタの生成・破棄によるメモリ増加を防ぐ。"""
    _, _, shortcut_key, _, overlay_key, overlay_enabled = load_config()

    hotkeys = [(shortcut_key, switch_audio)]
    if overlay_enabled == "1" and overlay_key:
        hotkeys.append((overlay_key, open_device_overlay))

    for key_combo, callback in hotkeys:
        try:
            keyboard.add_hotkey(key_combo, callback, trigger_on_release=False)
        except Exception as e:
            log_error("register_hotkey", e)

    threading.Thread(target=_rehook_watchdog, args=(hotkeys,), daemon=True).start()
    threading.Thread(target=_switch_worker, daemon=True).start()
    threading.Thread(target=_direct_switch_worker, daemon=True).start()
    threading.Thread(target=_update_check_loop, daemon=True).start()

    root = tk.Tk()
    root.withdraw()
    root.report_callback_exception = lambda exc_type, exc_value, tb: log_error("tk_callback", exc_value)
    toast_manager = _ToastManager(root)
    device_overlay = _DeviceOverlay(root)

    def _on_overlay_select(name, target_id):
        _direct_switch_queue.put((target_id, name))
    device_overlay.on_select = _on_overlay_select

    def poll_toast_queue():
        try:
            while True:
                device_name = _toast_queue.get_nowait()
                toast_manager.show(device_name)
        except queue.Empty:
            pass
        except Exception as e:
            log_error("poll_toast_queue", e)

        try:
            while True:
                latest_tag, release_url = _update_queue.get_nowait()
                toast_manager.show_update_available(latest_tag, release_url)
        except queue.Empty:
            pass
        except Exception as e:
            log_error("poll_update_queue", e)

        try:
            while True:
                _overlay_open_queue.get_nowait()
                device_dict, current_default = get_device_info()
                device_overlay.open(device_dict, current_default)
        except queue.Empty:
            pass
        except Exception as e:
            log_error("poll_overlay_queue", e)

        root.after(50, poll_toast_queue)

    root.after(50, poll_toast_queue)
    root.mainloop()

def open_setting_window():
    """2重起動防止・ハイブリッドスタートアップ自動生成付き設定UI。
    General(既存のA/Bトグル)とMulti-Device(3台以上から選ぶオーバーレイ)を
    タブで分けている。ひとつの画面に全部詰め込むとごちゃつくため。"""
    import customtkinter as ctk  # 設定画面を開く時だけ読み込む（バックグラウンド常駐時は不要なため）

    kill_previous_instances()

    device_dict, current_default = get_device_info()
    devices = list(device_dict.keys())

    if not devices:
        root_err = tk.Tk()
        root_err.withdraw()
        tk.messagebox.showerror(
            "Error",
            "再生デバイスを取得できませんでした。\nオーディオドライバの状態を確認してください。"
        )
        devices = ["Speakers", "Headphones"]

    saved_a, saved_b, saved_key, saved_startup, saved_overlay_key, saved_overlay_enabled = load_config()
    devices_b_options = ["Select device..."] + devices

    ctk.set_appearance_mode("dark")

    root = ctk.CTk(fg_color=COLOR_BG)
    root.title("AudioSwitcher")
    root.geometry("560x680")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    if os.path.exists(ICON_PATH):
        try:
            root.iconbitmap(ICON_PATH)
        except Exception:
            pass

    def make_hotkey_recorder(parent, initial_value):
        """ホットキー入力欄(クリックしてから新しいキーを押すと変更できる)を
        作る。メインのホットキーとオーバーレイのホットキーで全く同じ挙動が
        必要なため、ここに共通化している。戻り値: (key_box, get_value)"""
        key_box = ctk.CTkFrame(parent, fg_color=COLOR_INPUT, corner_radius=8,
                                border_width=1, border_color=COLOR_BORDER)

        entry = ctk.CTkEntry(
            key_box, font=FONT_MONO, justify="center", state="readonly",
            fg_color=COLOR_INPUT, text_color=COLOR_TEXT_MAIN,
            border_width=0, corner_radius=8,
        )
        entry.pack(fill="x", padx=2, pady=2, ipady=6)
        entry.configure(state="normal")
        entry.insert(0, initial_value)
        entry.configure(state="readonly")

        rec = {"recording": False, "keys": set(), "value": initial_value, "saved": initial_value}
        key_map = {"Control_L": "ctrl", "Control_R": "ctrl", "Alt_L": "alt", "Alt_R": "alt",
                   "Shift_L": "shift", "Shift_R": "shift", "Win_L": "win", "Win_R": "win"}

        def on_key_press(event):
            if not rec["recording"]: return "break"
            if event.keysym == "Escape":
                rec["value"] = rec["saved"]
                root.focus_set()
                return "break"
            key_name = key_map.get(event.keysym, event.keysym.lower())
            if "caps" in key_name or "lock" in key_name or "num_" in key_name: return "break"
            rec["keys"].add(key_name)
            ordered = [m for m in ["ctrl", "alt", "shift", "win"] if m in rec["keys"]]
            ordered += sorted(k for k in rec["keys"] if k not in ["ctrl", "alt", "shift", "win"])
            rec["value"] = "+".join(ordered)
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.insert(0, rec["value"])
            entry.configure(state="readonly")
            return "break"

        def on_key_release(event):
            if not rec["recording"]: return "break"
            if rec["value"] and rec["value"] != "Press Keys...":
                rec["recording"] = False
                root.focus_set()
            return "break"

        def on_focus_in(e):
            rec["recording"] = True
            rec["keys"] = set()
            key_box.configure(border_color=COLOR_ACCENT)
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.insert(0, "Press Keys...")
            entry.configure(state="readonly", text_color=COLOR_ACCENT)

        def on_focus_out(e):
            rec["recording"] = False
            key_box.configure(border_color=COLOR_BORDER)
            entry.configure(text_color=COLOR_TEXT_MAIN)
            entry.configure(state="normal")
            entry.delete(0, "end")
            if rec["value"] == "Press Keys..." or not rec["value"]:
                entry.insert(0, rec["saved"])
                rec["value"] = rec["saved"]
            else:
                entry.insert(0, rec["value"])
            entry.configure(state="readonly")

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<KeyPress>", on_key_press)
        entry.bind("<KeyRelease>", on_key_release)

        return key_box, (lambda: rec["value"])

    # ヘッダー
    header = ctk.CTkFrame(root, fg_color="transparent")
    header.pack(fill="x", padx=30, pady=(30, 18))
    ctk.CTkLabel(header, text="AudioSwitcher", font=FONT_DISPLAY,
                 text_color=COLOR_TEXT_MAIN, anchor="w").pack(fill="x")
    ctk.CTkLabel(header, text="信号の経路（ルーティング）をワンショートカットで切り替える。",
                 font=FONT_BODY, text_color=COLOR_TEXT_SUB, anchor="w").pack(fill="x", pady=(2, 0))

    tabview = ctk.CTkTabview(
        root, fg_color=COLOR_PANEL, border_width=1, border_color=COLOR_BORDER, corner_radius=14,
        segmented_button_fg_color=COLOR_INPUT, segmented_button_selected_color=COLOR_ACCENT,
        segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
        segmented_button_unselected_color=COLOR_INPUT, segmented_button_unselected_hover_color=COLOR_INPUT,
        text_color=COLOR_TEXT_MAIN, text_color_disabled=COLOR_TEXT_FAINT,
    )
    tabview.pack(fill="both", expand=True, padx=30, pady=(0, 18))
    tab_general = tabview.add("General")
    tab_multi = tabview.add("Multi-Device")

    # ルーティング・インジケーター（現在どちらのレイヤーがアクティブかをパッチベイ風に表示）
    indicator_panel = ctk.CTkFrame(tab_general, fg_color=COLOR_PANEL, corner_radius=14,
                                    border_width=0)
    indicator_panel.pack(fill="x", pady=(4, 14))

    indicator_canvas = tk.Canvas(indicator_panel, width=220, height=56,
                                  bg=COLOR_PANEL, highlightthickness=0)
    indicator_canvas.pack(pady=(14, 4))

    status_label = ctk.CTkLabel(indicator_panel, font=("Segoe UI", 9), text_color=COLOR_TEXT_SUB)
    status_label.pack(pady=(0, 2))

    settings_toast = _ToastManager(root)

    def refresh_indicator():
        a, b = combo_a.get(), combo_b.get()
        if current_default and current_default == a:
            active_side = "A"
        elif current_default and current_default == b:
            active_side = "B"
        else:
            active_side = None
        draw_routing_indicator(indicator_canvas, active_side)
        status_label.configure(
            text=f"現在の既定デバイス: {current_default}" if current_default
            else "現在の既定デバイスを検出できませんでした"
        )

    def on_test_switch():
        nonlocal current_default
        a, b = combo_a.get(), combo_b.get()
        if b == "Select device...":
            tk.messagebox.showwarning("Routing Incomplete", "Please select a valid target for Device Layer B.")
            return
        target_display_name = b if current_default == a else a
        target_id = device_dict.get(target_display_name)
        if not target_id:
            tk.messagebox.showerror("Error", "選択されたデバイスが見つかりませんでした。")
            return
        try:
            ensure_com_initialized()
            AudioUtilities.SetDefaultDevice(target_id)
        except Exception as e:
            log_error("test_switch", e)
            tk.messagebox.showerror("Error", f"切り替えに失敗しました: {e}")
            return
        current_default = target_display_name
        refresh_indicator()
        settings_toast.show(target_display_name)

    test_button = ctk.CTkButton(
        indicator_panel, text="今すぐテスト切り替え", command=on_test_switch,
        fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_ACCENT,
        border_width=1, border_color=COLOR_ACCENT_DIM, corner_radius=8,
        font=("Segoe UI", 9, "bold"), height=28,
    )
    test_button.pack(pady=(2, 14))

    # デバイス設定カード
    container = ctk.CTkFrame(tab_general, fg_color="transparent")
    container.pack(fill="both", expand=True)

    def section_label(parent, text):
        ctk.CTkLabel(parent, text=text, font=FONT_LABEL, text_color=COLOR_TEXT_SUB,
                     anchor="w").pack(fill="x", padx=24, pady=(18, 6))

    combo_kwargs = dict(
        state="readonly", fg_color=COLOR_INPUT, border_color=COLOR_BORDER, border_width=1,
        button_color=COLOR_INPUT, button_hover_color=COLOR_ACCENT_DIM,
        dropdown_fg_color=COLOR_INPUT, dropdown_text_color=COLOR_TEXT_MAIN,
        dropdown_hover_color=COLOR_ACCENT_DIM, text_color=COLOR_TEXT_MAIN,
        corner_radius=8, height=36,
    )

    section_label(container, "DEVICE LAYER A")
    combo_a = ctk.CTkComboBox(container, values=devices, **combo_kwargs)
    combo_a.pack(fill="x", padx=24, pady=(0, 16))
    if saved_a in devices: combo_a.set(saved_a)
    elif current_default in devices: combo_a.set(current_default)
    else: combo_a.set(devices[0])

    section_label(container, "DEVICE LAYER B")
    combo_b = ctk.CTkComboBox(container, values=devices_b_options, **combo_kwargs)
    combo_b.pack(fill="x", padx=24, pady=(0, 16))
    if saved_b in devices_b_options and saved_b != "Select device...": combo_b.set(saved_b)
    else: combo_b.set(devices_b_options[0])

    combo_a.configure(command=lambda _=None: refresh_indicator())
    combo_b.configure(command=lambda _=None: refresh_indicator())
    refresh_indicator()

    # ホットキー
    section_label(container, "GLOBAL TRIGGER HOTKEY")
    key_box, get_shortcut = make_hotkey_recorder(container, saved_key)
    key_box.pack(fill="x", padx=24, pady=(0, 8))

    ctk.CTkLabel(container, text="クリックしてから新しいキーの組み合わせを押すと変更できます",
                 font=("Segoe UI", 8), text_color=COLOR_TEXT_FAINT, anchor="w").pack(
        fill="x", padx=24, pady=(0, 14))

    # ---- Multi-Deviceタブ: 3台以上の出力先から選ぶオーバーレイの設定 ----
    container2 = ctk.CTkFrame(tab_multi, fg_color="transparent")
    container2.pack(fill="both", expand=True)

    ctk.CTkLabel(
        container2,
        text="ホットキーを押すと、今つながっている出力デバイスが画面中央に一覧表示され、\nそこから選んで切り替えられます。",
        font=FONT_BODY, text_color=COLOR_TEXT_SUB, justify="left", anchor="w"
    ).pack(fill="x", pady=(16, 14))

    overlay_enabled_var = tk.StringVar(value=saved_overlay_enabled)
    switch_overlay = ctk.CTkSwitch(
        container2, text="Enable device-select overlay", variable=overlay_enabled_var, onvalue="1", offvalue="0",
        font=FONT_BODY, text_color=COLOR_TEXT_SUB,
        fg_color=COLOR_INPUT, progress_color=COLOR_ACCENT,
        button_color=COLOR_TEXT_MAIN, button_hover_color=COLOR_ACCENT_HOVER,
    )
    switch_overlay.pack(anchor="w", pady=(0, 18))

    section_label(container2, "OVERLAY HOTKEY")
    overlay_key_box, get_overlay_shortcut = make_hotkey_recorder(container2, saved_overlay_key)
    overlay_key_box.pack(fill="x", padx=24, pady=(0, 8))
    ctk.CTkLabel(container2, text="クリックしてから新しいキーの組み合わせを押すと変更できます",
                 font=("Segoe UI", 8), text_color=COLOR_TEXT_FAINT, anchor="w").pack(
        fill="x", padx=24, pady=(0, 14))

    section_label(container2, f"DETECTED DEVICES ({len(devices)})")
    device_list_frame = ctk.CTkFrame(container2, fg_color=COLOR_INPUT, corner_radius=8,
                                      border_width=1, border_color=COLOR_BORDER)
    device_list_frame.pack(fill="x", padx=24, pady=(0, 10))
    for name in devices:
        row_text = f"● {name}" if name == current_default else name
        ctk.CTkLabel(
            device_list_frame, text=row_text, font=("Segoe UI", 9), anchor="w",
            text_color=(COLOR_ACCENT if name == current_default else COLOR_TEXT_SUB)
        ).pack(fill="x", padx=14, pady=4)

    # ボトム
    bottom_bar = ctk.CTkFrame(root, fg_color="transparent")
    bottom_bar.pack(fill="x", padx=30, pady=(0, 30))

    startup_var = tk.StringVar(value=saved_startup)
    switch_startup = ctk.CTkSwitch(
        bottom_bar, text="Start on boot", variable=startup_var, onvalue="1", offvalue="0",
        font=FONT_BODY, text_color=COLOR_TEXT_SUB,
        fg_color=COLOR_INPUT, progress_color=COLOR_ACCENT,
        button_color=COLOR_TEXT_MAIN, button_hover_color=COLOR_ACCENT_HOVER,
    )
    switch_startup.pack(side="left", anchor="w")

    btn_frame = ctk.CTkFrame(bottom_bar, fg_color="transparent")
    btn_frame.pack(side="right")

    def on_save_only():
        if combo_b.get() == "Select device...":
            tk.messagebox.showwarning("Routing Incomplete", "Please select a valid target for Device Layer B.")
            return
        save_config(combo_a.get(), combo_b.get(), get_shortcut(), startup_var.get(),
                    get_overlay_shortcut(), overlay_enabled_var.get())
        root.destroy()
        sys.exit(0)

    btn_save_only = ctk.CTkButton(
        btn_frame, text="Save only", command=on_save_only,
        fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
        border_width=1, border_color=COLOR_BORDER, corner_radius=8,
        font=FONT_LABEL, width=100, height=36,
    )
    btn_save_only.pack(side="left", padx=(0, 10))

    def on_save_and_run():
        if combo_b.get() == "Select device...":
            tk.messagebox.showwarning("Routing Incomplete", "Please select a valid target for Device Layer B.")
            return
        save_config(combo_a.get(), combo_b.get(), get_shortcut(), startup_var.get(),
                    get_overlay_shortcut(), overlay_enabled_var.get())
        root.destroy()
        start_backend()

    btn_save_run = ctk.CTkButton(
        btn_frame, text="Save & Run", command=on_save_and_run,
        fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
        corner_radius=8, font=FONT_LABEL, width=120, height=36,
    )
    btn_save_run.pack(side="right")

    root.mainloop()

if __name__ == "__main__":
    try:
        if "/background" in sys.argv:
            if os.path.exists(CONFIG_FILE): start_backend()
            else: open_setting_window()
        else:
            open_setting_window()
    except SystemExit:
        raise
    except Exception as e:
        log_error("main", e)
        raise