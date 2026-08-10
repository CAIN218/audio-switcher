import os
import subprocess
import keyboard
import tkinter as tk
from tkinter import messagebox
from tkinter import font as tkfont
import time
import sys
import warnings
import threading
import queue
import ctypes
import json
import tempfile
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

__version__ = "1.5.0"
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

def enable_dpi_awareness():
    """Per-Monitor DPIを意識させる。これを呼ばないとGetCursorPos/
    MonitorFromPointがDPI仮想化された座標を返すことがあり、マルチ
    モニター環境(特に解像度やDPIが混在する構成)でトーストやオーバーレイ
    が実際のカーソル位置とズレた場所(最悪、どのモニターにも表示されない
    座標)に出てしまう。Luminaで同じ症状(パネルが意図しない位置に出る)
    を修正した時と同じ対策。PyInstaller化した場合マニフェストでのDPI宣言
    が無いため、コード側で明示する必要がある。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception as e:
            log_error("enable_dpi_awareness", e)

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
    """設定ファイルの読み込み。5〜8行目(複数デバイス切り替えオーバーレイの
    ホットキー・有効/無効、A/Bトグル自体の有効/無効、自動更新の有効/無効)は
    後から追加した項目なので、無い場合はデフォルト値にフォールバックし、
    古い設定ファイルもそのまま読み込めるようにしている。A/Bトグルは元々
    常に有効だった機能なので、値が無い場合のデフォルトは"1"(有効)にして
    既存ユーザーの挙動を変えないようにしている。自動更新も同様に既定は
    "1"(有効)。"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
        dev_a = lines[0] if len(lines) > 0 else ""
        dev_b = lines[1] if len(lines) > 1 else ""
        key = lines[2] if len(lines) > 2 else "ctrl+alt+a"
        startup_val = lines[3] if len(lines) > 3 else "0"
        overlay_key = lines[4] if len(lines) > 4 else DEFAULT_OVERLAY_HOTKEY
        overlay_enabled = lines[5] if len(lines) > 5 else "0"
        ab_enabled = lines[6] if len(lines) > 6 else "1"
        auto_update_enabled = lines[7] if len(lines) > 7 else "1"
        return dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled, ab_enabled, auto_update_enabled
    return "", "", "ctrl+alt+a", "0", DEFAULT_OVERLAY_HOTKEY, "0", "1", "1"

def save_config(dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled, ab_enabled, auto_update_enabled="1"):
    """設定ファイルの保存とスタートアップ制御"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(f"{dev_a}\n{dev_b}\n{key}\n{startup_val}\n{overlay_key}\n{overlay_enabled}\n{ab_enabled}\n{auto_update_enabled}")
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

    DEVICE_FONT = ("Segoe UI", 10, "bold")

    def __init__(self, root):
        self.root = root
        self.window = None
        self.canvas = None
        self.fade_job = None
        self.marquee_job = None
        self.current_alpha = 0.0
        self.on_click = None
        self._hold_ms = 2600
        self._device_font = tkfont.Font(root=root, family=self.DEVICE_FONT[0],
                                         size=self.DEVICE_FONT[1], weight=self.DEVICE_FONT[2])
        self._update_subtitle_id = None
        self._pending_release_url = None
        self.auto_update_enabled = True

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
        if self.marquee_job:
            self.root.after_cancel(self.marquee_job)
            self.marquee_job = None

        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 16,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        # 左端のカッパーのアクセントバー
        canvas.create_rectangle(0, 12, 4, self.HEIGHT - 12, fill=COLOR_ACCENT, outline="")
        canvas.create_text(22, 22, anchor="w", text="AUDIO ROUTING SWITCHED",
                            fill=COLOR_ACCENT, font=("Segoe UI", 8, "bold"))

        # 長いデバイス名は「...」で切らず、収まりきらない分だけ左右に
        # ゆっくり往復させて全体を読めるようにする(Canvasは自身の外側を
        # 自動的に描画しないので、はみ出した部分は特別なクリップ処理
        # なしでそのまま隠れる)。
        left_x = 22
        right_bound = self.WIDTH - 14
        text_item = canvas.create_text(left_x, 44, anchor="w", text=device_name,
                                        fill=COLOR_TEXT_MAIN, font=self.DEVICE_FONT)
        text_width = self._device_font.measure(device_name)
        if text_width > (right_bound - left_x):
            self._start_marquee(text_item, text_width, left_x, right_bound)

        canvas.configure(cursor="")

        self.on_click = None
        self._update_subtitle_id = None
        self._hold_ms = 2600
        self._place_bottom_right()
        self.window.deiconify()
        self._place_bottom_right()
        self.current_alpha = 0.0
        self._fade_in()

    def _start_marquee(self, item, text_width, left_x, right_bound):
        """収まりきらないデバイス名を、はみ出した分だけ左へスクロールして
        端で少し止まり、また右へ戻って止まる、を繰り返す。"""
        canvas = self.canvas
        scroll_distance = text_width - (right_bound - left_x)
        min_x = left_x - scroll_distance
        state = {"x": left_x, "dir": -1, "pause": 18}  # 表示直後に少し静止してから動き出す

        def tick():
            if not canvas.winfo_exists() or self.window is None:
                return
            if state["pause"] > 0:
                state["pause"] -= 1
                self.marquee_job = self.root.after(40, tick)
                return
            state["x"] += state["dir"] * 1.4
            if state["dir"] < 0 and state["x"] <= min_x:
                state["x"] = min_x
                state["dir"] = 1
                state["pause"] = 18
            elif state["dir"] > 0 and state["x"] >= left_x:
                state["x"] = left_x
                state["dir"] = -1
                state["pause"] = 18
            canvas.coords(item, state["x"], 44)
            self.marquee_job = self.root.after(40, tick)

        self.marquee_job = self.root.after(40, tick)

    def show_update_available(self, latest_tag, release_url, installer_url=None):
        """新しいバージョンが見つかった時の控えめな通知。installer_urlがあれば
        クリックでインストーラーを直接ダウンロード・自動起動する。なければ
        (ポータブル版のみのリリース等)従来通りリリースページを開く。頻繁に
        出るものではないので、通常のトーストより長めに表示したままにする。"""
        self._ensure_window()
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None
        if self.marquee_job:
            self.root.after_cancel(self.marquee_job)
            self.marquee_job = None

        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 16,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        canvas.create_rectangle(0, 12, 4, self.HEIGHT - 12, fill=COLOR_ACCENT, outline="")
        canvas.create_text(22, 22, anchor="w", text="UPDATE AVAILABLE",
                            fill=COLOR_ACCENT, font=("Segoe UI", 8, "bold"))
        auto = self.auto_update_enabled and installer_url
        if auto:
            subtitle = f"{latest_tag} — installing automatically..."
        elif installer_url:
            subtitle = f"{latest_tag} — click to install"
        else:
            subtitle = f"{latest_tag} — click to view"
        self._update_subtitle_id = canvas.create_text(22, 44, anchor="w", text=subtitle,
                                                        fill=COLOR_TEXT_MAIN, font=("Segoe UI", 10, "bold"))

        self._pending_release_url = release_url
        if auto:
            canvas.configure(cursor="")
            self.on_click = None
        elif installer_url:
            canvas.configure(cursor="hand2")
            self.on_click = lambda: self._begin_update_download(installer_url, release_url)
        else:
            canvas.configure(cursor="hand2")
            self.on_click = lambda: webbrowser.open(release_url)
        self._hold_ms = 8000
        self._place_bottom_right()
        self.window.deiconify()
        self._place_bottom_right()
        self.current_alpha = 0.0
        self._fade_in()

        if auto:
            self._begin_update_download(installer_url, release_url)

    def _begin_update_download(self, installer_url, release_url):
        """アップデート通知がクリックされたら、インストーラーをバックグラ
        ウンドでダウンロードして自動起動する。Tkのメインスレッドをブロック
        しないよう別スレッドに投げ、結果は_install_status_queue経由で
        poll_toast_queue(メインスレッド)に伝えて表示を更新する。"""
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None
        self.on_click = None
        try:
            self.canvas.configure(cursor="")
            if self._update_subtitle_id is not None:
                self.canvas.itemconfigure(self._update_subtitle_id, text="Downloading installer...")
        except Exception:
            pass
        threading.Thread(target=_download_and_launch_installer,
                          args=(installer_url, release_url), daemon=True).start()

    def _finish_update_download(self, ok):
        """ダウンロード完了(成功/失敗)を受けて表示を更新する。失敗時は
        従来通りクリックでリリースページを開けるようにフォールバックする。"""
        try:
            if self._update_subtitle_id is not None and self.canvas is not None:
                if ok:
                    msg = "Installing update — AudioSwitcher will restart automatically."
                    self.canvas.itemconfigure(self._update_subtitle_id, text=msg)
                else:
                    self.canvas.itemconfigure(self._update_subtitle_id, text="Download failed — click to view")
                    self.canvas.configure(cursor="hand2")
                    self.on_click = lambda: webbrowser.open(self._pending_release_url)
        except Exception:
            pass
        self._hold_ms = 5000
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
        self.fade_job = self.root.after(self._hold_ms, self._fade_out)

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
            if self.marquee_job:
                self.root.after_cancel(self.marquee_job)
                self.marquee_job = None


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
_install_status_queue = queue.Queue()


def _fetch_latest_release_info():
    """GitHubの最新リリースを確認する。__version__より新しい場合だけ
    (latest_tag, release_url, installer_url)を返し、そうでなければNoneを
    返す。バックグラウンドの定期チェックと、設定画面を開くたびの即時チェック
    の両方から共通で使う。呼び出し側で例外を処理すること。"""
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
        headers={"Accept": "application/vnd.github+json", "User-Agent": "AudioSwitcher-update-check"},
    )
    with urllib.request.urlopen(req, timeout=5) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    latest_tag = data.get("tag_name", "")
    release_url = data.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases/latest"
    installer_url = None
    for asset in data.get("assets", []) or []:
        name = asset.get("name", "")
        if name.lower().startswith("audioswitcher-setup") and name.lower().endswith(".exe"):
            installer_url = asset.get("browser_download_url")
            break
    if latest_tag and _parse_version(latest_tag) > _parse_version(__version__):
        return (latest_tag, release_url, installer_url)
    return None


def _update_check_loop():
    """1日1回、GitHubの最新リリースを確認する。バックグラウンドの常駐
    ツールなので、失敗(オフライン等)は完全に握りつぶし通常動作に一切
    影響させない。新しいバージョンが見つかった時だけ、トースト表示用の
    キューに積んでメインスレッド側(poll_toast_queue)に伝える。"""
    while True:
        try:
            info = _fetch_latest_release_info()
            if info:
                _update_queue.put(info)
        except Exception as e:
            log_error("update_check", e)
        time.sleep(24 * 60 * 60)


def _download_and_launch_installer(installer_url, release_url):
    """アップデート通知のクリック(または自動更新)を受けて、インストーラー
    をダウンロードし、ウィザード画面を出さずサイレントインストールする。
    installer.iss側の[Code]がサイレント時も明示的にアプリを再起動するので、
    ここでは起動するところまでで良い。ネットワークエラー等は結果を
    _install_status_queue経由でメインスレッドに伝え、失敗時はリリース
    ページを開くフォールバックに切り替えてもらう。"""
    try:
        req = urllib.request.Request(installer_url, headers={"User-Agent": "AudioSwitcher-update-check"})
        installer_path = os.path.join(tempfile.gettempdir(), os.path.basename(installer_url))
        with urllib.request.urlopen(req, timeout=30) as resp, open(installer_path, "wb") as f:
            f.write(resp.read())
        os.startfile(installer_path, arguments="/VERYSILENT /SUPPRESSMSGBOX /NORESTART")
        _install_status_queue.put(("done", None))
    except Exception as e:
        log_error("download_installer", e)
        _install_status_queue.put(("error", release_url))


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
            dev_a, dev_b, _, _, _, _, _, _ = load_config()
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
    _, _, shortcut_key, _, overlay_key, overlay_enabled, ab_enabled, auto_update_enabled = load_config()

    hotkeys = []
    if ab_enabled == "1" and shortcut_key:
        hotkeys.append((shortcut_key, switch_audio))
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
    toast_manager.auto_update_enabled = (auto_update_enabled == "1")
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
                latest_tag, release_url, installer_url = _update_queue.get_nowait()
                toast_manager.show_update_available(latest_tag, release_url, installer_url)
        except queue.Empty:
            pass
        except Exception as e:
            log_error("poll_update_queue", e)

        try:
            while True:
                status, payload = _install_status_queue.get_nowait()
                if status == "error" and payload:
                    toast_manager._pending_release_url = payload
                toast_manager._finish_update_download(status == "done")
        except queue.Empty:
            pass
        except Exception as e:
            log_error("poll_install_status_queue", e)

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
    A/B SWITCH(既存のA/Bトグル)とMulti-Device(3台以上から選ぶオーバーレイ)を
    タブで分けている。ひとつの画面に全部詰め込むとごちゃつくため。"""
    import customtkinter as ctk  # 設定画面を開く時だけ読み込む（バックグラウンド常駐時は不要なため）

    is_first_run = not os.path.exists(CONFIG_FILE)

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

    saved_a, saved_b, saved_key, saved_startup, saved_overlay_key, saved_overlay_enabled, saved_ab_enabled, saved_auto_update = load_config()
    devices_b_options = ["Select device..."] + devices

    ctk.set_appearance_mode("dark")

    root = ctk.CTk(fg_color=COLOR_BG)
    # UPDATEタブの中に置くスイッチだが、変数自体はSave時に他の設定と一緒に
    # 保存する必要があるので、タブの有無に関わらずここで(root作成後に)
    # 用意しておく(StringVarはデフォルトのTkルートが無いと作れないため)。
    auto_update_var = tk.StringVar(value=saved_auto_update)
    root.title("AudioSwitcher")
    root.geometry("560x600")
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
                   "Shift_L": "shift", "Shift_R": "shift", "Win_L": "win", "Win_R": "win",
                   "Prior": "page up", "Next": "page down"}

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

        def set_enabled(enabled):
            """機能トグルがOFFの時、グレーアウトして触れないようにする"""
            if enabled:
                rec["saved"] = rec["value"]
                entry.configure(state="readonly", text_color=COLOR_TEXT_MAIN)
                key_box.configure(border_color=COLOR_BORDER)
            else:
                rec["recording"] = False
                entry.configure(state="disabled", text_color=COLOR_TEXT_FAINT)
                key_box.configure(border_color=COLOR_BORDER)

        return key_box, (lambda: rec["value"]), set_enabled

    # ヘッダー
    header = ctk.CTkFrame(root, fg_color="transparent")
    header.pack(fill="x", padx=30, pady=(22, 14))
    title_row = ctk.CTkFrame(header, fg_color="transparent")
    title_row.pack(fill="x")
    ctk.CTkLabel(title_row, text="AudioSwitcher", font=FONT_DISPLAY,
                 text_color=COLOR_TEXT_MAIN, anchor="w").pack(side="left")
    help_button = ctk.CTkButton(
        title_row, text="?", command=lambda: show_onboarding(),
        width=26, height=26, corner_radius=13,
        fg_color=COLOR_INPUT, hover_color=COLOR_ACCENT_DIM, text_color=COLOR_TEXT_SUB,
        border_width=1, border_color=COLOR_BORDER, font=("Segoe UI", 10, "bold"),
    )
    help_button.pack(side="right")
    ctk.CTkLabel(header, text="信号の経路（ルーティング）をワンショートカットで切り替える。",
                 font=FONT_BODY, text_color=COLOR_TEXT_SUB, anchor="w").pack(fill="x", pady=(2, 0))

    tabview = ctk.CTkTabview(
        root, width=1, height=1, fg_color=COLOR_PANEL, border_width=1, border_color=COLOR_BORDER, corner_radius=14,
        segmented_button_fg_color=COLOR_INPUT, segmented_button_selected_color=COLOR_ACCENT,
        segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
        segmented_button_unselected_color=COLOR_INPUT, segmented_button_unselected_hover_color=COLOR_ACCENT_DIM,
        segmented_button_font=("Segoe UI", 11, "bold"),
        text_color=COLOR_TEXT_MAIN, text_color_disabled=COLOR_TEXT_FAINT,
    )
    tabview.pack(fill="both", expand=True, padx=30, pady=(0, 14))
    tab_general = tabview.add("A/B SWITCH")
    tab_multi = tabview.add("MULTI-DEVICE")
    tab_update = tabview.add("UPDATE")

    # UPDATEタブは常に表示しておく(更新がある時しか出ないと、Auto-update
    # トグル自体を普段いじれなくなってしまうため)。中身(タイトル/ボタン)は
    # チェック結果に応じて後から書き換える。
    update_panel = ctk.CTkFrame(tab_update, fg_color=COLOR_PANEL, corner_radius=14, border_width=0)
    update_panel.pack(fill="x", pady=(2, 10))

    update_title_var = tk.StringVar(value="CHECKING FOR UPDATES...")
    ctk.CTkLabel(update_panel, textvariable=update_title_var, font=("Segoe UI", 9, "bold"),
                 text_color=COLOR_ACCENT, anchor="w").pack(fill="x", padx=20, pady=(18, 2))
    update_detail_var = tk.StringVar(value=f"Current version: v{__version__}")
    ctk.CTkLabel(update_panel, textvariable=update_detail_var, font=FONT_BODY, text_color=COLOR_TEXT_MAIN,
                 anchor="w").pack(fill="x", padx=20, pady=(0, 4))
    update_status_var = tk.StringVar(value="")
    ctk.CTkLabel(update_panel, textvariable=update_status_var, font=FONT_BODY, text_color=COLOR_TEXT_SUB,
                 anchor="w", wraplength=460, justify="left").pack(fill="x", padx=20, pady=(0, 10))

    _update_state = {"installer_url": None, "release_url": None}

    # インストール自体は(ONでもOFFでも)常にウィザードなしのサイレント
    # インストールで行う。ONは何もしなくても自動でこれが走り、OFFは
    # ボタンを押すことがそのトリガーになる、という違いだけ。
    def start_install():
        installer_url = _update_state["installer_url"]
        release_url = _update_state["release_url"]
        if not installer_url:
            if release_url:
                webbrowser.open(release_url)
            return
        install_btn.configure(state="disabled", text="Updating...")
        update_status_var.set("")

        def worker():
            try:
                req = urllib.request.Request(installer_url, headers={"User-Agent": "AudioSwitcher-update-check"})
                installer_path = os.path.join(tempfile.gettempdir(), os.path.basename(installer_url))
                with urllib.request.urlopen(req, timeout=30) as resp, open(installer_path, "wb") as f:
                    f.write(resp.read())
                os.startfile(installer_path, arguments="/VERYSILENT /SUPPRESSMSGBOX /NORESTART")
                msg = "Installing update in the background — AudioSwitcher will restart automatically."
                root.after(0, lambda: update_status_var.set(msg))
            except Exception as e:
                log_error("settings_download_installer", e)

                def on_fail():
                    install_btn.configure(state="normal", text="Open release page",
                                           command=lambda: webbrowser.open(release_url))
                    update_status_var.set("Download failed. Click to open the release page instead.")
                root.after(0, on_fail)
        threading.Thread(target=worker, daemon=True).start()

    def on_auto_toggle():
        if auto_update_var.get() == "1" and _update_state["installer_url"]:
            start_install()

    switch_auto_update = ctk.CTkSwitch(
        update_panel, text="Auto-update (always stay on the latest version)",
        variable=auto_update_var, onvalue="1", offvalue="0", command=on_auto_toggle,
        font=FONT_BODY, text_color=COLOR_TEXT_SUB,
        fg_color=COLOR_INPUT, progress_color=COLOR_ACCENT,
        button_color=COLOR_TEXT_MAIN, button_hover_color=COLOR_ACCENT_HOVER,
    )
    switch_auto_update.pack(fill="x", padx=20, pady=(4, 12), anchor="w")

    install_btn = ctk.CTkButton(
        update_panel, text="Up to date", state="disabled", command=start_install,
        fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
        corner_radius=8, font=FONT_LABEL, height=36,
    )
    install_btn.pack(fill="x", padx=20, pady=(0, 20))

    _update_info_queue = queue.Queue()

    def _check_for_updates():
        def worker():
            try:
                info = _fetch_latest_release_info()
            except Exception as e:
                log_error("settings_update_check", e)
                info = None
            _update_info_queue.put(info)
        threading.Thread(target=worker, daemon=True).start()

    def _poll_update_info_queue():
        try:
            info = _update_info_queue.get_nowait()
            if info:
                latest_tag, release_url, installer_url = info
                _update_state["installer_url"] = installer_url
                _update_state["release_url"] = release_url
                update_title_var.set("UPDATE AVAILABLE")
                update_detail_var.set(f"{latest_tag} is available (current: v{__version__})")
                install_btn.configure(state="normal", text="Download & Install")
                if auto_update_var.get() == "1" and installer_url:
                    start_install()
            else:
                update_title_var.set("UP TO DATE")
                update_detail_var.set(f"You're on the latest version (v{__version__})")
            return
        except queue.Empty:
            pass
        root.after(500, _poll_update_info_queue)

    _check_for_updates()
    root.after(500, _poll_update_info_queue)

    # ルーティング・インジケーター（現在どちらのレイヤーがアクティブかをパッチベイ風に表示）
    indicator_panel = ctk.CTkFrame(tab_general, fg_color=COLOR_PANEL, corner_radius=14,
                                    border_width=0)
    indicator_panel.pack(fill="x", pady=(2, 10))

    indicator_canvas = tk.Canvas(indicator_panel, width=220, height=52,
                                  bg=COLOR_PANEL, highlightthickness=0)
    indicator_canvas.pack(pady=(10, 2))

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
    test_button.pack(pady=(2, 10))

    # デバイス設定カード
    container = ctk.CTkFrame(tab_general, fg_color="transparent")
    container.pack(fill="both", expand=True)

    def section_label(parent, text):
        lbl = ctk.CTkLabel(parent, text=text, font=FONT_LABEL, text_color=COLOR_TEXT_SUB, anchor="w")
        lbl.pack(fill="x", padx=24, pady=(10, 4))
        return lbl

    ab_enabled_var = tk.StringVar(value=saved_ab_enabled)
    switch_ab = ctk.CTkSwitch(
        container, text="Enable A/B switch", variable=ab_enabled_var, onvalue="1", offvalue="0",
        font=FONT_BODY, text_color=COLOR_TEXT_SUB,
        fg_color=COLOR_INPUT, progress_color=COLOR_ACCENT,
        button_color=COLOR_TEXT_MAIN, button_hover_color=COLOR_ACCENT_HOVER,
    )
    switch_ab.pack(anchor="w", padx=24, pady=(0, 8))

    combo_kwargs = dict(
        state="readonly", fg_color=COLOR_INPUT, border_color=COLOR_BORDER, border_width=1,
        button_color=COLOR_INPUT, button_hover_color=COLOR_ACCENT_DIM,
        dropdown_fg_color=COLOR_INPUT, dropdown_text_color=COLOR_TEXT_MAIN,
        dropdown_hover_color=COLOR_ACCENT_DIM, text_color=COLOR_TEXT_MAIN,
        corner_radius=8, height=34,
    )

    def make_combo_fully_clickable(combo):
        """CTkComboBoxは標準だと右端の矢印(dropdown_arrow)部分しかクリックで
        開かない仕様(customtkinter側の_create_bindingsがright_parts/
        dropdown_arrowタグにしかBUtton-1を張っていないため)。本体のテキスト
        部分(Entry)をクリックしても同じ_clicked()ハンドラーが呼ばれるように
        バインドを足し、枠のどこを押しても開くようにする。
        またEntry標準のIビーム(テキスト選択)カーソルのままだと編集可能な
        欄に見えてクリックできることが伝わりづらいので、通常の矢印カーソル
        に戻しておく。"""
        combo._entry.bind("<Button-1>", combo._clicked)
        combo._entry.configure(cursor="arrow")

    label_a = section_label(container, "DEVICE LAYER A")
    combo_a = ctk.CTkComboBox(container, values=devices, **combo_kwargs)
    combo_a.pack(fill="x", padx=24, pady=(0, 10))
    if saved_a in devices: combo_a.set(saved_a)
    elif current_default in devices: combo_a.set(current_default)
    else: combo_a.set(devices[0])
    make_combo_fully_clickable(combo_a)

    label_b = section_label(container, "DEVICE LAYER B")
    combo_b = ctk.CTkComboBox(container, values=devices_b_options, **combo_kwargs)
    combo_b.pack(fill="x", padx=24, pady=(0, 10))
    if saved_b in devices_b_options and saved_b != "Select device...": combo_b.set(saved_b)
    else: combo_b.set(devices_b_options[0])
    make_combo_fully_clickable(combo_b)

    combo_a.configure(command=lambda _=None: refresh_indicator())
    combo_b.configure(command=lambda _=None: refresh_indicator())
    refresh_indicator()

    # ホットキー
    label_hotkey = section_label(container, "GLOBAL TRIGGER HOTKEY")
    key_box, get_shortcut, set_key_enabled = make_hotkey_recorder(container, saved_key)
    key_box.pack(fill="x", padx=24, pady=(0, 6))

    ctk.CTkLabel(container, text="クリックしてから新しいキーの組み合わせを押すと変更できます",
                 font=("Segoe UI", 8), text_color=COLOR_TEXT_FAINT, anchor="w").pack(
        fill="x", padx=24, pady=(0, 8))

    def apply_ab_enabled_style(*_):
        enabled = ab_enabled_var.get() == "1"
        combo_a.configure(state=("readonly" if enabled else "disabled"))
        combo_b.configure(state=("readonly" if enabled else "disabled"))
        set_key_enabled(enabled)
        dim = COLOR_TEXT_SUB if enabled else COLOR_TEXT_FAINT
        label_a.configure(text_color=dim)
        label_b.configure(text_color=dim)
        label_hotkey.configure(text_color=dim)

    switch_ab.configure(command=apply_ab_enabled_style)
    apply_ab_enabled_style()

    # ---- Multi-Deviceタブ: 3台以上の出力先から選ぶオーバーレイの設定 ----
    container2 = ctk.CTkFrame(tab_multi, fg_color="transparent")
    container2.pack(fill="both", expand=True)

    ctk.CTkLabel(
        container2,
        text="ホットキーを押すと、今つながっている出力デバイスが画面中央に一覧表示され、\nそこから選んで切り替えられます。",
        font=FONT_BODY, text_color=COLOR_TEXT_SUB, justify="left", anchor="w"
    ).pack(fill="x", padx=24, pady=(10, 10))

    overlay_enabled_var = tk.StringVar(value=saved_overlay_enabled)
    switch_overlay = ctk.CTkSwitch(
        container2, text="Enable device-select overlay", variable=overlay_enabled_var, onvalue="1", offvalue="0",
        font=FONT_BODY, text_color=COLOR_TEXT_SUB,
        fg_color=COLOR_INPUT, progress_color=COLOR_ACCENT,
        button_color=COLOR_TEXT_MAIN, button_hover_color=COLOR_ACCENT_HOVER,
    )
    switch_overlay.pack(anchor="w", padx=24, pady=(0, 10))

    label_overlay_key = section_label(container2, "OVERLAY HOTKEY")
    overlay_key_box, get_overlay_shortcut, set_overlay_key_enabled = make_hotkey_recorder(container2, saved_overlay_key)
    overlay_key_box.pack(fill="x", padx=24, pady=(0, 6))
    ctk.CTkLabel(container2, text="クリックしてから新しいキーの組み合わせを押すと変更できます",
                 font=("Segoe UI", 8), text_color=COLOR_TEXT_FAINT, anchor="w").pack(
        fill="x", padx=24, pady=(0, 8))

    def apply_overlay_enabled_style(*_):
        enabled = overlay_enabled_var.get() == "1"
        set_overlay_key_enabled(enabled)
        label_overlay_key.configure(text_color=(COLOR_TEXT_SUB if enabled else COLOR_TEXT_FAINT))

    switch_overlay.configure(command=apply_overlay_enabled_style)
    apply_overlay_enabled_style()

    section_label(container2, f"DETECTED DEVICES ({len(devices)})")
    device_list_frame = ctk.CTkFrame(container2, fg_color=COLOR_INPUT, corner_radius=8,
                                      border_width=1, border_color=COLOR_BORDER)
    device_list_frame.pack(fill="x", padx=24, pady=(0, 8))
    for name in devices:
        row_text = f"● {name}" if name == current_default else name
        ctk.CTkLabel(
            device_list_frame, text=row_text, font=("Segoe UI", 9), anchor="w",
            text_color=(COLOR_ACCENT if name == current_default else COLOR_TEXT_SUB)
        ).pack(fill="x", padx=14, pady=3)

    # ボトム
    bottom_bar = ctk.CTkFrame(root, fg_color="transparent")
    bottom_bar.pack(fill="x", padx=30, pady=(0, 22))

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
        if ab_enabled_var.get() == "1" and combo_b.get() == "Select device...":
            tk.messagebox.showwarning("Routing Incomplete", "Please select a valid target for Device Layer B.")
            return
        save_config(combo_a.get(), combo_b.get(), get_shortcut(), startup_var.get(),
                    get_overlay_shortcut(), overlay_enabled_var.get(), ab_enabled_var.get(),
                    auto_update_var.get())
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
        if ab_enabled_var.get() == "1" and combo_b.get() == "Select device...":
            tk.messagebox.showwarning("Routing Incomplete", "Please select a valid target for Device Layer B.")
            return
        save_config(combo_a.get(), combo_b.get(), get_shortcut(), startup_var.get(),
                    get_overlay_shortcut(), overlay_enabled_var.get(), ab_enabled_var.get(),
                    auto_update_var.get())
        root.destroy()
        start_backend()

    btn_save_run = ctk.CTkButton(
        btn_frame, text="Save & Run", command=on_save_and_run,
        fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
        corner_radius=8, font=FONT_LABEL, width=120, height=36,
    )
    btn_save_run.pack(side="right")

    # ウィンドウ高さは中身の実測サイズに合わせて自動調整する。CTkTabviewを
    # 導入したことで固定の見積もり値だと下端(ボトムバー)が切れることが
    # あったため、両タブの中身を組み終えたこのタイミングで測り直す。
    root.update_idletasks()
    required_height = root.winfo_reqheight() + 12
    root.geometry(f"560x{required_height}")

    # ---- 機能紹介ウォークスルー。初回起動時に自動で出るほか、ヘッダーの
    # ヘルプボタン(?)からいつでも呼び出せる。実際のトースト/オーバーレイを
    # その場で呼び出す方式は、モニター判定の問題(トーストが別モニターに
    # 出て見えない等)に左右されやすいため、ウィザード自身のCanvas上だけで
    # 完結する「キー操作のイラストアニメ」にしている
    # (実機・実際のウィンドウ位置に依存しない)。----
    def show_onboarding():
        wiz = {"step": 0, "window": None}

        def draw_keycaps(canvas, labels, cx, cy):
            """CTRL/ALT/Aのようなキーキャップ列を描き、後で押下アニメーション
            させるための(rect_id, text_id)のリストを返す。"""
            widths = [46 if len(l) <= 4 else 30 for l in labels]
            gap = 8
            total_w = sum(widths) + gap * (len(labels) - 1)
            x = cx - total_w / 2
            items = []
            for label, w in zip(labels, widths):
                rect = _rounded_rect(canvas, x, cy - 15, x + w, cy + 15, 6,
                                      fill=COLOR_INPUT, outline=COLOR_BORDER, width=1)
                text = canvas.create_text(x + w / 2, cy, text=label,
                                           fill=COLOR_TEXT_MAIN, font=("Segoe UI", 9, "bold"))
                items.append((rect, text))
                x += w + gap
            return items

        def _lerp_hex(c1, c2, t):
            """2つの#rrggbb色をtで線形補間する。トグルの色変化やクリック時の
            フラッシュを、パキッと切り替わるのではなく滑らかに見せるために使う。"""
            t = max(0.0, min(1.0, t))
            r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
            r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
            return f"#{round(r1+(r2-r1)*t):02x}{round(g1+(g2-g1)*t):02x}{round(b1+(b2-b1)*t):02x}"

        def tween(canvas, duration_ms, on_tick, on_done=None):
            """0〜1のイーズアウトされた進捗を毎フレームon_tickに渡す、
            汎用のなめらかアニメーション駆動役。カーソル移動・色変化・
            スライドインなど、このウィザードのアニメーション全般で使う。"""
            steps = max(1, duration_ms // 15)

            def frame(i):
                if not canvas.winfo_exists():
                    return
                t = i / steps
                eased = 1 - (1 - t) ** 3
                on_tick(eased)
                if i < steps:
                    canvas.after(15, lambda: frame(i + 1))
                elif on_done:
                    on_done()

            frame(0)

        def make_cursor(canvas):
            """Windowsの標準矢印カーソルに近い形・配色(白地に黒縁)のアイコン。
            set_posで即座に配置、move_toでなめらかに(tween経由で)移動できる。"""
            cur = canvas.create_polygon(
                0, 0, 0, 19, 5, 15, 8, 21.5, 11.5, 20, 7.5, 14.5, 13, 14.5,
                fill="#ffffff", outline="#000000", width=1.2,
            )
            canvas.itemconfig(cur, state="hidden")

            def set_pos(x, y):
                canvas.coords(
                    cur,
                    x, y, x, y + 19, x + 5, y + 15, x + 8, y + 21.5,
                    x + 11.5, y + 20, x + 7.5, y + 14.5, x + 13, y + 14.5,
                )
                # ドロップダウンなど、カーソル生成後に新しく描かれる要素が
                # 上に重なって隠してしまわないよう、動くたびに最前面へ。
                canvas.tag_raise(cur)

            def move_to(from_xy, to_xy, duration_ms, on_done=None):
                if not canvas.winfo_exists():
                    return
                canvas.itemconfig(cur, state="normal")

                def on_tick(e):
                    set_pos(from_xy[0] + (to_xy[0] - from_xy[0]) * e,
                            from_xy[1] + (to_xy[1] - from_xy[1]) * e)

                tween(canvas, duration_ms, on_tick, on_done)

            def hide():
                if canvas.winfo_exists():
                    canvas.itemconfig(cur, state="hidden")

            return set_pos, move_to, hide

        def make_settings_illustration(canvas):
            """ステップ1: 実際の設定画面のミニチュアの上で、マウスカーソルが
            なめらかに動きながら「デバイスを選ぶ→ホットキーを登録する→
            Multi-Deviceタブに切り替える→有効化する→保存する」という一連の
            操作を実演する。使った時どうなるかではなく、この設定画面自体を
            どう操作するかを見せるためのイラスト。"""
            W = 420
            cx = W / 2

            tab_w, tab_h, tab_gap = 150, 22, 6
            tx0 = cx - (tab_w * 2 + tab_gap) / 2
            tab_ab = _rounded_rect(canvas, tx0, 8, tx0 + tab_w, 8 + tab_h, 6, fill=COLOR_ACCENT, outline="")
            canvas.create_text(tx0 + tab_w / 2, 8 + tab_h / 2, text="A/B SWITCH",
                                fill=COLOR_BG, font=("Segoe UI", 8, "bold"))
            tab_multi = _rounded_rect(canvas, tx0 + tab_w + tab_gap, 8, tx0 + tab_w * 2 + tab_gap, 8 + tab_h, 6,
                                       fill=COLOR_INPUT, outline="")
            tab_multi_text = canvas.create_text(tx0 + tab_w + tab_gap + tab_w / 2, 8 + tab_h / 2, text="MULTI-DEVICE",
                                                 fill=COLOR_TEXT_SUB, font=("Segoe UI", 8, "bold"))

            _rounded_rect(canvas, 10, 36, W - 10, 250, 12, fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)

            def field(label_text, y, is_combo=False, placeholder="Select device..."):
                lbl = canvas.create_text(28, y, anchor="w", text=label_text,
                                          fill=COLOR_TEXT_SUB, font=("Segoe UI", 8, "bold"))
                box = _rounded_rect(canvas, 28, y + 8, W - 28, y + 34, 6,
                                     fill=COLOR_INPUT, outline=COLOR_BORDER, width=1)
                box_text = canvas.create_text(40, y + 21, anchor="w", text=placeholder,
                                               fill=COLOR_TEXT_FAINT, font=("Segoe UI", 9))
                items = [lbl, box, box_text]
                if is_combo:
                    caret = canvas.create_text(W - 40, y + 21, text="▾", fill=COLOR_TEXT_SUB, font=("Segoe UI", 9))
                    items.append(caret)
                return items, box_text, (28, y + 8, W - 28, y + 34)

            ab_items_a, box_text_a, box_a_coords = field("DEVICE LAYER A", 50, is_combo=True)
            ab_items_b, box_text_b, box_b_coords = field("DEVICE LAYER B", 96, is_combo=True)
            ab_items_key, box_text_key, box_key_coords = field("GLOBAL TRIGGER HOTKEY", 142, placeholder="ctrl+alt+a")
            box_a = ab_items_a[1]
            box_b = ab_items_b[1]
            box_key = ab_items_key[1]
            canvas.itemconfig(box_text_key, fill=COLOR_TEXT_MAIN, font=("Segoe UI", 9, "bold"))
            ab_group = ab_items_a + ab_items_b + ab_items_key

            toggle_track = _rounded_rect(canvas, 28, 54, 64, 70, 8, fill=COLOR_INPUT, outline=COLOR_BORDER, width=1)
            toggle_knob = canvas.create_oval(29, 55, 43, 69, fill=COLOR_TEXT_MAIN, outline="")
            toggle_label = canvas.create_text(72, 62, anchor="w", text="Enable device-select overlay",
                                               fill=COLOR_TEXT_SUB, font=("Segoe UI", 8))
            okey_items, box_text_okey, box_okey_coords = field("OVERLAY HOTKEY", 96, placeholder="ctrl+alt+s")
            box_okey = okey_items[1]
            canvas.itemconfig(box_text_okey, fill=COLOR_TEXT_MAIN, font=("Segoe UI", 9, "bold"))
            multi_group = [toggle_track, toggle_knob, toggle_label] + okey_items
            for item in multi_group:
                canvas.itemconfig(item, state="hidden")

            btn = _rounded_rect(canvas, W - 130, 216, W - 28, 244, 8, fill=COLOR_ACCENT, outline="")
            canvas.create_text(W - 79, 230, text="Save & Run", fill=COLOR_BG, font=("Segoe UI", 9, "bold"))

            set_cursor, move_cursor, hide_cursor = make_cursor(canvas)
            pos = {"xy": (W + 20, -10)}

            def flash(item, base, accent):
                def on_tick(e):
                    if not canvas.winfo_exists():
                        return
                    ratio = e * 2 if e < 0.5 else (1 - e) * 2
                    canvas.itemconfig(item, fill=_lerp_hex(base, accent, ratio))
                tween(canvas, 300, on_tick)

            def go(target_xy, on_done):
                move_cursor(pos["xy"], target_xy, 420, on_done)
                pos["xy"] = target_xy

            def open_dropdown_and_select(box_coords, box_text_item, options, choice_index, on_done):
                """実際のCTkComboBoxと同じく、枠をクリック→プルダウンが
                現れる→カーソルが降りて選択肢をクリック→枠の表示が更新される、
                という一連の操作をなめらかに実演する。"""
                x1, y1, x2, y2 = box_coords
                row_h = 24
                dd_top = y2 + 3
                dd_h = row_h * len(options) + 6
                dd_panel = _rounded_rect(canvas, x1, dd_top, x2, dd_top + dd_h, 6,
                                          fill=COLOR_INPUT, outline=COLOR_BORDER, width=1)
                dd_items = [dd_panel]
                dd_rows = []
                for i, opt in enumerate(options):
                    ry = dd_top + 3 + i * row_h
                    row_bg = canvas.create_rectangle(x1 + 2, ry, x2 - 2, ry + row_h - 2, fill=COLOR_INPUT, outline="")
                    row_text = canvas.create_text(x1 + 14, ry + (row_h - 2) / 2, anchor="w", text=opt,
                                                   fill=COLOR_TEXT_MAIN, font=("Segoe UI", 9))
                    dd_rows.append((row_bg, row_text))
                    dd_items.extend([row_bg, row_text])

                target_row_bg, _ = dd_rows[choice_index]
                target_xy = (x1 + 20, dd_top + 3 + choice_index * row_h + (row_h - 2) / 2 - 8)

                def highlight_and_pick():
                    def on_tick(e):
                        if not canvas.winfo_exists():
                            return
                        canvas.itemconfig(target_row_bg, fill=_lerp_hex(COLOR_INPUT, COLOR_ACCENT_DIM, e))
                    tween(canvas, 200, on_tick, finish)

                def finish():
                    canvas.itemconfig(box_text_item, text=options[choice_index], fill=COLOR_TEXT_MAIN)
                    for item in dd_items:
                        canvas.delete(item)
                    on_done()

                canvas.after(200, lambda: go(target_xy, highlight_and_pick))

            def start():
                set_cursor(*pos["xy"])
                canvas.after(500, lambda: go((cx, 50 + 21), click_device_a))

            def click_device_a():
                flash(box_a, COLOR_INPUT, COLOR_ACCENT_DIM)
                canvas.after(200, lambda: open_dropdown_and_select(
                    box_a_coords, box_text_a, ["Speakers", "Headphones", "Monitor Speakers"], 0,
                    lambda: canvas.after(300, lambda: go((cx, 96 + 21), click_device_b))
                ))

            def click_device_b():
                flash(box_b, COLOR_INPUT, COLOR_ACCENT_DIM)
                canvas.after(200, lambda: open_dropdown_and_select(
                    box_b_coords, box_text_b, ["Speakers", "Headphones", "Monitor Speakers"], 1,
                    lambda: canvas.after(300, lambda: go((cx, 142 + 21), click_hotkey))
                ))

            def click_hotkey():
                flash(box_key, COLOR_INPUT, COLOR_ACCENT_DIM)
                canvas.after(450, lambda: go((tx0 + tab_w + tab_gap + tab_w / 2, 8 + tab_h / 2), click_tab_multi))

            def click_tab_multi():
                def on_tick(e):
                    if not canvas.winfo_exists():
                        return
                    canvas.itemconfig(tab_ab, fill=_lerp_hex(COLOR_ACCENT, COLOR_INPUT, e))
                    canvas.itemconfig(tab_multi, fill=_lerp_hex(COLOR_INPUT, COLOR_ACCENT, e))
                    canvas.itemconfig(tab_multi_text, fill=_lerp_hex(COLOR_TEXT_SUB, COLOR_BG, e))
                tween(canvas, 250, on_tick)
                for item in ab_group:
                    canvas.itemconfig(item, state="hidden")
                for item in multi_group:
                    canvas.itemconfig(item, state="normal")
                canvas.after(400, lambda: go((46, 62), click_toggle))

            def click_toggle():
                def on_tick(e):
                    if not canvas.winfo_exists():
                        return
                    canvas.itemconfig(toggle_track, fill=_lerp_hex(COLOR_INPUT, COLOR_ACCENT, e))
                    canvas.coords(toggle_knob, 29 + 14 * e, 55, 43 + 14 * e, 69)
                tween(canvas, 250, on_tick)
                canvas.after(400, lambda: go((cx, 96 + 21), click_overlay_hotkey))

            def click_overlay_hotkey():
                flash(box_okey, COLOR_INPUT, COLOR_ACCENT_DIM)
                canvas.after(450, lambda: go((W - 79, 230), click_save))

            def click_save():
                def on_tick(e):
                    if not canvas.winfo_exists():
                        return
                    ratio = e * 2 if e < 0.5 else (1 - e) * 2
                    canvas.itemconfig(btn, fill=_lerp_hex(COLOR_ACCENT, COLOR_ACCENT_HOVER, ratio))
                tween(canvas, 200, on_tick)
                canvas.after(900, reset_loop)

            def reset_loop():
                if not canvas.winfo_exists():
                    return
                hide_cursor()
                for item in multi_group:
                    canvas.itemconfig(item, state="hidden")
                for item in ab_group:
                    canvas.itemconfig(item, state="normal")
                canvas.itemconfig(box_text_a, text="Select device...", fill=COLOR_TEXT_FAINT)
                canvas.itemconfig(box_text_b, text="Select device...", fill=COLOR_TEXT_FAINT)
                canvas.itemconfig(tab_ab, fill=COLOR_ACCENT)
                canvas.itemconfig(tab_multi, fill=COLOR_INPUT)
                canvas.itemconfig(tab_multi_text, fill=COLOR_TEXT_SUB)
                canvas.itemconfig(toggle_track, fill=COLOR_INPUT)
                canvas.coords(toggle_knob, 29, 55, 43, 69)
                pos["xy"] = (W + 20, -10)
                canvas.after(500, start)

            canvas.after(500, start)

        DESKTOP_WALLPAPER = "#233240"
        DESKTOP_TASKBAR = "#16212b"

        def make_desktop_mockup(canvas, x1, y1, x2, y2):
            """トースト/オーバーレイが実際にはデスクトップ画面のどこに出るかを
            示すための、ミニデスクトップ(壁紙+タスクバー)のモックアップ。"""
            desktop_bg = canvas.create_rectangle(x1, y1, x2, y2, fill=DESKTOP_WALLPAPER, outline=COLOR_BORDER, width=1)
            taskbar_h = 16
            taskbar = canvas.create_rectangle(x1, y2 - taskbar_h, x2, y2, fill=DESKTOP_TASKBAR, outline="")
            canvas.create_oval(x1 + 6, y2 - taskbar_h + 3, x1 + 16, y2 - 3, fill=COLOR_ACCENT, outline="")
            return desktop_bg, taskbar, taskbar_h

        def make_ab_illustration(canvas):
            """ステップ2: CTRL+ALT+Aを押す→実際の画面と同じく、デスクトップの
            右下(タスクバーの上)にトースト通知が現れる→デバイス名を入れ替えて
            ループ。"""
            cx = 210
            keys = draw_keycaps(canvas, ["CTRL", "ALT", "A"], cx, 22)

            desk_x1, desk_y1, desk_x2, desk_y2 = 20, 46, 400, 250
            _, _, taskbar_h = make_desktop_mockup(canvas, desk_x1, desk_y1, desk_x2, desk_y2)

            TOAST_W, TOAST_H = 190, 50
            tx = desk_x2 - TOAST_W - 10
            final_ty = desk_y2 - taskbar_h - TOAST_H - 8
            slide_offset = 14
            panel = _rounded_rect(canvas, tx, final_ty, tx + TOAST_W, final_ty + TOAST_H, 10,
                                   fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
            accent_bar = canvas.create_rectangle(tx, final_ty + 7, tx + 3, final_ty + TOAST_H - 7,
                                                  fill=COLOR_ACCENT, outline="")
            header = canvas.create_text(tx + 13, final_ty + 15, anchor="w", text="AUDIO ROUTING SWITCHED",
                                         fill=COLOR_ACCENT, font=("Segoe UI", 6, "bold"))
            device_text = canvas.create_text(tx + 13, final_ty + 31, anchor="w", text="Headphones",
                                              fill=COLOR_TEXT_MAIN, font=("Segoe UI", 9, "bold"))
            toast_items = (panel, accent_bar, header, device_text)
            for item in toast_items:
                canvas.itemconfig(item, state="hidden")

            state = {"device": "Speakers"}

            def press():
                if not canvas.winfo_exists():
                    return
                for rect, text in keys:
                    canvas.itemconfig(rect, fill=COLOR_ACCENT)
                    canvas.itemconfig(text, fill=COLOR_BG)
                canvas.after(160, release_and_show_toast)

            def release_and_show_toast():
                if not canvas.winfo_exists():
                    return
                for rect, text in keys:
                    canvas.itemconfig(rect, fill=COLOR_INPUT)
                    canvas.itemconfig(text, fill=COLOR_TEXT_MAIN)
                state["device"] = "Headphones" if state["device"] == "Speakers" else "Speakers"
                canvas.itemconfig(device_text, text=state["device"])
                # panelは角丸(_rounded_rect)のポリゴンで、通常の矩形より
                # 座標の個数が多い。canvas.coords()で絶対座標を与え直すと
                # 頂点数が合わず形が壊れて斜線が入るため、canvas.move()で
                # 差分だけ動かす(頂点数に関係なく形を保ったまま平行移動できる)。
                # まず表示直後に下側へ退避させ、そこからなめらかに元の位置へ。
                for item in toast_items:
                    canvas.itemconfig(item, state="normal")
                    canvas.move(item, 0, slide_offset)

                prev_dy = {"v": slide_offset}

                def on_tick_slide(e):
                    if not canvas.winfo_exists():
                        return
                    dy = slide_offset * (1 - e)
                    delta = dy - prev_dy["v"]
                    for item in toast_items:
                        canvas.move(item, 0, delta)
                    prev_dy["v"] = dy

                tween(canvas, 220, on_tick_slide, lambda: canvas.after(1500, hide_toast))

            def hide_toast():
                if not canvas.winfo_exists():
                    return
                for item in toast_items:
                    canvas.itemconfig(item, state="hidden")
                canvas.after(700, press)

            canvas.after(700, press)

        def make_overlay_illustration(canvas):
            """ステップ3: CTRL+ALT+Sを押す→実際のMulti-Deviceオーバーレイと
            同じく、デスクトップ全体が暗くなり中央にカード(「SELECT OUTPUT
            DEVICE」の見出し・番号付きの行)がスライドイン→マウスカーソルが
            降りてきて1行をクリックして選ぶ、を繰り返しループ。"""
            cx = 210
            keys = draw_keycaps(canvas, ["CTRL", "ALT", "S"], cx, 22)

            desk_x1, desk_y1, desk_x2, desk_y2 = 20, 46, 400, 250
            desktop_bg, taskbar, _ = make_desktop_mockup(canvas, desk_x1, desk_y1, desk_x2, desk_y2)
            dim_color = _lerp_hex(DESKTOP_WALLPAPER, "#000000", 0.55)
            dim_taskbar_color = _lerp_hex(DESKTOP_TASKBAR, "#000000", 0.55)

            CARD_W, ROW_H, HEADER_H = 240, 28, 28
            labels = ["Speakers", "Headphones", "Monitor Speakers"]
            card_h = HEADER_H + len(labels) * ROW_H + 8
            card_x = desk_x1 + ((desk_x2 - desk_x1) - CARD_W) / 2
            final_card_y = desk_y1 + ((desk_y2 - desk_y1) - card_h) / 2
            slide_offset = 14

            card_panel = _rounded_rect(canvas, card_x, final_card_y, card_x + CARD_W, final_card_y + card_h, 12,
                                        fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
            card_header = canvas.create_text(card_x + 14, final_card_y + HEADER_H / 2, anchor="w",
                                              text="SELECT OUTPUT DEVICE", fill=COLOR_TEXT_SUB,
                                              font=("Segoe UI", 7, "bold"))
            card_divider = canvas.create_line(card_x + 8, final_card_y + HEADER_H, card_x + CARD_W - 8,
                                               final_card_y + HEADER_H, fill=COLOR_BORDER)
            card_items = [card_panel, card_header, card_divider]

            rows = []
            for i, label in enumerate(labels):
                y0 = final_card_y + HEADER_H + i * ROW_H
                row_bg = _rounded_rect(canvas, card_x + 6, y0 + 3, card_x + CARD_W - 6, y0 + ROW_H - 3, 8,
                                        fill=COLOR_PANEL, outline="")
                text = canvas.create_text(card_x + 18, y0 + ROW_H / 2, anchor="w", text=label,
                                           fill=COLOR_TEXT_SUB, font=("Segoe UI", 8))
                num = canvas.create_text(card_x + CARD_W - 14, y0 + ROW_H / 2, anchor="e", text=str(i + 1),
                                          fill=COLOR_TEXT_FAINT, font=("Segoe UI", 8))
                rows.append((row_bg, text, num))
                card_items.extend([row_bg, text, num])

            for item in card_items:
                canvas.itemconfig(item, state="hidden")

            set_cursor, move_cursor, hide_cursor = make_cursor(canvas)

            def press():
                if not canvas.winfo_exists():
                    return
                for rect, text in keys:
                    canvas.itemconfig(rect, fill=COLOR_ACCENT)
                    canvas.itemconfig(text, fill=COLOR_BG)
                canvas.after(160, release_and_show_card)

            def release_and_show_card():
                if not canvas.winfo_exists():
                    return
                for rect, text in keys:
                    canvas.itemconfig(rect, fill=COLOR_INPUT)
                    canvas.itemconfig(text, fill=COLOR_TEXT_MAIN)
                # card_panelやrow_bgは角丸(_rounded_rect)のポリゴンで、通常の
                # 矩形より座標の個数が多い。canvas.coords()で絶対座標を与え
                # 直すと頂点数が合わず形が壊れて斜線が入るため、canvas.move()
                # で差分だけ動かす(頂点数に関係なく形を保ったまま平行移動)。
                # まず表示直後に下側へ退避させ、そこからなめらかに元の位置へ。
                for item in card_items:
                    canvas.itemconfig(item, state="normal")
                    canvas.move(item, 0, slide_offset)

                row_ys = [final_card_y + HEADER_H + i * ROW_H for i in range(len(labels))]
                prev_dy = {"v": slide_offset}

                def on_tick(e):
                    if not canvas.winfo_exists():
                        return
                    canvas.itemconfig(desktop_bg, fill=_lerp_hex(DESKTOP_WALLPAPER, dim_color, e))
                    canvas.itemconfig(taskbar, fill=_lerp_hex(DESKTOP_TASKBAR, dim_taskbar_color, e))
                    dy = slide_offset * (1 - e)
                    delta = dy - prev_dy["v"]
                    for item in card_items:
                        canvas.move(item, 0, delta)
                    prev_dy["v"] = dy

                def after_slide():
                    set_cursor(card_x + CARD_W - 30, final_card_y - 20)
                    move_cursor((card_x + CARD_W - 30, final_card_y - 20),
                                (card_x + 50, row_ys[1] + ROW_H / 2 - 8), 380,
                                lambda: canvas.after(150, lambda: select_row(1)))

                tween(canvas, 220, on_tick, after_slide)

            def select_row(row_index):
                if not canvas.winfo_exists():
                    return
                for i, (rect, text, num) in enumerate(rows):
                    is_sel = i == row_index
                    canvas.itemconfig(rect, fill=(COLOR_ACCENT_DIM if is_sel else COLOR_PANEL))
                    canvas.itemconfig(text, fill=(COLOR_ACCENT if is_sel else COLOR_TEXT_SUB))
                canvas.after(1000, reset_loop)

            def reset_loop():
                if not canvas.winfo_exists():
                    return
                hide_cursor()
                for item in card_items:
                    canvas.itemconfig(item, state="hidden")
                for rect, text, num in rows:
                    canvas.itemconfig(rect, fill=COLOR_PANEL)
                    canvas.itemconfig(text, fill=COLOR_TEXT_SUB)
                canvas.itemconfig(desktop_bg, fill=DESKTOP_WALLPAPER)
                canvas.itemconfig(taskbar, fill=DESKTOP_TASKBAR)
                canvas.after(700, press)

            canvas.after(700, press)

        steps = [
            {
                "title": "How to set it up",
                "body": "Pick your devices, set a hotkey, switch tabs to enable\nMulti-Device, and Save & Run — like this:",
                "illustration": make_settings_illustration,
            },
            {
                "title": "Instant A/B toggle",
                "body": "Once it's set up, pressing the hotkey you just\nconfigured instantly swaps between your two devices:",
                "illustration": make_ab_illustration,
            },
            {
                "title": "Switch between 3+ devices",
                "body": "With Multi-Device enabled, pressing its hotkey brings up\nevery connected device for you to pick from:",
                "illustration": make_overlay_illustration,
            },
            {
                "title": "You're all set",
                "body": "Pick your devices and hotkeys below, then Save & Run.",
                "illustration": None,
            },
        ]

        def close_wizard():
            if wiz["window"] is not None:
                try:
                    wiz["window"].destroy()
                except Exception:
                    pass
                wiz["window"] = None

        def advance():
            if wiz["step"] >= len(steps) - 1:
                close_wizard()
                return
            wiz["step"] += 1
            render()

        def go_back():
            if wiz["step"] > 0:
                wiz["step"] -= 1
                render()

        def render():
            for w in body_frame.winfo_children():
                w.destroy()
            step = steps[wiz["step"]]

            ctk.CTkLabel(body_frame, text=step["title"], font=FONT_DISPLAY,
                         text_color=COLOR_TEXT_MAIN, anchor="w").pack(fill="x", padx=32, pady=(26, 10))
            ctk.CTkLabel(body_frame, text=step["body"], font=FONT_BODY, text_color=COLOR_TEXT_SUB,
                         justify="left", anchor="w").pack(fill="x", padx=32, pady=(0, 18))

            illo_canvas = tk.Canvas(body_frame, width=420, height=260, bg=COLOR_PANEL, highlightthickness=0)
            illo_canvas.pack(pady=(0, 18))
            if step["illustration"]:
                step["illustration"](illo_canvas)

            n = len(steps)
            active_w, inactive_w, dot_h, gap = 20, 7, 7, 6
            dots_w = active_w + inactive_w * (n - 1) + gap * (n - 1) + 14
            dots_canvas = tk.Canvas(body_frame, width=dots_w, height=dot_h + 4, bg=COLOR_PANEL, highlightthickness=0)
            dots_canvas.pack(pady=(0, 18))
            dx = 4
            for i in range(n):
                is_active = i == wiz["step"]
                w = active_w if is_active else inactive_w
                color = COLOR_ACCENT if is_active else COLOR_ACCENT_DIM
                _rounded_rect(dots_canvas, dx, 2, dx + w, 2 + dot_h, dot_h / 2, fill=color, outline="")
                dx += w + gap

            bottom = ctk.CTkFrame(body_frame, fg_color="transparent")
            bottom.pack(fill="x", padx=32, pady=(0, 26))
            ctk.CTkButton(
                bottom, text="Skip", command=close_wizard,
                fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
                border_width=1, border_color=COLOR_BORDER, corner_radius=8,
                font=FONT_LABEL, width=90, height=36,
            ).pack(side="left")

            right_frame = ctk.CTkFrame(bottom, fg_color="transparent")
            right_frame.pack(side="right")
            if wiz["step"] > 0:
                ctk.CTkButton(
                    right_frame, text="Back", command=go_back,
                    fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
                    border_width=1, border_color=COLOR_BORDER, corner_radius=8,
                    font=FONT_LABEL, width=90, height=36,
                ).pack(side="left", padx=(0, 10))
            is_last = wiz["step"] == len(steps) - 1
            ctk.CTkButton(
                right_frame, text=("Get Started" if is_last else "Next"), command=advance,
                fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
                corner_radius=8, font=FONT_LABEL, width=120, height=36,
            ).pack(side="left")

        WIZ_W = 492

        win = ctk.CTkToplevel(root)
        win.title("Welcome to AudioSwitcher")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.transient(root)
        # 閉じる方法は下部の"Skip"だけにしたいので、タイトルバーの×では
        # 閉じられないようにする(WM_DELETE_WINDOWを無視するだけにする)。
        win.protocol("WM_DELETE_WINDOW", lambda: None)
        wiz["window"] = win

        body_frame = ctk.CTkFrame(win, fg_color=COLOR_PANEL, corner_radius=0)
        body_frame.pack(fill="both", expand=True)

        render()

        # ウィンドウサイズは実際に組み終えた中身の実測サイズに合わせる
        # (イラストの大きさを変えるたびに固定値を数え直さなくて済むように)。
        win.update_idletasks()
        wiz_h = win.winfo_reqheight() + 12
        root.update_idletasks()
        rx, ry = root.winfo_x(), root.winfo_y()
        rw, rh = root.winfo_width(), root.winfo_height()
        win.geometry(f"{WIZ_W}x{wiz_h}+{rx + (rw - WIZ_W) // 2}+{ry + (rh - wiz_h) // 2}")

    if is_first_run:
        show_onboarding()

    root.mainloop()

if __name__ == "__main__":
    enable_dpi_awareness()
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