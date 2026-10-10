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
from pycaw.api.policyconfig import IPolicyConfig
from pycaw.constants import DEVICE_STATE, EDataFlow, ERole, CLSID_CPolicyConfigClient

# customtkinter（設定画面専用のUIライブラリ）はここでは読み込まない。
# バックグラウンド常駐時はホットキー監視しかしないため、
# 使わない重いUIアセットをメモリに乗せないよう open_setting_window() 内で遅延importする。

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "app_config.txt")
ICON_PATH = os.path.join(BASE_DIR, "app_icon.ico")
# exe化した時はapp_icon.icoをexeの隣に置かないため、exeに同梱した物を使う
# (タイトルバーのアイコンも、アプリアイコンと同じヘッドホンにするため)。
if not os.path.exists(ICON_PATH) and hasattr(sys, "_MEIPASS"):
    ICON_PATH = os.path.join(sys._MEIPASS, "app_icon.ico")

__version__ = "1.7.0"
GITHUB_REPO = "CAIN218/audio-switcher"

DEFAULT_OVERLAY_HOTKEY = "ctrl+alt+s"

last_switch_time = 0
_com_thread_local = threading.local()

# ----------------------------------------------------------------------
# デザイントークン — Yuga's lab のアプリアイコン(暗い背景・白い線・差し色1色)
# に合わせた、黒鉛色×カッパー(銅)1色のシンプルな配色。設定画面とトースト通知の
# 両方で共有する。
# ----------------------------------------------------------------------
COLOR_BG          = "#18171c"  # アプリ全体の背景(アイコンの背景色と同じ)
COLOR_PANEL       = "#201f25"  # カードパネル
COLOR_INPUT       = "#2a282f"  # 入力欄・コンボボックス
COLOR_BORDER      = "#36333c"  # 通常の境界線
COLOR_ACCENT      = "#d08a52"  # カッパー(アイコンの差し色と同じ)
COLOR_ACCENT_HOVER = "#e0a06c"
COLOR_ACCENT_DIM  = "#5a4331"  # アクセントの控えめ版(非アクティブ表示用)
COLOR_TEXT_MAIN   = "#ece8e1"  # アイコンの白線に合わせたオフホワイト
COLOR_TEXT_SUB    = "#a29c90"
COLOR_TEXT_FAINT  = "#68645c"

# 画面には日本語と英語が混ざる。Segoe UIには日本語の字形が無く、日本語部分
# だけ別のフォントに置き換わって太さや高さがバラバラに見えていたため、
# 日本語・英数字の両方を同じ字形で描ける「Yu Gothic UI」(Windows 10/11標準)
# を優先して使う。無い環境では順に代わりのフォントを使う。
FONT_CANDIDATES = ("Yu Gothic UI", "Meiryo UI", "Segoe UI", "Noto Sans CJK JP")
FONT_FAMILY = FONT_CANDIDATES[0]


def F(size, weight="normal"):
    """UI全体で使うフォント指定を返す(ファミリーはresolve_font_familyで決まる)"""
    return (FONT_FAMILY, size, weight)


FONT_DISPLAY = F(16, "bold")
FONT_TITLE   = F(13, "bold")
FONT_LABEL   = F(10, "bold")
FONT_BODY    = F(10)
FONT_SMALL   = F(9)
FONT_MONO    = F(11, "bold")


def resolve_font_family(root):
    """インストールされているフォントを見て、使うファミリーを決める。
    Tkのルートが無いとフォント一覧を取れないので、ルート作成直後に呼ぶ。"""
    global FONT_FAMILY, FONT_DISPLAY, FONT_TITLE, FONT_LABEL, FONT_BODY, FONT_SMALL, FONT_MONO
    try:
        installed = set(tkfont.families(root))
    except Exception:
        installed = set()
    for name in FONT_CANDIDATES:
        if name in installed:
            FONT_FAMILY = name
            break
    FONT_DISPLAY = F(16, "bold")
    FONT_TITLE   = F(13, "bold")
    FONT_LABEL   = F(10, "bold")
    FONT_BODY    = F(10)
    FONT_SMALL   = F(9)
    FONT_MONO    = F(11, "bold")


def draw_headphones(canvas, cx, cy, size, color, width=None):
    """サイトのアプリアイコンと同じヘッドホンの図をCanvasに描く。
    元のSVGは幅168・高さ154の図形なので、sizeに合わせて縮小して描く。"""
    s = size / 168.0
    lw = max(1.5, (width or 16) * s)
    r = 70 * s
    oy = cy + 1 * s  # 図形の上下中心を合わせる補正
    items = [
        canvas.create_arc(cx - r, oy - r, cx + r, oy + r, start=0, extent=180,
                          style="arc", outline=color, width=lw),
        canvas.create_line(cx - r, oy, cx - r, oy + 20 * s, fill=color, width=lw, capstyle="round"),
        canvas.create_line(cx + r, oy, cx + r, oy + 20 * s, fill=color, width=lw, capstyle="round"),
    ]
    for x0 in (-84, 44):
        items.append(_rounded_rect(canvas, cx + x0 * s, oy + 6 * s, cx + (x0 + 40) * s, oy + 76 * s,
                                   14 * s, fill=color, outline=""))
    return items


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

def set_default_device_all_roles(target_id):
    """既定の再生デバイスを、eConsole/eMultimedia/eCommunicationsの全ロール
    に対して切り替える(SoundVolumeViewの"all"相当)。pycawのSetDefaultDevice
    はroles未指定だとeConsoleしか設定しない(ドキュメントやコード内の古い
    コメントとは異なる実際の挙動)ため、明示的に3ロール全てを渡す必要が
    ある。eCommunicationsを設定しないと、Discordのように「既定の通信
    デバイス」を優先して追従するアプリが切り替わらない。
    各ロールのSetDefaultEndpoint呼び出しは1回あたり数十ms かかる(COM
    オブジェクトの生成コストではなく、Windows Audioサービス側の処理自体が
    重い)ことが実測でわかったため、逐次3回呼ぶと切り替えの体感遅延が目立つ。
    各ロールを別スレッド(別COMアパートメント)で並列に呼び、合計の実時間を
    縮める。"""
    errors = []

    def _set_role(role):
        comtypes.CoInitialize()
        try:
            policy_config = comtypes.CoCreateInstance(
                CLSID_CPolicyConfigClient, IPolicyConfig, comtypes.CLSCTX_ALL
            )
            policy_config.SetDefaultEndpoint(target_id, role.value)
        except Exception as e:
            errors.append(e)
        finally:
            comtypes.CoUninitialize()

    threads = [threading.Thread(target=_set_role, args=(role,))
               for role in (ERole.eConsole, ERole.eMultimedia, ERole.eCommunications)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    if errors:
        raise errors[0]


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
    既存ユーザーの挙動を変えないようにしている。自動更新は、インストール
    直後に(既にそれより新しいリリースが存在していた場合)開いた瞬間に
    無言で更新・再起動されると初回体験として驚かせてしまうため、既定は
    "0"(無効)にしている。"""
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
        auto_update_enabled = lines[7] if len(lines) > 7 else "0"
        return dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled, ab_enabled, auto_update_enabled
    return "", "", "ctrl+alt+a", "0", DEFAULT_OVERLAY_HOTKEY, "0", "1", "0"

def load_overlay_excluded():
    """Multi-Deviceの一覧に「出さない」デバイス名のリスト(設定ファイルの9行目、
    JSON)。出すデバイスではなく出さないデバイスを覚えておくことで、後から
    新しくつないだデバイスは何もしなくても一覧に出るようにしている。
    9行目が無い古い設定ファイルでは空(=全部出す)として扱う。"""
    if not os.path.exists(CONFIG_FILE):
        return []
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
        if len(lines) > 8 and lines[8].strip():
            value = json.loads(lines[8])
            if isinstance(value, list):
                return [str(v) for v in value]
    except Exception as e:
        log_error("load_overlay_excluded", e)
    return []

def save_config(dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled, ab_enabled, auto_update_enabled="0",
                overlay_excluded=None):
    """設定ファイルの保存とスタートアップ制御。overlay_excludedを渡さない時は
    今の設定ファイルにある値をそのまま残す。"""
    if overlay_excluded is None:
        overlay_excluded = load_overlay_excluded()
    excluded_json = json.dumps(list(overlay_excluded), ensure_ascii=False)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(f"{dev_a}\n{dev_b}\n{key}\n{startup_val}\n{overlay_key}\n{overlay_enabled}\n{ab_enabled}\n{auto_update_enabled}\n{excluded_json}")
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
    WHATSNEW_WIDTH, WHATSNEW_HEIGHT = 360, 260

    def __init__(self, root):
        self.root = root
        self.DEVICE_FONT = F(11, "bold")
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
        self.whatsnew_window = None
        self.whatsnew_canvas = None
        self.whatsnew_alpha = 0.0
        self.whatsnew_fade_job = None

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
        self._draw_toast_frame(canvas, "Output switched")

        # 長いデバイス名は「...」で切らず、収まりきらない分だけ左右に
        # ゆっくり往復させて全体を読めるようにする(Canvasは自身の外側を
        # 自動的に描画しないので、はみ出した部分は特別なクリップ処理
        # なしでそのまま隠れる)。
        left_x = self.TEXT_X
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

    TEXT_X = 60

    def _draw_toast_frame(self, canvas, header_text):
        """トースト共通の土台: 角丸パネル + 左にアプリアイコンと同じヘッドホン
        (どのアプリからの通知かがひと目で分かるように) + 小さな見出し"""
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 16,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        _rounded_rect(canvas, 12, 12, 48, self.HEIGHT - 12 - (self.HEIGHT - 68), 10,
                      fill=COLOR_INPUT, outline="")
        draw_headphones(canvas, 30, 34, 22, COLOR_ACCENT)
        canvas.create_text(self.TEXT_X, 23, anchor="w", text=header_text,
                           fill=COLOR_ACCENT, font=F(9, "bold"))

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

    def show_update_available(self, latest_tag, release_url, installer_url=None, notes=""):
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
        self._draw_toast_frame(canvas, "Update available")
        auto = self.auto_update_enabled and installer_url
        if auto:
            subtitle = f"{latest_tag} — installing automatically..."
        elif installer_url:
            subtitle = f"{latest_tag} — click to install"
        else:
            subtitle = f"{latest_tag} — click to view"
        self._update_subtitle_id = canvas.create_text(self.TEXT_X, 44, anchor="w", text=subtitle,
                                                        fill=COLOR_TEXT_MAIN, font=F(10, "bold"))

        self._pending_release_url = release_url
        if auto:
            canvas.configure(cursor="")
            self.on_click = None
        elif installer_url:
            canvas.configure(cursor="hand2")
            self.on_click = lambda: self._begin_update_download(installer_url, release_url, latest_tag, notes)
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
            self._begin_update_download(installer_url, release_url, latest_tag, notes)

    def _begin_update_download(self, installer_url, release_url, latest_tag="", notes=""):
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
                          args=(installer_url, release_url, latest_tag, notes), daemon=True).start()

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

    def _ensure_whatsnew_window(self):
        if self.whatsnew_window is not None:
            return
        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.attributes("-alpha", 0.0)

        transparent_key = "#010203"
        win.configure(bg=transparent_key)
        win.attributes("-transparentcolor", transparent_key)

        canvas = tk.Canvas(win, width=self.WHATSNEW_WIDTH, height=self.WHATSNEW_HEIGHT,
                            bg=transparent_key, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        canvas.bind("<Button-1>", lambda e: self._dismiss_whats_new())
        win.withdraw()

        self.whatsnew_window = win
        self.whatsnew_canvas = canvas

    def _place_whatsnew_bottom_right(self):
        try:
            left, top, right, bottom = _active_monitor_work_area()
        except Exception:
            left, top = 0, 0
            right = self.whatsnew_window.winfo_screenwidth()
            bottom = self.whatsnew_window.winfo_screenheight()
        height = getattr(self, "whatsnew_height", self.WHATSNEW_HEIGHT)
        x = right - self.WHATSNEW_WIDTH - 16
        y = bottom - height - 6
        self.whatsnew_window.geometry(f"{self.WHATSNEW_WIDTH}x{height}+{x}+{y}")

    def show_whats_new(self, version, notes):
        """アップデート直後に一度だけ、他のトーストと同じ見た目(角丸パネル・
        ヘッドホン・フェード)のポップアップとして変更内容を表示する。
        設定画面を開かせたくない、という指摘を受けて、通知の仕組みを流用
        する形にした。内容を読む時間が必要なので自動では消さず、クリックで
        閉じる。以前は高さが固定で、項目が多いと文字が下にはみ出して
        「Got it」と重なっていたため、中身に合わせて高さを決めるようにした。"""
        self._ensure_whatsnew_window()
        if self.whatsnew_fade_job:
            self.root.after_cancel(self.whatsnew_fade_job)
            self.whatsnew_fade_job = None

        canvas = self.whatsnew_canvas
        canvas.delete("all")
        W = self.WHATSNEW_WIDTH
        MAX_H = 420
        FOOTER_H = 44

        draw_headphones(canvas, 34, 34, 24, COLOR_ACCENT)
        canvas.create_text(58, 26, anchor="w", text="Updated", fill=COLOR_ACCENT, font=F(9, "bold"))
        canvas.create_text(58, 45, anchor="w", text=f"Now on {version}",
                           fill=COLOR_TEXT_MAIN, font=F(13, "bold"))
        canvas.create_line(18, 66, W - 18, 66, fill=COLOR_BORDER)

        # Markdownの見出し(##)と箇条書き(-)だけを簡単に見た目調整して並べる
        y = 78
        truncated = False
        body = (notes or "").strip() or "No release notes."
        for raw in body.splitlines():
            line = raw.strip()
            if not line:
                continue
            if line.startswith("## "):
                item = canvas.create_text(22, y + 6, anchor="nw", text=line[3:], fill=COLOR_ACCENT,
                                          font=F(9, "bold"))
            elif line.startswith("- "):
                dot = canvas.create_oval(24, y + 7, 29, y + 12, fill=COLOR_TEXT_SUB, outline="")
                item = canvas.create_text(38, y, anchor="nw", text=line[2:], fill=COLOR_TEXT_MAIN,
                                          font=F(9), width=W - 58)
            else:
                item = canvas.create_text(22, y, anchor="nw", text=line, fill=COLOR_TEXT_SUB,
                                          font=F(9), width=W - 44)
            bbox = canvas.bbox(item)
            bottom = bbox[3] if bbox else y + 16
            if bottom > MAX_H - FOOTER_H:
                # 入りきらない分は省略して、続きがあることだけ伝える
                canvas.delete(item)
                if line.startswith("- "):
                    canvas.delete(dot)
                canvas.create_text(22, y, anchor="nw", text="…ほかの変更は GitHub のリリースノートへ",
                                   fill=COLOR_TEXT_FAINT, font=F(9))
                y += 20
                truncated = True
                break
            y = bottom + 5
        H = min(MAX_H, y + FOOTER_H)

        footer_y = H - 24
        _rounded_rect(canvas, W - 98, footer_y - 13, W - 18, footer_y + 13, 8,
                      fill=COLOR_INPUT, outline="")
        canvas.create_text(W - 58, footer_y, text="Got it", fill=COLOR_TEXT_MAIN, font=F(9, "bold"))
        canvas.create_text(22, footer_y, anchor="w", text="クリックで閉じる",
                           fill=COLOR_TEXT_FAINT, font=F(8))
        panel = _rounded_rect(canvas, 1, 1, W - 1, H - 1, 16, fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        canvas.tag_lower(panel)
        canvas.configure(cursor="hand2", height=H)
        self.whatsnew_height = H

        self._place_whatsnew_bottom_right()
        self.whatsnew_window.deiconify()
        self._place_whatsnew_bottom_right()
        self.whatsnew_alpha = 0.0
        self._whatsnew_fade_in()

    def _dismiss_whats_new(self):
        if self.whatsnew_fade_job:
            self.root.after_cancel(self.whatsnew_fade_job)
            self.whatsnew_fade_job = None
        self._whatsnew_fade_out()

    def _whatsnew_fade_in(self):
        if self.whatsnew_alpha < 0.97:
            self.whatsnew_alpha += 0.12
            self.whatsnew_window.attributes("-alpha", min(self.whatsnew_alpha, 0.97))
            self.whatsnew_fade_job = self.root.after(15, self._whatsnew_fade_in)
        else:
            self.whatsnew_fade_job = None

    def _whatsnew_fade_out(self):
        if self.whatsnew_alpha > 0.0:
            self.whatsnew_alpha -= 0.08
            self.whatsnew_window.attributes("-alpha", max(self.whatsnew_alpha, 0.0))
            self.whatsnew_fade_job = self.root.after(15, self._whatsnew_fade_out)
        else:
            self.whatsnew_window.withdraw()
            self.whatsnew_fade_job = None


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
                            text="Select output device", fill=COLOR_TEXT_SUB,
                            font=F(10, "bold"))
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
                fill=text_color, font=F(10, "bold" if is_selected else "normal")
            )

            items = [row_bg, text_id]
            if i < 9:
                items.append(canvas.create_text(
                    self.WIDTH - self.PADDING, (y0 + y1) / 2, anchor="e",
                    text=str(i + 1), fill=COLOR_TEXT_FAINT, font=F(9)
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
    (latest_tag, release_url, installer_url, notes)を返し、そうでなければ
    Noneを返す。notesはリリースの説明文(Markdown、gh release createの
    --notesで書いた内容)で、アップデート後の「What's New」表示に使う。
    バックグラウンドの定期チェックと、設定画面を開くたびの即時チェックの
    両方から共通で使う。呼び出し側で例外を処理すること。"""
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
        headers={"Accept": "application/vnd.github+json", "User-Agent": "AudioSwitcher-update-check"},
    )
    with urllib.request.urlopen(req, timeout=5) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    latest_tag = data.get("tag_name", "")
    release_url = data.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases/latest"
    notes = data.get("body") or ""
    installer_url = None
    for asset in data.get("assets", []) or []:
        name = asset.get("name", "")
        if name.lower().startswith("audioswitcher-setup") and name.lower().endswith(".exe"):
            installer_url = asset.get("browser_download_url")
            break
    if latest_tag and _parse_version(latest_tag) > _parse_version(__version__):
        return (latest_tag, release_url, installer_url, notes)
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


UPDATE_NOTICE_FILE = os.path.join(BASE_DIR, "update_notice.json")


def _write_update_notice(latest_tag, notes):
    """再起動後の新バージョンが「What's New」を表示できるよう、更新前の
    バージョンが自分で情報を書き残しておく。再インストール後もこのファイル
    自体は([Files]/[UninstallDelete]どちらにも入っていないので)そのまま
    残る。書き込みに失敗しても更新自体は継続する(What's Newが出ないだけ)。"""
    try:
        with open(UPDATE_NOTICE_FILE, "w", encoding="utf-8") as f:
            json.dump({"version": latest_tag, "notes": notes}, f, ensure_ascii=False)
    except Exception as e:
        log_error("write_update_notice", e)


def _read_pending_update_notice():
    """起動時に一度だけ呼ぶ。前回の更新時に書き残されたUPDATE_NOTICE_FILEが
    あれば(version, notes)を返してすぐ削除する(次回からは出さない)。無ければ
    None。常駐モード(start_backend)・設定画面(open_setting_window)の両方の
    起動経路から共通で使う。"""
    if not os.path.exists(UPDATE_NOTICE_FILE):
        return None
    try:
        with open(UPDATE_NOTICE_FILE, "r", encoding="utf-8") as f:
            notice = json.load(f)
    except Exception as e:
        log_error("read_update_notice", e)
        notice = None
    try:
        os.remove(UPDATE_NOTICE_FILE)
    except Exception:
        pass
    if not notice:
        return None
    return notice.get("version", ""), notice.get("notes", "")


def _download_and_launch_installer(installer_url, release_url, latest_tag="", notes=""):
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
        _write_update_notice(latest_tag, notes)
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
                    set_default_device_all_roles(target_id)
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
            set_default_device_all_roles(target_id)
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

    overlay_excluded = set(load_overlay_excluded())

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
    resolve_font_family(root)
    root.report_callback_exception = lambda exc_type, exc_value, tb: log_error("tk_callback", exc_value)
    toast_manager = _ToastManager(root)
    toast_manager.auto_update_enabled = (auto_update_enabled == "1")
    device_overlay = _DeviceOverlay(root)

    def _on_overlay_select(name, target_id):
        _direct_switch_queue.put((target_id, name))
    device_overlay.on_select = _on_overlay_select

    pending_notice = _read_pending_update_notice()
    if pending_notice:
        version, notes = pending_notice
        root.after(800, lambda: toast_manager.show_whats_new(version, notes))

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
                latest_tag, release_url, installer_url, notes = _update_queue.get_nowait()
                toast_manager.show_update_available(latest_tag, release_url, installer_url, notes)
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
                # 設定画面でOFFにしたデバイスは一覧に出さない(全部OFFの時は全部出す)
                shown = {name: dev_id for name, dev_id in device_dict.items() if name not in overlay_excluded}
                device_overlay.open(shown or device_dict, current_default)
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
    resolve_font_family(root)
    # Settingsタブの中に置くスイッチだが、変数自体はSave時に他の設定と一緒に
    # 保存する必要があるので、タブの有無に関わらずここで(root作成後に)
    # 用意しておく(StringVarはデフォルトのTkルートが無いと作れないため)。
    auto_update_var = tk.StringVar(value=saved_auto_update)
    root.title("AudioSwitcher")
    root.geometry("560x600")
    root.resizable(False, False)
    if os.path.exists(ICON_PATH):
        try:
            root.iconbitmap(ICON_PATH)
        except Exception:
            pass

    def show_whats_new_overlay(version, notes):
        """自動更新(またはSettingsタブでの更新)の直後に一度だけ、何が変わったか
        を表示する。別のToplevelを座標合わせして重ねる方式は、設定画面との
        位置ズレやタスクバーへの別アイコン表示などが繰り返し起きたため、
        別ウィンドウを一切使わず、設定画面(root)と同じウィンドウの中に直接
        不透明なフレームを重ねて表示する(常にピクセル単位で一致する)。"""
        # 余白を残して浮いているカードに見せる(全面に敷き詰めない)。backdropが
        # rootの背景色のまま少し見えることで「カードが浮いている」感じになる。
        backdrop = ctk.CTkFrame(root, fg_color=COLOR_BG, corner_radius=0)
        backdrop.place(relx=0, rely=0, relwidth=1, relheight=1)

        overlay = ctk.CTkFrame(backdrop, fg_color=COLOR_PANEL, corner_radius=16,
                                border_width=1, border_color=COLOR_BORDER)
        overlay.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.88, relheight=0.88)

        def close_overlay():
            backdrop.destroy()

        header = ctk.CTkFrame(overlay, fg_color="transparent")
        header.pack(fill="x", padx=28, pady=(26, 6))
        head_icon = tk.Canvas(header, width=44, height=44, bg=COLOR_PANEL, highlightthickness=0)
        head_icon.pack(side="left", padx=(0, 14))
        _rounded_rect(head_icon, 0, 0, 44, 44, 12, fill=COLOR_INPUT, outline="")
        draw_headphones(head_icon, 22, 22, 28, COLOR_ACCENT)
        head_text = ctk.CTkFrame(header, fg_color="transparent")
        head_text.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(head_text, text="✓ アップデートしました", font=F(9, "bold"),
                     text_color=COLOR_ACCENT, anchor="w", height=18).pack(fill="x")
        ctk.CTkLabel(head_text, text=f"What's new in {version}", font=F(17, "bold"),
                     text_color=COLOR_TEXT_MAIN, anchor="w", justify="left", height=28).pack(fill="x")

        ctk.CTkFrame(overlay, fg_color=COLOR_BORDER, height=1, corner_radius=0).pack(fill="x", padx=28, pady=(8, 0))

        body = ctk.CTkScrollableFrame(overlay, fg_color="transparent",
                                      scrollbar_button_color=COLOR_BORDER,
                                      scrollbar_button_hover_color=COLOR_TEXT_FAINT)
        body.pack(fill="both", expand=True, padx=22, pady=(4, 6))

        for raw in (notes or "").strip().splitlines():
            line = raw.strip()
            if not line:
                continue
            if line.startswith("## "):
                ctk.CTkLabel(body, text=line[3:], font=F(12, "bold"),
                             text_color=COLOR_ACCENT, anchor="w", height=22).pack(fill="x", padx=6, pady=(14, 4))
            elif line.startswith("- "):
                row = ctk.CTkFrame(body, fg_color="transparent")
                row.pack(fill="x", padx=6, pady=3)
                ctk.CTkLabel(row, text="•", font=F(11, "bold"), text_color=COLOR_TEXT_SUB,
                             width=16, height=22, anchor="w").pack(side="left", anchor="n")
                ctk.CTkLabel(row, text=line[2:], font=F(11), text_color=COLOR_TEXT_MAIN, height=22,
                             anchor="w", justify="left", wraplength=400).pack(side="left", fill="x", expand=True)
            else:
                ctk.CTkLabel(body, text=line, font=FONT_BODY, text_color=COLOR_TEXT_SUB,
                             anchor="w", justify="left", wraplength=420).pack(fill="x", padx=6, pady=2)

        # 文字が収まりきる時は右のスクロールバーを隠す(はみ出た時だけ出す)。
        def update_scrollbar_visibility():
            try:
                body.update_idletasks()
                canvas = body._parent_canvas
                bbox = canvas.bbox("all")
                content_h = bbox[3] - bbox[1] if bbox else 0
                if content_h <= canvas.winfo_height():
                    body._scrollbar.grid_remove()
                else:
                    body._scrollbar.grid()
            except Exception:
                pass
        root.after(50, update_scrollbar_visibility)

        footer = ctk.CTkFrame(overlay, fg_color="transparent")
        footer.pack(fill="x", padx=28, pady=(6, 26))
        release_url = f"https://github.com/{GITHUB_REPO}/releases"
        ctk.CTkButton(
            footer, text="GitHub で全部見る", command=lambda: webbrowser.open(release_url),
            fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
            border_width=0, font=FONT_LABEL, width=130, height=38,
        ).pack(side="left")
        ctk.CTkButton(
            footer, text="Got it", command=close_overlay,
            fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
            corner_radius=8, font=FONT_LABEL, width=120, height=38,
        ).pack(side="right")

        backdrop.lift()

    # 自動更新(またはSettingsタブでの更新)の直後に一度だけ、何が変わったかを
    # 見せる。設定画面と同じウィンドウの中に直接表示する。
    pending_notice = _read_pending_update_notice()
    if pending_notice:
        version, notes = pending_notice
        root.after(300, lambda: show_whats_new_overlay(version, notes))

    def make_hotkey_recorder(parent, initial_value):
        """ホットキー入力欄(クリックしてから新しいキーを押すと変更できる)を
        作る。メインのホットキーとオーバーレイのホットキーで全く同じ挙動が
        必要なため、ここに共通化している。戻り値: (key_box, get_value)"""
        key_box = ctk.CTkFrame(parent, fg_color=COLOR_INPUT, corner_radius=8,
                                border_width=1, border_color=COLOR_INPUT)

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
            key_box.configure(border_color=COLOR_INPUT)
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
                key_box.configure(border_color=COLOR_INPUT)
            else:
                rec["recording"] = False
                entry.configure(state="disabled", text_color=COLOR_TEXT_FAINT)
                key_box.configure(border_color=COLOR_INPUT)

        return key_box, (lambda: rec["value"]), set_enabled

    # ヘッダー: アプリアイコンと同じヘッドホン + タイトル + ヘルプ(?)ボタン
    header = ctk.CTkFrame(root, fg_color="transparent")
    header.pack(fill="x", padx=28, pady=(22, 14))
    title_row = ctk.CTkFrame(header, fg_color="transparent")
    title_row.pack(fill="x")
    logo = tk.Canvas(title_row, width=30, height=30, bg=COLOR_BG, highlightthickness=0)
    logo.pack(side="left", padx=(0, 10))
    draw_headphones(logo, 15, 15, 26, COLOR_ACCENT)
    title_box = ctk.CTkFrame(title_row, fg_color="transparent")
    title_box.pack(side="left", fill="x", expand=True)
    ctk.CTkLabel(title_box, text="AudioSwitcher", font=FONT_DISPLAY,
                 text_color=COLOR_TEXT_MAIN, anchor="w", height=24).pack(fill="x")
    ctk.CTkLabel(title_box, text="ショートカット1つで、音の出力先を切り替えます。",
                 font=FONT_SMALL, text_color=COLOR_TEXT_SUB, anchor="w", height=18).pack(fill="x")
    help_button = ctk.CTkButton(
        title_row, text="?", command=lambda: show_onboarding(),
        width=30, height=30, corner_radius=15,
        fg_color=COLOR_INPUT, hover_color=COLOR_ACCENT_DIM, text_color=COLOR_TEXT_MAIN,
        border_width=0, font=F(12, "bold"),
    )
    help_button.pack(side="right")

    tabview = ctk.CTkTabview(
        root, width=1, height=1, fg_color=COLOR_PANEL, border_width=0, corner_radius=14,
        segmented_button_fg_color=COLOR_INPUT, segmented_button_selected_color=COLOR_BORDER,
        segmented_button_selected_hover_color=COLOR_BORDER,
        segmented_button_unselected_color=COLOR_INPUT, segmented_button_unselected_hover_color=COLOR_PANEL,
        text_color=COLOR_TEXT_MAIN, text_color_disabled=COLOR_TEXT_FAINT,
    )
    tabview._segmented_button.configure(font=F(10, "bold"), height=30)
    tabview.pack(fill="both", expand=True, padx=28, pady=(0, 14))
    tab_general = tabview.add("A/B Switch")
    tab_multi = tabview.add("Multi-Device")
    tab_settings = tabview.add("Settings")

    def section_label(parent, text):
        lbl = ctk.CTkLabel(parent, text=text, font=FONT_LABEL, text_color=COLOR_TEXT_SUB,
                           anchor="w", height=20)
        lbl.pack(fill="x", padx=20, pady=(12, 4))
        return lbl

    def hint_label(parent, text):
        ctk.CTkLabel(parent, text=text, font=FONT_SMALL, text_color=COLOR_TEXT_FAINT,
                     anchor="w", justify="left", height=18).pack(fill="x", padx=20, pady=(0, 6))

    def make_switch(parent, text, variable, command=None):
        return ctk.CTkSwitch(
            parent, text=text, variable=variable, onvalue="1", offvalue="0", command=command,
            font=FONT_BODY, text_color=COLOR_TEXT_MAIN,
            fg_color=COLOR_BORDER, progress_color=COLOR_ACCENT,
            button_color=COLOR_TEXT_MAIN, button_hover_color="#ffffff",
        )

    # ---- Settingsタブ: 起動・アップデート・バージョン情報 ----
    # 以前はボトムバーにあった「Start on boot」と、UPDATEタブの中身を
    # ここにまとめた(A/B・Multi-Deviceのタブを本来の設定だけにするため)。
    settings_box = ctk.CTkFrame(tab_settings, fg_color="transparent")
    settings_box.pack(fill="both", expand=True)

    section_label(settings_box, "General")
    startup_var = tk.StringVar(value=saved_startup)
    switch_startup = make_switch(settings_box, "Start on boot（Windows起動時に自動で常駐）", startup_var)
    switch_startup.pack(anchor="w", padx=20, pady=(0, 4))

    # アップデート欄は常に表示しておく(更新がある時しか出ないと、Auto-update
    # トグル自体を普段いじれなくなってしまうため)。中身(タイトル/ボタン)は
    # チェック結果に応じて後から書き換える。
    section_label(settings_box, "Updates")
    update_panel = ctk.CTkFrame(settings_box, fg_color=COLOR_INPUT, corner_radius=10, border_width=0)
    update_panel.pack(fill="x", padx=20, pady=(0, 4))

    update_title_var = tk.StringVar(value="Checking for updates...")
    ctk.CTkLabel(update_panel, textvariable=update_title_var, font=FONT_LABEL,
                 text_color=COLOR_ACCENT, anchor="w", height=20).pack(fill="x", padx=16, pady=(12, 0))
    update_detail_var = tk.StringVar(value=f"Current version: v{__version__}")
    ctk.CTkLabel(update_panel, textvariable=update_detail_var, font=FONT_BODY, text_color=COLOR_TEXT_MAIN,
                 anchor="w", height=20).pack(fill="x", padx=16, pady=(0, 2))
    update_status_var = tk.StringVar(value="")
    ctk.CTkLabel(update_panel, textvariable=update_status_var, font=FONT_SMALL, text_color=COLOR_TEXT_SUB,
                 anchor="w", wraplength=420, justify="left", height=16).pack(fill="x", padx=16, pady=(0, 6))

    _update_state = {"installer_url": None, "release_url": None, "latest_tag": "", "notes": ""}

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
                _write_update_notice(_update_state["latest_tag"], _update_state["notes"])
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

    update_actions = ctk.CTkFrame(update_panel, fg_color="transparent")
    update_actions.pack(fill="x", padx=16, pady=(0, 14))
    switch_auto_update = make_switch(update_actions, "Auto-update", auto_update_var, on_auto_toggle)
    switch_auto_update.pack(side="left")

    install_btn = ctk.CTkButton(
        update_actions, text="Up to date", state="disabled", command=start_install,
        fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
        text_color_disabled=COLOR_TEXT_SUB,
        corner_radius=8, font=FONT_LABEL, width=150, height=32,
    )
    install_btn.pack(side="right")

    def refresh_install_btn_style():
        # 押せない時(最新版)はカッパーの塗りを外し、押せる時だけ目立たせる
        enabled = install_btn.cget("state") == "normal"
        install_btn.configure(fg_color=(COLOR_ACCENT if enabled else COLOR_BORDER))

    refresh_install_btn_style()

    section_label(settings_box, "Help")
    about_row = ctk.CTkFrame(settings_box, fg_color="transparent")
    about_row.pack(fill="x", padx=20, pady=(0, 16))
    ctk.CTkButton(
        about_row, text="使い方をもう一度見る", command=lambda: show_onboarding(),
        fg_color=COLOR_INPUT, hover_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN,
        corner_radius=8, font=FONT_LABEL, height=32,
    ).pack(side="left")
    ctk.CTkButton(
        about_row, text="GitHub", command=lambda: webbrowser.open(f"https://github.com/{GITHUB_REPO}"),
        fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
        corner_radius=8, font=FONT_LABEL, width=80, height=32,
    ).pack(side="left", padx=(8, 0))
    ctk.CTkLabel(about_row, text=f"v{__version__}", font=FONT_SMALL, text_color=COLOR_TEXT_FAINT,
                 anchor="e").pack(side="right")

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
                latest_tag, release_url, installer_url, notes = info
                _update_state["installer_url"] = installer_url
                _update_state["release_url"] = release_url
                _update_state["latest_tag"] = latest_tag
                _update_state["notes"] = notes
                update_title_var.set("Update available")
                update_detail_var.set(f"{latest_tag} is available (current: v{__version__})")
                install_btn.configure(state="normal", text="Download & Install")
                refresh_install_btn_style()
                if auto_update_var.get() == "1" and installer_url:
                    start_install()
            else:
                update_title_var.set("Up to date")
                update_detail_var.set(f"You're on the latest version (v{__version__})")
            return
        except queue.Empty:
            pass
        root.after(500, _poll_update_info_queue)

    _check_for_updates()
    root.after(500, _poll_update_info_queue)

    # ---- A/B Switchタブ ----
    # 今どちらのデバイスから音が出ているかを、デバイス名とA/Bのバッジで表示する
    now_card = ctk.CTkFrame(tab_general, fg_color=COLOR_INPUT, corner_radius=10)
    now_card.pack(fill="x", padx=20, pady=(10, 4))
    now_text = ctk.CTkFrame(now_card, fg_color="transparent")
    now_text.pack(side="left", fill="x", expand=True, padx=(16, 8), pady=12)
    ctk.CTkLabel(now_text, text="現在の出力", font=FONT_SMALL, text_color=COLOR_TEXT_SUB,
                 anchor="w", height=16).pack(fill="x")
    now_row = ctk.CTkFrame(now_text, fg_color="transparent")
    now_row.pack(fill="x")
    side_badge = ctk.CTkLabel(now_row, text="A", font=F(9, "bold"), width=22, height=22,
                              corner_radius=11, fg_color=COLOR_ACCENT, text_color=COLOR_BG)
    status_label = ctk.CTkLabel(now_row, text="", font=F(12, "bold"), text_color=COLOR_TEXT_MAIN,
                                anchor="w", height=24)
    status_label.pack(side="right", fill="x", expand=True)

    settings_toast = _ToastManager(root)

    def refresh_indicator():
        a, b = combo_a.get(), combo_b.get()
        if current_default and current_default == a:
            active_side = "A"
        elif current_default and current_default == b:
            active_side = "B"
        else:
            active_side = None
        if active_side:
            side_badge.configure(text=active_side)
            side_badge.pack(side="left", padx=(0, 8))
        else:
            side_badge.pack_forget()
        name = current_default or "検出できませんでした"
        if len(name) > 34:
            name = name[:32] + "…"
        status_label.configure(text=name)

    def on_test_switch():
        nonlocal current_default
        a, b = combo_a.get(), combo_b.get()
        if b == "Select device...":
            tk.messagebox.showwarning("Routing Incomplete", "Please select a valid target for Device B.")
            return
        target_display_name = b if current_default == a else a
        target_id = device_dict.get(target_display_name)
        if not target_id:
            tk.messagebox.showerror("Error", "選択されたデバイスが見つかりませんでした。")
            return
        try:
            ensure_com_initialized()
            set_default_device_all_roles(target_id)
        except Exception as e:
            log_error("test_switch", e)
            tk.messagebox.showerror("Error", f"切り替えに失敗しました: {e}")
            return
        current_default = target_display_name
        refresh_indicator()
        settings_toast.show(target_display_name)

    test_button = ctk.CTkButton(
        now_card, text="テスト切り替え", command=on_test_switch,
        fg_color="transparent", hover_color=COLOR_BORDER, text_color=COLOR_ACCENT,
        border_width=1, border_color=COLOR_ACCENT_DIM, corner_radius=8,
        font=FONT_LABEL, width=120, height=32,
    )
    test_button.pack(side="right", padx=(0, 14))

    # デバイス設定
    container = ctk.CTkFrame(tab_general, fg_color="transparent")
    container.pack(fill="both", expand=True)

    ab_enabled_var = tk.StringVar(value=saved_ab_enabled)
    switch_ab = make_switch(container, "Enable A/B switch", ab_enabled_var)
    switch_ab.pack(anchor="w", padx=20, pady=(12, 0))

    combo_kwargs = dict(
        state="readonly", fg_color=COLOR_INPUT, border_color=COLOR_INPUT, border_width=1,
        button_color=COLOR_INPUT, button_hover_color=COLOR_BORDER,
        dropdown_fg_color=COLOR_INPUT, dropdown_text_color=COLOR_TEXT_MAIN,
        dropdown_hover_color=COLOR_ACCENT_DIM, text_color=COLOR_TEXT_MAIN,
        text_color_disabled=COLOR_TEXT_FAINT,
        font=FONT_BODY, dropdown_font=FONT_BODY,
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

    label_a = section_label(container, "Device A")
    combo_a = ctk.CTkComboBox(container, values=devices, **combo_kwargs)
    combo_a.pack(fill="x", padx=20, pady=(0, 2))
    if saved_a in devices: combo_a.set(saved_a)
    elif current_default in devices: combo_a.set(current_default)
    else: combo_a.set(devices[0])
    make_combo_fully_clickable(combo_a)

    label_b = section_label(container, "Device B")
    combo_b = ctk.CTkComboBox(container, values=devices_b_options, **combo_kwargs)
    combo_b.pack(fill="x", padx=20, pady=(0, 2))
    if saved_b in devices_b_options and saved_b != "Select device...": combo_b.set(saved_b)
    else: combo_b.set(devices_b_options[0])
    make_combo_fully_clickable(combo_b)

    combo_a.configure(command=lambda _=None: refresh_indicator())
    combo_b.configure(command=lambda _=None: refresh_indicator())
    refresh_indicator()

    # ホットキー
    label_hotkey = section_label(container, "Hotkey（A ⇄ B を切り替え）")
    key_box, get_shortcut, set_key_enabled = make_hotkey_recorder(container, saved_key)
    key_box.pack(fill="x", padx=20, pady=(0, 4))
    hint_label(container, "クリックしてから、新しいキーの組み合わせを押すと変更できます")

    def apply_ab_enabled_style(*_):
        enabled = ab_enabled_var.get() == "1"
        combo_a.configure(state=("readonly" if enabled else "disabled"))
        combo_b.configure(state=("readonly" if enabled else "disabled"))
        set_key_enabled(enabled)
        test_button.configure(state=("normal" if enabled else "disabled"))
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
        text="ホットキーを押すと、つながっている出力デバイスが画面中央に一覧で出て、\nそこから選んで切り替えられます。",
        font=FONT_BODY, text_color=COLOR_TEXT_SUB, justify="left", anchor="w"
    ).pack(fill="x", padx=20, pady=(12, 8))

    overlay_enabled_var = tk.StringVar(value=saved_overlay_enabled)
    switch_overlay = make_switch(container2, "Enable device-select overlay", overlay_enabled_var)
    switch_overlay.pack(anchor="w", padx=20, pady=(0, 0))

    label_overlay_key = section_label(container2, "Hotkey（一覧を開く）")
    overlay_key_box, get_overlay_shortcut, set_overlay_key_enabled = make_hotkey_recorder(container2, saved_overlay_key)
    overlay_key_box.pack(fill="x", padx=20, pady=(0, 4))
    hint_label(container2, "クリックしてから、新しいキーの組み合わせを押すと変更できます")

    # 一覧に出すデバイスを1台ずつ選べるようにする(デバイスが多い人向け)。
    # 覚えておくのは「出さない」デバイスなので、新しくつないだ物は最初からON。
    saved_excluded = set(load_overlay_excluded())
    label_devices = section_label(container2, "一覧に出すデバイス")
    device_count_var = tk.StringVar()
    ctk.CTkLabel(label_devices.master, textvariable=device_count_var, font=FONT_SMALL,
                 text_color=COLOR_TEXT_FAINT, anchor="e", height=20).place(
        in_=label_devices, relx=1.0, rely=0.5, anchor="e")

    ROW_H = 30
    if len(devices) > 6:
        # 台数が多い時は高さを決めてスクロールさせる(ウィンドウが縦に伸びすぎないように)
        device_list_frame = ctk.CTkScrollableFrame(
            container2, fg_color=COLOR_INPUT, corner_radius=10, height=ROW_H * 5,
            scrollbar_button_color=COLOR_BORDER, scrollbar_button_hover_color=COLOR_TEXT_FAINT)
    else:
        device_list_frame = ctk.CTkFrame(container2, fg_color=COLOR_INPUT, corner_radius=10)
    device_list_frame.pack(fill="x", padx=20, pady=(0, 6))

    device_vars = {}
    device_checks = []
    for i, name in enumerate(devices):
        is_current = name == current_default
        var = tk.StringVar(value=("0" if name in saved_excluded else "1"))
        device_vars[name] = var
        row = ctk.CTkFrame(device_list_frame, fg_color="transparent", height=ROW_H)
        row.pack(fill="x", padx=12, pady=(6 if i == 0 else 0, 6 if i == len(devices) - 1 else 0))
        check = ctk.CTkCheckBox(
            row, text=name, variable=var, onvalue="1", offvalue="0",
            command=lambda: refresh_device_count(), font=FONT_BODY, text_color=COLOR_TEXT_MAIN,
            text_color_disabled=COLOR_TEXT_FAINT,
            fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_TEXT_FAINT,
            checkmark_color=COLOR_BG, checkbox_width=18, checkbox_height=18, corner_radius=5, border_width=2,
        )
        check.pack(side="left", pady=4)
        device_checks.append(check)
        if is_current:
            ctk.CTkLabel(row, text="使用中", font=FONT_SMALL, text_color=COLOR_ACCENT,
                         height=20).pack(side="right")

    def get_overlay_excluded():
        return [name for name, var in device_vars.items() if var.get() != "1"]

    def refresh_device_count():
        on = sum(1 for var in device_vars.values() if var.get() == "1")
        device_count_var.set(f"{on} / {len(devices)} 台")

    refresh_device_count()
    hint_label(container2, "チェックを外したデバイスは、ホットキーで開く一覧に出なくなります")

    def apply_overlay_enabled_style(*_):
        enabled = overlay_enabled_var.get() == "1"
        set_overlay_key_enabled(enabled)
        label_overlay_key.configure(text_color=(COLOR_TEXT_SUB if enabled else COLOR_TEXT_FAINT))
        label_devices.configure(text_color=(COLOR_TEXT_SUB if enabled else COLOR_TEXT_FAINT))
        for check in device_checks:
            check.configure(state=("normal" if enabled else "disabled"))

    switch_overlay.configure(command=apply_overlay_enabled_style)
    apply_overlay_enabled_style()

    # ボトム: 保存ボタンだけを右側に置く
    bottom_bar = ctk.CTkFrame(root, fg_color="transparent")
    bottom_bar.pack(fill="x", padx=28, pady=(0, 22))

    btn_frame = ctk.CTkFrame(bottom_bar, fg_color="transparent")
    btn_frame.pack(side="right")

    def on_save_only():
        if ab_enabled_var.get() == "1" and combo_b.get() == "Select device...":
            tk.messagebox.showwarning("Routing Incomplete", "Please select a valid target for Device B.")
            return
        save_config(combo_a.get(), combo_b.get(), get_shortcut(), startup_var.get(),
                    get_overlay_shortcut(), overlay_enabled_var.get(), ab_enabled_var.get(),
                    auto_update_var.get(), get_overlay_excluded())
        root.destroy()
        sys.exit(0)

    btn_save_only = ctk.CTkButton(
        btn_frame, text="Save only", command=on_save_only,
        fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
        border_width=1, border_color=COLOR_BORDER, corner_radius=8,
        font=FONT_LABEL, width=110, height=38,
    )
    btn_save_only.pack(side="left", padx=(0, 10))

    def on_save_and_run():
        if ab_enabled_var.get() == "1" and combo_b.get() == "Select device...":
            tk.messagebox.showwarning("Routing Incomplete", "Please select a valid target for Device B.")
            return
        save_config(combo_a.get(), combo_b.get(), get_shortcut(), startup_var.get(),
                    get_overlay_shortcut(), overlay_enabled_var.get(), ab_enabled_var.get(),
                    auto_update_var.get(), get_overlay_excluded())
        root.destroy()
        start_backend()

    btn_save_run = ctk.CTkButton(
        btn_frame, text="Save & Run", command=on_save_and_run,
        fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
        corner_radius=8, font=FONT_LABEL, width=130, height=38,
    )
    btn_save_run.pack(side="right")

    # ウィンドウ高さは中身の実測サイズに合わせて自動調整する。CTkTabviewを
    # 導入したことで固定の見積もり値だと下端(ボトムバー)が切れることが
    # あったため、全タブの中身を組み終えたこのタイミングで測り直す。
    # CTkTabviewの高さは「今開いているタブ」の中身でしか決まらないので、
    # 全タブを一度ずつ開いて一番高いものに合わせる(どのタブでも下の
    # Saveボタンが切れないように)。
    required_height = 0
    for tab_name in ("Multi-Device", "Settings", "A/B Switch"):
        tabview.set(tab_name)
        root.update_idletasks()
        required_height = max(required_height, root.winfo_reqheight() + 12)
    root.geometry(f"560x{required_height}")

    # ---- 機能紹介ウォークスルー。初回起動時に自動で出るほか、ヘッダーの
    # ヘルプボタン(?)や Settingsタブからいつでも呼び出せる。実際のトースト/
    # オーバーレイをその場で呼び出す方式は、モニター判定の問題(トーストが
    # 別モニターに出て見えない等)に左右されやすいため、ウィザード自身の
    # Canvas上だけで完結する「操作のイラストアニメ」にしている
    # (実機・実際のウィンドウ位置に依存しない)。
    # 以前のアニメはカーソルが説明なしに素早く動くだけで何をしているのか
    # 分かりづらかったため、各操作に「番号付きの字幕」「操作対象を囲む枠」
    # 「クリックの波紋」を付け、1つずつ間を取ってゆっくり見せるようにした。----
    def show_onboarding():
        if getattr(show_onboarding, "_open", None) is not None:
            try:
                show_onboarding._open.lift()
                show_onboarding._open.focus_force()
                return
            except Exception:
                pass
        wiz = {"step": 0, "window": None}

        ILLO_W, ILLO_H = 420, 300   # イラスト用Canvas(下の約50pxは字幕の帯)
        CAPTION_Y = 276
        MOVE_MS = 650                # カーソル移動の時間
        BEAT_MS = 650                # 操作と操作の間の「間」

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

        def later(canvas, ms, fn):
            """ウィザードを閉じた/ページを移った後に古いアニメが動かないよう、
            Canvasがまだ生きている時だけfnを呼ぶafter()。"""
            canvas.after(ms, lambda: fn() if canvas.winfo_exists() else None)

        def make_caption(canvas):
            """イラスト下部の字幕の帯。「今どの操作をしているか」を番号付きの
            一文で出す。文字は切り替わるたびにふわっと浮かび上がる。"""
            _rounded_rect(canvas, 10, CAPTION_Y - 18, ILLO_W - 10, CAPTION_Y + 18, 10,
                          fill=COLOR_INPUT, outline="")
            badge = canvas.create_oval(22, CAPTION_Y - 11, 44, CAPTION_Y + 11, fill=COLOR_ACCENT, outline="")
            badge_text = canvas.create_text(33, CAPTION_Y, text="1", fill=COLOR_BG, font=F(9, "bold"))
            text = canvas.create_text(56, CAPTION_Y, anchor="w", text="", fill=COLOR_TEXT_MAIN,
                                      font=F(10, "bold"))

            def set_caption(num, message):
                if not canvas.winfo_exists():
                    return
                canvas.itemconfig(badge_text, text=str(num))
                canvas.itemconfig(text, text=message)
                canvas.coords(text, 62, CAPTION_Y)
                prev = {"x": 62}

                def on_tick(e):
                    canvas.itemconfig(text, fill=_lerp_hex(COLOR_INPUT, COLOR_TEXT_MAIN, e))
                    canvas.itemconfig(badge, fill=_lerp_hex(COLOR_INPUT, COLOR_ACCENT, e))
                    x = 62 - 6 * e
                    canvas.move(text, x - prev["x"], 0)
                    prev["x"] = x
                tween(canvas, 260, on_tick)

            return set_caption

        def make_focus_ring(canvas):
            """操作対象を囲むカッパー色の枠。表示時に2回パルスして目を引く。
            show()で新しい場所に移ると前の枠は消える。"""
            ring = {"id": None}

            def clear():
                if ring["id"] is not None and canvas.winfo_exists():
                    canvas.delete(ring["id"])
                ring["id"] = None

            def show(x1, y1, x2, y2):
                clear()
                ring["id"] = _rounded_rect(canvas, x1 - 3, y1 - 3, x2 + 3, y2 + 3, 9,
                                           fill="", outline=COLOR_ACCENT, width=2)
                item = ring["id"]

                def on_tick(e):
                    if canvas.type(item) is None:
                        return
                    pulse = abs(((e * 2) % 1.0) * 2 - 1)  # 1→0→1→0→1 の2回パルス
                    canvas.itemconfig(item, outline=_lerp_hex(COLOR_ACCENT_DIM, COLOR_ACCENT, pulse))
                tween(canvas, 700, on_tick)

            return show, clear

        def ripple(canvas, x, y):
            """クリックした場所に広がる波紋(押したことがひと目で分かるように)"""
            ring_id = canvas.create_oval(x - 3, y - 3, x + 3, y + 3, outline=COLOR_ACCENT, width=2)

            def on_tick(e):
                if canvas.type(ring_id) is None:
                    return
                r = 3 + 15 * e
                canvas.coords(ring_id, x - r, y - r, x + r, y + r)
                canvas.itemconfig(ring_id, outline=_lerp_hex(COLOR_ACCENT, COLOR_PANEL, e))
            tween(canvas, 380, on_tick, lambda: canvas.delete(ring_id))

        def make_cursor(canvas):
            """Windowsの標準矢印カーソルに近い形・配色(白地に黒縁)のアイコン。
            set_posで即座に配置、move_toでなめらかに(tween経由で)移動、
            clickでその場に波紋を出す。"""
            cur = canvas.create_polygon(
                0, 0, 0, 19, 5, 15, 8, 21.5, 11.5, 20, 7.5, 14.5, 13, 14.5,
                fill="#ffffff", outline="#000000", width=1.2,
            )
            canvas.itemconfig(cur, state="hidden")
            pos = {"xy": (ILLO_W + 20, -10)}

            def set_pos(x, y):
                pos["xy"] = (x, y)
                canvas.coords(
                    cur,
                    x, y, x, y + 19, x + 5, y + 15, x + 8, y + 21.5,
                    x + 11.5, y + 20, x + 7.5, y + 14.5, x + 13, y + 14.5,
                )
                # ドロップダウンなど、カーソル生成後に新しく描かれる要素が
                # 上に重なって隠してしまわないよう、動くたびに最前面へ。
                canvas.tag_raise(cur)

            def move_to(to_xy, on_done=None, duration_ms=MOVE_MS):
                if not canvas.winfo_exists():
                    return
                canvas.itemconfig(cur, state="normal")
                fx, fy = pos["xy"]

                def on_tick(e):
                    set_pos(fx + (to_xy[0] - fx) * e, fy + (to_xy[1] - fy) * e)

                tween(canvas, duration_ms, on_tick, on_done)

            def click():
                x, y = pos["xy"]
                ripple(canvas, x, y)
                canvas.tag_raise(cur)

            def hide():
                if canvas.winfo_exists():
                    canvas.itemconfig(cur, state="hidden")
                pos["xy"] = (ILLO_W + 20, -10)

            return set_pos, move_to, click, hide

        def draw_keycaps(canvas, labels, cx, cy):
            """Ctrl + Alt + A のようなキーキャップ列を「+」付きで描き、
            press()で全部が沈み込んで光る、release()で戻る。"""
            widths = [48 if len(l) > 1 else 32 for l in labels]
            plus_w = 18
            total_w = sum(widths) + plus_w * (len(labels) - 1)
            x = cx - total_w / 2
            caps = []
            for i, (label, w) in enumerate(zip(labels, widths)):
                shadow = _rounded_rect(canvas, x, cy - 13, x + w, cy + 17, 7, fill="#0f0e12", outline="")
                rect = _rounded_rect(canvas, x, cy - 16, x + w, cy + 14, 7,
                                     fill=COLOR_INPUT, outline=COLOR_BORDER, width=1)
                text = canvas.create_text(x + w / 2, cy - 1, text=label,
                                          fill=COLOR_TEXT_MAIN, font=F(9, "bold"))
                caps.append((rect, text))
                x += w
                if i < len(labels) - 1:
                    canvas.create_text(x + plus_w / 2, cy - 1, text="+", fill=COLOR_TEXT_FAINT, font=F(10, "bold"))
                    x += plus_w

            def press():
                for rect, text in caps:
                    canvas.move(rect, 0, 3)
                    canvas.move(text, 0, 3)
                    canvas.itemconfig(rect, fill=COLOR_ACCENT, outline=COLOR_ACCENT)
                    canvas.itemconfig(text, fill=COLOR_BG)

            def release():
                for rect, text in caps:
                    canvas.move(rect, 0, -3)
                    canvas.move(text, 0, -3)
                    canvas.itemconfig(rect, fill=COLOR_INPUT, outline=COLOR_BORDER)
                    canvas.itemconfig(text, fill=COLOR_TEXT_MAIN)

            return press, release

        def make_settings_illustration(canvas):
            """ステップ1: 実際の設定画面のミニチュアの上で、
            「デバイスを選ぶ→ホットキーを登録する→Multi-Deviceタブに切り替える
            →有効化する→保存する」を、1操作ずつ字幕と枠で示しながら実演する。"""
            W = ILLO_W
            cx = W / 2
            set_caption = make_caption(canvas)
            show_ring, clear_ring = make_focus_ring(canvas)

            tab_w, tab_h, tab_gap = 130, 22, 6
            tx0 = cx - (tab_w * 2 + tab_gap) / 2
            tab_ab_coords = (tx0, 8, tx0 + tab_w, 8 + tab_h)
            tab_multi_coords = (tx0 + tab_w + tab_gap, 8, tx0 + tab_w * 2 + tab_gap, 8 + tab_h)
            tab_ab = _rounded_rect(canvas, *tab_ab_coords, 6, fill=COLOR_BORDER, outline="")
            canvas.create_text(tx0 + tab_w / 2, 8 + tab_h / 2, text="A/B Switch",
                               fill=COLOR_TEXT_MAIN, font=F(8, "bold"))
            tab_multi = _rounded_rect(canvas, *tab_multi_coords, 6, fill=COLOR_INPUT, outline="")
            canvas.create_text(tx0 + tab_w + tab_gap + tab_w / 2, 8 + tab_h / 2, text="Multi-Device",
                               fill=COLOR_TEXT_MAIN, font=F(8, "bold"))

            _rounded_rect(canvas, 10, 38, W - 10, 252, 12, fill=COLOR_BG, outline="")

            def field(label_text, y, is_combo=False, placeholder="Select device..."):
                lbl = canvas.create_text(28, y - 3, anchor="w", text=label_text,
                                         fill=COLOR_TEXT_SUB, font=F(8, "bold"))
                box = _rounded_rect(canvas, 28, y + 8, W - 28, y + 34, 6,
                                    fill=COLOR_INPUT, outline="")
                box_text = canvas.create_text(40, y + 21, anchor="w", text=placeholder,
                                              fill=COLOR_TEXT_FAINT, font=F(9))
                items = [lbl, box, box_text]
                if is_combo:
                    items.append(canvas.create_text(W - 40, y + 21, text="▾", fill=COLOR_TEXT_SUB, font=F(9)))
                return items, box_text, (28, y + 8, W - 28, y + 34)

            ab_items_a, box_text_a, box_a_coords = field("Device A", 54, is_combo=True)
            ab_items_b, box_text_b, box_b_coords = field("Device B", 100, is_combo=True)
            ab_items_key, box_text_key, box_key_coords = field("Hotkey", 146, placeholder="")
            canvas.itemconfig(box_text_key, fill=COLOR_TEXT_MAIN, font=F(9, "bold"))
            ab_group = ab_items_a + ab_items_b + ab_items_key

            toggle_coords = (28, 56, 64, 72)
            toggle_track = _rounded_rect(canvas, *toggle_coords, 8, fill=COLOR_BORDER, outline="")
            toggle_knob = canvas.create_oval(30, 58, 42, 70, fill=COLOR_TEXT_MAIN, outline="")
            toggle_label = canvas.create_text(74, 64, anchor="w", text="Enable device-select overlay",
                                              fill=COLOR_TEXT_MAIN, font=F(8))
            okey_items, box_text_okey, box_okey_coords = field("Hotkey", 100, placeholder="")
            canvas.itemconfig(box_text_okey, fill=COLOR_TEXT_MAIN, font=F(9, "bold"))
            multi_group = [toggle_track, toggle_knob, toggle_label] + okey_items
            for item in multi_group:
                canvas.itemconfig(item, state="hidden")

            btn_coords = (W - 130, 214, W - 28, 242)
            btn = _rounded_rect(canvas, *btn_coords, 8, fill=COLOR_ACCENT, outline="")
            canvas.create_text(W - 79, 228, text="Save & Run", fill=COLOR_BG, font=F(9, "bold"))

            set_cursor, move_cursor, click, hide_cursor = make_cursor(canvas)

            def center(c):
                return ((c[0] + c[2]) / 2, (c[1] + c[3]) / 2 - 6)

            def type_keys(text_item, combo, on_done):
                """ホットキー欄に「ctrl → ctrl+alt → ctrl+alt+a」と1キーずつ入る様子"""
                parts = combo.split("+")
                canvas.itemconfig(text_item, text="Press keys...", fill=COLOR_ACCENT)

                def step(i):
                    if i > len(parts):
                        canvas.itemconfig(text_item, fill=COLOR_TEXT_MAIN)
                        on_done()
                        return
                    if i > 0:
                        canvas.itemconfig(text_item, text="+".join(parts[:i]))
                    later(canvas, 320, lambda: step(i + 1))
                later(canvas, 380, lambda: step(0))

            def open_dropdown_and_select(box_coords, box_text_item, options, choice_index, on_done):
                """実際のCTkComboBoxと同じく、枠をクリック→プルダウンが
                現れる→カーソルが降りて選択肢をクリック→枠の表示が更新される、
                という一連の操作をなめらかに実演する。"""
                x1, y1, x2, y2 = box_coords
                row_h = 24
                dd_top = y2 + 3
                dd_h = row_h * len(options) + 6
                dd_items = [_rounded_rect(canvas, x1, dd_top, x2, dd_top + dd_h, 6,
                                          fill=COLOR_INPUT, outline=COLOR_BORDER, width=1)]
                dd_rows = []
                for i, opt in enumerate(options):
                    ry = dd_top + 3 + i * row_h
                    row_bg = canvas.create_rectangle(x1 + 2, ry, x2 - 2, ry + row_h - 2, fill=COLOR_INPUT, outline="")
                    row_text = canvas.create_text(x1 + 14, ry + (row_h - 2) / 2, anchor="w", text=opt,
                                                  fill=COLOR_TEXT_MAIN, font=F(9))
                    dd_rows.append(row_bg)
                    dd_items.extend([row_bg, row_text])

                target_row_bg = dd_rows[choice_index]
                target_xy = (x1 + 60, dd_top + 3 + choice_index * row_h + (row_h - 2) / 2 - 6)

                def highlight_and_pick():
                    def on_tick(e):
                        canvas.itemconfig(target_row_bg, fill=_lerp_hex(COLOR_INPUT, COLOR_ACCENT_DIM, e))
                    click()
                    tween(canvas, 220, on_tick, lambda: later(canvas, 250, finish))

                def finish():
                    canvas.itemconfig(box_text_item, text=options[choice_index], fill=COLOR_TEXT_MAIN)
                    for item in dd_items:
                        canvas.delete(item)
                    on_done()

                later(canvas, 300, lambda: move_cursor(target_xy, highlight_and_pick, 450))

            # 実演する操作の順番: (字幕, 操作対象, 実行する処理)
            def act_device_a(done):
                click()
                later(canvas, 150, lambda: open_dropdown_and_select(
                    box_a_coords, box_text_a, ["Speakers", "Headphones", "Monitor Speakers"], 0, done))

            def act_device_b(done):
                click()
                later(canvas, 150, lambda: open_dropdown_and_select(
                    box_b_coords, box_text_b, ["Speakers", "Headphones", "Monitor Speakers"], 1, done))

            def act_hotkey(done):
                click()
                type_keys(box_text_key, "ctrl+alt+a", done)

            def act_tab_multi(done):
                click()

                def on_tick(e):
                    canvas.itemconfig(tab_ab, fill=_lerp_hex(COLOR_BORDER, COLOR_INPUT, e))
                    canvas.itemconfig(tab_multi, fill=_lerp_hex(COLOR_INPUT, COLOR_BORDER, e))
                tween(canvas, 250, on_tick)
                for item in ab_group:
                    canvas.itemconfig(item, state="hidden")
                for item in multi_group:
                    canvas.itemconfig(item, state="normal")
                later(canvas, 300, done)

            def act_toggle(done):
                click()

                def on_tick(e):
                    canvas.itemconfig(toggle_track, fill=_lerp_hex(COLOR_BORDER, COLOR_ACCENT, e))
                    canvas.coords(toggle_knob, 30 + 22 * e, 58, 42 + 22 * e, 70)
                tween(canvas, 260, on_tick, done)

            def act_overlay_hotkey(done):
                click()
                type_keys(box_text_okey, "ctrl+alt+s", done)

            def act_save(done):
                click()

                def on_tick(e):
                    ratio = e * 2 if e < 0.5 else (1 - e) * 2
                    canvas.itemconfig(btn, fill=_lerp_hex(COLOR_ACCENT, COLOR_ACCENT_HOVER, ratio))
                tween(canvas, 260, on_tick, done)

            actions = [
                ("Device A に1台目の出力先を選ぶ", box_a_coords, act_device_a),
                ("Device B に2台目の出力先を選ぶ", box_b_coords, act_device_b),
                ("Hotkey をクリックして、使いたいキーを押す", box_key_coords, act_hotkey),
                ("3台以上使うなら Multi-Device タブへ", tab_multi_coords, act_tab_multi),
                ("スイッチを ON にする", toggle_coords, act_toggle),
                ("一覧を開く Hotkey も同じように登録", box_okey_coords, act_overlay_hotkey),
                ("最後に Save & Run で保存して常駐", btn_coords, act_save),
            ]

            def run(i):
                if not canvas.winfo_exists():
                    return
                if i >= len(actions):
                    later(canvas, 1400, reset_loop)
                    return
                message, coords, act = actions[i]
                set_caption(i + 1, message)
                show_ring(*coords)
                later(canvas, 350, lambda: move_cursor(center(coords), lambda: act(
                    lambda: later(canvas, BEAT_MS, lambda: run(i + 1)))))

            def reset_loop():
                hide_cursor()
                clear_ring()
                for item in multi_group:
                    canvas.itemconfig(item, state="hidden")
                for item in ab_group:
                    canvas.itemconfig(item, state="normal")
                canvas.itemconfig(box_text_a, text="Select device...", fill=COLOR_TEXT_FAINT)
                canvas.itemconfig(box_text_b, text="Select device...", fill=COLOR_TEXT_FAINT)
                canvas.itemconfig(box_text_key, text="")
                canvas.itemconfig(box_text_okey, text="")
                canvas.itemconfig(tab_ab, fill=COLOR_BORDER)
                canvas.itemconfig(tab_multi, fill=COLOR_INPUT)
                canvas.itemconfig(toggle_track, fill=COLOR_BORDER)
                canvas.coords(toggle_knob, 30, 58, 42, 70)
                later(canvas, 600, lambda: run(0))

            set_caption(1, "Device A に1台目の出力先を選ぶ")
            later(canvas, 700, lambda: run(0))

        DESKTOP_WALLPAPER = "#25303b"
        DESKTOP_TASKBAR = "#161d24"
        DESK = (20, 50, 400, 252)

        def make_desktop_mockup(canvas):
            """トースト/オーバーレイが実際にはデスクトップ画面のどこに出るかを
            示すための、ミニデスクトップ(壁紙+タスクバー)のモックアップ。"""
            x1, y1, x2, y2 = DESK
            desktop_bg = canvas.create_rectangle(x1, y1, x2, y2, fill=DESKTOP_WALLPAPER, outline="")
            taskbar_h = 16
            taskbar = canvas.create_rectangle(x1, y2 - taskbar_h, x2, y2, fill=DESKTOP_TASKBAR, outline="")
            canvas.create_oval(x1 + 6, y2 - taskbar_h + 4, x1 + 14, y2 - 4, fill=COLOR_TEXT_FAINT, outline="")
            return desktop_bg, taskbar, taskbar_h

        def make_toast(canvas, taskbar_h):
            """実際の画面と同じく、デスクトップの右下(タスクバーの上)に
            下からスッと出てくるトースト通知。show(name, on_done)/hide()"""
            x1, y1, x2, y2 = DESK
            TOAST_W, TOAST_H = 196, 52
            tx = x2 - TOAST_W - 10
            final_ty = y2 - taskbar_h - TOAST_H - 8
            slide_offset = 16
            panel = _rounded_rect(canvas, tx, final_ty, tx + TOAST_W, final_ty + TOAST_H, 10,
                                  fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
            accent_bar = canvas.create_rectangle(tx, final_ty + 8, tx + 3, final_ty + TOAST_H - 8,
                                                 fill=COLOR_ACCENT, outline="")
            header = canvas.create_text(tx + 14, final_ty + 16, anchor="w", text="Output switched",
                                        fill=COLOR_ACCENT, font=F(7, "bold"))
            device_text = canvas.create_text(tx + 14, final_ty + 34, anchor="w", text="",
                                             fill=COLOR_TEXT_MAIN, font=F(10, "bold"))
            items = (panel, accent_bar, header, device_text)
            for item in items:
                canvas.itemconfig(item, state="hidden")

            def show(name, on_done=None):
                canvas.itemconfig(device_text, text=name)
                # panelは角丸(_rounded_rect)のポリゴンで、通常の矩形より
                # 座標の個数が多い。canvas.coords()で絶対座標を与え直すと
                # 頂点数が合わず形が壊れて斜線が入るため、canvas.move()で
                # 差分だけ動かす(頂点数に関係なく形を保ったまま平行移動できる)。
                for item in items:
                    canvas.itemconfig(item, state="normal")
                    canvas.move(item, 0, slide_offset)
                    canvas.tag_raise(item)
                prev = {"v": slide_offset}

                def on_tick(e):
                    dy = slide_offset * (1 - e)
                    for item in items:
                        canvas.move(item, 0, dy - prev["v"])
                    prev["v"] = dy
                tween(canvas, 260, on_tick, on_done)

            def hide():
                for item in items:
                    canvas.itemconfig(item, state="hidden")

            return show, hide, (tx, final_ty, tx + TOAST_W, final_ty + TOAST_H)

        def make_ab_illustration(canvas):
            """ステップ2: Ctrl + Alt + A を押す→実際の画面と同じく、デスクトップの
            右下にトースト通知が出て、切り替わった先のデバイス名が分かる→ループ。"""
            set_caption = make_caption(canvas)
            show_ring, clear_ring = make_focus_ring(canvas)
            press, release = draw_keycaps(canvas, ["Ctrl", "Alt", "A"], ILLO_W / 2, 24)
            _, _, taskbar_h = make_desktop_mockup(canvas)
            show_toast, hide_toast, toast_coords = make_toast(canvas, taskbar_h)
            state = {"device": "Speakers"}

            def cycle():
                if not canvas.winfo_exists():
                    return
                clear_ring()
                set_caption(1, "Hotkey (Ctrl + Alt + A) を押すと…")
                later(canvas, 900, do_press)

            def do_press():
                press()
                later(canvas, 450, do_release)

            def do_release():
                release()
                state["device"] = "Headphones" if state["device"] == "Speakers" else "Speakers"
                set_caption(2, f"右下に通知が出て、{state['device']} に切り替わる")
                show_toast(state["device"], lambda: show_ring(*toast_coords))
                later(canvas, 2600, finish)

            def finish():
                hide_toast()
                clear_ring()
                later(canvas, 500, cycle)

            later(canvas, 300, cycle)

        def make_overlay_illustration(canvas):
            """ステップ3: Ctrl + Alt + S を押す→実際のMulti-Deviceオーバーレイと
            同じく、デスクトップ全体が暗くなり中央にデバイスの一覧が出る→
            マウスで1つ選ぶ→一覧が閉じて右下に切り替え通知、を繰り返しループ。"""
            set_caption = make_caption(canvas)
            show_ring, clear_ring = make_focus_ring(canvas)
            press, release = draw_keycaps(canvas, ["Ctrl", "Alt", "S"], ILLO_W / 2, 24)
            desktop_bg, taskbar, taskbar_h = make_desktop_mockup(canvas)
            dim_color = _lerp_hex(DESKTOP_WALLPAPER, "#000000", 0.55)
            dim_taskbar_color = _lerp_hex(DESKTOP_TASKBAR, "#000000", 0.55)

            x1, y1, x2, y2 = DESK
            CARD_W, ROW_H, HEADER_H = 240, 28, 28
            labels = ["Speakers", "Headphones", "Monitor Speakers"]
            card_h = HEADER_H + len(labels) * ROW_H + 8
            card_x = x1 + ((x2 - x1) - CARD_W) / 2
            final_card_y = y1 + ((y2 - y1 - 16) - card_h) / 2
            slide_offset = 16

            card_items = [
                _rounded_rect(canvas, card_x, final_card_y, card_x + CARD_W, final_card_y + card_h, 12,
                              fill=COLOR_PANEL, outline=COLOR_BORDER, width=1),
                canvas.create_text(card_x + 14, final_card_y + HEADER_H / 2, anchor="w",
                                   text="Select output device", fill=COLOR_TEXT_SUB, font=F(7, "bold")),
                canvas.create_line(card_x + 8, final_card_y + HEADER_H, card_x + CARD_W - 8,
                                   final_card_y + HEADER_H, fill=COLOR_BORDER),
            ]
            rows = []
            for i, label in enumerate(labels):
                y0 = final_card_y + HEADER_H + i * ROW_H
                row_bg = _rounded_rect(canvas, card_x + 6, y0 + 3, card_x + CARD_W - 6, y0 + ROW_H - 3, 8,
                                       fill=COLOR_PANEL, outline="")
                text = canvas.create_text(card_x + 18, y0 + ROW_H / 2, anchor="w", text=label,
                                          fill=COLOR_TEXT_SUB, font=F(8))
                num = canvas.create_text(card_x + CARD_W - 14, y0 + ROW_H / 2, anchor="e", text=str(i + 1),
                                         fill=COLOR_TEXT_FAINT, font=F(8))
                rows.append((row_bg, text))
                card_items.extend([row_bg, text, num])
            for item in card_items:
                canvas.itemconfig(item, state="hidden")

            show_toast, hide_toast, toast_coords = make_toast(canvas, taskbar_h)
            set_cursor, move_cursor, click, hide_cursor = make_cursor(canvas)
            row_ys = [final_card_y + HEADER_H + i * ROW_H for i in range(len(labels))]
            target = {"i": 1}

            def cycle():
                if not canvas.winfo_exists():
                    return
                set_caption(1, "Hotkey (Ctrl + Alt + S) を押すと…")
                later(canvas, 900, do_press)

            def do_press():
                press()
                later(canvas, 450, show_card)

            def show_card():
                release()
                set_caption(2, "画面の中央に、つながっているデバイスの一覧が出る")
                # card_panelやrow_bgは角丸のポリゴンなので、トーストと同じく
                # canvas.move()で差分だけ動かしてスライドインさせる。
                for item in card_items:
                    canvas.itemconfig(item, state="normal")
                    canvas.move(item, 0, slide_offset)
                prev = {"v": slide_offset}

                def on_tick(e):
                    canvas.itemconfig(desktop_bg, fill=_lerp_hex(DESKTOP_WALLPAPER, dim_color, e))
                    canvas.itemconfig(taskbar, fill=_lerp_hex(DESKTOP_TASKBAR, dim_taskbar_color, e))
                    dy = slide_offset * (1 - e)
                    for item in card_items:
                        canvas.move(item, 0, dy - prev["v"])
                    prev["v"] = dy

                tween(canvas, 260, on_tick, lambda: later(canvas, 1300, point))

            def point():
                i = target["i"]
                set_caption(3, "使いたいデバイスをクリック（数字キーでも選べる）")
                set_cursor(card_x + CARD_W - 30, final_card_y - 16)
                move_cursor((card_x + 70, row_ys[i] + ROW_H / 2 - 6), lambda: select_row(i))

            def select_row(i):
                for j, (rect, text) in enumerate(rows):
                    is_sel = j == i
                    canvas.itemconfig(rect, fill=(COLOR_ACCENT_DIM if is_sel else COLOR_PANEL))
                    canvas.itemconfig(text, fill=(COLOR_TEXT_MAIN if is_sel else COLOR_TEXT_SUB))
                later(canvas, 350, lambda: (click(), later(canvas, 450, lambda: switched(i))))

            def switched(i):
                hide_cursor()
                for item in card_items:
                    canvas.itemconfig(item, state="hidden")
                for rect, text in rows:
                    canvas.itemconfig(rect, fill=COLOR_PANEL)
                    canvas.itemconfig(text, fill=COLOR_TEXT_SUB)
                canvas.itemconfig(desktop_bg, fill=DESKTOP_WALLPAPER)
                canvas.itemconfig(taskbar, fill=DESKTOP_TASKBAR)
                set_caption(4, f"一覧が閉じて、{labels[i]} に切り替わる")
                show_toast(labels[i], lambda: show_ring(*toast_coords))
                later(canvas, 2400, finish)

            def finish():
                hide_toast()
                clear_ring()
                target["i"] = 2 if target["i"] == 1 else 1
                later(canvas, 500, cycle)

            later(canvas, 300, cycle)

        def make_done_illustration(canvas):
            """ステップ4: アプリアイコンのヘッドホンがふわっと浮かび、
            チェックマークが描かれる。"""
            cx, cy = ILLO_W / 2, 112
            glyph = draw_headphones(canvas, cx, cy, 130, COLOR_PANEL)
            check_bg = canvas.create_oval(cx - 19, cy + 22, cx + 19, cy + 60, fill=COLOR_PANEL, outline="")
            check = canvas.create_line(cx - 9, cy + 41, cx - 9, cy + 41, fill=COLOR_BG, width=4,
                                       capstyle="round", joinstyle="round")
            set_caption = make_caption(canvas)

            def fade_in(e):
                color = _lerp_hex(COLOR_PANEL, COLOR_TEXT_MAIN, e)
                for item in glyph:
                    kind = canvas.type(item)
                    if kind == "arc":
                        canvas.itemconfig(item, outline=color)
                    else:
                        canvas.itemconfig(item, fill=color)

            def draw_check():
                canvas.itemconfig(check_bg, fill=COLOR_ACCENT)

                def on_tick(e):
                    # ✓ を左下→右上の順に1本の線として描いていく
                    p1 = (cx - 9, cy + 41)
                    p2 = (cx - 2, cy + 48)
                    p3 = (cx + 10, cy + 33)
                    if e < 0.4:
                        t = e / 0.4
                        canvas.coords(check, p1[0], p1[1], p1[0] + (p2[0] - p1[0]) * t, p1[1] + (p2[1] - p1[1]) * t)
                    else:
                        t = (e - 0.4) / 0.6
                        canvas.coords(check, *p1, *p2, p2[0] + (p3[0] - p2[0]) * t, p2[1] + (p3[1] - p2[1]) * t)
                tween(canvas, 420, on_tick)

            set_caption(1, "下の Get Started で設定画面に戻り、Save & Run で完了")
            tween(canvas, 600, fade_in, lambda: later(canvas, 150, draw_check))

        steps = [
            {
                "title": "設定のしかた",
                "body": "デバイスとホットキーを選んで、Save & Run を押すだけです。\n下のアニメで、操作を1つずつ順番に見せます。",
                "illustration": make_settings_illustration,
            },
            {
                "title": "A/B をワンキーで切り替え",
                "body": "設定したホットキーを押すたびに、Device A と Device B が\n入れ替わります。切り替え先は右下の通知で分かります。",
                "illustration": make_ab_illustration,
            },
            {
                "title": "3台以上から選ぶ（Multi-Device）",
                "body": "Multi-Device を ON にすると、もう1つのホットキーで\nつながっているデバイスの一覧を出して、そこから選べます。",
                "illustration": make_overlay_illustration,
            },
            {
                "title": "準備完了",
                "body": "この説明は、右上の ? ボタンや Settings タブから\nいつでも見直せます。",
                "illustration": make_done_illustration,
            },
        ]

        def close_wizard():
            win = wiz["window"]
            if win is None:
                return
            wiz["window"] = None
            show_onboarding._open = None

            # 閉じる時もふわっと消す
            def fade(a):
                try:
                    if a <= 0:
                        win.destroy()
                        return
                    win.attributes("-alpha", a)
                    win.after(15, lambda: fade(a - 0.15))
                except Exception:
                    try:
                        win.destroy()
                    except Exception:
                        pass
            fade(0.85)

        def advance():
            if wiz["step"] >= len(steps) - 1:
                close_wizard()
                return
            wiz["step"] += 1
            render(direction=1)

        def go_back():
            if wiz["step"] > 0:
                wiz["step"] -= 1
                render(direction=-1)

        def slide_in(canvas, direction):
            """ページを移った時、イラストを進む方向から横にスッと入れる。
            (どちらに進んだのかが分かるように)"""
            offset = 40 * direction
            canvas.move("all", offset, 0)
            prev = {"v": offset}

            def on_tick(e):
                dx = offset * (1 - e)
                canvas.move("all", dx - prev["v"], 0)
                prev["v"] = dx
            tween(canvas, 280, on_tick)

        def render(direction=0):
            for w in body_frame.winfo_children():
                w.destroy()
            step = steps[wiz["step"]]
            n = len(steps)

            top = ctk.CTkFrame(body_frame, fg_color="transparent")
            top.pack(fill="x", padx=32, pady=(24, 0))
            ctk.CTkLabel(top, text=f"STEP {wiz['step'] + 1} / {n}", font=F(9, "bold"),
                         text_color=COLOR_ACCENT, anchor="w", height=18).pack(side="left")

            ctk.CTkLabel(body_frame, text=step["title"], font=FONT_DISPLAY,
                         text_color=COLOR_TEXT_MAIN, anchor="w").pack(fill="x", padx=32, pady=(2, 6))
            ctk.CTkLabel(body_frame, text=step["body"], font=FONT_BODY, text_color=COLOR_TEXT_SUB,
                         justify="left", anchor="w").pack(fill="x", padx=32, pady=(0, 16))

            illo_canvas = tk.Canvas(body_frame, width=ILLO_W, height=ILLO_H, bg=COLOR_PANEL, highlightthickness=0)
            illo_canvas.pack(pady=(0, 16))
            if step["illustration"]:
                step["illustration"](illo_canvas)
            if direction:
                slide_in(illo_canvas, direction)

            # 進捗ドット: 今のページだけ横長。ページを移るとなめらかに伸び縮みする
            active_w, inactive_w, dot_h, gap = 22, 8, 8, 6
            dots_w = active_w + inactive_w * (n - 1) + gap * (n - 1) + 8
            dots_canvas = tk.Canvas(body_frame, width=dots_w, height=dot_h + 4, bg=COLOR_PANEL, highlightthickness=0)
            dots_canvas.pack(pady=(0, 18))
            prev_step = wiz["step"] - direction

            def draw_dots(e):
                dots_canvas.delete("all")
                dx = 4
                for i in range(n):
                    w_now = active_w if i == wiz["step"] else inactive_w
                    w_prev = active_w if i == prev_step else inactive_w
                    w = w_prev + (w_now - w_prev) * e
                    color = COLOR_ACCENT if i == wiz["step"] else COLOR_BORDER
                    _rounded_rect(dots_canvas, dx, 2, dx + w, 2 + dot_h, dot_h / 2, fill=color, outline="")
                    dx += w + gap
            if direction:
                tween(dots_canvas, 280, draw_dots)
            else:
                draw_dots(1)

            bottom = ctk.CTkFrame(body_frame, fg_color="transparent")
            bottom.pack(fill="x", padx=32, pady=(0, 26))
            ctk.CTkButton(
                bottom, text="Skip", command=close_wizard,
                fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
                border_width=0, corner_radius=8,
                font=FONT_LABEL, width=80, height=38,
            ).pack(side="left")

            right_frame = ctk.CTkFrame(bottom, fg_color="transparent")
            right_frame.pack(side="right")
            if wiz["step"] > 0:
                ctk.CTkButton(
                    right_frame, text="Back", command=go_back,
                    fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
                    border_width=1, border_color=COLOR_BORDER, corner_radius=8,
                    font=FONT_LABEL, width=90, height=38,
                ).pack(side="left", padx=(0, 10))
            is_last = wiz["step"] == len(steps) - 1
            ctk.CTkButton(
                right_frame, text=("Get Started" if is_last else "Next"), command=advance,
                fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
                corner_radius=8, font=FONT_LABEL, width=120, height=38,
            ).pack(side="left")

        WIZ_W = 492

        win = ctk.CTkToplevel(root)
        win.title("AudioSwitcher の使い方")
        win.resizable(False, False)
        win.transient(root)
        win.configure(fg_color=COLOR_PANEL)
        # 開く時に一瞬で出すのではなく、ふわっとフェードインさせる
        try:
            win.attributes("-alpha", 0.0)
        except Exception:
            pass
        # 閉じる方法は下部の"Skip"だけにしたいので、タイトルバーの×では
        # 閉じられないようにする(WM_DELETE_WINDOWを無視するだけにする)。
        win.protocol("WM_DELETE_WINDOW", lambda: None)
        wiz["window"] = win
        show_onboarding._open = win

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
        final_y = ry + (rh - wiz_h) // 2
        win.geometry(f"{WIZ_W}x{wiz_h}+{rx + (rw - WIZ_W) // 2}+{final_y}")

        def fade_in(a=0.0):
            try:
                if wiz["window"] is not win:
                    return
                a = min(1.0, a + 0.1)
                win.attributes("-alpha", a)
                if a < 1.0:
                    win.after(15, lambda: fade_in(a))
            except Exception:
                pass
        win.after(30, fade_in)

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