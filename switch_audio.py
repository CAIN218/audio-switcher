import os
import re
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
import urllib.parse
import hashlib
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

__version__ = "1.8.3"
GITHUB_REPO = "CAIN218/audio-switcher"

DEFAULT_OVERLAY_HOTKEY = "ctrl+alt+s"

last_switch_time = 0
_com_thread_local = threading.local()

# ----------------------------------------------------------------------
# デザイントークン — 設定画面・トースト通知・一覧オーバーレイで共有する、
# 青1色のシンプルな配色。ライト/ダークの2種類を持ち、設定の「表示テーマ」
# (ライト / ダーク / デバイスに合わせる)で切り替える。
# 各COLOR_*は apply_theme() が書き換えるモジュール変数。描く瞬間に参照
# されるので、画面を作り直せば新しい配色になる。
# ----------------------------------------------------------------------
THEME_PALETTES = {
    "dark": {
        "COLOR_BG": "#1c1c1f",          # アプリ全体の背景
        "COLOR_PANEL": "#26262a",       # サイドバー・カード
        "COLOR_INPUT": "#2f2f34",       # 入力欄・コンボボックス
        "COLOR_BORDER": "#3d3d43",      # 境界線
        "COLOR_ACCENT": "#6ea0ff",      # 差し色(青)
        "COLOR_ACCENT_HOVER": "#8db3ff",
        "COLOR_ACCENT_DIM": "#25324f",  # アクセントの控えめ版(選択行・ホバー)
        "COLOR_ON_ACCENT": "#0b1220",   # アクセント色の上に載せる文字
        "COLOR_TEXT_MAIN": "#ececef",
        "COLOR_TEXT_SUB": "#9a9aa3",
        "COLOR_TEXT_FAINT": "#6c6c75",
    },
    "light": {
        "COLOR_BG": "#f7f7f8",
        "COLOR_PANEL": "#ffffff",
        "COLOR_INPUT": "#efeff2",
        "COLOR_BORDER": "#dcdce1",
        "COLOR_ACCENT": "#2563eb",
        "COLOR_ACCENT_HOVER": "#1d4fd0",
        "COLOR_ACCENT_DIM": "#dbe6fb",
        "COLOR_ON_ACCENT": "#ffffff",
        "COLOR_TEXT_MAIN": "#1d1d20",
        "COLOR_TEXT_SUB": "#6a6a72",
        "COLOR_TEXT_FAINT": "#9a9aa3",
    },
}
# CustomTkinterの(ライト, ダーク)の組。設定画面のウィジェットに渡すと、
# set_appearance_mode()で開いたまま即座に色が切り替わる。
THEME_TUPLES = {k: (THEME_PALETTES["light"][k], THEME_PALETTES["dark"][k])
                for k in THEME_PALETTES["dark"]}
THEME_CHOICES = ("light", "dark", "system")  # 保存する値。systemは「デバイスに合わせる」
THEME = "system"
EFFECTIVE_THEME = "dark"


def _device_prefers_dark():
    """Windowsの「アプリのモード」がダークかどうか。読めない時はダーク。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return value == 0
    except Exception:
        return True


def _res(color):
    """(ライト, ダーク)の組なら今のテーマの1色に直す。素のtkウィジェットや
    Canvasは組を受け取れないので、それらに渡す時に使う。"""
    if isinstance(color, tuple):
        return color[1] if EFFECTIVE_THEME == "dark" else color[0]
    return color


def apply_theme(choice):
    """表示テーマ(light/dark/system)を配色に反映する。"""
    global THEME, EFFECTIVE_THEME
    THEME = choice if choice in THEME_CHOICES else "system"
    if THEME == "system":
        EFFECTIVE_THEME = "dark" if _device_prefers_dark() else "light"
    else:
        EFFECTIVE_THEME = THEME
    globals().update(THEME_PALETTES[EFFECTIVE_THEME])


apply_theme("dark")

# ----------------------------------------------------------------------
# 表示言語 — 英語(既定)と日本語を切り替えられる。画面上の文言はすべて
# L("English", "日本語") の形で両方を並べて書き、表示する瞬間に選ぶ。
# ----------------------------------------------------------------------
LANG = "en"


def L(en, ja):
    """今の表示言語に合わせて、英語/日本語のどちらかを返す"""
    return ja if LANG == "ja" else en


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


def draw_speaker(canvas, cx, cy, size, color):
    """サイトのアイコンに描かれているスピーカーの図(本体+音の波2本)。
    元の図形は幅約148・高さ約132なので、sizeに合わせて縮小して描く。"""
    s = size / 148.0
    pts = [(-70, -26), (-40, -26), (0, -62), (0, 62), (-40, 26), (-70, 26)]
    items = [canvas.create_polygon([v for x, y in pts for v in (cx + x * s, cy + y * s)],
                                   fill=color, outline=color, width=max(1, 8 * s), joinstyle="round")]
    lw = max(1.5, 14 * s)
    for ox, r, ang in ((-3.4, 42, 45.6), (-4.3, 78, 45.9)):
        x0 = cx + ox * s
        items.append(canvas.create_arc(x0 - r * s, cy - r * s, x0 + r * s, cy + r * s,
                                       start=-ang, extent=ang * 2, style="arc", outline=color, width=lw))
    return items


_HEADPHONE_WORDS = ("headphone", "headset", "earphone", "earbud", "buds", "airpods", "ヘッドホン",
                    "ヘッドフォン", "ヘッドセット", "イヤホン")


def looks_like_headphones(device_name):
    """デバイス名からヘッドホン系かどうかを推測する(通知のアイコンの出し分け用)"""
    name = (device_name or "").lower()
    return any(word in name for word in _HEADPHONE_WORDS)


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

def load_language():
    """表示言語(設定ファイルの10行目、"en"/"ja")。10行目が無い古い設定
    ファイルや、設定ファイル自体が無い初回起動時は英語にする。"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
            if len(lines) > 9 and lines[9].strip() in ("en", "ja"):
                return lines[9].strip()
    except Exception as e:
        log_error("load_language", e)
    return "en"


def load_theme():
    """表示テーマ(設定ファイルの11行目、"light"/"dark"/"system")。無ければ
    「デバイスに合わせる」。"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
            if len(lines) > 10 and lines[10].strip() in THEME_CHOICES:
                return lines[10].strip()
    except Exception as e:
        log_error("load_theme", e)
    return "system"


def _write_config_file(dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled, ab_enabled,
                       auto_update_enabled, overlay_excluded):
    """10行目に表示言語(今のLANG)、11行目に表示テーマ(今のTHEME)を書く以外は、従来と同じ9行の形式"""
    excluded_json = json.dumps(list(overlay_excluded), ensure_ascii=False)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(f"{dev_a}\n{dev_b}\n{key}\n{startup_val}\n{overlay_key}\n{overlay_enabled}\n{ab_enabled}\n{auto_update_enabled}\n{excluded_json}\n{LANG}\n{THEME}")


def save_language(lang):
    """言語だけを切り替えて保存する(設定画面の言語スイッチ用)。まだ一度も
    保存していない初回起動時は、ここで設定ファイルを作ってしまうと「初回」の
    判定が狂うため、メモリ上だけで切り替えておき、Save時に一緒に保存する。"""
    global LANG
    LANG = lang
    if os.path.exists(CONFIG_FILE):
        try:
            _write_config_file(*load_config(), load_overlay_excluded())
        except Exception as e:
            log_error("save_language", e)


def save_theme(choice):
    """表示テーマだけを切り替えて保存する(設定画面の表示テーマ用)。言語と同じく、
    初回起動時は設定ファイルをまだ作らない。"""
    apply_theme(choice)
    if os.path.exists(CONFIG_FILE):
        try:
            _write_config_file(*load_config(), load_overlay_excluded())
        except Exception as e:
            log_error("save_theme", e)


def save_config(dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled, ab_enabled, auto_update_enabled="0",
                overlay_excluded=None):
    """設定ファイルの保存とスタートアップ制御。overlay_excludedを渡さない時は
    今の設定ファイルにある値をそのまま残す。"""
    if overlay_excluded is None:
        overlay_excluded = load_overlay_excluded()
    _write_config_file(dev_a, dev_b, key, startup_val, overlay_key, overlay_enabled, ab_enabled,
                       auto_update_enabled, overlay_excluded)
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
        tk.messagebox.showerror(L("Startup Error", "スタートアップのエラー"),
                                L(f"Failed to update startup folder: {str(e)}", f"スタートアップの設定に失敗しました: {str(e)}"))

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
    WIDTH, HEIGHT = 340, 76
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
        self._small_font = tkfont.Font(root=root, font=F(10))
        self.theme_source = lambda: THEME  # 常駐側はload_themeに差し替える
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
        x = right - self.WIDTH - 10
        y = bottom - self.HEIGHT - 10
        self._base_xy = (x, y)
        self._apply_position()

    def _apply_position(self):
        x, y = getattr(self, "_base_xy", (0, 0))
        offset = int(round(getattr(self, "_slide", 0)))
        self.window.geometry(f"{self.WIDTH}x{self.HEIGHT}+{x}+{y + offset}")

    def _refresh_theme(self):
        """表示のたびに配色を引き直す。設定で変えたテーマや、Windowsの
        アプリモードの切り替えが、再起動なしで次の通知から反映される。"""
        try:
            apply_theme(self.theme_source())
            if self.theme_source is load_theme:
                global LANG
                LANG = load_language()  # 常駐側は、設定で変えた言語も次の通知から反映する
        except Exception as e:
            log_error("toast_refresh_theme", e)

    def _fit_text(self, text, width):
        """幅に収まらない文字を「…」で切る(前のデバイス名用)"""
        if self._small_font.measure(text) <= width:
            return text
        while text and self._small_font.measure(text + "…") > width:
            text = text[:-1]
        return text + "…"

    def show(self, device_name, previous=None):
        self._refresh_theme()
        self._ensure_window()
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None
        if self.marquee_job:
            self.root.after_cancel(self.marquee_job)
            self.marquee_job = None

        canvas = self.canvas
        canvas.delete("all")
        # 前のデバイスを薄く上に、切り替え先を太字で下に並べて、何から何へ
        # 変わったかが分かるようにする。前が分からない時は切り替え先だけ。
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 16,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        dot_x, text_x = 24, 42
        text_w = self.WIDTH - 18 - text_x
        if previous and previous != device_name:
            prev_y, new_y = 27, 51
            canvas.create_oval(dot_x - 4, prev_y - 4, dot_x + 4, prev_y + 4, fill=COLOR_TEXT_FAINT, outline="")
            canvas.create_text(text_x, prev_y, anchor="w", text=self._fit_text(previous, text_w),
                               fill=COLOR_TEXT_FAINT, font=F(10))
        else:
            new_y = self.HEIGHT / 2
        canvas.create_oval(dot_x - 5, new_y - 5, dot_x + 5, new_y + 5, fill=COLOR_ACCENT, outline="")

        # 長いデバイス名は「...」で切らず、収まりきらない分だけ左右に
        # ゆっくり往復させて全体を読めるようにする(Canvasは自身の外側を
        # 自動的に描画しないので、はみ出した部分は特別なクリップ処理
        # なしでそのまま隠れる)。
        # 文字は専用の小さなCanvasに描く。メインのCanvasに直接描くと、長い名前を
        # スクロールさせた時に左のアイコンの上や枠の外まで文字がはみ出すため。
        text_canvas = tk.Canvas(canvas, width=text_w, height=26, bg=COLOR_PANEL, highlightthickness=0)
        text_canvas.bind("<Button-1>", lambda e: self.on_click() if self.on_click else None)
        canvas.create_window(text_x, new_y, anchor="w", window=text_canvas)
        text_item = text_canvas.create_text(0, 13, anchor="w", text=device_name,
                                            fill=COLOR_TEXT_MAIN, font=self.DEVICE_FONT)
        text_width = self._device_font.measure(device_name)
        if text_width > text_w:
            self._start_marquee(text_item, text_width, 0, text_w, text_canvas, 13)

        canvas.configure(cursor="")
        self._progress_item = None

        self.on_click = None
        self._update_subtitle_id = None
        self._hold_ms = 2600
        self._place_bottom_right()
        self.window.deiconify()
        self._place_bottom_right()
        self.current_alpha = 0.0
        self._fade_in()

    TEXT_X = 74
    SUB_Y = 27
    MAIN_Y = 47

    def _draw_toast_frame(self, canvas, header_text, glyph="headphones"):
        """トースト共通の土台: 角丸パネル + 左にカッパーの丸いアイコン
        (切り替え先がヘッドホンならヘッドホン、それ以外はスピーカー)+ 小さな見出し。
        以前の「左端の線+英字の見出し」はただの帯に見えて味気なかったため、
        何に切り替わったかがアイコンでひと目で分かる形にした。"""
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 20,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        cy = self.HEIGHT / 2
        canvas.create_oval(16, cy - 22, 60, cy + 22, fill=COLOR_ACCENT, outline="")
        if glyph == "speaker":
            draw_speaker(canvas, 38, cy, 24, COLOR_ON_ACCENT)
        else:
            draw_headphones(canvas, 38, cy, 24, COLOR_ON_ACCENT)
        canvas.create_text(self.TEXT_X, self.SUB_Y, anchor="w", text=header_text,
                           fill=COLOR_TEXT_SUB, font=F(9))

    def _start_marquee(self, item, text_width, left_x, right_bound, canvas=None, text_y=None):
        """収まりきらないデバイス名を、はみ出した分だけ左へスクロールして
        端で少し止まり、また右へ戻って止まる、を繰り返す。"""
        canvas = canvas or self.canvas
        text_y = self.MAIN_Y if text_y is None else text_y
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
            canvas.coords(item, state["x"], text_y)
            self.marquee_job = self.root.after(40, tick)

        self.marquee_job = self.root.after(40, tick)

    def show_update_available(self, latest_tag, release_url, installer_url=None, notes=""):
        """新しいバージョンが見つかった時の控えめな通知。installer_urlがあれば
        クリックでインストーラーを直接ダウンロード・自動起動する。なければ
        (ポータブル版のみのリリース等)従来通りリリースページを開く。頻繁に
        出るものではないので、通常のトーストより長めに表示したままにする。"""
        self._refresh_theme()
        self._ensure_window()
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None
        if self.marquee_job:
            self.root.after_cancel(self.marquee_job)
            self.marquee_job = None

        canvas = self.canvas
        canvas.delete("all")
        self._progress_item = None
        self._draw_toast_frame(canvas, L("Update available", "アップデートがあります"))
        auto = self.auto_update_enabled and installer_url
        if auto:
            subtitle = L(f"{latest_tag} — installing automatically...", f"{latest_tag} — 自動で更新しています...")
        elif installer_url:
            subtitle = L(f"{latest_tag} — click to install", f"{latest_tag} — クリックで更新")
        else:
            subtitle = L(f"{latest_tag} — click to view", f"{latest_tag} — クリックで詳細を表示")
        self._update_subtitle_id = canvas.create_text(self.TEXT_X, self.MAIN_Y, anchor="w", text=subtitle,
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
                self.canvas.itemconfigure(self._update_subtitle_id, text=L("Downloading installer...", "ダウンロードしています..."))
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
                    msg = L("Installing update — AudioSwitcher will restart automatically.", "更新しています — 自動で再起動します")
                    self.canvas.itemconfigure(self._update_subtitle_id, text=msg)
                else:
                    self.canvas.itemconfigure(self._update_subtitle_id, text=L("Download failed — click to view", "更新に失敗 — クリックで詳細を表示"))
                    self.canvas.configure(cursor="hand2")
                    self.on_click = lambda: webbrowser.open(self._pending_release_url)
        except Exception:
            pass
        self._hold_ms = 5000
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
        self.fade_job = self.root.after(self._hold_ms, self._fade_out)

    def _fade_in(self):
        # フェードしながら下から少しせり上がる(_slideは表示開始時に14px)
        if self.current_alpha == 0.0:
            self._slide = 14
        if self.current_alpha < 0.98:
            self.current_alpha += 0.1
            self._slide = max(0.0, self._slide * 0.72)
            self.window.attributes("-alpha", min(self.current_alpha, 0.98))
            self._apply_position()
            self.fade_job = self.root.after(15, self._fade_in)
        else:
            self._slide = 0
            self._apply_position()
            if getattr(self, "_progress_item", None):
                self._hold_started = time.time()
                self._hold_tick()
            else:
                self.fade_job = self.root.after(self._hold_ms, self._fade_out)

    def _hold_tick(self):
        """表示中、残り時間のラインを右から縮めていき、なくなったら消す"""
        elapsed = (time.time() - self._hold_started) * 1000
        ratio = max(0.0, 1.0 - elapsed / self._hold_ms)
        try:
            x1, y, x2, _ = self.canvas.coords(self._progress_item)
            full = self.WIDTH - 14 - self.TEXT_X
            self.canvas.coords(self._progress_item, x1, y, x1 + full * ratio, y)
        except Exception:
            pass
        if ratio <= 0:
            self._fade_out()
        else:
            self.fade_job = self.root.after(30, self._hold_tick)

    def _fade_out(self):
        if self.current_alpha > 0.0:
            self.current_alpha -= 0.08
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
        self._refresh_theme()
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
        canvas.create_text(58, 26, anchor="w", text=L("Updated", "アップデートしました"), fill=COLOR_ACCENT, font=F(9, "bold"))
        canvas.create_text(58, 45, anchor="w", text=L(f"Now on {version}", f"{version} になりました"),
                           fill=COLOR_TEXT_MAIN, font=F(13, "bold"))
        canvas.create_line(18, 66, W - 18, 66, fill=COLOR_BORDER)

        # Markdownの見出し(##)と箇条書き(-)だけを簡単に見た目調整して並べる
        y = 78
        truncated = False
        body = (notes or "").strip() or L("No release notes.", "リリースノートはありません。")
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
                canvas.create_text(22, y, anchor="nw", text=L("…more in the release notes on GitHub", "…ほかの変更は GitHub のリリースノートへ"),
                                   fill=COLOR_TEXT_FAINT, font=F(9))
                y += 20
                truncated = True
                break
            y = bottom + 5
        H = min(MAX_H, y + FOOTER_H)

        footer_y = H - 24
        _rounded_rect(canvas, W - 98, footer_y - 13, W - 18, footer_y + 13, 8,
                      fill=COLOR_INPUT, outline="")
        canvas.create_text(W - 58, footer_y, text=L("Got it", "閉じる"), fill=COLOR_TEXT_MAIN, font=F(9, "bold"))
        canvas.create_text(22, footer_y, anchor="w", text=L("Click anywhere to close", "クリックで閉じる"),
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
                            text=L("Select output device", "出力デバイスを選択"), fill=COLOR_TEXT_SUB,
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


_CHECKSUM_ASSET = "SHA256SUMS"


def _is_github_release_url(url):
    """更新ファイルは必ずGitHubのリリースアセットから取る。それ以外のホスト・
    http:は(APIの応答が改ざんされていても)ダウンロードしない。"""
    try:
        p = urllib.parse.urlparse(url)
    except Exception:
        return False
    return p.scheme == "https" and p.hostname == "github.com" and p.path.lower().startswith(f"/{GITHUB_REPO.lower()}/releases/download/")


def _has_checksum_asset(assets):
    return any(a.get("name", "") == _CHECKSUM_ASSET for a in assets or [])


def _download_verified(url, dest_path, timeout):
    """urlをdest_pathに保存する。同じリリースのSHA256SUMSに載っているハッシュと
    一致しなければ保存したファイルを消して例外を投げる(実行も展開もしない)。"""
    if not _is_github_release_url(url):
        raise RuntimeError("untrusted update URL")
    headers = {"User-Agent": "AudioSwitcher-update-check"}
    sums_url = url.rsplit("/", 1)[0] + "/" + _CHECKSUM_ASSET
    with urllib.request.urlopen(urllib.request.Request(sums_url, headers=headers), timeout=10) as resp:
        sums = resp.read().decode("utf-8", "replace")
    name = url.rsplit("/", 1)[1]
    expected = None
    for line in sums.splitlines():
        parts = line.split()
        if len(parts) == 2 and parts[1].lstrip("*") == name:
            expected = parts[0].lower()
            break
    if not expected:
        raise RuntimeError("no checksum for update file")
    h = hashlib.sha256()
    with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=timeout) as resp, \
            open(dest_path, "wb") as f:
        for chunk in iter(lambda: resp.read(65536), b""):
            h.update(chunk)
            f.write(chunk)
    if h.hexdigest() != expected:
        try:
            os.remove(dest_path)
        except OSError:
            pass
        raise RuntimeError("update checksum mismatch")


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
    # SHA256SUMSが無いリリースは自動インストールせず、リリースページを開く
    # フォールバックに任せる(検証できないファイルは実行しない)。
    if _has_checksum_asset(data.get("assets")):
        for asset in data.get("assets", []) or []:
            name = asset.get("name", "")
            if name.lower().startswith("audioswitcher-setup") and name.lower().endswith(".exe"):
                url = asset.get("browser_download_url")
                if _is_github_release_url(url or ""):
                    installer_url = url
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


def pick_release_notes(notes):
    """リリースノートから今の言語の分を選ぶ。形式は、英語の本文の後に
    `<!-- ja -->` だけの行を置き、その後ろに日本語を続ける。目印が無ければ
    そのまま返す。日本語側が空なら英語にする。インストール手順の節は、更新
    した後に見せても意味がないので外す。"""
    text = notes or ""
    parts = re.split(r"(?m)^[ \t]*<!--[ \t]*ja[ \t]*-->[ \t]*$", text, maxsplit=1)
    if len(parts) == 2:
        en, ja = parts
        text = ja if (LANG == "ja" and ja.strip()) else en
    out, skip = [], False
    for line in text.splitlines():
        if line.startswith("## "):
            skip = line[3:].strip().lower() in ("install", "インストール")
        if not skip:
            out.append(line)
    return "\n".join(out).strip()


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
    return notice.get("version", ""), pick_release_notes(notice.get("notes", ""))


def _download_and_launch_installer(installer_url, release_url, latest_tag="", notes=""):
    """アップデート通知のクリック(または自動更新)を受けて、インストーラー
    をダウンロードし、ウィザード画面を出さずサイレントインストールする。
    installer.iss側の[Code]がサイレント時も明示的にアプリを再起動するので、
    ここでは起動するところまでで良い。ネットワークエラー等は結果を
    _install_status_queue経由でメインスレッドに伝え、失敗時はリリース
    ページを開くフォールバックに切り替えてもらう。"""
    try:
        installer_path = os.path.join(tempfile.mkdtemp(prefix="audioswitcher-update-"), os.path.basename(installer_url))
        _download_verified(installer_url, installer_path, 30)
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
                _toast_queue.put((target_display_name, current_default))
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
            try:
                previous = get_device_info()[1]
            except Exception:
                previous = None
            ensure_com_initialized()
            set_default_device_all_roles(target_id)
            _toast_queue.put((target_display_name, previous))
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
    toast_manager.theme_source = load_theme
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
                item = _toast_queue.get_nowait()
                name, previous = item if isinstance(item, tuple) else (item, None)
                toast_manager.show(name, previous)
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

def open_setting_window(prefill=None):
    """prefill: 言語を切り替えて画面を作り直す時に、保存前の入力内容と開いて
    いたタブを引き継ぐための辞書(通常の起動ではNone)。
    2重起動防止・ハイブリッドスタートアップ自動生成付き設定UI。
    A/B SWITCH(既存のA/Bトグル)とMulti-Device(3台以上から選ぶオーバーレイ)を
    タブで分けている。ひとつの画面に全部詰め込むとごちゃつくため。"""
    # この画面の中では色を(ライト, ダーク)の組として扱う。CustomTkinterの部品は
    # 組を受け取り、テーマを変えると開いたまま即座に色が切り替わる。素のtk部品
    # (Canvasなど)に渡す所だけ _res() で今のテーマの1色に直す。
    COLOR_BG = THEME_TUPLES["COLOR_BG"]
    COLOR_PANEL = THEME_TUPLES["COLOR_PANEL"]
    COLOR_INPUT = THEME_TUPLES["COLOR_INPUT"]
    COLOR_BORDER = THEME_TUPLES["COLOR_BORDER"]
    COLOR_ACCENT = THEME_TUPLES["COLOR_ACCENT"]
    COLOR_ACCENT_HOVER = THEME_TUPLES["COLOR_ACCENT_HOVER"]
    COLOR_ACCENT_DIM = THEME_TUPLES["COLOR_ACCENT_DIM"]
    COLOR_ON_ACCENT = THEME_TUPLES["COLOR_ON_ACCENT"]
    COLOR_TEXT_MAIN = THEME_TUPLES["COLOR_TEXT_MAIN"]
    COLOR_TEXT_SUB = THEME_TUPLES["COLOR_TEXT_SUB"]
    COLOR_TEXT_FAINT = THEME_TUPLES["COLOR_TEXT_FAINT"]

    import customtkinter as ctk  # 設定画面を開く時だけ読み込む（バックグラウンド常駐時は不要なため）

    is_first_run = not os.path.exists(CONFIG_FILE) and prefill is None

    kill_previous_instances()

    device_dict, current_default = get_device_info()
    devices = list(device_dict.keys())

    if not devices:
        root_err = tk.Tk()
        root_err.withdraw()
        tk.messagebox.showerror(
            L("Error", "エラー"),
            L("Couldn't get the playback devices.\nPlease check your audio driver.",
              "再生デバイスを取得できませんでした。\nオーディオドライバの状態を確認してください。")
        )
        devices = ["Speakers", "Headphones"]

    saved_a, saved_b, saved_key, saved_startup, saved_overlay_key, saved_overlay_enabled, saved_ab_enabled, saved_auto_update = load_config()
    if prefill:
        (saved_a, saved_b, saved_key, saved_startup, saved_overlay_key, saved_overlay_enabled,
         saved_ab_enabled, saved_auto_update) = prefill["values"]
    devices_b_options = [L("Select device...", "デバイスを選択...")] + devices

    apply_theme(THEME)
    ctk.set_appearance_mode(EFFECTIVE_THEME)

    root = ctk.CTk(fg_color=COLOR_BG)
    resolve_font_family(root)
    # Settingsタブの中に置くスイッチだが、変数自体はSave時に他の設定と一緒に
    # 保存する必要があるので、タブの有無に関わらずここで(root作成後に)
    # 用意しておく(StringVarはデフォルトのTkルートが無いと作れないため)。
    auto_update_var = tk.StringVar(value=saved_auto_update)
    root.title("AudioSwitcher")
    root.geometry("700x600")
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
        head_icon = tk.Canvas(header, width=44, height=44, bg=_res(COLOR_PANEL), highlightthickness=0)
        head_icon.pack(side="left", padx=(0, 14))
        _rounded_rect(head_icon, 0, 0, 44, 44, 12, fill=_res(COLOR_INPUT), outline="")
        draw_headphones(head_icon, 22, 22, 28, _res(COLOR_ACCENT))
        head_text = ctk.CTkFrame(header, fg_color="transparent")
        head_text.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(head_text, text=L("✓ Updated", "✓ アップデートしました"), font=F(9, "bold"),
                     text_color=COLOR_ACCENT, anchor="w", height=18).pack(fill="x")
        ctk.CTkLabel(head_text, text=L(f"What's new in {version}", f"{version} の更新内容"), font=F(17, "bold"),
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
            footer, text=L("See all on GitHub", "GitHub で全部見る"), command=lambda: webbrowser.open(release_url),
            fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
            border_width=0, font=FONT_LABEL, width=130, height=38,
        ).pack(side="left")
        ctk.CTkButton(
            footer, text=L("Got it", "閉じる"), command=close_overlay,
            fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_ON_ACCENT,
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
            if rec["value"] and rec["value"] != L("Press keys...", "キーを押してください..."):
                rec["recording"] = False
                root.focus_set()
            return "break"

        def on_focus_in(e):
            rec["recording"] = True
            rec["keys"] = set()
            key_box.configure(border_color=COLOR_ACCENT)
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.insert(0, L("Press keys...", "キーを押してください..."))
            entry.configure(state="readonly", text_color=COLOR_ACCENT)

        def on_focus_out(e):
            rec["recording"] = False
            key_box.configure(border_color=COLOR_INPUT)
            entry.configure(text_color=COLOR_TEXT_MAIN)
            entry.configure(state="normal")
            entry.delete(0, "end")
            if rec["value"] == L("Press keys...", "キーを押してください...") or not rec["value"]:
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

    # レイアウト: 左にメニュー(切り替え / 一覧 / 全般)、右にそのページ、下に
    # 「閉じる」(保存して常駐)。変更は開いている間も自動で保存される。
    shell = ctk.CTkFrame(root, fg_color="transparent", corner_radius=0)
    shell.pack(fill="both", expand=True)

    sidebar = ctk.CTkFrame(shell, fg_color=COLOR_PANEL, corner_radius=0, width=176)
    sidebar.pack(side="left", fill="y")
    sidebar.pack_propagate(False)
    ctk.CTkFrame(shell, fg_color=COLOR_BORDER, corner_radius=0, width=1).pack(side="left", fill="y")

    brand = ctk.CTkFrame(sidebar, fg_color="transparent")
    brand.pack(fill="x", padx=16, pady=(20, 14))
    logo = tk.Canvas(brand, width=28, height=28, bg=_res(COLOR_PANEL), highlightthickness=0)
    logo.pack(side="left", padx=(0, 8))
    draw_headphones(logo, 14, 14, 24, _res(COLOR_ACCENT))

    def redraw_logo():
        logo.configure(bg=_res(COLOR_PANEL))
        logo.delete("all")
        draw_headphones(logo, 14, 14, 24, _res(COLOR_ACCENT))
    ctk.CTkLabel(brand, text="AudioSwitcher", font=F(12, "bold"), text_color=COLOR_TEXT_MAIN,
                 anchor="w").pack(side="left")

    main_area = ctk.CTkFrame(shell, fg_color="transparent", corner_radius=0)
    main_area.pack(side="left", fill="both", expand=True)

    # 下の帯(状態表示 + 閉じるボタン)。中身はウィジェットを作り終えてから入れる。
    # 先にside="bottom"で確保しておかないと、ページが広がって押し出されてしまう。
    bottom_bar = ctk.CTkFrame(main_area, fg_color="transparent", corner_radius=0)
    bottom_bar.pack(side="bottom", fill="x")
    ctk.CTkFrame(bottom_bar, fg_color=COLOR_BORDER, corner_radius=0, height=1).pack(fill="x")

    page_title = ctk.CTkLabel(main_area, text="", font=F(16, "bold"), text_color=COLOR_TEXT_MAIN, anchor="w")
    page_title.pack(fill="x", padx=28, pady=(22, 6))
    pages_box = ctk.CTkFrame(main_area, fg_color="transparent", corner_radius=0)
    pages_box.pack(fill="both", expand=True, padx=8)

    page_names = [L("A/B Switch", "切り替え"), L("Device list", "デバイス一覧"), L("General", "全般")]
    tab_general = ctk.CTkFrame(pages_box, fg_color="transparent", corner_radius=0)
    tab_multi = ctk.CTkFrame(pages_box, fg_color="transparent", corner_radius=0)
    tab_settings = ctk.CTkFrame(pages_box, fg_color="transparent", corner_radius=0)
    page_frames = [tab_general, tab_multi, tab_settings]
    nav_buttons = []

    def show_page(index):
        for frame in page_frames:
            frame.pack_forget()
        page_frames[index].pack(fill="both", expand=True)
        page_title.configure(text=page_names[index])
        for i, btn in enumerate(nav_buttons):
            btn.configure(fg_color=(COLOR_ACCENT_DIM if i == index else "transparent"),
                          text_color=(COLOR_TEXT_MAIN if i == index else COLOR_TEXT_SUB))

    for i, name in enumerate(page_names):
        btn = ctk.CTkButton(
            sidebar, text=name, anchor="w", command=lambda i=i: show_page(i),
            fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
            corner_radius=8, font=F(11, "bold"), height=36,
        )
        btn.pack(fill="x", padx=10, pady=1)
        nav_buttons.append(btn)

    ctk.CTkButton(
        sidebar, text=L("How to use", "使い方"), anchor="w", command=lambda: show_onboarding(),
        fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
        corner_radius=8, font=FONT_BODY, height=34,
    ).pack(side="bottom", fill="x", padx=10, pady=(0, 14))

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
            button_color="#ffffff", button_hover_color="#ffffff",
        )

    # ---- Settingsタブ: 起動・アップデート・バージョン情報 ----
    # 以前はボトムバーにあった「Start on boot」と、UPDATEタブの中身を
    # ここにまとめた(A/B・Multi-Deviceのタブを本来の設定だけにするため)。
    settings_box = ctk.CTkFrame(tab_settings, fg_color="transparent")
    settings_box.pack(fill="both", expand=True)

    section_label(settings_box, L("General", "一般"))
    startup_var = tk.StringVar(value=saved_startup)
    switch_startup = make_switch(settings_box, L("Start on boot (runs in the background)", "Windows 起動時に自動で常駐する"), startup_var)
    switch_startup.pack(anchor="w", padx=20, pady=(0, 4))

    # 表示言語・表示テーマ。切り替えたら、入力途中の内容を引き継いだまま画面を
    # 作り直す(全部の文言や色を個別に書き換えるより確実で、見落としが出ない)。
    reopen = {"prefill": None}

    def reopen_prefill():
        return {
            "values": (combo_a.get(), combo_b.get(), get_shortcut(), startup_var.get(),
                       get_overlay_shortcut(), overlay_enabled_var.get(), ab_enabled_var.get(),
                       auto_update_var.get()),
            "excluded": get_overlay_excluded(),
            "tab": "settings",
        }

    theme_row = ctk.CTkFrame(settings_box, fg_color="transparent")
    theme_row.pack(fill="x", padx=20, pady=(6, 4))
    ctk.CTkLabel(theme_row, text=L("Appearance", "表示テーマ"), font=FONT_BODY, text_color=COLOR_TEXT_MAIN,
                 anchor="w").pack(side="left")
    THEME_LABELS = {"light": L("Light", "ライト"), "dark": L("Dark", "ダーク"),
                    "system": L("Match device", "デバイスに合わせる")}

    def apply_theme_live(choice):
        """画面を作り直さず、開いたまま配色を切り替える"""
        guide = getattr(show_onboarding, "_open", None)
        if guide:
            guide["close"]()
        save_theme(choice)
        ctk.set_appearance_mode(EFFECTIVE_THEME)
        redraw_logo()

    def watch_device_theme():
        """「デバイスに合わせる」の時、Windowsのアプリモードが変わったら追従する"""
        if THEME == "system":
            now = "dark" if _device_prefers_dark() else "light"
            if now != EFFECTIVE_THEME:
                apply_theme("system")
                ctk.set_appearance_mode(EFFECTIVE_THEME)
                redraw_logo()
        root.after(2000, watch_device_theme)

    def on_theme_change(label):
        choice = next(k for k, v in THEME_LABELS.items() if v == label)
        if choice == THEME:
            return
        apply_theme_live(choice)

    theme_switch = ctk.CTkSegmentedButton(
        theme_row, values=list(THEME_LABELS.values()), command=on_theme_change,
        font=FONT_LABEL, height=28,
        fg_color=COLOR_INPUT, selected_color=COLOR_ACCENT_DIM, selected_hover_color=COLOR_ACCENT_DIM,
        unselected_color=COLOR_INPUT, unselected_hover_color=COLOR_BORDER,
        text_color=COLOR_TEXT_MAIN,
    )
    theme_switch.set(THEME_LABELS[THEME])
    theme_switch.pack(side="right")
    lang_row = ctk.CTkFrame(settings_box, fg_color="transparent")
    lang_row.pack(fill="x", padx=20, pady=(6, 4))
    ctk.CTkLabel(lang_row, text="Language / 言語", font=FONT_BODY, text_color=COLOR_TEXT_MAIN,
                 anchor="w").pack(side="left")
    LANG_LABELS = {"en": "English", "ja": "日本語"}

    def on_language_change(label):
        lang = next(k for k, v in LANG_LABELS.items() if v == label)
        if lang == LANG:
            return
        reopen["prefill"] = reopen_prefill()
        save_language(lang)
        root.destroy()

    lang_switch = ctk.CTkSegmentedButton(
        lang_row, values=list(LANG_LABELS.values()), command=on_language_change,
        font=FONT_LABEL, height=28,
        fg_color=COLOR_INPUT, selected_color=COLOR_ACCENT_DIM, selected_hover_color=COLOR_ACCENT_DIM,
        unselected_color=COLOR_INPUT, unselected_hover_color=COLOR_BORDER,
        text_color=COLOR_TEXT_MAIN,
    )
    lang_switch.set(LANG_LABELS[LANG])
    lang_switch.pack(side="right")

    # アップデート欄は常に表示しておく(更新がある時しか出ないと、Auto-update
    # トグル自体を普段いじれなくなってしまうため)。中身(タイトル/ボタン)は
    # チェック結果に応じて後から書き換える。
    section_label(settings_box, L("Updates", "アップデート"))
    update_panel = ctk.CTkFrame(settings_box, fg_color=COLOR_INPUT, corner_radius=10, border_width=0)
    update_panel.pack(fill="x", padx=20, pady=(0, 4))

    update_title_var = tk.StringVar(value=L("Checking for updates...", "アップデートを確認しています..."))
    ctk.CTkLabel(update_panel, textvariable=update_title_var, font=FONT_LABEL,
                 text_color=COLOR_ACCENT, anchor="w", height=20).pack(fill="x", padx=16, pady=(12, 0))
    update_detail_var = tk.StringVar(value=L(f"Current version: v{__version__}", f"現在のバージョン: v{__version__}"))
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
        install_btn.configure(state="disabled", text=L("Updating...", "更新中..."))
        update_status_var.set("")

        def worker():
            try:
                installer_path = os.path.join(tempfile.mkdtemp(prefix="audioswitcher-update-"), os.path.basename(installer_url))
                _download_verified(installer_url, installer_path, 30)
                _write_update_notice(_update_state["latest_tag"], _update_state["notes"])
                os.startfile(installer_path, arguments="/VERYSILENT /SUPPRESSMSGBOX /NORESTART")
                msg = L("Installing update in the background — AudioSwitcher will restart automatically.",
                        "バックグラウンドで更新しています — 自動で再起動します")
                root.after(0, lambda: update_status_var.set(msg))
            except Exception as e:
                log_error("settings_download_installer", e)

                def on_fail():
                    install_btn.configure(state="normal", text=L("Open release page", "リリースページを開く"),
                                           command=lambda: webbrowser.open(release_url))
                    update_status_var.set(L("Download failed. Click to open the release page instead.",
                                         "ダウンロードに失敗しました。代わりにリリースページを開けます。"))
                root.after(0, on_fail)
        threading.Thread(target=worker, daemon=True).start()

    def on_auto_toggle():
        if auto_update_var.get() == "1" and _update_state["installer_url"]:
            start_install()

    update_actions = ctk.CTkFrame(update_panel, fg_color="transparent")
    update_actions.pack(fill="x", padx=16, pady=(0, 14))
    switch_auto_update = make_switch(update_actions, L("Auto-update", "自動アップデート"), auto_update_var, on_auto_toggle)
    switch_auto_update.pack(side="left")

    install_btn = ctk.CTkButton(
        update_actions, text=L("Up to date", "最新です"), state="disabled", command=start_install,
        fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_ON_ACCENT,
        text_color_disabled=COLOR_TEXT_SUB,
        corner_radius=8, font=FONT_LABEL, width=150, height=32,
    )
    install_btn.pack(side="right")

    def refresh_install_btn_style():
        # 押せない時(最新版)はカッパーの塗りを外し、押せる時だけ目立たせる
        enabled = install_btn.cget("state") == "normal"
        install_btn.configure(fg_color=(COLOR_ACCENT if enabled else COLOR_BORDER))

    refresh_install_btn_style()

    section_label(settings_box, L("Help", "ヘルプ"))
    about_row = ctk.CTkFrame(settings_box, fg_color="transparent")
    about_row.pack(fill="x", padx=20, pady=(0, 16))
    ctk.CTkButton(
        about_row, text=L("Show the tutorial again", "使い方をもう一度見る"), command=lambda: show_onboarding(),
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
                update_title_var.set(L("Update available", "アップデートがあります"))
                update_detail_var.set(L(f"{latest_tag} is available (current: v{__version__})",
                                         f"{latest_tag} を利用できます（現在: v{__version__}）"))
                install_btn.configure(state="normal", text=L("Download & Install", "ダウンロードして更新"))
                refresh_install_btn_style()
                if auto_update_var.get() == "1" and installer_url:
                    start_install()
            else:
                update_title_var.set(L("Up to date", "最新です"))
                update_detail_var.set(L(f"You're on the latest version (v{__version__})", f"最新のバージョンです（v{__version__}）"))
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
    ctk.CTkLabel(now_text, text=L("Current output", "現在の出力"), font=FONT_SMALL, text_color=COLOR_TEXT_SUB,
                 anchor="w", height=16).pack(fill="x")
    now_row = ctk.CTkFrame(now_text, fg_color="transparent")
    now_row.pack(fill="x")
    side_badge = ctk.CTkLabel(now_row, text="A", font=F(9, "bold"), width=22, height=22,
                              corner_radius=11, fg_color=COLOR_ACCENT, text_color=COLOR_ON_ACCENT)
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
        name = current_default or L("Not detected", "検出できませんでした")
        if len(name) > 34:
            name = name[:32] + "…"
        status_label.configure(text=name)

    def on_test_switch():
        nonlocal current_default
        a, b = combo_a.get(), combo_b.get()
        if b == L("Select device...", "デバイスを選択..."):
            tk.messagebox.showwarning(L("Routing Incomplete", "設定が未完了です"),
                                       L("Please select a valid target for Device B.", "デバイス B の出力先を選んでください。"))
            return
        target_display_name = b if current_default == a else a
        target_id = device_dict.get(target_display_name)
        if not target_id:
            tk.messagebox.showerror(L("Error", "エラー"), L("The selected device wasn't found.", "選択されたデバイスが見つかりませんでした。"))
            return
        try:
            ensure_com_initialized()
            set_default_device_all_roles(target_id)
        except Exception as e:
            log_error("test_switch", e)
            tk.messagebox.showerror(L("Error", "エラー"), L(f"Switch failed: {e}", f"切り替えに失敗しました: {e}"))
            return
        previous, current_default = current_default, target_display_name
        refresh_indicator()
        settings_toast.show(target_display_name, previous)

    test_button = ctk.CTkButton(
        now_card, text=L("Test switch", "テスト切り替え"), command=on_test_switch,
        fg_color="transparent", hover_color=COLOR_BORDER, text_color=COLOR_ACCENT,
        border_width=1, border_color=COLOR_ACCENT_DIM, corner_radius=8,
        font=FONT_LABEL, width=120, height=32,
    )
    test_button.pack(side="right", padx=(0, 14))

    # デバイス設定
    container = ctk.CTkFrame(tab_general, fg_color="transparent")
    container.pack(fill="both", expand=True)

    ab_enabled_var = tk.StringVar(value=saved_ab_enabled)
    switch_ab = make_switch(container, L("Enable A/B switch", "A/B 切り替えを使う"), ab_enabled_var)
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

    label_a = section_label(container, L("Device A", "デバイス A"))
    combo_a = ctk.CTkComboBox(container, values=devices, **combo_kwargs)
    combo_a.pack(fill="x", padx=20, pady=(0, 2))
    if saved_a in devices: combo_a.set(saved_a)
    elif current_default in devices: combo_a.set(current_default)
    else: combo_a.set(devices[0])
    make_combo_fully_clickable(combo_a)

    label_b = section_label(container, L("Device B", "デバイス B"))
    combo_b = ctk.CTkComboBox(container, values=devices_b_options, **combo_kwargs)
    combo_b.pack(fill="x", padx=20, pady=(0, 2))
    if saved_b in devices_b_options and saved_b != L("Select device...", "デバイスを選択..."): combo_b.set(saved_b)
    else: combo_b.set(devices_b_options[0])
    make_combo_fully_clickable(combo_b)

    combo_a.configure(command=lambda _=None: refresh_indicator())
    combo_b.configure(command=lambda _=None: refresh_indicator())
    refresh_indicator()

    # ホットキー
    label_hotkey = section_label(container, L("Hotkey (switch A ⇄ B)", "ホットキー（A ⇄ B を切り替え）"))
    key_box, get_shortcut, set_key_enabled = make_hotkey_recorder(container, saved_key)
    key_box.pack(fill="x", padx=20, pady=(0, 4))
    hint_label(container, L("Click, then press a new key combination to change it", "クリックしてから、新しいキーの組み合わせを押すと変更できます"))

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
        text=L("Press the hotkey to show your connected output devices in the\ncenter of the screen, then pick one to switch to it.",
               "ホットキーを押すと、つながっている出力デバイスが画面中央に一覧で出て、\nそこから選んで切り替えられます。"),
        font=FONT_BODY, text_color=COLOR_TEXT_SUB, justify="left", anchor="w"
    ).pack(fill="x", padx=20, pady=(12, 8))

    overlay_enabled_var = tk.StringVar(value=saved_overlay_enabled)
    switch_overlay = make_switch(container2, L("Enable device-select overlay", "デバイス一覧を使う"), overlay_enabled_var)
    switch_overlay.pack(anchor="w", padx=20, pady=(0, 0))

    label_overlay_key = section_label(container2, L("Hotkey (open the list)", "ホットキー（一覧を開く）"))
    overlay_key_box, get_overlay_shortcut, set_overlay_key_enabled = make_hotkey_recorder(container2, saved_overlay_key)
    overlay_key_box.pack(fill="x", padx=20, pady=(0, 4))
    hint_label(container2, L("Click, then press a new key combination to change it", "クリックしてから、新しいキーの組み合わせを押すと変更できます"))

    # 一覧に出すデバイスを1台ずつ選べるようにする(デバイスが多い人向け)。
    # 覚えておくのは「出さない」デバイスなので、新しくつないだ物は最初からON。
    saved_excluded = set(prefill["excluded"] if prefill else load_overlay_excluded())
    label_devices = section_label(container2, L("Devices shown in the list", "一覧に出すデバイス"))
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
            checkmark_color=COLOR_ON_ACCENT, checkbox_width=18, checkbox_height=18, corner_radius=5, border_width=2,
        )
        check.pack(side="left", pady=4)
        device_checks.append(check)
        if is_current:
            ctk.CTkLabel(row, text=L("In use", "使用中"), font=FONT_SMALL, text_color=COLOR_ACCENT,
                         height=20).pack(side="right")

    def get_overlay_excluded():
        return [name for name, var in device_vars.items() if var.get() != "1"]

    def refresh_device_count():
        on = sum(1 for var in device_vars.values() if var.get() == "1")
        device_count_var.set(L(f"{on} / {len(devices)}", f"{on} / {len(devices)} 台"))

    refresh_device_count()
    hint_label(container2, L("Unchecked devices won't appear in the hotkey list",
                              "チェックを外したデバイスは、ホットキーで開く一覧に出なくなります"))

    def apply_overlay_enabled_style(*_):
        enabled = overlay_enabled_var.get() == "1"
        set_overlay_key_enabled(enabled)
        label_overlay_key.configure(text_color=(COLOR_TEXT_SUB if enabled else COLOR_TEXT_FAINT))
        label_devices.configure(text_color=(COLOR_TEXT_SUB if enabled else COLOR_TEXT_FAINT))
        for check in device_checks:
            check.configure(state=("normal" if enabled else "disabled"))

    switch_overlay.configure(command=apply_overlay_enabled_style)
    apply_overlay_enabled_style()

    # ボトム: 状態表示(自動保存) + 「閉じる」。閉じる = 保存して常駐を始める。
    bar_inner = ctk.CTkFrame(bottom_bar, fg_color="transparent")
    bar_inner.pack(fill="x", padx=28, pady=14)
    default_status = L("Changes are saved automatically", "変更は自動で保存されます")
    status_var = tk.StringVar(value=default_status)
    ctk.CTkLabel(bar_inner, textvariable=status_var, font=FONT_SMALL, text_color=COLOR_TEXT_SUB,
                 anchor="w").pack(side="left")

    def is_routing_incomplete():
        return ab_enabled_var.get() == "1" and combo_b.get() == L("Select device...", "デバイスを選択...")

    def current_values():
        return (combo_a.get(), combo_b.get(), get_shortcut(), startup_var.get(),
                get_overlay_shortcut(), overlay_enabled_var.get(), ab_enabled_var.get(),
                auto_update_var.get(), tuple(get_overlay_excluded()))

    def on_close():
        if is_routing_incomplete():
            tk.messagebox.showwarning(L("Routing Incomplete", "設定が未完了です"),
                                       L("Please select a valid target for Device B.", "デバイス B の出力先を選んでください。"))
            show_page(0)
            return
        values = current_values()
        save_config(*values[:8], list(values[8]))
        root.destroy()
        start_backend()

    close_button = ctk.CTkButton(
        bar_inner, text=L("Close", "閉じる"), command=on_close,
        fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_ON_ACCENT,
        corner_radius=8, font=FONT_LABEL, width=110, height=36,
    )
    close_button.pack(side="right")

    # 自動保存: 値が変わったら少し待って保存する(押し忘れで消えないように)。
    # 変更の検出を各ウィジェットのコールバックでなく定期チェックにしているのは、
    # ホットキー入力欄・スイッチ・コンボなど種類が多く、取りこぼしを避けるため。
    autosave = {"last": current_values()}

    def autosave_tick():
        try:
            values = current_values()
            if values != autosave["last"] and not is_routing_incomplete():
                autosave["last"] = values
                save_config(*values[:8], list(values[8]))
                status_var.set(L("Saved", "保存しました"))
                root.after(1500, lambda: status_var.set(default_status))
        except Exception as e:
            log_error("autosave", e)
        root.after(700, autosave_tick)

    root.after(700, autosave_tick)

    # ウィンドウ高さは中身の実測サイズに合わせて自動調整する(一番高いページに
    # 合わせ、どのページでも下の「閉じる」が切れないようにする)。
    show_page(0)
    root.update_idletasks()
    page_h = max(frame.winfo_reqheight() for frame in page_frames)
    chrome_h = page_title.winfo_reqheight() + 28 + bottom_bar.winfo_reqheight() + 16
    root.geometry(f"700x{max(560, page_h + chrome_h)}")

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
        """使い方ガイド。本物の設定画面の上に、触る場所を囲む枠と短い説明を重ねて、
        1か所ずつ案内する。ミニチュアを描く方式をやめたのは、描画が粗くなるのと、
        実際の画面と見た目がずれるため。枠と説明カードは普通のウィジェットなので
        文字も図形も画面の拡大率にそのまま合わせて鮮明に出る。"""
        if getattr(show_onboarding, "_open", None) is not None:
            try:
                show_onboarding._open["lift"]()
                return
            except Exception:
                show_onboarding._open = None

        try:
            scale = float(root._get_widget_scaling())
        except Exception:
            scale = 1.0

        def px(v):
            return int(round(v * scale))

        wiz = {"step": 0, "token": 0, "ring_rect": None}
        ring = [tk.Frame(root, bg=_res(COLOR_ACCENT)) for _ in range(4)]
        CW, CH = 276, 132
        card = tk.Frame(root, bg=_res(COLOR_PANEL), highlightthickness=1, highlightbackground=_res(COLOR_BORDER),
                        width=px(CW), height=px(CH))
        card.pack_propagate(False)
        title_label = ctk.CTkLabel(card, text="", font=F(11, "bold"), text_color=COLOR_TEXT_MAIN, anchor="w")
        title_label.pack(fill="x", padx=14, pady=(12, 0))
        body_label = ctk.CTkLabel(card, text="", font=FONT_BODY, text_color=COLOR_TEXT_SUB, anchor="nw",
                                  justify="left", width=CW - 28, wraplength=px(CW - 28))
        body_label.pack(fill="both", expand=True, padx=14, pady=(2, 0))
        foot = tk.Frame(card, bg=_res(COLOR_PANEL))
        foot.pack(fill="x", padx=12, pady=(0, 10))
        dots = tk.Canvas(foot, width=px(56), height=px(12), bg=_res(COLOR_PANEL), highlightthickness=0)
        dots.pack(side="left")
        btn_next = ctk.CTkButton(foot, text="", width=64, height=28, corner_radius=8, font=FONT_LABEL,
                                 fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER,
                                 text_color=COLOR_ON_ACCENT, command=lambda: advance())
        btn_next.pack(side="right")
        btn_back = ctk.CTkButton(foot, text=L("Back", "戻る"), width=56, height=28, corner_radius=8,
                                 font=FONT_LABEL, fg_color="transparent", hover_color=COLOR_INPUT,
                                 text_color=COLOR_TEXT_SUB, command=lambda: go_back())
        btn_back.pack(side="right", padx=(0, 4))
        btn_skip = ctk.CTkButton(card, text=L("Skip", "スキップ"), width=56, height=22, corner_radius=6,
                                 font=FONT_SMALL, fg_color="transparent", hover_color=COLOR_INPUT,
                                 text_color=COLOR_TEXT_FAINT, command=lambda: close_wizard())
        btn_skip.place(relx=1.0, x=-8, y=8, anchor="ne")

        toast = tk.Frame(root, bg=_res(COLOR_PANEL), highlightthickness=2, highlightbackground=_res(COLOR_ACCENT),
                         width=px(250), height=px(64))
        toast.pack_propagate(False)
        ctk.CTkLabel(toast, text="B", width=30, height=30, corner_radius=15, fg_color=COLOR_ACCENT,
                     text_color=COLOR_ON_ACCENT, font=F(11, "bold")).pack(side="left", padx=(12, 10))
        tbox = tk.Frame(toast, bg=_res(COLOR_PANEL))
        tbox.pack(side="left")
        ctk.CTkLabel(tbox, text=L("Output switched", "出力を切り替えました"), font=FONT_SMALL,
                     text_color=COLOR_TEXT_SUB, anchor="w", height=16).pack(anchor="w")
        ctk.CTkLabel(tbox, text="Headphones", font=F(12, "bold"), text_color=COLOR_TEXT_MAIN,
                     anchor="w", height=22).pack(anchor="w")

        def rect_of(*widgets, pad=6):
            root.update_idletasks()
            x1 = min(w.winfo_rootx() for w in widgets) - root.winfo_rootx() - pad
            y1 = min(w.winfo_rooty() for w in widgets) - root.winfo_rooty() - pad
            x2 = max(w.winfo_rootx() + w.winfo_width() for w in widgets) - root.winfo_rootx() + pad
            y2 = max(w.winfo_rooty() + w.winfo_height() for w in widgets) - root.winfo_rooty() + pad
            return (x1, y1, x2, y2)

        def place_ring(r):
            x1, y1, x2, y2 = [int(v) for v in r]
            t = max(2, px(2))
            ring[0].place(x=x1, y=y1, width=x2 - x1, height=t)
            ring[1].place(x=x1, y=y2 - t, width=x2 - x1, height=t)
            ring[2].place(x=x1, y=y1, width=t, height=y2 - y1)
            ring[3].place(x=x2 - t, y=y1, width=t, height=y2 - y1)
            for f in ring:
                f.lift()
            card.lift()

        def hide_ring():
            for f in ring:
                f.place_forget()
            wiz["ring_rect"] = None

        def move_ring(target):
            """前の枠から次の枠へ、短く滑らせて動かす"""
            wiz["token"] += 1
            token = wiz["token"]
            start = wiz["ring_rect"]
            wiz["ring_rect"] = target
            if start is None:
                place_ring(target)
                return

            def frame(i):
                if token != wiz["token"]:
                    return
                t = i / 8
                e = 1 - (1 - t) ** 3
                place_ring([a + (b - a) * e for a, b in zip(start, target)])
                if i < 8:
                    root.after(15, lambda: frame(i + 1))
            frame(1)

        def place_card(target, prefer="below"):
            rw, rh = root.winfo_width(), root.winfo_height()
            cw, ch = px(CW), px(CH)
            if target is None:
                x, y = px(205), px(250)
            else:
                x1, y1, x2, y2 = target
                x = max(px(190), min(x1, rw - cw - px(16)))
                below, above = y2 + px(12), y1 - ch - px(12)
                if prefer == "below" and below + ch <= rh - px(8):
                    y = below
                elif above >= px(50):
                    y = above
                else:
                    y = max(px(50), min(below, rh - ch - px(8)))
            card.place(x=x, y=y, width=cw, height=ch)
            card.lift()

        def show_toast():
            rw, rh = root.winfo_width(), root.winfo_height()
            tw = px(250)
            fx, y = rw - tw - px(24), rh - px(64) - px(90)
            token = wiz["token"]

            def frame(i):
                if token != wiz["token"]:
                    return
                e = 1 - (1 - i / 12) ** 3
                toast.place(x=int(rw - (rw - fx) * e), y=y, width=tw, height=px(64))
                toast.lift()
                if i < 12:
                    root.after(15, lambda: frame(i + 1))
            frame(1)

        steps = [
            {"title": L("Choose two devices", "切り替える2台を選ぶ"),
             "body": L("Pick Device A and Device B. The hotkey switches between them.",
                       "デバイスAとBを選びます。ホットキーで、この2台を入れ替えます。"),
             "page": 0, "target": lambda: rect_of(label_a, combo_b)},
            {"title": L("Set a hotkey", "ホットキーを決める"),
             "body": L("Click the box, then press the keys you want to use.",
                       "枠をクリックして、使いたいキーを押します。"),
             "page": 0, "target": lambda: rect_of(key_box)},
            {"title": L("Switch with one key", "ワンキーで切り替え"),
             "body": L("Each press swaps A and B. The device you switched to appears at the bottom right.",
                       "押すたびにAとBが入れ替わり、切り替え先が右下に出ます。"),
             "page": 0, "target": None, "toast": True},
            {"title": L("Pick from 3+ devices", "3台以上から選ぶ"),
             "body": L("Turn on Device list to open a list with a second hotkey and pick one.",
                       "「デバイス一覧」をONにすると、別のホットキーで一覧が出て、そこから選べます。"),
             "page": 1, "target": lambda: rect_of(switch_overlay, overlay_key_box)},
            {"title": L("That's all", "準備完了"),
             "body": L("Press Close to save and start. Open this guide again from “How to use”.",
                       "「閉じる」を押すと保存されて、常駐が始まります。この説明は「使い方」からまた見られます。"),
             "page": 1, "target": lambda: rect_of(close_button, pad=4), "prefer": "above"},
        ]

        def draw_dots():
            dots.delete("all")
            n = len(steps)
            for i in range(n):
                on = i == wiz["step"]
                cx = px(6 + i * 11)
                r = px(3.5 if on else 2.5)
                dots.create_oval(cx - r, px(6) - r, cx + r, px(6) + r,
                                 fill=_res(COLOR_ACCENT if on else COLOR_BORDER), outline="")

        def close_wizard():
            wiz["token"] += 1
            show_onboarding._open = None
            for w in (*ring, card, toast):
                try:
                    w.destroy()
                except Exception:
                    pass

        def play():
            wiz["token"] += 1
            step = steps[wiz["step"]]
            root.update_idletasks()
            show_page(step["page"])
            root.update_idletasks()
            title_label.configure(text=step["title"])
            body_label.configure(text=step["body"])
            btn_back.configure(state=("normal" if wiz["step"] else "disabled"))
            btn_next.configure(text=(L("Start", "はじめる") if wiz["step"] == len(steps) - 1 else L("Next", "次へ")))
            draw_dots()
            toast.place_forget()
            target = step["target"]() if step["target"] else None
            if target is None:
                hide_ring()
            else:
                move_ring(target)
            place_card(target, step.get("prefer", "below"))
            if step.get("toast"):
                root.after(300, show_toast)

        def advance():
            if wiz["step"] >= len(steps) - 1:
                close_wizard()
                return
            wiz["step"] += 1
            play()

        def go_back():
            if wiz["step"] > 0:
                wiz["step"] -= 1
                play()

        root.bind("<Escape>", lambda e: close_wizard() if show_onboarding._open else None)
        show_onboarding._open = {"lift": lambda: card.lift(), "close": close_wizard}
        root.update()
        play()

    root.after(2000, watch_device_theme)
    if is_first_run:
        show_onboarding()
    if prefill and prefill.get("tab") == "settings":
        show_page(2)

    root.mainloop()

    if reopen["prefill"]:
        open_setting_window(reopen["prefill"])

if __name__ == "__main__":
    enable_dpi_awareness()
    LANG = load_language()
    apply_theme(load_theme())
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