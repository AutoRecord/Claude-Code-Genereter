"""ポモドーロタイマー - 透過オーバーレイ版"""

import tkinter as tk
import json
import os
import time
from datetime import date

try:
    import winsound
except ImportError:
    winsound = None


class PomodoroApp:
    """透過オーバーレイ型のポモドーロタイマーアプリケーション"""

    # ---- デフォルト設定 ----
    DEFAULT_SETTINGS = {
        "work_minutes": 25,
        "short_break_minutes": 5,
        "long_break_minutes": 15,
        "cycles_before_long_break": 4,
        "dark_mode": True,
        "opacity": 0.75,
        "ui_scale": 100,
    }

    # ---- 設定値の上下限 ----
    MIN_MINUTES = 1
    MAX_MINUTES = 120
    MIN_CYCLES = 1
    MAX_CYCLES = 10
    UI_SCALE_MIN = 70
    UI_SCALE_MAX = 150

    # ---- モード → 設定キーのマッピング ----
    MODE_KEY_MAP = {
        "work": "work_minutes",
        "short_break": "short_break_minutes",
        "long_break": "long_break_minutes",
    }

    # ---- テーマカラー ----
    LIGHT_THEME = {
        "bg": "#ECECEC",
        "fg": "#222222",
        "accent": "#E74C3C",
        "accent_break": "#27AE60",
        "accent_long_break": "#2980B9",
        "button_bg": "#DCDCDC",
        "button_fg": "#222222",
        "button_hover": "#C8C8C8",
        "panel_bg": "#E0E0E0",
        "arc_bg": "#CCCCCC",
        "entry_bg": "#FFFFFF",
        "entry_fg": "#222222",
        "label_fg": "#555555",
        "close_fg": "#888888",
        "close_hover": "#E74C3C",
    }

    DARK_THEME = {
        "bg": "#1E1E2E",
        "fg": "#CDD6F4",
        "accent": "#F38BA8",
        "accent_break": "#A6E3A1",
        "accent_long_break": "#89B4FA",
        "button_bg": "#313244",
        "button_fg": "#CDD6F4",
        "button_hover": "#45475A",
        "panel_bg": "#181825",
        "arc_bg": "#313244",
        "entry_bg": "#313244",
        "entry_fg": "#CDD6F4",
        "label_fg": "#A6ADC8",
        "close_fg": "#7F849C",
        "close_hover": "#F38BA8",
    }

    MODES = {
        "work": "作業中",
        "short_break": "短い休憩",
        "long_break": "長い休憩",
    }

    # ---- フォント定数 ----
    FONT_JP = "Meiryo UI"
    FONT_MONO = "Cascadia Mono"

    # ---- サイズ定数 ----
    MINI_W, MINI_H = 300, 400
    EXPANDED_W, EXPANDED_H = 460, 700
    MINI_MIN_W, MINI_MIN_H = 250, 340
    EXPANDED_MIN_W, EXPANDED_MIN_H = 400, 550

    # ---- アニメーション定数 ----
    FADE_INTERVAL_MS = 30
    FADE_LERP_FACTOR = 0.25
    FADE_THRESHOLD = 0.02
    HOVER_OPACITY = 0.95
    END_NOTIFY_DURATION_MS = 2000
    END_NOTIFY_FLASH_COUNT = 4
    END_NOTIFY_FLASH_INTERVAL_MS = 300
    SCALE_DEBOUNCE_MS = 300
    TITLEBAR_HEIGHT = 32
    SCREEN_EDGE_MARGIN = 50
    INIT_MARGIN_X = 30
    INIT_MARGIN_Y = 60
    ARC_BASE_RADIUS = 80
    ARC_MIN_RADIUS = 60
    ARC_BASE_WIDTH = 12
    ARC_MIN_WIDTH = 7

    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

        # データディレクトリ
        self.data_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "data"
        )
        os.makedirs(self.data_dir, exist_ok=True)

        # 設定読み込み
        self.settings = self._load_settings()

        # タイマー状態
        self.mode = "work"
        self.running = False
        self.time_left = self.settings["work_minutes"] * 60
        self.total_time = self.time_left
        self.completed_cycles = 0
        self.timer_id = None
        self._timer_start_mono = 0.0
        self._timer_start_left = 0

        # UI状態
        self.expanded = False
        self.drag_x = 0
        self.drag_y = 0
        self.fade_id = None
        self.end_fade_id = None
        self._flash_count = 0
        self._scale_debounce_id = None
        self._save_debounce_id = None
        self.target_alpha = self.settings["opacity"]
        self.current_alpha = self.settings["opacity"]

        # 統計 / タスク
        self.stats = self._load_stats()
        self.tasks = self._load_tasks()

        # テーマ
        self.theme = (
            self.DARK_THEME if self.settings["dark_mode"] else self.LIGHT_THEME
        )

        # ウィンドウ属性
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", self.current_alpha)

        # 起動位置: デスクトップ右下
        init_w, init_h = self._calc_window_size(expanded=False)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = sw - init_w - self.INIT_MARGIN_X
        y = sh - init_h - self.INIT_MARGIN_Y
        self.root.geometry(f"{init_w}x{init_h}+{x}+{y}")

        # UI構築
        self._build_ui()
        self._apply_theme()
        self._update_display()
        self._update_stats_display()

        # ホバー透過
        self.root.bind("<Enter>", self._on_hover_enter)
        self.root.bind("<Leave>", self._on_hover_leave)

        # キーボードショートカット
        self.root.bind("<Escape>", lambda e: self._on_close())

        self.root.deiconify()

    # ============================
    #  ユーティリティ
    # ============================
    def _data_path(self, name):
        """データディレクトリ内のファイルパスを返す"""
        return os.path.join(self.data_dir, name)

    def _scaled_size(self, base_size):
        """UIスケールに応じたフォントサイズを返す"""
        return max(7, int(base_size * self.settings["ui_scale"] / 100))

    def _calc_window_size(self, expanded):
        """現在のUIスケールに応じたウィンドウサイズを計算する"""
        scale = self.settings["ui_scale"] / 100.0
        if expanded:
            w = max(self.EXPANDED_MIN_W, int(self.EXPANDED_W * scale))
            h = max(self.EXPANDED_MIN_H, int(self.EXPANDED_H * scale))
        else:
            w = max(self.MINI_MIN_W, int(self.MINI_W * scale))
            h = max(self.MINI_MIN_H, int(self.MINI_H * scale))
        return w, h

    def _clamp(self, value, min_val, max_val):
        """値を範囲内にクランプする"""
        return max(min_val, min(value, max_val))

    # ============================
    #  永続化
    # ============================
    def _load_settings(self):
        """設定ファイルを読み込み、型バリデーション付きで返す"""
        try:
            with open(self._data_path("settings.json"), "r", encoding="utf-8") as f:
                saved = json.load(f)
                if not isinstance(saved, dict):
                    return dict(self.DEFAULT_SETTINGS)
                s = dict(self.DEFAULT_SETTINGS)
                s.update(saved)
                # 型バリデーション
                return self._validate_settings(s)
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return dict(self.DEFAULT_SETTINGS)

    def _validate_settings(self, s):
        """設定値の型と範囲を検証・修正する"""
        defaults = self.DEFAULT_SETTINGS
        for key in ("work_minutes", "short_break_minutes", "long_break_minutes"):
            try:
                s[key] = self._clamp(int(s[key]), self.MIN_MINUTES, self.MAX_MINUTES)
            except (TypeError, ValueError):
                s[key] = defaults[key]
        try:
            s["cycles_before_long_break"] = self._clamp(
                int(s["cycles_before_long_break"]), self.MIN_CYCLES, self.MAX_CYCLES
            )
        except (TypeError, ValueError):
            s["cycles_before_long_break"] = defaults["cycles_before_long_break"]
        try:
            s["opacity"] = self._clamp(float(s["opacity"]), 0.3, 1.0)
        except (TypeError, ValueError):
            s["opacity"] = defaults["opacity"]
        try:
            s["ui_scale"] = self._clamp(
                int(s["ui_scale"]), self.UI_SCALE_MIN, self.UI_SCALE_MAX
            )
        except (TypeError, ValueError):
            s["ui_scale"] = defaults["ui_scale"]
        s["dark_mode"] = bool(s.get("dark_mode", defaults["dark_mode"]))
        return s

    def _save_settings_debounced(self):
        """デバウンス付きで設定を保存する（スライダー操作向け）"""
        if self._save_debounce_id:
            self.root.after_cancel(self._save_debounce_id)
        self._save_debounce_id = self.root.after(200, self._save_settings)

    def _save_settings(self):
        """設定をファイルに保存する"""
        self._save_debounce_id = None
        self._safe_write("settings.json", self.settings)

    def _load_stats(self):
        """統計ファイルを読み込む"""
        try:
            with open(self._data_path("stats.json"), "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return {}

    def _save_stats(self):
        self._safe_write("stats.json", self.stats)

    def _load_tasks(self):
        """タスクファイルを読み込む"""
        try:
            with open(self._data_path("tasks.json"), "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return []

    def _save_tasks(self):
        self._safe_write("tasks.json", self.tasks)

    def _safe_write(self, filename, data):
        """一時ファイル経由でアトミックに書き込む"""
        path = self._data_path(filename)
        tmp_path = path + ".tmp"
        try:
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, path)
        except OSError:
            # 一時ファイルが残っていれば掃除
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    # ============================
    #  UI 構築
    # ============================
    def _build_ui(self):
        """メインUIを構築する"""
        self.container = tk.Frame(self.root, bd=0)
        self.container.pack(fill=tk.BOTH, expand=True)

        font_jp = self.FONT_JP

        # ---- タイトルバー ----
        self.title_bar = tk.Frame(
            self.container, height=self.TITLEBAR_HEIGHT
        )
        self.title_bar.pack(fill=tk.X)
        self.title_bar.pack_propagate(False)

        self.toggle_btn = tk.Label(
            self.title_bar, text="≡", font=(font_jp, self._scaled_size(15)),
            cursor="hand2", padx=8,
        )
        self.toggle_btn.pack(side=tk.LEFT)
        self.toggle_btn.bind("<Button-1>", lambda e: self._toggle_expand())

        self.title_label = tk.Label(
            self.title_bar, text="🍅 ポモドーロ",
            font=(font_jp, self._scaled_size(11)), anchor="w",
        )
        self.title_label.pack(side=tk.LEFT, padx=4)

        self.close_btn = tk.Label(
            self.title_bar, text="✕", font=(font_jp, self._scaled_size(13)),
            cursor="hand2", padx=8,
        )
        self.close_btn.pack(side=tk.RIGHT)
        self.close_btn.bind("<Button-1>", lambda e: self._on_close())

        for widget in [self.title_bar, self.title_label]:
            widget.bind("<Button-1>", self._drag_start)
            widget.bind("<B1-Motion>", self._drag_motion)

        # ---- ミニモード領域 ----
        self.mini_frame = tk.Frame(self.container)
        self.mini_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

        self.mode_label = tk.Label(
            self.mini_frame, text=self.MODES[self.mode],
            font=(font_jp, self._scaled_size(15), "bold"),
        )
        self.mode_label.pack(pady=(6, 0))

        self.cycle_label = tk.Label(
            self.mini_frame, text=self._cycle_text(),
            font=(font_jp, self._scaled_size(11)),
        )
        self.cycle_label.pack()

        # 円形プログレス
        self.canvas_size = max(160, int(200 * self.settings["ui_scale"] / 100))
        self.canvas = tk.Canvas(
            self.mini_frame,
            width=self.canvas_size, height=self.canvas_size,
            highlightthickness=0, bd=0,
        )
        self.canvas.pack(pady=6)
        self.canvas.bind("<Button-1>", self._drag_start)
        self.canvas.bind("<B1-Motion>", self._drag_motion)

        # ---- 時間調整ボタン ----
        time_adj_frame = tk.Frame(self.mini_frame)
        time_adj_frame.pack(pady=(0, 4))

        self.time_down_btn = tk.Button(
            time_adj_frame, text="−1分", width=5,
            font=(font_jp, self._scaled_size(10)), relief=tk.FLAT, bd=0,
            cursor="hand2", command=lambda: self._adjust_time(-1),
        )
        self.time_down_btn.pack(side=tk.LEFT, padx=4)

        self.time_display_label = tk.Label(
            time_adj_frame, text=self._time_setting_text(),
            font=(font_jp, self._scaled_size(10)),
        )
        self.time_display_label.pack(side=tk.LEFT, padx=8)

        self.time_up_btn = tk.Button(
            time_adj_frame, text="+1分", width=5,
            font=(font_jp, self._scaled_size(10)), relief=tk.FLAT, bd=0,
            cursor="hand2", command=lambda: self._adjust_time(1),
        )
        self.time_up_btn.pack(side=tk.LEFT, padx=4)

        # ---- 操作ボタン ----
        btn_frame = tk.Frame(self.mini_frame)
        btn_frame.pack(pady=(4, 0))

        btn_defs = [
            ("start_btn", "▶", self._start_timer, "開始"),
            ("pause_btn", "⏸", self._pause_timer, "一時停止"),
            ("reset_btn", "⏹", self._reset_timer, "リセット"),
            ("skip_btn", "⏭", self._skip_to_next, "スキップ"),
        ]
        for attr, text, cmd, tooltip_text in btn_defs:
            btn = tk.Button(
                btn_frame, text=text, width=4,
                font=(font_jp, self._scaled_size(13)), relief=tk.FLAT, bd=0,
                cursor="hand2", command=cmd,
            )
            btn.pack(side=tk.LEFT, padx=4)
            self._bind_hover(btn)
            self._bind_tooltip(btn, tooltip_text)
            setattr(self, attr, btn)

        self.pause_btn.config(state=tk.DISABLED)

        # ホバーエフェクトを時間調整ボタンにも適用
        self._bind_hover(self.time_down_btn)
        self._bind_hover(self.time_up_btn)

        # ---- 展開パネル（初期非表示） ----
        self.expanded_frame = tk.Frame(self.container)
        self._build_expanded_panel()

    def _bind_hover(self, btn):
        """ボタンにホバーエフェクトをバインドする"""
        def on_enter(e):
            if str(btn["state"]) != "disabled":
                btn.config(bg=self.theme["button_hover"])
        def on_leave(e):
            if str(btn["state"]) != "disabled":
                btn.config(bg=self.theme["button_bg"])
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

    def _bind_tooltip(self, widget, text):
        """ウィジェットにツールチップをバインドする"""
        tip_window = None

        def show_tip(e):
            nonlocal tip_window
            if tip_window:
                return
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 4
            tip_window = tw = tk.Toplevel(widget)
            tw.wm_overrideredirect(True)
            tw.wm_attributes("-topmost", True)
            tw.wm_geometry(f"+{x}+{y}")
            label = tk.Label(
                tw, text=text, font=(self.FONT_JP, 9),
                bg="#333333", fg="#FFFFFF", padx=6, pady=2,
                relief=tk.SOLID, bd=1,
            )
            label.pack()

        def hide_tip(e):
            nonlocal tip_window
            if tip_window:
                tip_window.destroy()
                tip_window = None

        widget.bind("<Enter>", show_tip, add="+")
        widget.bind("<Leave>", hide_tip, add="+")

    def _build_expanded_panel(self):
        """展開パネル（タスク・統計・設定）を構築する"""
        panel = self.expanded_frame
        font_jp = self.FONT_JP

        # --- タスクパネル ---
        task_panel = tk.LabelFrame(
            panel, text=" 📋 タスク ",
            font=(font_jp, self._scaled_size(11), "bold"),
            bd=1, relief=tk.GROOVE,
        )
        task_panel.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))

        inp_frame = tk.Frame(task_panel)
        inp_frame.pack(fill=tk.X, padx=8, pady=(8, 4))
        self.task_entry = tk.Entry(
            inp_frame, font=(font_jp, self._scaled_size(11)),
            relief=tk.FLAT, bd=1,
        )
        self.task_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        self.task_entry.bind("<Return>", lambda e: self._add_task())

        self.add_task_btn = tk.Button(
            inp_frame, text="+", width=3,
            font=(font_jp, self._scaled_size(11), "bold"),
            relief=tk.FLAT, bd=0, cursor="hand2", command=self._add_task,
        )
        self.add_task_btn.pack(side=tk.LEFT, padx=(6, 0))
        self._bind_hover(self.add_task_btn)

        self.task_listbox = tk.Listbox(
            task_panel, font=(font_jp, self._scaled_size(11)),
            relief=tk.FLAT, bd=0, activestyle="none",
            height=5, selectmode=tk.SINGLE,
        )
        self.task_listbox.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)

        task_btn_frame = tk.Frame(task_panel)
        task_btn_frame.pack(fill=tk.X, padx=8, pady=(0, 6))
        self.complete_task_btn = tk.Button(
            task_btn_frame, text="✓ 完了",
            font=(font_jp, self._scaled_size(10)),
            relief=tk.FLAT, bd=0, cursor="hand2",
            command=self._toggle_task_complete,
        )
        self.complete_task_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._bind_hover(self.complete_task_btn)

        self.delete_task_btn = tk.Button(
            task_btn_frame, text="✕ 削除",
            font=(font_jp, self._scaled_size(10)),
            relief=tk.FLAT, bd=0, cursor="hand2",
            command=self._delete_task,
        )
        self.delete_task_btn.pack(side=tk.LEFT)
        self._bind_hover(self.delete_task_btn)

        # --- 統計パネル ---
        stats_panel = tk.LabelFrame(
            panel, text=" 📊 統計 ",
            font=(font_jp, self._scaled_size(11), "bold"),
            bd=1, relief=tk.GROOVE,
        )
        stats_panel.pack(fill=tk.X, padx=12, pady=4)

        self.today_pomodoros_label = tk.Label(
            stats_panel, text="今日: 0 ポモドーロ",
            font=(font_jp, self._scaled_size(11)), anchor="w",
        )
        self.today_pomodoros_label.pack(fill=tk.X, padx=10, pady=(6, 2))
        self.today_time_label = tk.Label(
            stats_panel, text="作業時間: 0分",
            font=(font_jp, self._scaled_size(11)), anchor="w",
        )
        self.today_time_label.pack(fill=tk.X, padx=10, pady=2)
        self.total_pomodoros_label = tk.Label(
            stats_panel, text="累計: 0 ポモドーロ",
            font=(font_jp, self._scaled_size(11)), anchor="w",
        )
        self.total_pomodoros_label.pack(fill=tk.X, padx=10, pady=(2, 6))

        # --- 設定パネル ---
        settings_panel = tk.LabelFrame(
            panel, text=" ⚙️ 設定 ",
            font=(font_jp, self._scaled_size(11), "bold"),
            bd=1, relief=tk.GROOVE,
        )
        settings_panel.pack(fill=tk.X, padx=12, pady=(4, 8))

        self._setting_vars = {}
        settings_grid = tk.Frame(settings_panel)
        settings_grid.pack(fill=tk.X, padx=10, pady=6)
        for i, (label, key) in enumerate([
            ("作業(分)", "work_minutes"),
            ("短休(分)", "short_break_minutes"),
            ("長休(分)", "long_break_minutes"),
            ("サイクル", "cycles_before_long_break"),
        ]):
            row, col = divmod(i, 2)
            frame = tk.Frame(settings_grid)
            frame.grid(row=row, column=col, padx=6, pady=3, sticky="ew")
            tk.Label(
                frame, text=label,
                font=(font_jp, self._scaled_size(10)), width=7, anchor="w",
            ).pack(side=tk.LEFT)
            var = tk.StringVar(value=str(self.settings[key]))
            tk.Entry(
                frame, textvariable=var,
                font=(font_jp, self._scaled_size(10)),
                width=4, relief=tk.FLAT, bd=1, justify=tk.CENTER,
            ).pack(side=tk.LEFT, ipady=2)
            self._setting_vars[key] = var
        settings_grid.columnconfigure(0, weight=1)
        settings_grid.columnconfigure(1, weight=1)

        # 透過度スライダー
        opacity_frame = tk.Frame(settings_panel)
        opacity_frame.pack(fill=tk.X, padx=10, pady=(0, 4))
        tk.Label(
            opacity_frame, text="透過度",
            font=(font_jp, self._scaled_size(10)),
        ).pack(side=tk.LEFT)
        self.opacity_scale = tk.Scale(
            opacity_frame, from_=30, to=100, orient=tk.HORIZONTAL,
            font=(font_jp, self._scaled_size(9)), length=140, sliderlength=16,
            command=self._on_opacity_change, showvalue=True,
        )
        self.opacity_scale.set(int(self.settings["opacity"] * 100))
        self.opacity_scale.pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0)
        )

        # UIスケールスライダー
        scale_frame = tk.Frame(settings_panel)
        scale_frame.pack(fill=tk.X, padx=10, pady=(0, 4))
        tk.Label(
            scale_frame, text="UIサイズ",
            font=(font_jp, self._scaled_size(10)),
        ).pack(side=tk.LEFT)
        self.ui_scale_slider = tk.Scale(
            scale_frame, from_=self.UI_SCALE_MIN, to=self.UI_SCALE_MAX,
            orient=tk.HORIZONTAL,
            font=(font_jp, self._scaled_size(9)), length=140, sliderlength=16,
            command=self._on_ui_scale_change, showvalue=True,
        )
        self.ui_scale_slider.set(self.settings["ui_scale"])
        self.ui_scale_slider.pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0)
        )

        # ボタン行
        btn_row = tk.Frame(settings_panel)
        btn_row.pack(fill=tk.X, padx=10, pady=(2, 8))

        self.apply_btn = tk.Button(
            btn_row, text="適用",
            font=(font_jp, self._scaled_size(10)),
            relief=tk.FLAT, bd=0, cursor="hand2",
            command=self._apply_settings,
        )
        self.apply_btn.pack(side=tk.LEFT, padx=(0, 10))
        self._bind_hover(self.apply_btn)

        self.dark_mode_var = tk.BooleanVar(value=self.settings["dark_mode"])
        self.dark_mode_check = tk.Checkbutton(
            btn_row, text="🌙 ダークモード",
            font=(font_jp, self._scaled_size(10)),
            variable=self.dark_mode_var, command=self._toggle_dark_mode,
        )
        self.dark_mode_check.pack(side=tk.LEFT)

        self._refresh_task_list()

    # ============================
    #  展開 / 折りたたみ
    # ============================
    def _toggle_expand(self):
        """ミニ/展開モードを切り替える"""
        if self.expanded:
            self.expanded_frame.pack_forget()
            self.expanded = False
            self.toggle_btn.config(text="≡")
        else:
            self.expanded_frame.pack(
                fill=tk.BOTH, expand=True, after=self.mini_frame
            )
            self.expanded = True
            self.toggle_btn.config(text="△")
            self._update_stats_display()
            self._refresh_task_list()

        w, h = self._calc_window_size(self.expanded)
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        if x + w > sw:
            x = sw - w - 10
        if y + h > sh:
            y = sh - h - 40
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ============================
    #  ドラッグ移動
    # ============================
    def _drag_start(self, event):
        self.drag_x = event.x_root - self.root.winfo_x()
        self.drag_y = event.y_root - self.root.winfo_y()

    def _drag_motion(self, event):
        x = event.x_root - self.drag_x
        y = event.y_root - self.drag_y
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        margin = self.SCREEN_EDGE_MARGIN
        x = max(-w + margin, min(x, sw - margin))
        y = max(0, min(y, sh - margin))
        self.root.geometry(f"+{x}+{y}")

    # ============================
    #  ホバー透過アニメーション
    # ============================
    def _on_hover_enter(self, event=None):
        self._cancel_end_fade()
        self._fade_to(self.HOVER_OPACITY)

    def _on_hover_leave(self, event=None):
        self._cancel_end_fade()
        self._fade_to(self.settings["opacity"])

    def _cancel_end_fade(self):
        """タイマー終了時のフェード予約をキャンセルする"""
        if self.end_fade_id:
            self.root.after_cancel(self.end_fade_id)
            self.end_fade_id = None

    def _fade_to(self, target):
        """指定した透過度に滑らかにフェードする"""
        self.target_alpha = target
        if self.fade_id:
            self.root.after_cancel(self.fade_id)
        self._fade_step()

    def _fade_step(self):
        diff = self.target_alpha - self.current_alpha
        if abs(diff) < self.FADE_THRESHOLD:
            self.current_alpha = self.target_alpha
            self.root.attributes("-alpha", self.current_alpha)
            self.fade_id = None
            return
        self.current_alpha += diff * self.FADE_LERP_FACTOR
        self.root.attributes("-alpha", self.current_alpha)
        self.fade_id = self.root.after(self.FADE_INTERVAL_MS, self._fade_step)

    def _on_opacity_change(self, val):
        self.settings["opacity"] = int(val) / 100.0
        self.target_alpha = self.settings["opacity"]
        self.current_alpha = self.settings["opacity"]
        self.root.attributes("-alpha", self.current_alpha)
        self._save_settings_debounced()

    def _on_ui_scale_change(self, val):
        self.settings["ui_scale"] = int(val)
        self._save_settings_debounced()
        if self._scale_debounce_id:
            self.root.after_cancel(self._scale_debounce_id)
        self._scale_debounce_id = self.root.after(
            self.SCALE_DEBOUNCE_MS, self._rebuild_ui
        )

    def _rebuild_ui(self):
        """UIスケール変更時にウィジェットを再構築する"""
        self._scale_debounce_id = None
        was_expanded = self.expanded

        # 実行中の after コールバックをキャンセル
        for attr in ("timer_id", "fade_id", "end_fade_id"):
            after_id = getattr(self, attr, None)
            if after_id:
                self.root.after_cancel(after_id)
                setattr(self, attr, None)

        # 全ウィジェット削除
        self.container.destroy()
        self.expanded = False

        # 再構築
        self._build_ui()
        self._apply_theme()
        self._update_display()
        self._update_stats_display()

        # 展開状態を復元
        if was_expanded:
            self._toggle_expand()

        # ウィンドウサイズ
        w, h = self._calc_window_size(self.expanded)
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        # タイマーが動いていたなら再開
        if self.running:
            self._tick()

    # ============================
    #  タイマーロジック
    # ============================
    def _cycle_text(self):
        """サイクル表示テキストを返す"""
        n = self.settings["cycles_before_long_break"]
        if self.mode == "long_break":
            return f"サイクル {n}/{n} ✓"
        return f"サイクル {self.completed_cycles % n + 1}/{n}"

    def _time_setting_text(self):
        """現在のモードに対応する設定時間のラベルテキストを返す"""
        key = self.MODE_KEY_MAP[self.mode]
        label_map = {
            "work_minutes": "作業",
            "short_break_minutes": "短休",
            "long_break_minutes": "長休",
        }
        return f"{label_map[key]} {self.settings[key]}分"

    def _adjust_time(self, delta):
        """現在のモードの設定時間を±delta分で調整する"""
        key = self.MODE_KEY_MAP[self.mode]
        new_val = self._clamp(
            self.settings[key] + delta, self.MIN_MINUTES, self.MAX_MINUTES
        )
        self.settings[key] = new_val
        self._save_settings()

        # 展開パネルの入力欄も同期
        if key in self._setting_vars:
            self._setting_vars[key].set(str(new_val))

        # total_time と time_left を更新（実行中も対応）
        new_total = new_val * 60
        if self.running:
            elapsed = self.total_time - self.time_left
            self.total_time = new_total
            self.time_left = max(0, new_total - elapsed)
            self._timer_start_mono = time.monotonic()
            self._timer_start_left = self.time_left
            # 調整結果で0以下になった場合は即座に終了処理
            if self.time_left <= 0:
                self._update_display()
                self._on_timer_end()
                return
        else:
            self.time_left = new_total
            self.total_time = new_total
        self._update_display()
        self.time_display_label.config(text=self._time_setting_text())

    def _start_timer(self):
        """タイマーを開始する"""
        if not self.running:
            self.running = True
            self._timer_start_mono = time.monotonic()
            self._timer_start_left = self.time_left
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.NORMAL)
            self._tick()

    def _pause_timer(self):
        """タイマーを一時停止する"""
        self.running = False
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)

    def _reset_timer(self):
        """タイマーをリセットする（作業モードに戻る）"""
        self.running = False
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        self.mode = "work"
        self.time_left = self.settings["work_minutes"] * 60
        self.total_time = self.time_left
        self.completed_cycles = 0
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self._update_display()

    def _skip_to_next(self):
        """現在のモードをスキップして次に進む"""
        was_running = self.running
        self.running = False
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        self._switch_mode()
        if was_running:
            self._start_timer()
        else:
            self.start_btn.config(state=tk.NORMAL)
            self.pause_btn.config(state=tk.DISABLED)

    def _tick(self):
        """1秒ごとのタイマー更新処理"""
        if self.running:
            elapsed = time.monotonic() - self._timer_start_mono
            self.time_left = max(0, self._timer_start_left - int(elapsed))
            self._update_display()
            if self.time_left <= 0:
                self._on_timer_end()
            else:
                self.timer_id = self.root.after(1000, self._tick)

    def _on_timer_end(self):
        """タイマー終了時の処理（通知・統計・自動モード遷移）"""
        # ★ 重要: running を False にしてから _start_timer を呼ぶ
        self.running = False
        self.timer_id = None

        # 通知音
        try:
            if winsound:
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            else:
                self.root.bell()
        except Exception:
            self.root.bell()

        # 統計記録（作業モード完了時のみ）
        if self.mode == "work":
            today = date.today().isoformat()
            if today not in self.stats:
                self.stats[today] = {"pomodoros": 0, "minutes": 0}
            self.stats[today]["pomodoros"] += 1
            self.stats[today]["minutes"] += self.settings["work_minutes"]
            self._save_stats()
            if self.expanded:
                self._update_stats_display()
            self.completed_cycles += 1

        # 終了通知: 不透明化 + 点滅パルス
        self._fade_to(1.0)
        self._cancel_end_fade()
        self._flash_count = 0
        self._flash_notify()

        # 次のモードへ自動遷移
        self._switch_mode()
        self._start_timer()

    def _flash_notify(self):
        """タイマー終了時にプログレスリングを点滅させて注意を引く"""
        if self._flash_count >= self.END_NOTIFY_FLASH_COUNT * 2:
            # 点滅完了 → 通常の透過度に戻す
            self.end_fade_id = self.root.after(
                self.END_NOTIFY_DURATION_MS, self._restore_opacity_after_end
            )
            return
        self._flash_count += 1
        # 不透明度をパルスさせる
        if self._flash_count % 2 == 0:
            self.root.attributes("-alpha", 1.0)
            self.current_alpha = 1.0
        else:
            self.root.attributes("-alpha", 0.6)
            self.current_alpha = 0.6
        self.end_fade_id = self.root.after(
            self.END_NOTIFY_FLASH_INTERVAL_MS, self._flash_notify
        )

    def _restore_opacity_after_end(self):
        """タイマー終了後に元の透過度に戻す"""
        self.end_fade_id = None
        self._fade_to(self.settings["opacity"])

    def _switch_mode(self):
        """次のモード(作業/短休/長休)に遷移する"""
        n = self.settings["cycles_before_long_break"]
        if self.mode == "work":
            if self.completed_cycles > 0 and self.completed_cycles % n == 0:
                self.mode = "long_break"
            else:
                self.mode = "short_break"
        else:
            self.mode = "work"

        key = self.MODE_KEY_MAP[self.mode]
        self.time_left = self.settings[key] * 60
        self.total_time = self.time_left
        self._update_display()

    # ============================
    #  表示更新
    # ============================
    def _update_display(self):
        """タイマーの表示を更新する"""
        mins, secs = divmod(self.time_left, 60)
        time_str = f"{mins:02d}:{secs:02d}"
        self.mode_label.config(text=self.MODES[self.mode])
        self.cycle_label.config(text=self._cycle_text())
        self.time_display_label.config(text=self._time_setting_text())
        self._draw_progress(time_str)

    def _get_accent(self):
        """現在のモードに対応するアクセントカラーを返す"""
        if self.mode == "work":
            return self.theme["accent"]
        elif self.mode == "short_break":
            return self.theme["accent_break"]
        return self.theme["accent_long_break"]

    def _draw_progress(self, time_str):
        """円形プログレスバーとタイマー数字を描画する"""
        canvas = self.canvas
        canvas.delete("all")
        size = self.canvas_size
        cx, cy = size / 2, size / 2
        scale = self.settings["ui_scale"] / 100.0
        r = max(self.ARC_MIN_RADIUS, int(self.ARC_BASE_RADIUS * scale))
        lw = max(self.ARC_MIN_WIDTH, int(self.ARC_BASE_WIDTH * scale))
        color = self._get_accent()

        # 背景円弧
        canvas.create_arc(
            cx - r, cy - r, cx + r, cy + r,
            start=90, extent=-360,
            outline=self.theme["arc_bg"], width=lw, style=tk.ARC,
        )

        # プログレス
        progress = self.time_left / self.total_time if self.total_time > 0 else 0
        extent = -360 * progress
        if abs(extent) > 0.5:
            canvas.create_arc(
                cx - r, cy - r, cx + r, cy + r,
                start=90, extent=extent,
                outline=color, width=lw, style=tk.ARC,
            )

        # タイマー数字
        canvas.create_text(
            cx, cy,
            text=time_str,
            font=(self.FONT_MONO, self._scaled_size(32), "bold"),
            fill=self.theme["fg"],
        )

    def _update_stats_display(self):
        """統計パネルの表示を更新する"""
        today = date.today().isoformat()
        today_data = self.stats.get(today, {"pomodoros": 0, "minutes": 0})
        total_p = sum(d.get("pomodoros", 0) for d in self.stats.values())
        total_m = sum(d.get("minutes", 0) for d in self.stats.values())
        self.today_pomodoros_label.config(
            text=f"今日: {today_data['pomodoros']} ポモドーロ"
        )
        self.today_time_label.config(
            text=f"作業時間: {today_data['minutes']}分"
        )
        self.total_pomodoros_label.config(
            text=f"累計: {total_p} ポモドーロ ({total_m}分)"
        )

    # ============================
    #  タスク管理
    # ============================
    def _add_task(self):
        text = self.task_entry.get().strip()
        if not text:
            return
        # 長さ制限
        if len(text) > 200:
            text = text[:200]
        self.tasks.append({"text": text, "done": False})
        self.task_entry.delete(0, tk.END)
        self._save_tasks()
        self._refresh_task_list()

    def _delete_task(self):
        sel = self.task_listbox.curselection()
        if not sel:
            return
        self.tasks.pop(sel[0])
        self._save_tasks()
        self._refresh_task_list()

    def _toggle_task_complete(self):
        sel = self.task_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.tasks[idx]["done"] = not self.tasks[idx]["done"]
        self._save_tasks()
        self._refresh_task_list()

    def _refresh_task_list(self):
        self.task_listbox.delete(0, tk.END)
        for task in self.tasks:
            prefix = "✓ " if task.get("done") else "○ "
            self.task_listbox.insert(tk.END, prefix + task.get("text", ""))
        for i, task in enumerate(self.tasks):
            if task.get("done"):
                self.task_listbox.itemconfig(i, fg=self.theme["label_fg"])

    # ============================
    #  設定
    # ============================
    def _apply_settings(self):
        """展開パネルの設定を適用する"""
        try:
            new = {}
            for key, var in self._setting_vars.items():
                val = int(var.get())
                if val <= 0:
                    raise ValueError
                new[key] = val
        except ValueError:
            return

        # 範囲制限
        for key in ("work_minutes", "short_break_minutes", "long_break_minutes"):
            if key in new:
                new[key] = self._clamp(new[key], self.MIN_MINUTES, self.MAX_MINUTES)
        if "cycles_before_long_break" in new:
            new["cycles_before_long_break"] = self._clamp(
                new["cycles_before_long_break"], self.MIN_CYCLES, self.MAX_CYCLES
            )

        self.settings.update(new)
        self._save_settings()

        if self.running:
            # 実行中: 現在のモードの設定時間が変わった場合のみ更新
            key = self.MODE_KEY_MAP[self.mode]
            new_total = self.settings[key] * 60
            if new_total != self.total_time:
                elapsed = self.total_time - self.time_left
                self.total_time = new_total
                self.time_left = max(0, new_total - elapsed)
                self._timer_start_mono = time.monotonic()
                self._timer_start_left = self.time_left
                if self.time_left <= 0:
                    self._update_display()
                    self._on_timer_end()
                    return
                self._update_display()
        else:
            # 停止中: タイマーをリセットするが、サイクルカウントは維持
            mode_key = self.MODE_KEY_MAP[self.mode]
            self.time_left = self.settings[mode_key] * 60
            self.total_time = self.time_left
            self.start_btn.config(state=tk.NORMAL)
            self.pause_btn.config(state=tk.DISABLED)
            self._update_display()

        # 入力欄の値を正規化後の値で更新
        for key, var in self._setting_vars.items():
            var.set(str(self.settings[key]))
        self.time_display_label.config(text=self._time_setting_text())

    def _toggle_dark_mode(self):
        self.settings["dark_mode"] = self.dark_mode_var.get()
        self.theme = (
            self.DARK_THEME if self.settings["dark_mode"] else self.LIGHT_THEME
        )
        self._save_settings()
        self._apply_theme()
        self._update_display()
        self._refresh_task_list()

    # ============================
    #  テーマ適用
    # ============================
    def _apply_theme(self):
        """全ウィジェットにテーマカラーを適用する"""
        theme = self.theme
        self.root.config(bg=theme["bg"])
        self.container.config(bg=theme["bg"])

        # タイトルバー
        self.title_bar.config(bg=theme["bg"])
        self.title_label.config(bg=theme["bg"], fg=theme["label_fg"])
        self.toggle_btn.config(bg=theme["bg"], fg=theme["fg"])
        self.close_btn.config(bg=theme["bg"], fg=theme["close_fg"])

        # 閉じるボタンのホバーは self.theme を参照するラムダにする
        self.close_btn.bind(
            "<Enter>",
            lambda e: self.close_btn.config(fg=self.theme["close_hover"]),
        )
        self.close_btn.bind(
            "<Leave>",
            lambda e: self.close_btn.config(fg=self.theme["close_fg"]),
        )

        # ミニフレーム
        self.mini_frame.config(bg=theme["bg"])
        self.mode_label.config(bg=theme["bg"], fg=theme["fg"])
        self.cycle_label.config(bg=theme["bg"], fg=theme["label_fg"])
        self.canvas.config(bg=theme["bg"])

        # ボタン
        for btn in [
            self.start_btn, self.pause_btn, self.reset_btn, self.skip_btn,
            self.time_down_btn, self.time_up_btn,
        ]:
            btn.config(
                bg=theme["button_bg"], fg=theme["button_fg"],
                activebackground=theme["button_hover"],
                activeforeground=theme["button_fg"],
            )

        self.time_display_label.config(bg=theme["bg"], fg=theme["label_fg"])

        # ミニフレーム内の子Frame
        for child in self.mini_frame.winfo_children():
            if isinstance(child, tk.Frame):
                child.config(bg=theme["bg"])

        # 展開パネル
        self.expanded_frame.config(bg=theme["bg"])
        self._theme_children(self.expanded_frame, theme)

    def _theme_children(self, widget, theme):
        """ウィジェットツリーを再帰的に走査してテーマを適用する"""
        cls = widget.winfo_class()
        # 親がLabelFrameかどうかでパネル内かを判定
        in_panel = self._is_inside_labelframe(widget)
        panel_bg = theme["panel_bg"] if in_panel else theme["bg"]

        try:
            if cls == "Frame":
                widget.config(bg=panel_bg)
            elif cls == "Labelframe":
                widget.config(bg=theme["panel_bg"], fg=theme["fg"])
            elif cls == "Label":
                widget.config(bg=panel_bg, fg=theme["fg"])
            elif cls == "Button":
                widget.config(
                    bg=theme["button_bg"], fg=theme["button_fg"],
                    activebackground=theme["button_hover"],
                    activeforeground=theme["button_fg"],
                )
            elif cls == "Entry":
                widget.config(
                    bg=theme["entry_bg"], fg=theme["entry_fg"],
                    insertbackground=theme["entry_fg"],
                )
            elif cls == "Listbox":
                widget.config(
                    bg=theme["entry_bg"], fg=theme["entry_fg"],
                    selectbackground=theme["accent"],
                    selectforeground="#FFFFFF",
                )
            elif cls == "Checkbutton":
                widget.config(
                    bg=panel_bg, fg=theme["fg"],
                    activebackground=panel_bg, activeforeground=theme["fg"],
                    selectcolor=theme["entry_bg"],
                )
            elif cls == "Scale":
                widget.config(
                    bg=panel_bg, fg=theme["fg"],
                    troughcolor=theme["arc_bg"],
                    activebackground=theme["accent"],
                    highlightthickness=0,
                )
        except tk.TclError:
            pass

        for child in widget.winfo_children():
            self._theme_children(child, theme)

    def _is_inside_labelframe(self, widget):
        """ウィジェットがLabelFrameの子孫かどうかを判定する"""
        parent = widget.master
        while parent:
            if parent.winfo_class() == "Labelframe":
                return True
            if parent == self.expanded_frame:
                return False
            parent = parent.master
        return False

    # ============================
    #  終了
    # ============================
    def _on_close(self):
        """アプリケーションを終了する"""
        # 全ての after コールバックをキャンセル
        for after_id in [
            self.timer_id, self.fade_id, self.end_fade_id,
            self._scale_debounce_id, self._save_debounce_id,
        ]:
            if after_id:
                try:
                    self.root.after_cancel(after_id)
                except Exception:
                    pass
        self._save_settings()
        self._save_stats()
        self._save_tasks()
        self.root.destroy()

    def run(self):
        """アプリケーションのメインループを開始する"""
        self.root.mainloop()


if __name__ == "__main__":
    app = PomodoroApp()
    app.run()
