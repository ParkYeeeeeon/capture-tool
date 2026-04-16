"""
캡처 툴  v6  —  알캡처 스타일
캡처 : 사각형 · 단위영역 · 전체화면 · 스크롤 · 지정사이즈
편집 : 펜 · 형광펜 · 텍스트 · 사각형 · 타원 · 화살표 · 선 · 모자이크 · 크롭 · 지우개
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, simpledialog
import threading, time, io, math, os, sys, subprocess, json
from datetime import datetime

import ctypes

import mss
from PIL import Image, ImageTk, ImageDraw, ImageFont
import numpy as np
import win32clipboard, win32con, win32gui, win32api

# 작업표시줄 우클릭 → 종료 메뉴 커스텀 ID
_ID_TRAY_QUIT = 0x0601


# ─── 유틸리티 ────────────────────────────────────────────────────────────────

def image_to_clipboard(img: Image.Image):
    output = io.BytesIO()
    img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_DIB, data)
    win32clipboard.CloseClipboard()


def frame_diff(a: Image.Image, b: Image.Image) -> float:
    aa = np.array(a.convert("RGB"), dtype=np.float32)
    bb = np.array(b.convert("RGB"), dtype=np.float32)
    return float(np.mean(np.abs(aa - bb)))


def _find_new_content_start(accumulated: Image.Image, frame: Image.Image):
    """누적 이미지 하단 영역을 다음 프레임에서 직접 찾아 새 콘텐츠 시작 행을 반환.

    원리: accumulated의 마지막 PROBE_H 행(프로브)이 frame의 어느 위치에 있는지 탐색.
    프로브 위치 + PROBE_H = 새 콘텐츠 시작 행.
    매칭 실패(동일 프레임 / 품질 불량) 시 None 반환 → 해당 프레임 건너뜀.
    """
    PROBE_H = 80          # 프로브 높이 (픽셀)
    MAX_SCORE = 12.0      # 행 평균 SAD 허용 최대값 (0~255 스케일)

    w    = min(accumulated.width, frame.width)
    H    = frame.height
    A    = accumulated.height

    if A < PROBE_H or H < PROBE_H + 10:
        return None

    # 프로브: 누적 이미지 하단 PROBE_H 행의 열 평균 (1D 신호)
    probe_mean = (
        np.array(accumulated.crop((0, A - PROBE_H, w, A)).convert("L"), dtype=np.float32)
        .mean(axis=1)
    )

    # 다음 프레임 전체의 열 평균
    frame_mean = (
        np.array(frame.crop((0, 0, w, H)).convert("L"), dtype=np.float32)
        .mean(axis=1)
    )

    search_end = H - PROBE_H - 5
    if search_end <= 0:
        return None

    best_score, best_pos = float('inf'), -1
    for pos in range(search_end + 1):
        score = float(np.mean(np.abs(frame_mean[pos:pos + PROBE_H] - probe_mean)))
        if score < best_score:
            best_score, best_pos = score, pos

    if best_pos < 0 or best_score > MAX_SCORE:
        return None   # 매칭 실패

    new_start = best_pos + PROBE_H
    if new_start >= H - 3:
        return None   # 새 콘텐츠 없음 (동일 프레임)

    return new_start


def stitch_images(images: list) -> Image.Image:
    """프로브 매칭으로 프레임을 순서대로 이어 붙임.
    각 프레임에서 누적 이미지 하단과 일치하는 위치를 찾아 정확하게 새 픽셀만 추가."""
    if not images:
        return None
    if len(images) == 1:
        return images[0]

    result = images[0]
    for frame in images[1:]:
        new_start = _find_new_content_start(result, frame)
        if new_start is None:
            continue
        w        = min(result.width, frame.width)
        new_strip = frame.crop((0, new_start, w, frame.height))
        if new_strip.height < 3:
            continue
        combined = Image.new("RGB", (w, result.height + new_strip.height))
        combined.paste(result.crop((0, 0, w, result.height)), (0, 0))
        combined.paste(new_strip, (0, result.height))
        result = combined

    return result


def get_work_area():
    try:
        return win32gui.SystemParametersInfo(0x30)
    except Exception:
        import ctypes
        return (0, 0, ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1))


def clip_to_work_area(l, t, r, b):
    wl, wt, wr, wb = get_work_area()
    return max(l, wl), max(t, wt), min(r, wr), min(b, wb)


_BROWSER_CONTENT_CLASSES = frozenset({
    "Chrome_RenderWidgetHostHWND",
    "Internet Explorer_Server",
})


def get_window_rect_at(x: int, y: int):
    try:
        hwnd = win32gui.WindowFromPoint((x, y))
        cls  = win32gui.GetClassName(hwnd)
        if cls in _BROWSER_CONTENT_CLASSES:
            l, t, r, b = win32gui.GetWindowRect(hwnd)
        else:
            hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
            l, t, r, b = win32gui.GetWindowRect(hwnd)
        if r - l > 50 and b - t > 50:
            return (l, t, r, b)
    except Exception:
        pass
    return None


def do_scroll(cx, cy, clicks):
    win32api.SetCursorPos((cx, cy))
    time.sleep(0.03)
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -120 * clicks, 0)


def do_scroll_hwnd(hwnd, cx, cy, clicks):
    """커서를 움직이지 않고 창 핸들에 직접 WM_MOUSEWHEEL 전송.
    사용자가 중지 버튼을 클릭할 수 있도록 커서 위치를 건드리지 않음."""
    try:
        delta = -120 * clicks          # 음수 = 아래로 스크롤
        wparam = (delta & 0xFFFF) << 16            # HIWORD = wheel delta
        lparam = ((cy & 0xFFFF) << 16) | (cx & 0xFFFF)  # screen 좌표
        win32gui.SendMessage(hwnd, 0x020A, wparam, lparam)  # WM_MOUSEWHEEL
    except Exception:
        # 폴백: 전통적 방식 (커서 이동 포함)
        win32api.SetCursorPos((cx, cy))
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -120 * clicks, 0)


def _grab_region(x1: int, y1: int, x2: int, y2: int) -> Image.Image | None:
    """mss로 지정 영역 캡처. 크기가 유효하지 않으면 None 반환."""
    w, h = x2 - x1, y2 - y1
    if w < 1 or h < 1:
        return None
    with mss.mss() as sct:
        shot = sct.grab({"left": x1, "top": y1, "width": w, "height": h})
        return Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")


def _tb_sep(parent):
    """툴바 세로 구분선"""
    tk.Frame(parent, bg="#555", width=1).pack(side="left", fill="y", padx=4, pady=3)


def _tb_btn(parent, text, cmd, active=False, **kw):
    bg = "#5566cc" if active else "#3d3d3d"
    b = tk.Button(parent, text=text, command=cmd,
                  bg=bg, fg="white", relief="flat",
                  font=("맑은 고딕", 9), padx=5, pady=3,
                  activebackground="#6677dd", activeforeground="white",
                  **kw)
    b.pack(side="left", padx=1, pady=2)
    return b


# ─── 영역 선택 오버레이 ────────────────────────────────────────────────────────

class RegionSelector:
    def __init__(self, on_select):
        self.on_select = on_select
        self.start_x = self.start_y = 0
        self.rect_id = self.size_id = None
        self.dragging = False

        with mss.mss() as sct:
            mon  = sct.monitors[0]   # 모든 모니터를 합친 가상 데스크탑
            shot = sct.grab(mon)
            screen = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
            self.sw, self.sh = shot.width, shot.height
            self.vx, self.vy = mon["left"], mon["top"]   # 가상 데스크탑 원점 (음수 가능)

        ov   = Image.new("RGBA", screen.size, (0, 0, 30, 110))
        dark = Image.alpha_composite(screen.convert("RGBA"), ov).convert("RGB")

        self.win = tk.Toplevel()
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        # 가상 데스크탑 전체를 덮도록 위치 지정 (멀티모니터 대응)
        self.win.geometry(f"{self.sw}x{self.sh}+{self.vx}+{self.vy}")

        self.cv = tk.Canvas(self.win, width=self.sw, height=self.sh,
                            cursor="cross", highlightthickness=0, bd=0)
        self.cv.pack(fill="both", expand=True)
        self.photo = ImageTk.PhotoImage(dark)
        self.cv.create_image(0, 0, anchor="nw", image=self.photo)

        cx = self.sw // 2
        self.cv.create_rectangle(cx - 330, 10, cx + 330, 58,
                                  fill="#000000", outline="", stipple="gray50")
        self.cv.create_text(cx, 34,
            text="드래그: 영역 선택   |   클릭: 창 자동 선택   |   ESC: 취소",
            fill="white", font=("맑은 고딕", 12, "bold"))

        self.cv.bind("<ButtonPress-1>",   self._press)
        self.cv.bind("<B1-Motion>",        self._drag)
        self.cv.bind("<ButtonRelease-1>",  self._release)
        self.win.bind("<Escape>", lambda e: self._cancel())
        self.win.after(50, lambda: (self.win.lift(), self.win.focus_force()))

    def _press(self, e):
        self.start_x, self.start_y = e.x, e.y
        self.dragging = False
        if self.rect_id: self.cv.delete(self.rect_id)
        self.rect_id = self.cv.create_rectangle(e.x, e.y, e.x, e.y,
            outline="#00d4ff", width=2, fill="#00d4ff", stipple="gray25")

    def _drag(self, e):
        self.dragging = True
        if self.rect_id:
            self.cv.coords(self.rect_id, self.start_x, self.start_y, e.x, e.y)
        w, h = abs(e.x - self.start_x), abs(e.y - self.start_y)
        if self.size_id: self.cv.delete(self.size_id)
        self.size_id = self.cv.create_text(e.x + 8, e.y + 16,
            text=f"{w}×{h}", fill="white", font=("맑은 고딕", 10), anchor="nw")

    def _cancel(self):
        self.win.destroy()
        self.on_select(None)

    def _release(self, e):
        # 캔버스 좌표 → 실제 화면 절대 좌표 (vx/vy 오프셋 포함)
        sx = self.vx + e.x
        sy = self.vy + e.y
        is_click = (not self.dragging or
                    (abs(e.x - self.start_x) < 8 and abs(e.y - self.start_y) < 8))
        x1 = self.vx + min(self.start_x, e.x); y1 = self.vy + min(self.start_y, e.y)
        x2 = self.vx + max(self.start_x, e.x); y2 = self.vy + max(self.start_y, e.y)
        self.win.destroy()
        if is_click:
            rect = get_window_rect_at(sx, sy)
            self.on_select(rect)
        elif x2 - x1 > 10 and y2 - y1 > 10:
            self.on_select((x1, y1, x2, y2))
        else:
            self.on_select(None)


# ─── 지정 사이즈 선택 ──────────────────────────────────────────────────────────

class FixedSizeSelector:
    """지정한 크기의 박스를 커서와 함께 이동 → 클릭하여 캡처"""

    def __init__(self, cap_w, cap_h, on_select):
        self.cap_w, self.cap_h = cap_w, cap_h
        self.on_select = on_select

        with mss.mss() as sct:
            mon  = sct.monitors[0]
            shot = sct.grab(mon)
            screen = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
            self.sw, self.sh = shot.width, shot.height
            self.vx, self.vy = mon["left"], mon["top"]

        ov   = Image.new("RGBA", screen.size, (0, 0, 30, 90))
        dark = Image.alpha_composite(screen.convert("RGBA"), ov).convert("RGB")

        self.win = tk.Toplevel()
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.geometry(f"{self.sw}x{self.sh}+{self.vx}+{self.vy}")

        self.cv = tk.Canvas(self.win, width=self.sw, height=self.sh,
                            cursor="crosshair", highlightthickness=0)
        self.cv.pack(fill="both", expand=True)
        self.photo = ImageTk.PhotoImage(dark)
        self.cv.create_image(0, 0, anchor="nw", image=self.photo)

        cx = self.sw // 2
        self.cv.create_text(cx, 30,
            text=f"클릭하여  {cap_w}×{cap_h}  영역 캡처   |   ESC: 취소",
            fill="white", font=("맑은 고딕", 12, "bold"))

        self._box_id = None
        self.cv.bind("<Motion>",          self._move)
        self.cv.bind("<ButtonPress-1>",   self._click)
        self.win.bind("<Escape>", lambda e: self._cancel())
        self.win.after(50, lambda: (self.win.lift(), self.win.focus_force()))

    def _move(self, e):
        if self._box_id: self.cv.delete(self._box_id)
        x1, y1 = e.x - self.cap_w // 2, e.y - self.cap_h // 2
        x2, y2 = x1 + self.cap_w, y1 + self.cap_h
        self._box_id = self.cv.create_rectangle(x1, y1, x2, y2,
                                                  outline="#00d4ff", width=2)

    def _click(self, e):
        # 캔버스 좌표 → 실제 화면 절대 좌표
        x = self.vx + e.x - self.cap_w // 2
        y = self.vy + e.y - self.cap_h // 2
        self.win.destroy()
        self.on_select((x, y, x + self.cap_w, y + self.cap_h))

    def _cancel(self):
        self.win.destroy()
        self.on_select(None)


# ─── 이미지 편집기 ─────────────────────────────────────────────────────────────

class EditorWindow:
    """알캡처 스타일 이미지 편집기"""

    ZOOMS = [0.1, 0.15, 0.2, 0.25, 0.33, 0.5, 0.67, 0.75,
             1.0, 1.25, 1.5, 2.0, 3.0, 4.0]

    TOOLS = [
        ("pen",       "✏ 펜"),
        ("highlight", "🖌 형광펜"),
        ("text",      "A 텍스트"),
        ("rect",      "□ 사각형"),
        ("ellipse",   "○ 타원"),
        ("arrow",     "→ 화살표"),
        ("line",      "/ 선"),
        ("mosaic",    "▒ 모자이크"),
        ("crop",      "✂ 크롭"),
        ("eraser",    "◻ 지우개"),
    ]

    def __init__(self, parent, img: Image.Image,
                 on_update=None, on_close=None, history=None):
        self.parent    = parent
        self.img       = img.copy()
        self.orig      = img.copy()
        self.on_update = on_update
        self.on_close  = on_close
        self._history  = history or []     # CaptureApp.history 참조
        self.undo_stack: list[Image.Image] = []
        self.redo_stack: list[Image.Image] = []

        self.tool    = "pen"
        self.color   = "#ff0000"
        self.size    = 3
        self.fill    = False          # 채우기 모드 (사각형·타원)
        self.dash    = False          # 점선 모드 (사각형·타원·선)
        self.zi      = self._fit_idx()
        self._img_ox = 0              # 이미지 캔버스 내 X 오프셋 (중앙 배치)
        self._img_oy = 0              # 이미지 캔버스 내 Y 오프셋

        self._sx = self._sy = 0.0
        self._pts: list[tuple] = []
        self._prev_id = None

        # 텍스트 오버레이 목록 (PIL에 확정되기 전 드래그 가능한 상태)
        # 각 항목: {text, ix, iy, fs, color, cid}
        self._text_items: list[dict] = []
        self._txt_drag: dict | None = None   # 드래그 중인 항목 정보

        self.win = tk.Toplevel(parent)
        self.win.title(f"편집기  [{img.width}×{img.height}px]")
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_toolbar()
        self._build_main()          # 캔버스 + 이력 패널을 나란히 배치
        self._set_tool("pen")
        self._refresh()
        # 창이 실제로 렌더링된 뒤 한 번 더 렌더해 중앙 배치 정확히 적용
        self.win.after(60, self._refresh)
        # 렌더 완료 후 캡처툴과 겹치지 않는 위치로 이동
        self.win.after(80, self._place_window)
        # 창 열리자마자 포커스 확보 → 클릭 없이 Ctrl+C 바로 사용 가능
        self.win.after(120, lambda: (self.win.lift(), self.win.focus_force()))

    # ── 프로퍼티 ─────────────────────────────────────────────────────────────

    @property
    def zoom(self):
        return self.ZOOMS[self.zi]

    # ── 초기화 ───────────────────────────────────────────────────────────────

    def _fit_idx(self):
        """원본(100%)이 기본. 2000×760 캔버스를 벗어날 때만 축소."""
        MAX_W, MAX_H = 1980, 755
        idx_100 = self.ZOOMS.index(1.0)
        if self.img.width <= MAX_W and self.img.height <= MAX_H:
            return idx_100   # 원본 크기로 표시
        # 100% 미만에서 맞는 가장 큰 줌 찾기
        best = 0
        for i, z in enumerate(self.ZOOMS):
            if z <= 1.0 and self.img.width * z <= MAX_W and self.img.height * z <= MAX_H:
                best = i
        return best

    def _build_toolbar(self):
        tb = tk.Frame(self.win, bg="#2d2d2d", pady=3)
        tb.pack(fill="x")

        # 파일
        _tb_btn(tb, "💾 저장", self._save)
        _tb_btn(tb, "📋 복사", self._copy)
        _tb_sep(tb)

        # 실행취소 / 다시실행
        self.btn_undo = _tb_btn(tb, "↩ 취소", self._undo)
        self.btn_redo = _tb_btn(tb, "↪ 재실행", self._redo)
        _tb_sep(tb)

        # 도구 버튼
        self._tool_btns = {}
        for key, lbl in self.TOOLS:
            b = _tb_btn(tb, lbl, lambda k=key: self._set_tool(k))
            self._tool_btns[key] = b
        _tb_sep(tb)

        # 색상
        self._color_btn = tk.Button(tb, text="  ■  ", command=self._pick_color,
                                     bg=self.color, relief="flat", padx=6, pady=3)
        self._color_btn.pack(side="left", padx=3, pady=2)

        # 채우기 토글 (사각형·타원에만 효과)
        self._fill_btn = tk.Button(tb, text="☐ 채우기", command=self._toggle_fill,
                                    bg="#3d3d3d", fg="white", relief="flat",
                                    font=("맑은 고딕", 9), padx=5, pady=3,
                                    activebackground="#6677dd", activeforeground="white",
                                    cursor="hand2")
        self._fill_btn.pack(side="left", padx=1, pady=2)

        # 점선 토글 (사각형·타원·선에 효과)
        self._dash_btn = tk.Button(tb, text="☐ 점선", command=self._toggle_dash,
                                    bg="#3d3d3d", fg="white", relief="flat",
                                    font=("맑은 고딕", 9), padx=5, pady=3,
                                    activebackground="#6677dd", activeforeground="white",
                                    cursor="hand2")
        self._dash_btn.pack(side="left", padx=1, pady=2)
        _tb_sep(tb)

        # 두께
        tk.Label(tb, text="두께", bg="#2d2d2d", fg="#aaa",
                 font=("맑은 고딕", 8)).pack(side="left")
        self._size_var = tk.IntVar(value=self.size)
        size_cb = ttk.Combobox(tb, textvariable=self._size_var,
                                values=[1, 2, 3, 5, 8, 12, 20], width=3, state="readonly")
        size_cb.pack(side="left", padx=2)
        size_cb.bind("<<ComboboxSelected>>",
                     lambda e: setattr(self, "size", self._size_var.get()))
        _tb_sep(tb)

        # 글꼴 크기 (텍스트 도구)
        tk.Label(tb, text="글꼴", bg="#2d2d2d", fg="#aaa",
                 font=("맑은 고딕", 8)).pack(side="left")
        self._tsize_var = tk.IntVar(value=20)
        ts_cb = ttk.Combobox(tb, textvariable=self._tsize_var,
                              values=[12, 16, 20, 24, 32, 40, 56, 72], width=3, state="readonly")
        ts_cb.pack(side="left", padx=2)
        # (줌 컨트롤은 캔버스 호버 오버레이로 이동)

    def _build_main(self):
        outer = tk.Frame(self.win, bg="#1e1e1e")
        outer.pack(fill="both", expand=True)

        # ── 캔버스 영역
        cf = tk.Frame(outer, bg="#1e1e1e")
        cf.pack(side="left", fill="both", expand=True)
        hs = ttk.Scrollbar(cf, orient="horizontal")
        vs = ttk.Scrollbar(cf, orient="vertical")
        self.cv = tk.Canvas(cf, bg="#1e1e1e", cursor="crosshair",
                             highlightthickness=0,
                             xscrollcommand=hs.set, yscrollcommand=vs.set)
        hs.config(command=self.cv.xview)
        vs.config(command=self.cv.yview)
        hs.pack(side="bottom", fill="x")
        vs.pack(side="right",  fill="y")
        self.cv.pack(fill="both", expand=True)

        self.cv.bind("<ButtonPress-1>",    self._press)
        self.cv.bind("<B1-Motion>",         self._drag)
        self.cv.bind("<ButtonRelease-1>",   self._release)
        self.cv.bind("<MouseWheel>",          self._wheel)
        self.cv.bind("<Control-MouseWheel>",  self._cwheel)
        self.win.bind("<Control-z>", lambda e: self._undo())
        self.win.bind("<Control-y>", lambda e: self._redo())

        # ── 줌 오버레이 (캔버스 좌하단, 호버 시에만 표시) ──────────────
        self._zoom_hide_job = None

        zbar = tk.Frame(cf, bg="#1a1a28", bd=0)
        zbar.place(relx=0.0, rely=1.0, anchor="sw", x=10, y=-10)
        zbar.place_forget()   # 처음엔 숨김
        self._zoom_bar = zbar

        def _zbtn(text, cmd):
            b = tk.Button(zbar, text=text, command=cmd,
                          bg="#2a2a3a", fg="white", relief="flat",
                          font=("맑은 고딕", 8), padx=7, pady=3,
                          cursor="hand2",
                          activebackground=_C["btn_act"], activeforeground="white")
            b.pack(side="left", padx=1, pady=2)
            return b

        _zbtn("−", self._zout)
        self.zlbl = tk.Label(zbar, text="", bg="#1a1a28", fg="#bbbbdd",
                             font=("맑은 고딕", 9), width=5)
        self.zlbl.pack(side="left")
        _zbtn("+", self._zin)
        _zbtn("맞춤", self._zfit)
        _zbtn("100%", self._z100)

        # 호버 show / delayed hide
        def _zoom_show(e=None):
            if self._zoom_hide_job:
                self.win.after_cancel(self._zoom_hide_job)
                self._zoom_hide_job = None
            zbar.place(relx=0.0, rely=1.0, anchor="sw", x=10, y=-10)

        def _zoom_hide_soon(e=None):
            self._zoom_hide_job = self.win.after(
                220, lambda: zbar.place_forget()
            )

        for w in [self.cv, zbar] + zbar.winfo_children():
            w.bind("<Enter>", lambda e: _zoom_show())
            w.bind("<Leave>", lambda e: _zoom_hide_soon())

        # ── 우측 캡처 이력 패널
        self._build_history_panel(outer)

        # ── 히스토리 단축키 ──────────────────────────────────────────────
        self.win.bind("<Control-c>", lambda e: self._hist_copy())
        self.win.bind("<Control-s>", lambda e: self._hist_save())

    # ── 이력 패널 ────────────────────────────────────────────────────────────

    HIST_TW, HIST_TH = 160, 94   # 이력 썸네일 크기

    def _build_history_panel(self, parent):
        panel = tk.Frame(parent, bg=_C["toolbar"], width=186)
        panel.pack(side="right", fill="y")
        panel.pack_propagate(False)

        tk.Label(panel, text="캡처 기록", bg=_C["toolbar"], fg=_C["txt_dim"],
                 font=("맑은 고딕", 8)).pack(pady=(8, 4))

        # 스크롤 가능한 썸네일 목록
        gcv = tk.Canvas(panel, bg=_C["toolbar"], highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(panel, orient="vertical", command=gcv.yview)
        gcv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        gcv.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(gcv, bg=_C["toolbar"])
        win_id = gcv.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda e: gcv.configure(scrollregion=gcv.bbox("all")))
        gcv.bind("<Configure>",
                 lambda e: gcv.itemconfig(win_id, width=e.width))
        def _scroll(e):
            gcv.yview_scroll(-e.delta // 120, "units")

        gcv.bind("<MouseWheel>", _scroll)
        inner.bind("<MouseWheel>", _scroll)
        self._hist_scroll = _scroll   # _populate_history 에서 카드에도 바인딩

        self._hist_inner = inner
        self._hist_gcv   = gcv
        self._hist_cards = []   # (item, frame, img_lbl) 튜플 목록
        self._hist_sel   = None

        self._populate_history()

    def _make_hist_photo(self, img: Image.Image) -> ImageTk.PhotoImage:
        t = img.copy()
        t.thumbnail((self.HIST_TW, self.HIST_TH), Image.LANCZOS)
        bg = Image.new("RGB", (self.HIST_TW, self.HIST_TH), (24, 24, 37))
        bg.paste(t, ((self.HIST_TW - t.width) // 2,
                     (self.HIST_TH - t.height) // 2))
        return ImageTk.PhotoImage(bg)

    def _populate_history(self):
        """history 목록 전체를 이력 패널에 렌더링"""
        for w in self._hist_inner.winfo_children():
            w.destroy()
        self._hist_cards.clear()

        for item in self._history:
            photo = self._make_hist_photo(item["img"])

            # 카드 — 상대 배치로 삭제 버튼 오버레이
            card = tk.Frame(self._hist_inner, bg=_C["card"], cursor="hand2")
            card.pack(fill="x", padx=4, pady=3)

            il = tk.Label(card, image=photo, bg=_C["card"])
            il.image = photo
            il.pack(padx=3, pady=(4, 1))

            tl = tk.Label(card, text=item["label"].split()[0],
                          bg=_C["card"], fg=_C["txt_dim"],
                          font=("맑은 고딕", 7))
            tl.pack(pady=(0, 4))

            # 삭제 버튼 — 카드 우상단 오버레이
            del_btn = tk.Button(
                card, text="✕",
                bg="#c0392b", fg="white", relief="flat",
                font=("맑은 고딕", 7, "bold"),
                padx=3, pady=0, cursor="hand2",
                activebackground="#e74c3c", activeforeground="white",
                command=lambda it=item: self._delete_history(it)
            )
            del_btn.place(relx=1.0, rely=0.0, anchor="ne", x=-3, y=3)

            self._hist_cards.append((item, card, il, tl))

            for w in (card, il, tl):
                w.bind("<Button-1>",
                       lambda e, it=item, c=card: self._load_history(it, c))
                w.bind("<MouseWheel>", self._hist_scroll)
            del_btn.bind("<MouseWheel>", self._hist_scroll)

    def _hist_selected(self):
        """현재 선택된 히스토리 카드의 (item, card, img_label) 반환."""
        for item, card, il, tl in self._hist_cards:
            if card is self._hist_sel:
                return item, card, il
        return None, None, None

    def _canvas_blink(self):
        """이미지 위에 반투명 딤 오버레이를 잠시 표시해 복사 피드백을 준다.
        캡처 직전 화면이 살짝 어두워지는 느낌과 동일한 방식."""
        iw = int(self.img.width  * self.zoom)
        ih = int(self.img.height * self.zoom)
        x1, y1 = self._img_ox, self._img_oy
        x2, y2 = x1 + iw, y1 + ih
        self.cv.delete("copy_dim")
        self.cv.create_rectangle(
            x1, y1, x2, y2,
            fill="#000000", stipple="gray25",
            outline="", tags="copy_dim"
        )
        self.win.after(280, lambda: self.cv.delete("copy_dim"))

    def _hist_copy(self):
        """현재 편집기에 표시된 이미지를 클립보드에 복사."""
        image_to_clipboard(self.img)
        self._canvas_blink()

    def _hist_save(self):
        """현재 편집기에 표시된 이미지를 파일로 저장."""
        name = f"capture_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        path = filedialog.asksaveasfilename(
            defaultextension=".png", initialfile=name,
            filetypes=[("PNG", "*.png"), ("JPEG", "*.jpg"),
                       ("BMP", "*.bmp"), ("모든 파일", "*.*")],
            parent=self.win)
        if path:
            self.img.save(path)

    def _delete_history(self, item: dict):
        """이력 항목 삭제 — 목록·파일·인덱스 모두 제거."""
        # 현재 편집 중인 이미지와 같은 항목이면 선택 해제
        if self._hist_sel is not None:
            try:
                sel_items = [it for it, c, *_ in self._hist_cards if c == self._hist_sel]
                if sel_items and sel_items[0] is item:
                    self._hist_sel = None
            except Exception:
                pass

        # history 리스트에서 제거
        self._history[:] = [it for it in self._history if it is not item]

        # 디스크 파일 삭제 (CaptureApp 참조 통해)
        try:
            app_history = self.parent.history  # type: ignore
            app_history[:] = [it for it in app_history if it is not item]
            if item.get("_file"):
                fp = os.path.join(self.parent._cache_dir, item["_file"])  # type: ignore
                if os.path.exists(fp):
                    os.unlink(fp)
            self.parent._persist_index()  # type: ignore
        except Exception:
            pass

        self._populate_history()

    def _load_history(self, item: dict, card: tk.Frame):
        """이력 항목 클릭 → 편집기에 해당 이미지 로드"""
        # 선택 하이라이트
        if self._hist_sel is not None:
            try:
                for w in self._hist_sel.winfo_children():
                    w.config(bg=_C["card"])
                self._hist_sel.config(bg=_C["card"])
            except Exception:
                pass
        card.config(bg=_C["card_sel"])
        for w in card.winfo_children():
            w.config(bg=_C["card_sel"])
        self._hist_sel = card

        # 편집기 이미지 교체
        self._push_undo()
        self.img  = item["img"].copy()
        self.orig = item["img"].copy()
        self.zi   = self._fit_idx()
        self.win.title(f"편집기  [{self.img.width}×{self.img.height}px]")
        self._refresh()

    # ── 도구 선택 ────────────────────────────────────────────────────────────

    def _set_tool(self, key):
        if key != "text":
            self._commit_text_items()
        self.tool = key
        for k, b in self._tool_btns.items():
            b.config(bg="#5566cc" if k == key else "#3d3d3d")
        cursor = {"text": "xterm", "eraser": "dotbox",
                  "crop": "crosshair"}.get(key, "crosshair")
        self.cv.config(cursor=cursor)

    def _pick_color(self):
        _, hx = colorchooser.askcolor(color=self.color, parent=self.win)
        if hx:
            self.color = hx
            self._color_btn.config(bg=hx)

    def _toggle_fill(self):
        self.fill = not self.fill
        self._fill_btn.config(
            text="☑ 채우기" if self.fill else "☐ 채우기",
            bg="#5566cc" if self.fill else "#3d3d3d")

    def _toggle_dash(self):
        self.dash = not self.dash
        self._dash_btn.config(
            text="☑ 점선" if self.dash else "☐ 점선",
            bg="#5566cc" if self.dash else "#3d3d3d")

    # ── 실행취소 ─────────────────────────────────────────────────────────────

    def _txt_snapshot(self) -> list:
        """현재 텍스트 오버레이 목록의 복사본 (cid 제외)"""
        return [
            {k: v for k, v in item.items() if k != "cid"}
            for item in self._text_items
        ]

    def _push_undo(self):
        self.undo_stack.append((self.img.copy(), self._txt_snapshot()))
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _restore_state(self, state):
        """(img, text_items_snap) 상태를 복원. _refresh가 오버레이를 재생성함."""
        img, txt_snap = state
        self.img = img
        self._text_items = txt_snap   # cid 없는 스냅샷, _refresh에서 재생성
        self._refresh()

    def _undo(self):
        if not self.undo_stack:
            return
        self.redo_stack.append((self.img.copy(), self._txt_snapshot()))
        self._restore_state(self.undo_stack.pop())

    def _redo(self):
        if not self.redo_stack:
            return
        self.undo_stack.append((self.img.copy(), self._txt_snapshot()))
        self._restore_state(self.redo_stack.pop())

    # ── 렌더링 ───────────────────────────────────────────────────────────────

    def _refresh(self):
        z   = self.zoom
        iw  = max(1, int(self.img.width  * z))
        ih  = max(1, int(self.img.height * z))
        algo = Image.LANCZOS if z <= 1 else Image.NEAREST
        self.photo = ImageTk.PhotoImage(self.img.resize((iw, ih), algo))

        # 캔버스 실제 크기 (렌더링 전이면 2000×800 기준 캔버스 추정)
        cw = self.cv.winfo_width()
        ch = self.cv.winfo_height()
        if cw < 10: cw = 1160   # 1350 - 역사패널(186) - 스크롤바(4)
        if ch < 10: ch = 760    # 800  - 툴바(~40)

        # 이미지 중앙 배치 오프셋 (이미지가 캔버스보다 작을 때만 여백 추가)
        ox = max((cw - iw) // 2, 0)
        oy = max((ch - ih) // 2, 0)
        self._img_ox = ox
        self._img_oy = oy

        sr_w = max(iw + ox * 2, cw)
        sr_h = max(ih + oy * 2, ch)

        self.cv.delete("img")
        self.cv.create_image(ox, oy, anchor="nw", image=self.photo, tags="img")
        self.cv.config(scrollregion=(0, 0, sr_w, sr_h))

        # 이미지가 캔버스보다 크면 중앙으로 스크롤
        if iw > cw:
            self.cv.xview_moveto(max(0.0, (iw / 2 - cw / 2)) / sr_w)
        if ih > ch:
            self.cv.yview_moveto(max(0.0, (ih / 2 - ch / 2)) / sr_h)

        self.zlbl.config(text=f"{int(z * 100)}%")
        self.win.geometry("1350x800")

        # 텍스트 오버레이 위치 갱신 (zoom 변경 시 캔버스 좌표 재계산)
        self.cv.delete("txt_overlay")
        for item in self._text_items:
            cx, cy = self._i2c(item["ix"], item["iy"])
            fs = max(8, int(item["fs"] * self.zoom))
            item["cid"] = self.cv.create_text(
                cx, cy, text=item["text"], fill=item["color"],
                anchor="nw", font=("맑은 고딕", -fs, "bold"),  # 음수 = 픽셀 단위
                tags="txt_overlay")

    def _place_window(self):
        """캡처툴(부모 창)과 겹치지 않는 위치에 편집기를 배치."""
        self.win.update_idletasks()
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        ew = self.win.winfo_width()
        eh = self.win.winfo_height()
        root = self.win.master
        rx = root.winfo_x()
        ry = root.winfo_y()
        rw = root.winfo_width()
        rh = root.winfo_height()
        # 편집기 x: 캡처툴과 같은 x 정렬, 화면 밖으로 나가지 않게 클램프
        ex = max(0, min(rx, sw - ew))
        # 아래 공간이 충분하면 캡처툴 아래, 아니면 위
        ey = ry + rh + 6
        if ey + eh > sh:
            ey = max(0, ry - eh - 6)
        self.win.geometry(f"+{ex}+{ey}")

    def _cv_pos(self, ex, ey):
        """위젯 좌표 → 캔버스 절대 좌표"""
        return self.cv.canvasx(ex), self.cv.canvasy(ey)

    def _c2i(self, cx, cy):
        """캔버스 좌표 → 이미지 픽셀 좌표 (중앙 오프셋 보정 + 클램프)"""
        ix = max(0, min(self.img.width  - 1, int((cx - self._img_ox) / self.zoom)))
        iy = max(0, min(self.img.height - 1, int((cy - self._img_oy) / self.zoom)))
        return ix, iy

    # ── 마우스 이벤트 ─────────────────────────────────────────────────────────

    def _press(self, e):
        self._sx, self._sy = self._cv_pos(e.x, e.y)
        self._pts = [(self._sx, self._sy)]
        self._prev_id = None
        self._txt_drag = None
        if self.tool == "text":
            # 기존 텍스트 아이템 클릭 → 드래그 모드
            cx, cy = self._sx, self._sy
            hit = self.cv.find_overlapping(cx - 6, cy - 6, cx + 6, cy + 6)
            hit_ids = set(hit)
            for i, item in enumerate(self._text_items):
                if item["cid"] in hit_ids:
                    self._push_undo()   # 드래그 전 위치 저장
                    self._txt_drag = dict(idx=i, last_cx=cx, last_cy=cy)
                    return
            # 빈 공간 클릭 → 새 텍스트 입력
            self._do_text(e.x, e.y)

    def _drag(self, e):
        cx, cy = self._cv_pos(e.x, e.y)
        lw = max(1, int(self.size * self.zoom))

        # 텍스트 드래그
        if self._txt_drag is not None:
            item = self._text_items[self._txt_drag["idx"]]
            dx = cx - self._txt_drag["last_cx"]
            dy = cy - self._txt_drag["last_cy"]
            self.cv.move(item["cid"], dx, dy)
            item["ix"] = max(0, item["ix"] + int(dx / self.zoom))
            item["iy"] = max(0, item["iy"] + int(dy / self.zoom))
            self._txt_drag["last_cx"] = cx
            self._txt_drag["last_cy"] = cy
            return

        if self._prev_id:
            self.cv.delete(self._prev_id)

        if self.tool in ("pen", "eraser"):
            self._pts.append((cx, cy))
            if len(self._pts) >= 2:
                flat = [c for p in self._pts for c in p]
                col  = self.color if self.tool == "pen" else "#ffffff"
                self._prev_id = self.cv.create_line(
                    *flat, fill=col, width=lw, capstyle="round", joinstyle="round")

        elif self.tool == "highlight":
            self._pts.append((cx, cy))
            if len(self._pts) >= 2:
                flat = [c for p in self._pts for c in p]
                self._prev_id = self.cv.create_line(
                    *flat, fill=self.color,
                    width=max(1, int(self.size * 6 * self.zoom)),
                    capstyle="round", joinstyle="round", stipple="gray50")

        elif self.tool == "rect":
            fc = self.color if self.fill else ""
            dk = (max(6, lw*4), max(4, lw*2)) if self.dash else ()
            self._prev_id = self.cv.create_rectangle(
                self._sx, self._sy, cx, cy,
                outline=self.color, fill=fc, width=lw, dash=dk)

        elif self.tool == "ellipse":
            fc = self.color if self.fill else ""
            dk = (max(6, lw*4), max(4, lw*2)) if self.dash else ()
            self._prev_id = self.cv.create_oval(
                self._sx, self._sy, cx, cy,
                outline=self.color, fill=fc, width=lw, dash=dk)

        elif self.tool == "line":
            self._prev_id = self.cv.create_line(
                self._sx, self._sy, cx, cy,
                fill=self.color, width=lw, capstyle="round")

        elif self.tool == "arrow":
            self._prev_id = self.cv.create_line(
                self._sx, self._sy, cx, cy,
                fill=self.color, width=lw,
                arrow=tk.LAST,
                arrowshape=(max(10, lw * 3), max(14, lw * 4), max(4, lw)))

        elif self.tool in ("mosaic", "crop"):
            self._prev_id = self.cv.create_rectangle(
                self._sx, self._sy, cx, cy,
                outline="#00d4ff", width=2, dash=(4, 4))

    def _release(self, e):
        cx, cy = self._cv_pos(e.x, e.y)
        if self._txt_drag is not None:
            self._txt_drag = None
            return

        if self._prev_id:
            self.cv.delete(self._prev_id)
            self._prev_id = None

        if self.tool == "text":
            return

        ix1, iy1 = self._c2i(self._sx, self._sy)
        ix2, iy2 = self._c2i(cx, cy)
        pts_img  = [(self._c2i(px, py)) for px, py in self._pts]
        # 화살표/선은 방향이 중요하므로 정렬 전 좌표 보존
        ax1, ay1, ax2, ay2 = ix1, iy1, ix2, iy2
        # rect/ellipse/mosaic/crop 은 PIL이 x1≤x2, y1≤y2 필요
        ix1, ix2 = min(ix1, ix2), max(ix1, ix2)
        iy1, iy2 = min(iy1, iy2), max(iy1, iy2)

        self._push_undo()

        if self.tool == "pen":
            if len(pts_img) >= 2:
                draw = ImageDraw.Draw(self.img)
                draw.line(pts_img, fill=self.color, width=self.size, joint="curve")

        elif self.tool == "eraser":
            if len(pts_img) >= 2:
                # 원본 이미지에서 해당 영역을 복원 (흰색 덮기 X)
                mask = Image.new("L", self.img.size, 0)
                ImageDraw.Draw(mask).line(pts_img, fill=255,
                                          width=self.size * 4, joint="curve")
                self.img = Image.composite(self.orig, self.img, mask)

        elif self.tool == "highlight":
            if len(pts_img) >= 2:
                self._draw_highlight(pts_img)

        elif self.tool == "rect":
            draw = ImageDraw.Draw(self.img)
            if self.dash:
                self._draw_dashed_rect(draw, ix1, iy1, ix2, iy2)
            else:
                fc = self.color if self.fill else None
                draw.rectangle([ix1, iy1, ix2, iy2],
                               outline=self.color, fill=fc, width=self.size)

        elif self.tool == "ellipse":
            draw = ImageDraw.Draw(self.img)
            if self.dash:
                self._draw_dashed_ellipse(draw, ix1, iy1, ix2, iy2)
            else:
                fc = self.color if self.fill else None
                draw.ellipse([ix1, iy1, ix2, iy2],
                             outline=self.color, fill=fc, width=self.size)

        elif self.tool == "line":
            draw = ImageDraw.Draw(self.img)
            draw.line([ax1, ay1, ax2, ay2], fill=self.color, width=self.size)

        elif self.tool == "arrow":
            self._draw_arrow(ax1, ay1, ax2, ay2)

        elif self.tool == "mosaic":
            self._do_mosaic(ix1, iy1, ix2, iy2)

        elif self.tool == "crop":
            self._do_crop(ix1, iy1, ix2, iy2)
            return  # crop 내부에서 _refresh 호출

        self._refresh()

    # ── 도구 구현 ────────────────────────────────────────────────────────────

    def _draw_highlight(self, pts):
        overlay = Image.new("RGBA", self.img.size, (0, 0, 0, 0))
        draw    = ImageDraw.Draw(overlay)
        r = int(self.color[1:3], 16)
        g = int(self.color[3:5], 16)
        b = int(self.color[5:7], 16)
        flat = [c for p in pts for c in p]
        if len(flat) >= 4:
            draw.line(flat, fill=(r, g, b, 120),
                      width=self.size * 6, joint="curve")
        self.img = Image.alpha_composite(
            self.img.convert("RGBA"), overlay).convert("RGB")

    # ── 점선 헬퍼 ────────────────────────────────────────────────────────────

    @staticmethod
    def _dash_segments(x1, y1, x2, y2, dash, gap):
        """직선 구간을 (on/off) 세그먼트 좌표 리스트로 반환."""
        import math
        length = math.hypot(x2 - x1, y2 - y1)
        if length < 1:
            return []
        dx, dy = (x2 - x1) / length, (y2 - y1) / length
        segs, pos, on = [], 0.0, True
        while pos < length:
            end = min(pos + (dash if on else gap), length)
            if on:
                segs.append((x1 + dx*pos, y1 + dy*pos,
                             x1 + dx*end, y1 + dy*end))
            pos, on = end, not on
        return segs

    def _draw_dashed_rect(self, draw, x1, y1, x2, y2):
        d = max(8, self.size * 3)
        g = max(5, self.size * 2)
        fc = self.color if self.fill else None
        if fc:
            draw.rectangle([x1, y1, x2, y2], fill=fc)
        for side in [(x1,y1,x2,y1),(x2,y1,x2,y2),(x2,y2,x1,y2),(x1,y2,x1,y1)]:
            for seg in self._dash_segments(*side, d, g):
                draw.line(seg, fill=self.color, width=self.size)

    def _draw_dashed_ellipse(self, draw, x1, y1, x2, y2):
        cx, cy = (x1+x2)/2, (y1+y2)/2
        rx, ry = abs(x2-x1)/2, abs(y2-y1)/2
        if rx < 1 or ry < 1:
            return
        fc = self.color if self.fill else None
        if fc:
            draw.ellipse([x1, y1, x2, y2], fill=fc)
        # 타원 둘레 근사 (Ramanujan)
        h = ((rx-ry)/(rx+ry))**2
        perim = math.pi*(rx+ry)*(1 + 3*h/(10+math.sqrt(4-3*h)))
        n = max(200, int(perim * 2))
        d = max(8, self.size * 3)
        g = max(5, self.size * 2)
        on, seg_len = True, 0.0
        pts = [(cx + rx*math.cos(2*math.pi*i/n),
                cy + ry*math.sin(2*math.pi*i/n)) for i in range(n+1)]
        for i in range(len(pts)-1):
            px,py,qx,qy = *pts[i], *pts[i+1]
            step = math.hypot(qx-px, qy-py)
            if on:
                draw.line([px,py,qx,qy], fill=self.color, width=self.size)
            seg_len += step
            if on and seg_len >= d:
                on, seg_len = False, 0.0
            elif not on and seg_len >= g:
                on, seg_len = True, 0.0

    def _draw_arrow(self, x1, y1, x2, y2):
        draw = ImageDraw.Draw(self.img)
        draw.line([x1, y1, x2, y2], fill=self.color, width=self.size)
        if x1 == x2 and y1 == y2:
            return
        angle = math.atan2(y2 - y1, x2 - x1)
        alen  = max(12, self.size * 5)
        for side in (-0.45, 0.45):
            ax = x2 - alen * math.cos(angle + side)
            ay = y2 - alen * math.sin(angle + side)
            draw.line([x2, y2, int(ax), int(ay)], fill=self.color, width=self.size)

    def _i2c(self, ix, iy):
        """이미지 좌표 → 캔버스 절대 좌표"""
        return ix * self.zoom + self._img_ox, iy * self.zoom + self._img_oy

    def _ask_text(self):
        """다크 테마 텍스트 입력 다이얼로그. 입력값 또는 None 반환."""
        result = [None]
        dlg = tk.Toplevel(self.win)
        dlg.title("텍스트 입력")
        dlg.configure(bg=_C["bg"])
        dlg.resizable(False, False)
        dlg.transient(self.win)
        dlg.grab_set()
        dlg.attributes("-topmost", True)
        dlg.update_idletasks()
        wx = self.win.winfo_x() + self.win.winfo_width()  // 2 - 200
        wy = self.win.winfo_y() + self.win.winfo_height() // 2 - 75
        dlg.geometry(f"400x150+{wx}+{wy}")

        tk.Label(dlg, text="텍스트 입력", bg=_C["bg"], fg=_C["accent"],
                 font=("맑은 고딕", 10, "bold")).pack(pady=(14, 6))

        entry = tk.Entry(dlg, bg=_C["card"], fg=_C["txt"],
                         insertbackground=_C["txt"], relief="flat",
                         font=("맑은 고딕", 11), width=34,
                         highlightthickness=1, highlightcolor=_C["accent"],
                         highlightbackground=_C["btn"])
        entry.pack(padx=20, pady=2, ipady=7)
        entry.focus_set()

        def _ok(e=None):
            result[0] = entry.get()
            dlg.destroy()

        entry.bind("<Return>", _ok)
        entry.bind("<Escape>", lambda e: dlg.destroy())

        bf = tk.Frame(dlg, bg=_C["bg"])
        bf.pack(pady=10)
        tk.Button(bf, text="확인", command=_ok,
                  bg=_C["btn_act"], fg="white", relief="flat",
                  font=("맑은 고딕", 9), width=8, cursor="hand2").pack(side="left", padx=6)
        tk.Button(bf, text="취소", command=dlg.destroy,
                  bg=_C["btn"], fg=_C["txt"], relief="flat",
                  font=("맑은 고딕", 9), width=8, cursor="hand2").pack(side="left", padx=6)

        dlg.wait_window()
        return result[0] if result[0] else None

    def _do_text(self, ex, ey):
        """클릭 위치에 텍스트 오버레이 추가 (드래그 가능, 미확정 상태)."""
        text = self._ask_text()
        if not text:
            return
        self._push_undo()          # ← 텍스트 추가 직전에 상태 저장 (개별 취소 가능)
        ix, iy = self._c2i(*self._cv_pos(ex, ey))
        cx, cy = self._i2c(ix, iy)
        fs = self._tsize_var.get()
        disp_fs = max(8, int(fs * self.zoom))
        cid = self.cv.create_text(
            cx, cy, text=text, fill=self.color, anchor="nw",
            font=("맑은 고딕", -disp_fs, "bold"), tags="txt_overlay")  # 음수 = 픽셀 단위
        self._text_items.append(
            dict(text=text, ix=ix, iy=iy, fs=fs, color=self.color, cid=cid))

    def _commit_text_items(self):
        """오버레이 텍스트를 PIL 이미지에 확정하고 캔버스 항목 제거."""
        if not self._text_items:
            return
        # _push_undo는 _do_text에서 항목별로 이미 호출됨 — 여기서는 생략
        draw = ImageDraw.Draw(self.img)
        for item in self._text_items:
            font = None
            for name in ("malgunbd.ttf", "arialbd.ttf", "malgun.ttf", "arial.ttf"):
                try:
                    font = ImageFont.truetype(name, item["fs"])
                    break
                except Exception:
                    pass
            if font is None:
                font = ImageFont.load_default()
            draw.text((item["ix"], item["iy"]), item["text"],
                      fill=item["color"], font=font)
            try:
                self.cv.delete(item["cid"])
            except Exception:
                pass
        self._text_items.clear()
        self._txt_drag = None
        self._refresh()

    def _do_mosaic(self, x1, y1, x2, y2, block=14):
        x1, x2 = min(x1, x2), max(x1, x2)
        y1, y2 = min(y1, y2), max(y1, y2)
        x1, y1 = max(0, x1), max(0, y1)
        x2, y2 = min(self.img.width, x2), min(self.img.height, y2)
        if x2 - x1 < 2 or y2 - y1 < 2:
            return
        region = self.img.crop((x1, y1, x2, y2))
        sw = max(1, (x2 - x1) // block)
        sh = max(1, (y2 - y1) // block)
        pix = region.resize((sw, sh), Image.BOX).resize((x2 - x1, y2 - y1), Image.NEAREST)
        self.img.paste(pix, (x1, y1))

    def _do_crop(self, x1, y1, x2, y2):
        x1, x2 = min(x1, x2), max(x1, x2)
        y1, y2 = min(y1, y2), max(y1, y2)
        x1, y1 = max(0, x1), max(0, y1)
        x2, y2 = min(self.img.width, x2), min(self.img.height, y2)
        if x2 - x1 < 2 or y2 - y1 < 2:
            return
        self.img  = self.img.crop((x1, y1, x2, y2))
        self.orig = self.orig.crop((x1, y1, x2, y2))  # 지우개 기준도 함께 크롭
        self._refresh()

    # ── 줌 ──────────────────────────────────────────────────────────────────

    def _zin(self):
        if self.zi < len(self.ZOOMS) - 1:
            self.zi += 1; self._refresh()

    def _zout(self):
        if self.zi > 0:
            self.zi -= 1; self._refresh()

    def _z100(self):
        self.zi = self.ZOOMS.index(1.0); self._refresh()

    def _zfit(self):
        self.zi = self._fit_idx(); self._refresh()

    def _wheel(self, e):
        self.cv.yview_scroll(int(-e.delta / 120), "units")

    def _cwheel(self, e):
        self._zin() if e.delta > 0 else self._zout()

    # ── 파일 / 닫기 ──────────────────────────────────────────────────────────

    def _save(self):
        self._commit_text_items()
        name = f"capture_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        path = filedialog.asksaveasfilename(
            defaultextension=".png", initialfile=name,
            filetypes=[("PNG", "*.png"), ("JPEG", "*.jpg"),
                       ("BMP", "*.bmp"), ("모든 파일", "*.*")],
            parent=self.win)
        if path:
            self.img.save(path)
            messagebox.showinfo("저장", f"저장:\n{path}", parent=self.win)

    def _copy(self):
        image_to_clipboard(self.img)
        messagebox.showinfo("복사", "클립보드에 복사되었습니다.", parent=self.win)

    def _on_close(self):
        self._commit_text_items()
        if self.on_update and self.undo_stack:
            if messagebox.askyesno("편집 반영",
                                    "편집된 이미지를 히스토리에 반영하시겠습니까?",
                                    parent=self.win):
                self.on_update(self.img.copy())
        self.win.destroy()
        if self.on_close:
            try: self.on_close()
            except Exception: pass


# ─── 스크롤 캡처 ─────────────────────────────────────────────────────────────

class ScrollCapture:
    SCROLL_CLICKS  = 5
    SCROLL_DELAY   = 0.35
    BOTTOM_DIFF    = 2.0   # 픽셀 평균차 이하면 "프레임 동일" 판정
    BOTTOM_CONSEC  = 4     # 연속 N회 동일 프레임 → 바닥 확정
    MAX_FRAMES     = 300

    def __init__(self, parent, region, on_done):
        self.parent  = parent
        self.region  = region
        self.on_done = on_done
        self.running = True
        self.images  = []
        print(f"[SC] __init__ region={region}")

        x1, y1, x2, y2 = region
        sh = parent.winfo_screenheight()
        sw = parent.winfo_screenwidth()
        mini_h = 46
        my = y2 + 4 if y2 + mini_h + 4 <= sh else max(0, y1 - mini_h - 4)
        mx = min(max(0, x1), sw - 250)
        print(f"[SC] mini pos mx={mx} my={my}  screen={sw}x{sh}")

        self.mini = tk.Toplevel(parent)
        self.mini.overrideredirect(True)
        self.mini.attributes("-topmost", True)
        self.mini.geometry(f"244x{mini_h}+{mx}+{my}")
        self.mini.configure(bg="#1a1a2e")
        self.mini.deiconify()          # iconify된 부모에서 확실히 표시
        self.mini.lift()
        print(f"[SC] mini window created")

        self.lbl = tk.Label(self.mini, text="3초 후 시작 — 캡처 창 클릭!",
                            bg="#1a1a2e", fg="white", font=("맑은 고딕", 9))
        self.lbl.pack(side="left", padx=8, fill="x", expand=True)
        tk.Button(self.mini, text="■ 중지", command=self._stop,
                  bg="#e74c3c", fg="white", font=("맑은 고딕", 9, "bold"),
                  relief="flat", padx=8, pady=3).pack(side="right", padx=6, pady=5)

        threading.Thread(target=self._run, daemon=True).start()
        print(f"[SC] thread started")

    def _ui(self, fn):
        try:
            self.parent.after(0, fn)
        except Exception as e:
            print(f"[SC] _ui error: {e}")

    def _set(self, msg):
        print(f"[SC] _set: {msg}")
        self._ui(lambda m=msg: self._set_label(m))

    def _set_label(self, msg):
        try:
            self.lbl.config(text=msg)
        except Exception as e:
            print(f"[SC] _set_label error: {e}")

    def _run(self):
        print("[SC] _run start")
        self._set("1초 후 시작 — 지금 캡처 창 클릭!")
        time.sleep(1)
        if not self.running: return
        if not self.running: return
        print("[SC] countdown done, starting capture")

        x1, y1, x2, y2 = self.region
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
        mon = {"left": x1, "top": y1, "width": x2 - x1, "height": y2 - y1}
        print(f"[SC] mon={mon}  center=({cx},{cy})")

        # 커서를 캡처 영역 중앙에 한 번만 이동 후, 이후 스크롤은 hwnd에 직접 전송
        win32api.SetCursorPos((cx, cy))
        time.sleep(0.15)
        target_hwnd = win32gui.WindowFromPoint((cx, cy))
        print(f"[SC] target_hwnd={target_hwnd}")

        same_count = 0
        prev_img   = None

        try:
            with mss.mss() as sct:
                while self.running and len(self.images) < self.MAX_FRAMES:
                    shot = sct.grab(mon)
                    img  = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")

                    self.images.append(img)
                    n = len(self.images)
                    self._set(f"캡처 중... {n}장")

                    if n >= 2 and prev_img is not None:
                        diff = frame_diff(prev_img, img)
                        print(f"[SC] frame {n}  diff={diff:.2f}  same_count={same_count}")
                        # 바닥 판정: 프레임이 거의 동일 (페이지가 더 이상 안 움직임)
                        if diff < self.BOTTOM_DIFF:
                            same_count += 1
                            if same_count >= self.BOTTOM_CONSEC:
                                # 중복 프레임 제거
                                for _ in range(min(same_count, len(self.images) - 1)):
                                    self.images.pop()
                                self._set(f"바닥 도달! 총 {len(self.images)}장")
                                print(f"[SC] bottom detected  diff={diff:.2f}")
                                time.sleep(0.3)
                                break
                        else:
                            same_count = 0

                    prev_img = img
                    # 커서를 건드리지 않고 hwnd에 직접 스크롤 전송
                    do_scroll_hwnd(target_hwnd, cx, cy, self.SCROLL_CLICKS)
                    time.sleep(self.SCROLL_DELAY)
        except Exception as e:
            print(f"[SC] _run exception: {e}")
            import traceback; traceback.print_exc()

        print(f"[SC] _run done, images={len(self.images)}")
        self._finish()

    def _stop(self):
        print("[SC] _stop called")
        self.running = False

    def _finish(self):
        print(f"[SC] _finish images={len(self.images)}")
        images = list(self.images)

        def _done():
            print("[SC] _done on main thread")
            try:
                self.mini.destroy()
            except Exception as e:
                print(f"[SC] mini.destroy error: {e}")
            if not images:
                print("[SC] no images, calling on_done(None)")
                self.on_done(None)
                return
            def _stitch():
                print(f"[SC] stitching {len(images)} frames...")
                result = stitch_images(images)
                print(f"[SC] stitch done: {result.size if result else None}")
                self.parent.after(0, lambda: self.on_done(result))
            threading.Thread(target=_stitch, daemon=True).start()

        self._ui(_done)


# ─── OCR ─────────────────────────────────────────────────────────────────────

def _do_ocr(pil_img: Image.Image) -> tuple[str | None, str]:
    """이미지에서 텍스트 추출.
    반환: (text | None, debug_info)
    ① PowerShell + Windows 내장 OCR
       - GetFileFromPathAsync 대신 MemoryStream + AsRandomAccessStream 사용
       - AsTask 리플렉션으로 WinRT async 처리 (Status 폴링보다 안정적)
    ② pytesseract 폴백
    """
    import tempfile
    img_tmp = ps_tmp = None
    debug = ""
    try:
        img_tmp = tempfile.mktemp(suffix=".png")
        ps_tmp  = tempfile.mktemp(suffix=".ps1")
        pil_img.save(img_tmp)

        img_esc = img_tmp.replace("'", "''")

        # WinRT async 문제 원인:
        #   1) PS IDispatch로 COM 객체 Status 읽기 불가 → 폴링 방식 실패
        #   2) STA 스레드 blocking → 콜백 데드락
        # 해결: MTA Runspace + AsTask<T>(명시적 타입) 반영
        ps = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[void][Windows.Media.Ocr.OcrEngine,            Windows.Foundation, ContentType=WindowsRuntime]
[void][Windows.Media.Ocr.OcrResult,             Windows.Foundation, ContentType=WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder,  Windows.Foundation, ContentType=WindowsRuntime]
[void][Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType=WindowsRuntime]

$imgPath = '{img_esc}'

$rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
$rs.ApartmentState = [System.Threading.ApartmentState]::MTA
$rs.Open()
$ps2 = [System.Management.Automation.PowerShell]::Create()
$ps2.Runspace = $rs
[void]$ps2.AddScript({{
    param([string]$imgPath)
    $ErrorActionPreference = 'Stop'
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    [void][Windows.Media.Ocr.OcrEngine,            Windows.Foundation, ContentType=WindowsRuntime]
    [void][Windows.Media.Ocr.OcrResult,             Windows.Foundation, ContentType=WindowsRuntime]
    [void][Windows.Graphics.Imaging.BitmapDecoder,  Windows.Foundation, ContentType=WindowsRuntime]
    [void][Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType=WindowsRuntime]

    # AsTask<T> via reflection — Status 폴링 대신 .NET Task 사용 (IDispatch 우회)
    $atM = [System.WindowsRuntimeSystemExtensions].GetMethods() |
           Where-Object {{ $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 }} |
           Select-Object -First 1

    function AwaitTask($op, [Type]$rt) {{
        $t = $atM.MakeGenericMethod($rt).Invoke($null, @($op))
        if (-not $t.Wait(25000)) {{ throw "AsTask timeout: $($rt.Name)" }}
        if ($t.IsFaulted) {{ throw $t.Exception.GetBaseException().Message }}
        $t.Result
    }}

    $bytes = [System.IO.File]::ReadAllBytes($imgPath)
    $ms    = [System.IO.MemoryStream]::new($bytes)
    $ras   = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($ms)

    $decoder = AwaitTask ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($ras))  ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap  = AwaitTask ($decoder.GetSoftwareBitmapAsync())                             ([Windows.Graphics.Imaging.SoftwareBitmap])
    $ras.Dispose(); $ms.Dispose()

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) {{ throw 'OCR engine unavailable' }}

    $result = AwaitTask ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $result.Lines | ForEach-Object {{ $_.Text }}
}})
[void]$ps2.AddParameter('imgPath', $imgPath)

$output = $ps2.Invoke()
foreach ($e in $ps2.Streams.Error) {{ [Console]::Error.WriteLine($e.ToString()) }}
$ps2.Dispose()
$rs.Close()
$output
"""
        with open(ps_tmp, "w", encoding="utf-8") as f:
            f.write(ps)

        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-File", ps_tmp],
            capture_output=True, text=True, timeout=90,
            encoding="utf-8", errors="replace")

        debug = (r.stderr or "").strip()
        text  = (r.stdout or "").strip()
        if text:
            return text, ""
        # stdout 비어있음 → stderr 포함해 반환
        return None, debug or "(empty output)"

    except Exception as ex:
        debug = str(ex)
    finally:
        for p in (img_tmp, ps_tmp):
            if p:
                try: os.unlink(p)
                except Exception: pass

    # ② pytesseract 폴백
    try:
        import pytesseract
        text = pytesseract.image_to_string(pil_img, lang="kor+eng").strip()
        return (text or None), ""
    except ImportError:
        pass
    except Exception as ex:
        debug += f"\npytesseract: {ex}"

    return None, debug


# ─── 화면 녹화 ───────────────────────────────────────────────────────────────

class ScreenRecorder:
    """mss + opencv-python 으로 지정 영역을 녹화 (고화질 30fps)"""

    FPS = 30

    # 시도할 코덱 순서: H.264 → XVID → mp4v
    _CODECS = [
        ("avc1", ".mp4"),
        ("XVID", ".avi"),
        ("mp4v", ".mp4"),
    ]

    def __init__(self, region: tuple, on_done):
        self.region  = region
        self.on_done = on_done
        self.running = False
        self._thread = None
        self.out_path: str | None = None

    def start(self):
        self.running = True
        self._thread = threading.Thread(target=self._record, daemon=True)
        self._thread.start()

    def stop(self):
        self.running = False

    def _record(self):
        try:
            import cv2
        except ImportError:
            self.on_done(None, "opencv-python 패키지가 없습니다.\npip install opencv-python")
            return

        x1, y1, x2, y2 = self.region
        w, h = x2 - x1, y2 - y1
        if w < 4 or h < 4:
            self.on_done(None, "녹화 영역이 너무 작습니다.")
            return

        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.join(os.path.expanduser("~"), "Desktop", f"recording_{ts}")

        # 사용 가능한 코덱 탐색
        out = None
        path = None
        for codec, ext in self._CODECS:
            path = base + ext
            fourcc = cv2.VideoWriter_fourcc(*codec)
            out    = cv2.VideoWriter(path, fourcc, self.FPS, (w, h))
            if out.isOpened():
                break
            out.release()
            out = None

        if out is None:
            self.on_done(None, "사용 가능한 비디오 코덱을 찾을 수 없습니다.")
            return

        frame_time = 1.0 / self.FPS
        with mss.mss() as sct:
            mon = {"left": x1, "top": y1, "width": w, "height": h}
            while self.running:
                t0    = time.time()
                shot  = sct.grab(mon)
                # BGRA → BGR (alpha 채널 제거)
                frame = np.array(shot, dtype=np.uint8)[:, :, :3]
                out.write(frame)
                elapsed = time.time() - t0
                rem = frame_time - elapsed
                if rem > 0:
                    time.sleep(rem)

        out.release()
        self.out_path = path
        self.on_done(path, None)


class RecordingHighlight:
    """녹화 중 선택 영역을 테두리+주변 어둡게 하여 강조 표시"""

    BORDER = 3      # 테두리 두께(px)
    BORDER_COLOR = "#ff3333"

    def __init__(self, parent, region: tuple):
        x1, y1, x2, y2 = region
        b = self.BORDER
        c = self.BORDER_COLOR

        self._wins = []

        # ── 4개 반투명 어두운 블록 (선택 영역 외부를 어둡게)
        sw = parent.winfo_screenwidth()
        sh = parent.winfo_screenheight()

        dark_regions = [
            (0,   0,   sw,      y1),        # 위
            (0,   y2,  sw,      sh - y2),   # 아래
            (0,   y1,  x1,      y2 - y1),   # 좌
            (x2,  y1,  sw - x2, y2 - y1),   # 우
        ]
        for (dx, dy, dw, dh) in dark_regions:
            if dw > 0 and dh > 0:
                w = tk.Toplevel(parent)
                w.overrideredirect(True)
                w.attributes("-topmost", True)
                w.attributes("-alpha", 0.45)
                w.configure(bg="#000000")
                w.geometry(f"{dw}x{dh}+{dx}+{dy}")
                self._wins.append(w)

        # ── 선택 영역 테두리 (빨간 테두리선)
        border_rects = [
            (x1 - b, y1 - b, x2 - x1 + b * 2, b),   # 위
            (x1 - b, y2,     x2 - x1 + b * 2, b),   # 아래
            (x1 - b, y1,     b,      y2 - y1),       # 좌
            (x2,     y1,     b,      y2 - y1),       # 우
        ]
        for (bx, by, bw, bh) in border_rects:
            if bw > 0 and bh > 0:
                w = tk.Toplevel(parent)
                w.overrideredirect(True)
                w.attributes("-topmost", True)
                w.configure(bg=c)
                w.geometry(f"{bw}x{bh}+{bx}+{by}")
                self._wins.append(w)

    def destroy(self):
        for w in self._wins:
            try: w.destroy()
            except Exception: pass
        self._wins.clear()


class RecordingOverlay:
    """녹화 중 플로팅 인디케이터 (항상 위)"""

    def __init__(self, parent, on_stop):
        self._start = time.time()
        self._on_stop = on_stop

        self.win = tk.Toplevel(parent)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.configure(bg="#cc2222")
        self.win.geometry("180x38+20+20")

        tk.Label(self.win, text="⏺", bg="#cc2222", fg="white",
                 font=("맑은 고딕", 11)).pack(side="left", padx=(8, 2), pady=6)
        self._lbl = tk.Label(self.win, text="녹화중 00:00",
                              bg="#cc2222", fg="white",
                              font=("맑은 고딕", 10, "bold"))
        self._lbl.pack(side="left", padx=4)
        tk.Button(self.win, text="■ 중지", command=self._stop,
                  bg="#991111", fg="white", relief="flat",
                  font=("맑은 고딕", 9), padx=6, pady=2,
                  cursor="hand2").pack(side="left", padx=8)

        self._tick()

    def _tick(self):
        if not self.win.winfo_exists():
            return
        s = int(time.time() - self._start)
        self._lbl.config(text=f"녹화중 {s // 60:02d}:{s % 60:02d}")
        self.win.after(1000, self._tick)

    def _stop(self):
        self._on_stop()

    def destroy(self):
        try:
            self.win.destroy()
        except Exception:
            pass


# ─── 메인 앱 ─────────────────────────────────────────────────────────────────

MAX_HISTORY = 100

# 색상 팔레트
_C = {
    "bg":       "#1e1e2e",   # 앱 배경
    "toolbar":  "#181825",   # 툴바/헤더
    "card":     "#2a2a3e",   # 썸네일 카드
    "card_sel": "#3d3d6b",   # 선택된 카드
    "btn":      "#313152",   # 버튼
    "btn_act":  "#5566cc",   # 버튼 활성화
    "txt":      "#cdd6f4",   # 주 텍스트
    "txt_dim":  "#6c7086",   # 보조 텍스트
    "accent":   "#89b4fa",   # 강조색
    "status":   "#11111b",   # 상태바
}

THUMB_W, THUMB_H = 214, 126   # 썸네일 표시 크기 (~17:10)


class CaptureApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("캡처 툴")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=_C["bg"])

        # ── 전역 폰트: 맑은 고딕 ────────────────────────────────────────────
        self.root.option_add("*Font", ("맑은 고딕", 9))
        self.root.option_add("*Label.Font",   ("맑은 고딕", 9))
        self.root.option_add("*Button.Font",  ("맑은 고딕", 9))
        self.root.option_add("*Entry.Font",   ("맑은 고딕", 9))
        self.root.option_add("*Text.Font",    ("맑은 고딕", 9))
        self.root.option_add("*Menu.Font",    ("맑은 고딕", 9))
        _style = ttk.Style(self.root)
        _style.configure(".",          font=("맑은 고딕", 9))
        _style.configure("TCombobox",  font=("맑은 고딕", 8))
        _style.configure("TLabel",     font=("맑은 고딕", 9))
        _style.configure("TButton",    font=("맑은 고딕", 9))

        # ── 카메라 아이콘 ────────────────────────────────────────────────────
        self._icon_photo = self._make_camera_icon()
        self.root.wm_iconphoto(True, self._icon_photo)

        self._cache_dir = self._get_cache_dir()
        self.history: list[dict] = self._load_saved_history()
        self._editor             = None
        self._mode               = "capture"   # "capture" | "record"
        self._delay              = 0           # 지연 초 (0/3/5/10)
        self._recorder:   ScreenRecorder    | None = None
        self._rec_overlay: RecordingOverlay | None = None
        self._rec_highlight: RecordingHighlight | None = None

        self._build_ui()
        # X 버튼 → 최소화 (작업표시줄에 남김). 우클릭 메뉴로 완전 종료.
        self.root.protocol("WM_DELETE_WINDOW", self.root.iconify)
        self._build_tray_menu()
        self.root.mainloop()

    # ── UI 구성 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        def _hovbtn(parent, text, cmd, accent=False, **kw):
            bg0 = _C["btn_act"] if accent else _C["btn"]
            fg0 = "white"
            b = tk.Button(parent, text=text, command=cmd,
                          bg=bg0, fg=fg0, relief="flat",
                          font=("맑은 고딕", 8), cursor="hand2",
                          padx=8, pady=3, **kw)
            b.bind("<Enter>", lambda e: b.config(bg=_C["btn_act"], fg="white"))
            b.bind("<Leave>", lambda e: b.config(bg=bg0, fg=fg0))
            return b

        # ── 행 1: 캡처 ↔ 녹화 토글 + 편집기 열기 ───────────────────────────
        row1 = tk.Frame(self.root, bg=_C["toolbar"], pady=5)
        row1.pack(fill="x")

        self._btn_cap = _hovbtn(row1, "📷 캡처",
                                lambda: self._switch_mode("capture"), accent=True)
        self._btn_cap.pack(side="left", padx=(8, 2))

        self._btn_rec = _hovbtn(row1, "⏺ 녹화",
                                lambda: self._switch_mode("record"))
        self._btn_rec.pack(side="left", padx=2)

        # 편집기 열기 (항상 표시 — 최근 캡처를 편집기로)
        _hovbtn(row1, "캡처 히스토리 »", self._open_last_in_editor
                ).pack(side="right", padx=(0, 8))

        # ── 행 2: 유틸 (캡처 모드 전용 — 지연 + 텍스트 추출) ───────────────
        self._util_frame = tk.Frame(self.root, bg=_C["toolbar"], pady=3)
        self._util_frame.pack(fill="x")

        # 지연 콤보
        self._delay_frame = tk.Frame(self._util_frame, bg=_C["toolbar"])
        self._delay_frame.pack(side="left", padx=(10, 0))
        tk.Label(self._delay_frame, text="지연:", bg=_C["toolbar"],
                 fg=_C["txt_dim"], font=("맑은 고딕", 8)).pack(side="left")
        self._delay_var = tk.StringVar(value="없음")
        ttk.Combobox(self._delay_frame, textvariable=self._delay_var,
                     values=["없음", "3초", "5초", "10초"],
                     width=4, state="readonly",
                     font=("맑은 고딕", 8)).pack(side="left", padx=2)

        # 텍스트 추출
        self._ocr_btn = _hovbtn(self._util_frame, "🔤 텍스트추출",
                                self._capture_text)
        self._ocr_btn.pack(side="right", padx=(0, 8))

        # ── 행 3: 기능 버튼 (캡처 모드 5개 / 녹화 모드 3개) ─────────────────
        self._btn_frame = tk.Frame(self.root, bg=_C["toolbar"], pady=6)
        self._btn_frame.pack(fill="x")

        self._render_mode_buttons()

        # 단축키
        for key, cmd in [("<F1>", self._capture_region),
                          ("<F2>", self._capture_unit),
                          ("<F3>", self._capture_fullscreen),
                          ("<F4>", self._capture_scroll),
                          ("<F5>", self._capture_fixed)]:
            self.root.bind(key, lambda e, c=cmd: c())

        self.root.geometry("320x148")

    def _render_mode_buttons(self):
        for w in self._btn_frame.winfo_children():
            w.destroy()

        if self._mode == "capture":
            modes = [
                ("□\n사각형",   self._capture_region,    "F1"),
                ("⊡\n단위영역", self._capture_unit,      "F2"),
                ("■\n전체화면", self._capture_fullscreen, "F3"),
                ("↕\n스크롤",   self._capture_scroll,    "F4"),
                ("⊞\n지정크기", self._capture_fixed,     "F5"),
            ]
            n = 5
        else:
            modes = [
                ("□\n사각형",   self._record_region,    ""),
                ("⊡\n단위영역", self._record_unit,      ""),
                ("■\n전체화면", self._record_fullscreen, ""),
            ]
            n = 3

        for col, (lbl, cmd, key) in enumerate(modes):
            f = tk.Frame(self._btn_frame, bg=_C["toolbar"])
            f.grid(row=0, column=col, padx=3, pady=0)
            b = tk.Button(f, text=lbl, command=cmd,
                          bg=_C["btn"], fg=_C["txt"],
                          activebackground=_C["btn_act"], activeforeground="white",
                          relief="flat", font=("맑은 고딕", 8), width=6,
                          pady=4, cursor="hand2")
            b.pack()
            if key:
                tk.Label(f, text=key, bg=_C["toolbar"], fg=_C["txt_dim"],
                         font=("맑은 고딕", 7)).pack()
            b.bind("<Enter>", lambda e, w=b: w.config(bg=_C["btn_act"]))
            b.bind("<Leave>", lambda e, w=b: w.config(bg=_C["btn"]))
        for c in range(n):
            self._btn_frame.columnconfigure(c, weight=1)

    def _build_tray_menu(self):
        """우클릭 컨텍스트 메뉴 — 창 복원 / 완전 종료."""
        menu = tk.Menu(self.root, tearoff=0,
                       bg=_C["card"], fg=_C["txt"],
                       activebackground=_C["btn_act"], activeforeground="white",
                       font=("맑은 고딕", 9), relief="flat", bd=0)
        menu.add_command(label="🪟  창 열기", command=self.root.deiconify)
        menu.add_separator()
        menu.add_command(label="❌  종료",    command=self.root.destroy)

        def _show(e):
            try:
                menu.tk_popup(e.x_root, e.y_root)
            finally:
                menu.grab_release()

        self.root.bind("<Button-3>", _show)

        # ── 작업표시줄 우클릭 시스템 메뉴에 "종료" 추가 ─────────────────
        # win32gui.SetWindowLong 은 Python 콜백을 직접 받으며 64비트 안전.
        # 단, 윈도우가 완전히 그려진 뒤에야 HWND 가 유효하므로 after 로 예약.
        self.root.after(300, self._attach_sysmenu_quit)

    # ── 카메라 아이콘 생성 ────────────────────────────────────────────────────

    @staticmethod
    def _make_camera_icon() -> ImageTk.PhotoImage:
        """64×64 카메라 아이콘을 PIL로 생성해 PhotoImage로 반환."""
        SZ = 64
        img = Image.new("RGBA", (SZ, SZ), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)

        BODY  = (70,  70,  75,  255)   # 짙은 회색 본체
        LENS1 = (30,  30,  35,  255)   # 렌즈 바깥
        LENS2 = (55,  95, 175,  255)   # 렌즈 안쪽 (파란빛)
        LENS3 = (120, 160, 230, 255)   # 렌즈 하이라이트
        FLASH = (240, 220,  80,  255)  # 플래시
        SHINE = (255, 255, 255, 140)   # 반사광

        # 뷰파인더 돌출부
        d.rounded_rectangle([22, 12, 40, 21], radius=3, fill=BODY)
        # 본체
        d.rounded_rectangle([6, 19, 58, 50], radius=7, fill=BODY)
        # 플래시
        d.rounded_rectangle([44, 23, 53, 29], radius=2, fill=FLASH)
        # 렌즈 외곽
        cx, cy = 30, 34
        d.ellipse([cx-13, cy-13, cx+13, cy+13], fill=LENS1)
        # 렌즈 중간
        d.ellipse([cx-9,  cy-9,  cx+9,  cy+9],  fill=LENS2)
        # 렌즈 안쪽
        d.ellipse([cx-5,  cy-5,  cx+5,  cy+5],  fill=LENS3)
        # 렌즈 반사광
        d.ellipse([cx-8,  cy-8,  cx-4,  cy-5],  fill=SHINE)

        return ImageTk.PhotoImage(img)

    # ── 작업표시줄 우클릭 → 종료 (시스템 메뉴) ──────────────────────────────

    def _attach_sysmenu_quit(self):
        """시스템 메뉴에 '종료' 추가 + win32gui 서브클래싱.

        win32gui.SetWindowLong(GWL_WNDPROC) 은 Python callable 을 직접 받아
        pywin32 내부에서 thunk 처리 → ctypes 포인터 잘림 문제 없음.
        """
        try:
            hwnd = int(self.root.wm_frame(), 16)
            if not hwnd:
                return

            # 시스템 메뉴에 구분선 + 종료
            hmenu = win32gui.GetSystemMenu(hwnd, False)
            win32gui.AppendMenu(hmenu, win32con.MF_SEPARATOR, 0,             "")
            win32gui.AppendMenu(hmenu, win32con.MF_STRING,    _ID_TRAY_QUIT, "종료")

            # WndProc 서브클래싱 (pywin32 방식)
            def _proc(h, msg, wp, lp):
                if msg == win32con.WM_SYSCOMMAND and wp == _ID_TRAY_QUIT:
                    self.root.after(0, self.root.destroy)
                    return 0
                return win32gui.CallWindowProc(self._sysmenu_old_proc, h, msg, wp, lp)

            self._sysmenu_proc     = _proc          # GC 방지용 참조 유지
            self._sysmenu_old_proc = win32gui.SetWindowLong(
                hwnd, win32con.GWL_WNDPROC, _proc
            )
        except Exception as exc:
            print(f"[sysmenu] {exc}")

    def _switch_mode(self, mode: str):
        self._mode = mode
        self._btn_cap.config(bg=_C["btn_act"] if mode == "capture" else _C["btn"],
                              fg="white" if mode == "capture" else _C["txt"])
        self._btn_rec.config(bg=_C["btn_act"] if mode == "record" else _C["btn"],
                              fg="white" if mode == "record" else _C["txt"])
        # 유틸 행(지연·OCR)은 캡처 모드에서만 표시
        if mode == "capture":
            self._util_frame.pack(fill="x", before=self._btn_frame)
            self.root.geometry("320x148")
        else:
            self._util_frame.pack_forget()
            self.root.geometry("320x120")
        self._render_mode_buttons()

    # ── 패키지 자동 설치 ─────────────────────────────────────────────────────

    def _install_and_run(self, pkg: str, action, label: str = ""):
        """pip으로 패키지 백그라운드 설치 후 action() 실행"""
        dlg = tk.Toplevel(self.root)
        dlg.title("패키지 설치")
        dlg.resizable(False, False)
        dlg.attributes("-topmost", True)
        dlg.configure(bg=_C["bg"])
        rx, ry = self.root.winfo_x(), self.root.winfo_y()
        rh = self.root.winfo_height()
        dlg.geometry(f"280x90+{rx}+{ry + rh + 6}")
        dlg.grab_set()

        tk.Label(dlg, text=f"설치 중: {pkg}",
                 bg=_C["bg"], fg=_C["accent"],
                 font=("맑은 고딕", 10, "bold")).pack(pady=(16, 4))
        status_lbl = tk.Label(dlg, text="잠시 기다려 주세요...",
                               bg=_C["bg"], fg=_C["txt_dim"],
                               font=("맑은 고딕", 8))
        status_lbl.pack()

        def _run():
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", pkg, "-q"],
                    check=True, capture_output=True)
                self.root.after(0, lambda: (dlg.destroy(), action()))
            except Exception as e:
                self.root.after(0, lambda: (
                    dlg.destroy(),
                    messagebox.showerror("설치 실패",
                                         f"'{pkg}' 설치에 실패했습니다.\n수동 설치:\npip install {pkg}",
                                         parent=self.root)))

        threading.Thread(target=_run, daemon=True).start()

    # ── 지연 캡처 헬퍼 ──────────────────────────────────────────────────────

    def _get_delay(self) -> int:
        return {"없음": 0, "3초": 3, "5초": 5, "10초": 10}.get(
            self._delay_var.get(), 0)

    def _delayed(self, fn):
        """지연 없으면 즉시, 있으면 1초 단위 카운트다운 후 fn() 실행"""
        d = self._get_delay()
        if d == 0:
            fn()
            return
        # 카운트다운 오버레이
        overlay = tk.Toplevel(self.root)
        overlay.overrideredirect(True)
        overlay.attributes("-topmost", True)
        overlay.configure(bg="#000000")
        overlay.attributes("-alpha", 0.75)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        overlay.geometry(f"120x120+{sw // 2 - 60}+{sh // 2 - 60}")
        lbl = tk.Label(overlay, text=str(d), bg="#000000", fg="white",
                        font=("맑은 고딕", 64, "bold"))
        lbl.pack(expand=True)

        def _tick(remaining):
            if remaining <= 0:
                # 항상 fn() 호출 — overlay 상태와 무관하게
                try: overlay.destroy()
                except Exception: pass
                fn()
                return
            try:
                lbl.config(text=str(remaining))
            except Exception:
                pass
            # root.after() 사용: overlay가 닫혀도 타이머 계속 동작
            self.root.after(1000, lambda: _tick(remaining - 1))

        _tick(d)

    # ── 녹화 ─────────────────────────────────────────────────────────────────

    def _record_region(self):
        self._hide_root()
        self.root.after(150, lambda: RegionSelector(self._on_record_region))

    def _on_record_region(self, region):
        self._show_root()
        if region:
            self._start_recording(region)

    def _record_unit(self):
        self._hide_root()
        self.root.after(150, lambda: RegionSelector(self._on_record_region))

    def _record_fullscreen(self):
        with mss.mss() as sct:
            m = sct.monitors[0]
            region = (m["left"], m["top"],
                      m["left"] + m["width"], m["top"] + m["height"])
        self._start_recording(region)

    def _start_recording(self, region):
        if self._recorder and self._recorder.running:
            return
        try:
            import cv2  # noqa: F401
        except ImportError:
            self._install_and_run(
                "opencv-python",
                lambda: self._start_recording(region),
                "opencv-python")
            return
        self._recorder    = ScreenRecorder(region, self._on_record_done)
        self._rec_highlight = RecordingHighlight(self.root, region)
        self._rec_overlay = RecordingOverlay(self.root, self._stop_recording)
        self._recorder.start()

    def _stop_recording(self):
        if self._rec_highlight:
            self._rec_highlight.destroy()
            self._rec_highlight = None
        if self._recorder:
            self._recorder.stop()

    def _on_record_done(self, path: str | None, err: str | None):
        def _ui():
            if self._rec_highlight:
                self._rec_highlight.destroy()
                self._rec_highlight = None
            if self._rec_overlay:
                self._rec_overlay.destroy()
                self._rec_overlay = None
            if err:
                messagebox.showerror("녹화 오류", err, parent=self.root)
            elif path:
                messagebox.showinfo("녹화 완료",
                                    f"저장 완료:\n{path}", parent=self.root)
                os.startfile(os.path.dirname(path))
        self.root.after(0, _ui)

    # ── OCR 텍스트 추출 ───────────────────────────────────────────────────────

    def _capture_text(self):
        self._hide_root()
        self.root.after(150, lambda: RegionSelector(self._on_text_region))

    def _on_text_region(self, region):
        self._show_root()
        if not region:
            return
        x1, y1, x2, y2 = region
        img = _grab_region(x1, y1, x2, y2)
        if img is None:
            return

        self._run_ocr(img)

    def _run_ocr(self, img):
        def _bg():
            text, debug = _do_ocr(img)
            self.root.after(0, lambda: self._ocr_done(text, debug))
        threading.Thread(target=_bg, daemon=True).start()

    def _ocr_done(self, text: str | None, debug: str = ""):
        if text is None:
            msg = "텍스트를 인식하지 못했습니다."
            if debug:
                msg += f"\n\n[디버그]\n{debug[:400]}"
            msg += "\n\nPowerShell 실행 정책 문제일 경우:\n  Set-ExecutionPolicy RemoteSigned"
            messagebox.showerror("텍스트 추출 실패", msg, parent=self.root)
            return
        if not text.strip():
            messagebox.showinfo("텍스트 추출", "인식된 텍스트가 없습니다.",
                                parent=self.root)
            return
        # 클립보드에 복사
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        # 결과 미리보기 다이얼로그
        dlg = tk.Toplevel(self.root)
        dlg.title("텍스트 추출 결과")
        dlg.resizable(True, True)
        dlg.attributes("-topmost", True)
        dlg.configure(bg=_C["bg"])
        rx, ry = self.root.winfo_x(), self.root.winfo_y()
        dlg.geometry(f"360x220+{rx}+{ry + self.root.winfo_height() + 6}")

        tk.Label(dlg, text="클립보드에 복사됨  ✓",
                 bg=_C["bg"], fg=_C["accent"],
                 font=("맑은 고딕", 9, "bold")).pack(pady=(10, 4))

        txt = tk.Text(dlg, bg=_C["card"], fg=_C["txt"],
                      insertbackground=_C["txt"],
                      relief="flat", font=("맑은 고딕", 10),
                      wrap="word", padx=8, pady=6)
        txt.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        txt.insert("1.0", text)
        txt.config(state="disabled")

        tk.Button(dlg, text="닫기", command=dlg.destroy,
                  bg=_C["btn_act"], fg="white", relief="flat",
                  padx=12, pady=4, cursor="hand2").pack(pady=(0, 10))

    # ── 히스토리 관리 ────────────────────────────────────────────────────────

    def _add(self, img: Image.Image):
        ts    = datetime.now().strftime("%H:%M:%S")
        label = f"{ts}   {img.width}×{img.height}"
        fname = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3] + ".png"
        item  = {"img": img, "label": label, "edited": False, "_file": fname}
        self.history.insert(0, item)

        # 디스크에 이미지 저장 (백그라운드)
        threading.Thread(target=self._save_item, args=(img, fname), daemon=True).start()

        # MAX_HISTORY 초과 시 가장 오래된 항목·파일 삭제
        if len(self.history) > MAX_HISTORY:
            old = self.history.pop()
            self._delete_file(old.get("_file"))

        self._persist_index()

    def _update_history(self, item: dict, img: Image.Image):
        """편집 후 히스토리 반영: 이미지·레이블 갱신"""
        item["img"]    = img
        item["edited"] = True
        ts = item["label"].split()[0]
        item["label"]  = f"{ts}   {img.width}×{img.height} ✏"
        # 편집된 이미지 덮어쓰기 저장
        if item.get("_file"):
            threading.Thread(target=self._save_item,
                             args=(img, item["_file"]), daemon=True).start()
        self._persist_index()

    # ── 히스토리 영속화 ──────────────────────────────────────────────────────────

    def _get_cache_dir(self) -> str:
        import hashlib, uuid
        # MAC 주소 기준으로 디바이스별 히스토리 격리
        mac   = uuid.getnode()                          # 48-bit int
        token = hashlib.sha256(str(mac).encode()).hexdigest()[:16]
        d = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")),
                         "CaptureApp", "history", token)
        os.makedirs(d, exist_ok=True)
        return d

    def _load_saved_history(self) -> list:
        idx = os.path.join(self._cache_dir, "index.json")
        if not os.path.exists(idx):
            return []
        try:
            with open(idx, "r", encoding="utf-8") as f:
                entries = json.load(f)
            result = []
            for e in entries:
                fp = os.path.join(self._cache_dir, e["filename"])
                if not os.path.exists(fp):
                    continue
                try:
                    img = Image.open(fp).copy()
                    result.append({
                        "img":     img,
                        "label":   e.get("label", ""),
                        "edited":  e.get("edited", False),
                        "_file":   e["filename"],
                    })
                except Exception:
                    pass
            return result
        except Exception:
            return []

    def _save_item(self, img: Image.Image, fname: str):
        try:
            img.save(os.path.join(self._cache_dir, fname))
        except Exception:
            pass

    def _delete_file(self, fname: str | None):
        if not fname:
            return
        try:
            os.unlink(os.path.join(self._cache_dir, fname))
        except Exception:
            pass

    def _persist_index(self):
        """현재 history 순서대로 index.json 갱신"""
        idx = os.path.join(self._cache_dir, "index.json")
        entries = [
            {"filename": it["_file"], "label": it["label"], "edited": it["edited"]}
            for it in self.history if it.get("_file")
        ]
        try:
            with open(idx, "w", encoding="utf-8") as f:
                json.dump(entries, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── 편집기 직접 열기 ─────────────────────────────────────────────────────

    def _open_last_in_editor(self):
        """가장 최근 캡처를 편집기로 열기. 이력이 없으면 안내."""
        if not self.history:
            messagebox.showinfo("편집기", "캡처 기록이 없습니다.\n먼저 화면을 캡처해주세요.",
                                parent=self.root)
            return
        item = self.history[0]
        self._open_editor(item["img"].copy(), item)

    # ── 편집기 열기 ──────────────────────────────────────────────────────────

    def _open_editor(self, img: Image.Image, item: dict):
        if self._editor is not None:
            try: self._editor.win.destroy()
            except Exception: pass
        self._editor = EditorWindow(
            self.root, img,
            on_update=lambda edited, it=item: self._update_history(it, edited),
            on_close=lambda: setattr(self, "_editor", None),
            history=self.history)

    # ── 루트 창 숨김/복원 헬퍼 ─────────────────────────────────────────────────
    # withdraw/iconify 대신 화면 밖으로 이동: 자식 Toplevel이 정상 표시됨 (Windows 제약 우회)

    def _hide_root(self):
        self._root_x = self.root.winfo_x()
        self._root_y = self.root.winfo_y()
        self.root.geometry("+99999+99999")
        self.root.update_idletasks()

    def _show_root(self):
        x = getattr(self, "_root_x", 100)
        y = getattr(self, "_root_y", 100)
        self.root.geometry(f"+{x}+{y}")

    # ── 캡처 모드 ────────────────────────────────────────────────────────────

    def _capture_fullscreen(self):
        self._delayed(self._do_fullscreen_now)

    def _do_fullscreen_now(self):
        self.root.withdraw()
        self.root.after(200, self._do_fullscreen)

    def _do_fullscreen(self):
        with mss.mss() as sct:
            shot = sct.grab(sct.monitors[0])
            img  = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
        self.root.deiconify()
        self._show(img)

    def _capture_region(self):
        self._delayed(self._do_capture_region)

    def _do_capture_region(self):
        self._hide_root()
        self.root.after(150, lambda: RegionSelector(self._on_region))

    def _on_region(self, region):
        if not region:
            self._show_root(); return
        x1, y1, x2, y2 = region
        img = _grab_region(x1, y1, x2, y2)
        self._show_root()
        if img: self._show(img)

    def _capture_unit(self):
        self._delayed(self._do_capture_unit)

    def _do_capture_unit(self):
        self._hide_root()
        self.root.after(150, lambda: RegionSelector(self._on_unit))

    def _on_unit(self, region):
        if not region:
            self._show_root(); return
        x1, y1, x2, y2 = region
        img = _grab_region(x1, y1, x2, y2)
        self._show_root()
        if img: self._show(img)

    def _capture_scroll(self):
        self._delayed(self._do_capture_scroll)

    def _do_capture_scroll(self):
        self._hide_root()
        self.root.after(200, lambda: RegionSelector(self._on_scroll_region))

    def _on_scroll_region(self, region):
        if not region:
            self._show_root(); return
        ScrollCapture(self.root, region, self._on_scroll_done)

    def _on_scroll_done(self, img):
        self._show_root()
        if img: self._show(img)

    def _capture_fixed(self):
        self._do_capture_fixed()

    def _do_capture_fixed(self):
        """지정 사이즈 캡처"""
        dlg = tk.Toplevel(self.root)
        dlg.title("지정 사이즈 캡처")
        dlg.resizable(False, False)
        dlg.attributes("-topmost", True)
        dlg.configure(bg=_C["bg"])
        dlg.grab_set()

        # 캡처 툴 근처에 배치
        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        rh = self.root.winfo_height()
        dlg.geometry(f"+{rx}+{ry + rh + 6}")

        _lbl_style = dict(bg=_C["bg"], fg=_C["txt"], font=("맑은 고딕", 9))
        _entry_style = dict(bg=_C["card"], fg=_C["txt"], insertbackground=_C["txt"],
                            relief="flat", bd=4)

        tk.Label(dlg, text="캡처할 크기를 입력하세요",
                 bg=_C["bg"], fg=_C["accent"],
                 font=("맑은 고딕", 10, "bold")).grid(
                     row=0, column=0, columnspan=3, padx=16, pady=(14, 6))

        tk.Label(dlg, text="너비:", **_lbl_style).grid(row=1, column=0, padx=8, pady=4, sticky="e")
        w_var = tk.IntVar(value=800)
        tk.Entry(dlg, textvariable=w_var, width=7, **_entry_style).grid(row=1, column=1)
        tk.Label(dlg, text="px", **_lbl_style).grid(row=1, column=2, sticky="w")

        tk.Label(dlg, text="높이:", **_lbl_style).grid(row=2, column=0, padx=8, pady=4, sticky="e")
        h_var = tk.IntVar(value=600)
        tk.Entry(dlg, textvariable=h_var, width=7, **_entry_style).grid(row=2, column=1)
        tk.Label(dlg, text="px", **_lbl_style).grid(row=2, column=2, sticky="w")

        def _ok():
            try:
                cw, ch = int(w_var.get()), int(h_var.get())
                if cw < 10 or ch < 10:
                    raise ValueError
            except Exception:
                messagebox.showerror("오류", "유효한 크기를 입력하세요.", parent=dlg)
                return
            dlg.destroy()
            def _do():
                self.root.withdraw()
                self.root.after(200, lambda: FixedSizeSelector(cw, ch, self._on_fixed))
            self._delayed(_do)

        tk.Button(dlg, text="확인", command=_ok,
                  bg=_C["btn_act"], fg="white", activebackground=_C["btn_act"],
                  relief="flat", width=8, pady=4,
                  cursor="hand2").grid(row=3, column=0, columnspan=3, pady=12)
        dlg.bind("<Return>", lambda e: _ok())

    def _on_fixed(self, region):
        if not region:
            self.root.deiconify(); return
        x1, y1, x2, y2 = region
        img = _grab_region(x1, y1, x2, y2)
        self.root.deiconify()
        if img: self._show(img)

    def _show(self, img: Image.Image):
        self._add(img)
        item = self.history[0]
        # 열린 편집기가 있으면 이력 패널만 갱신하고 이미지 교체
        if self._editor is not None:
            try:
                self._editor._history = self.history
                self._editor._populate_history()
                self._editor.img  = img.copy()
                self._editor.orig = img.copy()
                self._editor.zi   = self._editor._fit_idx()
                self._editor.win.title(f"편집기  [{img.width}×{img.height}px]")
                self._editor._refresh()
                return
            except Exception:
                pass
        self._open_editor(img, item)



if __name__ == "__main__":
    CaptureApp()
