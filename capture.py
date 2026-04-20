"""
Ebook Screenshot Capture Tool
- 드래그로 캡처 영역 선택
- 수동 모드: Space / → 키로 캡처
- 자동 모드: 랜덤 간격 자동 캡처, N페이지마다 휴식
- P 키: 자동 모드 일시정지/재개
- Q / Esc: 종료
- PDF 변환: 폴더의 이미지를 하나의 PDF로 합치기
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ctypes
import ctypes.wintypes
from PIL import Image
import keyboard
import os
import random
import threading
import time
import glob as _glob


class RegionSelector:
    """마우스 드래그로 화면 영역을 선택하는 전체화면 오버레이"""

    def __init__(self, parent):
        self.region = None
        self.start_x = self.start_y = 0

        self.top = tk.Toplevel(parent)
        self.top.attributes("-fullscreen", True)
        self.top.attributes("-alpha", 0.3)
        self.top.attributes("-topmost", True)
        self.top.configure(bg="black")
        self.top.title("영역 선택")

        self.canvas = tk.Canvas(self.top, cursor="cross", bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            self.top,
            text="캡처할 영역을 드래그하세요  |  ESC: 취소",
            fg="white", bg="black", font=("맑은 고딕", 14),
        ).place(relx=0.5, rely=0.05, anchor="center")

        self.rect = None
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.top.bind("<Escape>", lambda e: self.top.destroy())

    def _on_press(self, event):
        self.start_x, self.start_y = event.x, event.y
        if self.rect:
            self.canvas.delete(self.rect)

    def _on_drag(self, event):
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, event.x, event.y,
            outline="red", width=2, fill="white", stipple="gray12",
        )

    def _on_release(self, event):
        x1, x2 = min(self.start_x, event.x), max(self.start_x, event.x)
        y1, y2 = min(self.start_y, event.y), max(self.start_y, event.y)
        if x2 - x1 > 10 and y2 - y1 > 10:
            self.region = {"left": x1, "top": y1, "width": x2 - x1, "height": y2 - y1}
        self.top.destroy()

    def select(self, parent):
        parent.wait_window(self.top)
        return self.region


class CaptureApp:
    def __init__(self):
        self.region = None
        self.click_pos = None
        self.target_hwnd = None
        self.save_dir = os.path.join(os.path.expanduser("~"), "Pictures", "ebook_capture")
        self.counter = 1
        self.capture_count = 0
        self.running = False
        self.pre_delay = 1.0

        self._auto_thread = None
        self._auto_paused = False
        self._auto_stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()

        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.root = tk.Tk()
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        self.root.title(f"Ebook 캡처 도구 {'[관리자]' if is_admin else '[권한 없음]'}")
        self.root.resizable(False, False)

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        capture_tab = ttk.Frame(notebook)
        pdf_tab = ttk.Frame(notebook)
        notebook.add(capture_tab, text="  캡처  ")
        notebook.add(pdf_tab, text="  PDF 변환  ")

        self._build_capture_tab(capture_tab)
        self._build_pdf_tab(pdf_tab)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_capture_tab(self, parent):
        pad = {"padx": 10, "pady": 4}

        # 저장 폴더
        f = ttk.LabelFrame(parent, text="저장 폴더")
        f.pack(fill=tk.X, padx=10, pady=5)
        self.dir_var = tk.StringVar(value=self.save_dir)
        ttk.Entry(f, textvariable=self.dir_var, width=42).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(f, text="찾아보기", command=self._browse_dir).pack(side=tk.LEFT, padx=5)

        # 파일 설정
        f = ttk.LabelFrame(parent, text="파일 설정")
        f.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(f, text="파일명 접두사:").grid(row=0, column=0, **pad)
        self.prefix_var = tk.StringVar(value="page")
        ttk.Entry(f, textvariable=self.prefix_var, width=12).grid(row=0, column=1, **pad)
        ttk.Label(f, text="시작 번호:").grid(row=0, column=2, **pad)
        self.start_var = tk.StringVar(value="1")
        ttk.Entry(f, textvariable=self.start_var, width=6).grid(row=0, column=3, **pad)
        ttk.Label(f, text="형식:").grid(row=0, column=4, **pad)
        self.fmt_var = tk.StringVar(value="png")
        ttk.Combobox(f, textvariable=self.fmt_var, values=["png", "jpg"],
                     width=6, state="readonly").grid(row=0, column=5, **pad)

        # 자동 모드 설정
        f = ttk.LabelFrame(parent, text="자동 캡처 설정")
        f.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(f, text="간격:").grid(row=0, column=0, **pad)
        self.auto_min_var = tk.StringVar(value="1.0")
        self.auto_max_var = tk.StringVar(value="3.0")
        ttk.Entry(f, textvariable=self.auto_min_var, width=5).grid(row=0, column=1, **pad)
        ttk.Label(f, text="~").grid(row=0, column=2)
        ttk.Entry(f, textvariable=self.auto_max_var, width=5).grid(row=0, column=3, **pad)
        ttk.Label(f, text="초  |").grid(row=0, column=4, **pad)
        ttk.Label(f, text="휴식:").grid(row=0, column=5, **pad)
        self.rest_every_var = tk.StringVar(value="30")
        ttk.Entry(f, textvariable=self.rest_every_var, width=4).grid(row=0, column=6, **pad)
        ttk.Label(f, text="페이지마다").grid(row=0, column=7)
        self.rest_sec_var = tk.StringVar(value="10")
        ttk.Entry(f, textvariable=self.rest_sec_var, width=4).grid(row=0, column=8, **pad)
        ttk.Label(f, text="초  |").grid(row=0, column=9)
        ttk.Label(f, text="캡처 전 대기:").grid(row=0, column=10, **pad)
        self.pre_delay_var = tk.StringVar(value="1.0")
        ttk.Entry(f, textvariable=self.pre_delay_var, width=5).grid(row=0, column=11, **pad)
        ttk.Label(f, text="초").grid(row=0, column=12)

        # [1단계] 대상 창 선택 (row 1)
        self.window_label = tk.StringVar(value="대상 창: 미선택")
        ttk.Label(f, textvariable=self.window_label, foreground="gray").grid(
            row=1, column=0, columnspan=5, **pad, sticky="w"
        )
        self.pick_window_btn = ttk.Button(
            f, text="① 대상 창 선택 (3초 후)", command=self._pick_target_window
        )
        self.pick_window_btn.grid(row=1, column=5, columnspan=3, **pad, sticky="w")

        # [2단계] 클릭 위치 지정 (row 2)
        self.click_label = tk.StringVar(value="클릭 위치: 미지정")
        ttk.Label(f, textvariable=self.click_label, foreground="gray").grid(
            row=2, column=0, columnspan=5, **pad, sticky="w"
        )
        self.pick_click_btn = ttk.Button(
            f, text="② 클릭 위치 지정 (3초 후)", command=self._pick_click_pos, state=tk.DISABLED
        )
        self.pick_click_btn.grid(row=2, column=5, columnspan=3, **pad, sticky="w")

        # 총 페이지 + 최소화 (row 3)
        ttk.Label(f, text="총 페이지:").grid(row=3, column=0, **pad)
        self.total_pages_var = tk.StringVar(value="0")
        ttk.Entry(f, textvariable=self.total_pages_var, width=5).grid(row=3, column=1, **pad)
        ttk.Label(f, text="(0=무한)").grid(row=3, column=2)
        self.minimize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            f, text="자동 시작 시 창 최소화", variable=self.minimize_var
        ).grid(row=3, column=3, columnspan=4, sticky="w", **pad)

        # 영역 선택
        f = ttk.LabelFrame(parent, text="캡처 영역")
        f.pack(fill=tk.X, padx=10, pady=5)
        self.region_label = tk.StringVar(value="영역이 선택되지 않았습니다")
        ttk.Label(f, textvariable=self.region_label).pack(side=tk.LEFT, padx=10, pady=5)
        ttk.Button(f, text="영역 선택", command=self._select_region).pack(side=tk.RIGHT, padx=10, pady=5)

        # 단축키 안내
        f = ttk.LabelFrame(parent, text="단축키")
        f.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(
            f,
            text="수동: Space / →  캡처    |    자동: P  일시정지/재개    |    Q / Esc  종료",
            font=("맑은 고딕", 9),
        ).pack(padx=10, pady=4)

        # 상태
        f = ttk.Frame(parent)
        f.pack(fill=tk.X, padx=10, pady=2)
        self.status_var = tk.StringVar(value="대기 중...")
        ttk.Label(f, textvariable=self.status_var, foreground="gray").pack(side=tk.LEFT)
        self.count_var = tk.StringVar(value="저장된 페이지: 0")
        ttk.Label(f, textvariable=self.count_var).pack(side=tk.RIGHT)

        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(parent, variable=self.progress_var, maximum=100).pack(
            fill=tk.X, padx=10, pady=2
        )

        # 버튼
        f = ttk.Frame(parent)
        f.pack(pady=8)
        self.start_manual_btn = ttk.Button(f, text="수동 시작", command=self._start_manual, width=12)
        self.start_manual_btn.pack(side=tk.LEFT, padx=4)
        self.start_auto_btn = ttk.Button(f, text="자동 시작", command=self._start_auto, width=12)
        self.start_auto_btn.pack(side=tk.LEFT, padx=4)
        self.pause_btn = ttk.Button(f, text="일시정지 (P)", command=self._toggle_pause,
                                    width=14, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=4)
        self.stop_btn = ttk.Button(f, text="중지", command=self._stop, width=8, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=4)
        ttk.Button(f, text="폴더 열기", command=self._open_folder, width=10).pack(side=tk.LEFT, padx=4)

    def _build_pdf_tab(self, parent):
        pad = {"padx": 10, "pady": 6}

        # 소스 폴더
        f = ttk.LabelFrame(parent, text="이미지 폴더")
        f.pack(fill=tk.X, padx=10, pady=8)
        self.pdf_src_var = tk.StringVar(value=self.save_dir)
        ttk.Entry(f, textvariable=self.pdf_src_var, width=42).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(f, text="찾아보기", command=self._pdf_browse_src).pack(side=tk.LEFT, padx=5)

        # 출력 PDF
        f = ttk.LabelFrame(parent, text="저장할 PDF 파일")
        f.pack(fill=tk.X, padx=10, pady=5)
        default_pdf = os.path.join(os.path.expanduser("~"), "Pictures", "ebook.pdf")
        self.pdf_out_var = tk.StringVar(value=default_pdf)
        ttk.Entry(f, textvariable=self.pdf_out_var, width=42).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(f, text="찾아보기", command=self._pdf_browse_out).pack(side=tk.LEFT, padx=5)

        # 옵션
        f = ttk.LabelFrame(parent, text="옵션")
        f.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(f, text="포함 확장자:").grid(row=0, column=0, **pad)
        self.pdf_ext_var = tk.StringVar(value="png, jpg")
        ttk.Entry(f, textvariable=self.pdf_ext_var, width=16).grid(row=0, column=1, **pad)
        ttk.Label(f, text="(쉼표 구분)").grid(row=0, column=2, padx=4)
        ttk.Label(f, text="정렬:").grid(row=0, column=3, **pad)
        self.pdf_sort_var = tk.StringVar(value="파일명")
        ttk.Combobox(f, textvariable=self.pdf_sort_var,
                     values=["파일명", "수정 시간"], width=10,
                     state="readonly").grid(row=0, column=4, **pad)

        # 상태
        f = ttk.Frame(parent)
        f.pack(fill=tk.X, padx=10, pady=4)
        self.pdf_status_var = tk.StringVar(value="대기 중...")
        ttk.Label(f, textvariable=self.pdf_status_var, foreground="gray").pack(side=tk.LEFT)
        self.pdf_count_var = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.pdf_count_var).pack(side=tk.RIGHT)

        self.pdf_progress_var = tk.DoubleVar(value=0)
        self.pdf_progressbar = ttk.Progressbar(
            parent, variable=self.pdf_progress_var, maximum=100
        )
        self.pdf_progressbar.pack(fill=tk.X, padx=10, pady=2)

        # 버튼
        f = ttk.Frame(parent)
        f.pack(pady=10)
        self.pdf_convert_btn = ttk.Button(
            f, text="PDF 변환 시작", command=self._start_pdf_convert, width=16
        )
        self.pdf_convert_btn.pack(side=tk.LEFT, padx=6)
        ttk.Button(f, text="PDF 폴더 열기", command=self._pdf_open_folder, width=14).pack(
            side=tk.LEFT, padx=6
        )

    # ── 공통 헬퍼 ──────────────────────────────────────────────────────────────

    def _browse_dir(self):
        d = filedialog.askdirectory(initialdir=self.dir_var.get())
        if d:
            self.dir_var.set(d)

    def _select_region(self):
        self.root.withdraw()
        time.sleep(0.3)
        region = RegionSelector(self.root).select(self.root)
        self.root.deiconify()
        if region:
            self.region = region
            self.region_label.set(
                f"X:{region['left']}  Y:{region['top']}  크기:{region['width']}×{region['height']}"
            )
        else:
            self.region_label.set("취소됨")

    def _apply_settings(self):
        self.save_dir = self.dir_var.get()
        self.prefix = self.prefix_var.get() or "page"
        self.fmt = self.fmt_var.get()
        try:
            self.counter = int(self.start_var.get())
        except ValueError:
            self.counter = 1
        try:
            self.pre_delay = float(self.pre_delay_var.get())
        except ValueError:
            self.pre_delay = 1.0
        os.makedirs(self.save_dir, exist_ok=True)

    def _set_buttons_running(self, auto_mode: bool):
        self.start_manual_btn.config(state=tk.DISABLED)
        self.start_auto_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.NORMAL if auto_mode else tk.DISABLED)

    def _set_buttons_idle(self):
        self.start_manual_btn.config(state=tk.NORMAL)
        self.start_auto_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.DISABLED)

    # ── 대상 창 / 클릭 위치 선택 ──────────────────────────────────────────────

    def _get_window_title(self, hwnd) -> str:
        buf = ctypes.create_unicode_buffer(256)
        ctypes.windll.user32.GetWindowTextW(hwnd, buf, 256)
        return buf.value or f"(hwnd={hwnd})"

    def _pick_target_window(self):
        self.pick_window_btn.config(state=tk.DISABLED)
        self.window_label.set("3... 대상 앱 위로 마우스를 이동하세요")
        self.root.after(1000, lambda: self.window_label.set("2..."))
        self.root.after(2000, lambda: self.window_label.set("1..."))
        self.root.after(3000, self._do_pick_target_window)

    def _do_pick_target_window(self):
        pt = ctypes.wintypes.POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        hwnd = ctypes.windll.user32.WindowFromPoint(pt)
        self.target_hwnd = ctypes.windll.user32.GetAncestor(hwnd, 2)
        title = self._get_window_title(self.target_hwnd)
        self.window_label.set(f"대상 창: [{title}]  (hwnd={self.target_hwnd})")
        self.pick_window_btn.config(state=tk.NORMAL)
        # 창이 선택됐으면 클릭 위치 지정 버튼 활성화
        self.pick_click_btn.config(state=tk.NORMAL)

    def _pick_click_pos(self):
        if not hasattr(self, "target_hwnd") or not self.target_hwnd:
            messagebox.showwarning("알림", "먼저 ① 대상 창을 선택하세요.")
            return
        self.pick_click_btn.config(state=tk.DISABLED)
        self.click_label.set("3... 다음 페이지 버튼 위로 마우스를 이동하세요")
        self.root.after(1000, lambda: self.click_label.set("2..."))
        self.root.after(2000, lambda: self.click_label.set("1..."))
        self.root.after(3000, self._do_pick_click_pos)

    def _do_pick_click_pos(self):
        pt = ctypes.wintypes.POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        self.click_pos = (pt.x, pt.y)
        self.click_label.set(f"클릭 위치: ({pt.x}, {pt.y})")
        self.pick_click_btn.config(state=tk.NORMAL)

    def _real_click(self, abs_x, abs_y):
        ctypes.windll.user32.SetCursorPos(int(abs_x), int(abs_y))
        time.sleep(0.1)
        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFT DOWN
        time.sleep(0.05)
        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFT UP

    def _do_capture(self, with_pre_delay: bool = False):
        if with_pre_delay and self.pre_delay > 0:
            self._interruptible_sleep(self.pre_delay, label="캡처 전 대기")
        if not self.running:
            return
        try:
            # 캡처 전 ebook 창을 앞으로
            if self.target_hwnd:
                ctypes.windll.user32.SetForegroundWindow(self.target_hwnd)
                time.sleep(0.1)
            filename = f"{self.prefix}_{self.counter:04d}.{self.fmt}"
            filepath = os.path.join(self.save_dir, filename)
            img = self._grab_region(self.region)
            img.save(filepath, "JPEG", quality=95) if self.fmt == "jpg" else img.save(filepath)
            self.counter += 1
            self.capture_count += 1
            self.root.after(0, self._update_status, filename)
        except Exception as e:
            self.root.after(0, messagebox.showerror, "캡처 오류", str(e))

    def _grab_region(self, region: dict) -> Image.Image:
        import dxcam
        if not hasattr(self, "_dxcam") or self._dxcam is None:
            self._dxcam = dxcam.create(output_color="RGB")
        left   = region["left"]
        top    = region["top"]
        right  = left + region["width"]
        bottom = top  + region["height"]
        frame = self._dxcam.grab(region=(left, top, right, bottom))
        if frame is None:
            raise RuntimeError("dxcam: 프레임을 가져오지 못했습니다 (화면 변화 없음)")
        return Image.fromarray(frame)

    def _grab_region_DXGI_DISABLED(self, region: dict) -> Image.Image:
        """DXGI Desktop Duplication으로 전체 화면 캡처 후 지정 영역 크롭."""
        import numpy as np

        # ── COM 헬퍼 ──────────────────────────────────────────────────────────
        PTR = ctypes.sizeof(ctypes.c_void_p)

        def _vtfn(com_ptr, idx, fn_type):
            # COM 객체 첫 필드 = vtable 포인터, vtable[idx] = 함수 포인터
            # from_address 로 두 번 역참조
            obj_addr = com_ptr.value if isinstance(com_ptr, ctypes.c_void_p) else int(com_ptr)
            vtable_addr = ctypes.c_void_p.from_address(obj_addr).value
            fn_addr     = ctypes.c_void_p.from_address(vtable_addr + idx * PTR).value
            return fn_type(fn_addr)

        def _qi(com_ptr, iid_bytes):
            IID = (ctypes.c_byte * 16)(*iid_bytes)
            result = ctypes.c_void_p()
            fn = _vtfn(com_ptr, 0, ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p,
                ctypes.POINTER(ctypes.c_byte * 16),
                ctypes.POINTER(ctypes.c_void_p),
            ))
            hr = fn(com_ptr, ctypes.byref(IID), ctypes.byref(result))
            if hr != 0:
                raise RuntimeError(f"QueryInterface 실패: 0x{hr & 0xFFFFFFFF:08X}")
            return result

        def _release(com_ptr):
            if com_ptr and com_ptr.value:
                _vtfn(com_ptr, 2, ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p))(com_ptr)

        # ── 상수 ──────────────────────────────────────────────────────────────
        D3D_DRIVER_TYPE_HARDWARE = 1
        D3D11_SDK_VERSION        = 7
        D3D11_USAGE_STAGING      = 3
        D3D11_CPU_ACCESS_READ    = 0x20000
        D3D11_MAP_READ           = 1
        DXGI_FORMAT_BGRA8        = 87

        # ── 구조체 ────────────────────────────────────────────────────────────
        class D3D11_TEXTURE2D_DESC(ctypes.Structure):
            _fields_ = [
                ("Width",           ctypes.c_uint),
                ("Height",          ctypes.c_uint),
                ("MipLevels",       ctypes.c_uint),
                ("ArraySize",       ctypes.c_uint),
                ("Format",          ctypes.c_uint),
                ("SampleDescCount", ctypes.c_uint),
                ("SampleDescQuality", ctypes.c_uint),
                ("Usage",           ctypes.c_uint),
                ("BindFlags",       ctypes.c_uint),
                ("CPUAccessFlags",  ctypes.c_uint),
                ("MiscFlags",       ctypes.c_uint),
            ]

        # BOOL은 4바이트(c_int), 구조체 전체 48바이트
        class DXGI_OUTDUPL_FRAME_INFO(ctypes.Structure):
            _fields_ = [
                ("LastPresentTime",             ctypes.c_int64),
                ("LastMouseUpdateTime",         ctypes.c_int64),
                ("AccumulatedFrames",           ctypes.c_uint),
                ("RectsCoalesced",              ctypes.c_int),
                ("ProtectedContentMaskedOut",   ctypes.c_int),
                ("PointerPositionX",            ctypes.c_int),
                ("PointerPositionY",            ctypes.c_int),
                ("PointerPositionVisible",      ctypes.c_int),
                ("TotalMetadataBufferSize",     ctypes.c_uint),
                ("PointerShapeBufferSize",      ctypes.c_uint),
            ]

        class D3D11_MAPPED_SUBRESOURCE(ctypes.Structure):
            _fields_ = [
                ("pData",      ctypes.c_void_p),
                ("RowPitch",   ctypes.c_uint),
                ("DepthPitch", ctypes.c_uint),
            ]

        # ── IID (GUID → 리틀엔디언 바이트) ───────────────────────────────────
        # IDXGIDevice  {54ec77fa-1377-44e6-8c32-88fd5f44c84c}
        IID_IDXGIDevice  = [0xFA,0x77,0xEC,0x54,0x77,0x13,0xE6,0x44,
                            0x8C,0x32,0x88,0xFD,0x5F,0x44,0xC8,0x4C]
        # IDXGIOutput1 {00cddea8-939b-4b83-a340-a685226666cc}
        IID_IDXGIOutput1 = [0xA8,0xDE,0xCD,0x00,0x9B,0x93,0x83,0x4B,
                            0xA3,0x40,0xA6,0x85,0x22,0x66,0x66,0xCC]
        # ID3D11Texture2D {6f15aaf2-d208-4e89-9ab4-489535d34f9c}
        IID_Tex2D        = [0xF2,0xAA,0x15,0x6F,0x08,0xD2,0x89,0x4E,
                            0x9A,0xB4,0x48,0x95,0x35,0xD3,0x4F,0x9C]

        # ── D3D11 디바이스 생성 ───────────────────────────────────────────────
        p_device  = ctypes.c_void_p()
        p_context = ctypes.c_void_p()
        feat_lvl  = ctypes.c_uint(0)
        hr = ctypes.windll.d3d11.D3D11CreateDevice(
            None, D3D_DRIVER_TYPE_HARDWARE, None, 0, None, 0,
            D3D11_SDK_VERSION,
            ctypes.byref(p_device), ctypes.byref(feat_lvl), ctypes.byref(p_context),
        )
        if hr != 0:
            raise RuntimeError(f"D3D11CreateDevice 실패: 0x{hr & 0xFFFFFFFF:08X}")

        try:
            # IDXGIDevice → IDXGIDevice::GetAdapter (idx 7) → IDXGIAdapter
            p_dxgi_dev = _qi(p_device, IID_IDXGIDevice)
            p_adapter  = ctypes.c_void_p()
            hr = _vtfn(p_dxgi_dev, 7, ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)
            ))(p_dxgi_dev, ctypes.byref(p_adapter))
            _release(p_dxgi_dev)
            if hr != 0:
                raise RuntimeError(f"GetAdapter 실패: 0x{hr & 0xFFFFFFFF:08X}")

            # IDXGIAdapter::EnumOutputs(0) (idx 7)
            p_output = ctypes.c_void_p()
            hr = _vtfn(p_adapter, 7, ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p, ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p)
            ))(p_adapter, 0, ctypes.byref(p_output))
            _release(p_adapter)
            if hr != 0:
                raise RuntimeError(f"EnumOutputs 실패: 0x{hr & 0xFFFFFFFF:08X}")

            # IDXGIOutput → IDXGIOutput1::DuplicateOutput (idx 22)
            p_output1 = _qi(p_output, IID_IDXGIOutput1)
            _release(p_output)
            p_dupl = ctypes.c_void_p()
            hr = _vtfn(p_output1, 22, ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)
            ))(p_output1, p_device, ctypes.byref(p_dupl))
            _release(p_output1)
            if hr != 0:
                raise RuntimeError(f"DuplicateOutput 실패: 0x{hr & 0xFFFFFFFF:08X}")

            # AcquireNextFrame (idx 8)
            frame_info = DXGI_OUTDUPL_FRAME_INFO()
            p_resource = ctypes.c_void_p()
            acq_fn = _vtfn(p_dupl, 8, ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p, ctypes.c_uint,
                ctypes.POINTER(DXGI_OUTDUPL_FRAME_INFO),
                ctypes.POINTER(ctypes.c_void_p),
            ))
            hr = acq_fn(p_dupl, 300, ctypes.byref(frame_info), ctypes.byref(p_resource))
            if hr == -2005270489:  # DXGI_ERROR_WAIT_TIMEOUT
                hr = acq_fn(p_dupl, 1000, ctypes.byref(frame_info), ctypes.byref(p_resource))
            if hr != 0:
                _release(p_dupl)
                raise RuntimeError(f"AcquireNextFrame 실패: 0x{hr & 0xFFFFFFFF:08X}")

            # ID3D11Texture2D::GetDesc (idx 10)
            p_tex = _qi(p_resource, IID_Tex2D)
            _release(p_resource)
            desc = D3D11_TEXTURE2D_DESC()
            _vtfn(p_tex, 10, ctypes.WINFUNCTYPE(
                None, ctypes.c_void_p, ctypes.POINTER(D3D11_TEXTURE2D_DESC)
            ))(p_tex, ctypes.byref(desc))
            w, h = desc.Width, desc.Height

            # 스테이징 텍스처 생성 — ID3D11Device::CreateTexture2D (idx 5)
            # 포맷은 소스와 동일하게 해야 CopyResource 동작
            stage_desc = D3D11_TEXTURE2D_DESC()
            stage_desc.Width             = w
            stage_desc.Height            = h
            stage_desc.MipLevels         = 1
            stage_desc.ArraySize         = 1
            stage_desc.Format            = desc.Format   # 소스 포맷 그대로
            stage_desc.SampleDescCount   = 1
            stage_desc.SampleDescQuality = 0
            stage_desc.Usage             = D3D11_USAGE_STAGING
            stage_desc.CPUAccessFlags    = D3D11_CPU_ACCESS_READ
            print(f"[DXGI] w={w} h={h} fmt={desc.Format}")
            p_stage = ctypes.c_void_p()
            hr = _vtfn(p_device, 5, ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p,
                ctypes.POINTER(D3D11_TEXTURE2D_DESC),
                ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p),
            ))(p_device, ctypes.byref(stage_desc), None, ctypes.byref(p_stage))
            if hr != 0:
                _release(p_tex)
                _vtfn(p_dupl, 14, ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p))(p_dupl)
                _release(p_dupl)
                raise RuntimeError(f"CreateTexture2D 실패: 0x{hr & 0xFFFFFFFF:08X}")

            # CopyResource — ID3D11DeviceContext (idx 36)
            _vtfn(p_context, 36, ctypes.WINFUNCTYPE(
                None, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p
            ))(p_context, p_stage, p_tex)
            _release(p_tex)

            # ReleaseFrame — IDXGIOutputDuplication (idx 14)
            _vtfn(p_dupl, 14, ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p))(p_dupl)
            _release(p_dupl)

            # Map — ID3D11DeviceContext (idx 14)
            mapped = D3D11_MAPPED_SUBRESOURCE()
            hr = _vtfn(p_context, 14, ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p,
                ctypes.c_uint, ctypes.c_uint, ctypes.c_uint,
                ctypes.POINTER(D3D11_MAPPED_SUBRESOURCE),
            ))(p_context, p_stage, 0, D3D11_MAP_READ, 0, ctypes.byref(mapped))
            if hr != 0:
                _release(p_stage)
                raise RuntimeError(f"Map 실패: 0x{hr & 0xFFFFFFFF:08X}")

            row = mapped.RowPitch
            buf = (ctypes.c_byte * (row * h))()
            ctypes.memmove(buf, mapped.pData, row * h)

            # Unmap — ID3D11DeviceContext (idx 15)
            _vtfn(p_context, 15, ctypes.WINFUNCTYPE(
                None, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint
            ))(p_context, p_stage, 0)
            _release(p_stage)

            # 픽셀 변환: BGRA(87) or RGBA → RGB
            arr = np.frombuffer(buf, dtype=np.uint8).reshape((h, row // 4, 4))
            arr = arr[:, :w, :]
            if desc.Format == 87:   # DXGI_FORMAT_B8G8R8A8_UNORM
                img = Image.fromarray(arr[:, :, [2, 1, 0]], "RGB")
            elif desc.Format in (28, 29):  # DXGI_FORMAT_R8G8B8A8_UNORM / SRGB
                img = Image.fromarray(arr[:, :, :3], "RGB")
            else:
                # 알 수 없는 포맷: BGRA로 시도
                print(f"[DXGI] 알 수 없는 포맷 {desc.Format}, BGRA로 시도")
                img = Image.fromarray(arr[:, :, [2, 1, 0]], "RGB")
            l = region["left"];  t = region["top"]
            return img.crop((l, t, l + region["width"], t + region["height"]))

        finally:
            if p_device:
                _release(p_device)
            if p_context:
                _release(p_context)

    def _update_status(self, filename):
        self.status_var.set(f"저장됨: {filename}")
        self.count_var.set(f"저장된 페이지: {self.capture_count}")

    # ── 수동 모드 ──────────────────────────────────────────────────────────────

    def _start_manual(self):
        if not self.region:
            messagebox.showwarning("알림", "먼저 캡처 영역을 선택하세요.")
            return
        self._apply_settings()
        self.running = True
        self.capture_count = 0
        self._set_buttons_running(auto_mode=False)
        self.status_var.set("수동 모드 — Space 또는 → 키를 누르세요")
        self.progress_var.set(0)

        keyboard.add_hotkey("space", self._manual_capture, suppress=False)
        keyboard.add_hotkey("right", self._manual_capture, suppress=False)
        keyboard.add_hotkey("q", self._stop, suppress=False)
        keyboard.add_hotkey("esc", self._stop, suppress=False)

    def _manual_capture(self):
        if self.running:
            threading.Thread(target=self._do_capture, args=(True,), daemon=True).start()

    # ── 자동 모드 ──────────────────────────────────────────────────────────────

    def _start_auto(self):
        if not self.region:
            messagebox.showwarning("알림", "먼저 캡처 영역을 선택하세요.")
            return
        self._apply_settings()
        try:
            auto_min = float(self.auto_min_var.get())
            auto_max = float(self.auto_max_var.get())
            rest_every = int(self.rest_every_var.get())
            rest_sec = int(self.rest_sec_var.get())
            total_pages = int(self.total_pages_var.get())
        except ValueError:
            messagebox.showwarning("알림", "자동 캡처 설정값을 확인하세요.")
            return

        self.running = True
        self.capture_count = 0
        self._auto_paused = False
        self._auto_stop_event.clear()
        self._pause_event.set()
        self._set_buttons_running(auto_mode=True)
        self.status_var.set("자동 모드 시작...")

        keyboard.add_hotkey("p", self._toggle_pause, suppress=False)
        keyboard.add_hotkey("q", self._stop, suppress=False)
        keyboard.add_hotkey("esc", self._stop, suppress=False)

        if self.minimize_var.get():
            self.root.after(500, self.root.iconify)

        self._auto_thread = threading.Thread(
            target=self._auto_loop,
            args=(auto_min, auto_max, rest_every, rest_sec, total_pages),
            daemon=True,
        )
        self._auto_thread.start()

    def _auto_loop(self, auto_min, auto_max, rest_every, rest_sec, total_pages=0):
        try:
            # 첫 페이지: 현재 화면 그대로 캡처
            self._pause_event.wait()
            if not self._auto_stop_event.is_set():
                self._do_capture()

            while not self._auto_stop_event.is_set():
                self._pause_event.wait()
                if self._auto_stop_event.is_set():
                    break

                # 페이지 넘김 클릭 전 ebook 창을 앞으로
                if self.click_pos:
                    if self.target_hwnd:
                        ctypes.windll.user32.SetForegroundWindow(self.target_hwnd)
                        time.sleep(0.15)
                    self._real_click(*self.click_pos)

                # 페이지 로딩 대기
                self._interruptible_sleep(self.pre_delay, label="페이지 로딩 대기")
                if self._auto_stop_event.is_set():
                    break

                # 캡처
                self._do_capture()

                # 목표 페이지 도달 시 종료
                if total_pages > 0 and self.capture_count >= total_pages:
                    self.root.after(0, self.status_var.set,
                                    f"완료! {self.capture_count}페이지 저장됨")
                    break

                # N페이지마다 휴식 (0 제외, 직후 체크)
                if self.capture_count > 0 and self.capture_count % rest_every == 0:
                    self.root.after(
                        0, self.status_var.set,
                        f"{self.capture_count}페이지 완료 — {rest_sec}초 휴식 중...",
                    )
                    self._interruptible_sleep(rest_sec, label="휴식")
                    if self._auto_stop_event.is_set():
                        break
                    continue

                # 랜덤 추가 대기
                delay = random.uniform(auto_min, auto_max)
                self._interruptible_sleep(delay, label="다음까지 대기")

        except Exception as e:
            self.root.after(0, messagebox.showerror, "자동 캡처 오류", str(e))
        finally:
            self.root.after(0, self._on_auto_finish)

    def _interruptible_sleep(self, seconds: float, label: str = ""):
        steps = max(int(seconds * 20), 1)
        for i in range(steps):
            if self._auto_stop_event.is_set():
                return
            self._pause_event.wait()
            pct = (i + 1) / steps * 100
            self.root.after(0, self.progress_var.set, pct)
            time.sleep(seconds / steps)
        self.root.after(0, self.progress_var.set, 0)

    def _toggle_pause(self):
        if not self.running or self._auto_thread is None:
            return
        if self._auto_paused:
            self._auto_paused = False
            self._pause_event.set()
            self.root.after(0, self.pause_btn.config, {"text": "일시정지 (P)"})
            self.root.after(0, self.status_var.set, "자동 모드 재개")
        else:
            self._auto_paused = True
            self._pause_event.clear()
            self.root.after(0, self.pause_btn.config, {"text": "재개 (P)"})
            self.root.after(0, self.status_var.set, "일시정지됨 — P 키로 재개")

    def _on_auto_finish(self):
        self.progress_var.set(0)
        if self.running:
            self._stop()

    # ── 종료 ──────────────────────────────────────────────────────────────────

    def _stop(self):
        self.running = False
        self._auto_stop_event.set()
        self._pause_event.set()
        keyboard.unhook_all()
        self.root.after(0, self._set_buttons_idle)
        self.root.after(0, self.progress_var.set, 0)
        self.root.after(0, self.status_var.set, f"중지됨. 총 {self.capture_count}페이지 저장")
        self.root.after(0, self.pause_btn.config, {"text": "일시정지 (P)"})
        self.root.after(0, self.root.deiconify)  # 최소화 상태면 복원

    def _open_folder(self):
        d = self.dir_var.get()
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    # ── PDF 변환 ──────────────────────────────────────────────────────────────

    def _pdf_browse_src(self):
        d = filedialog.askdirectory(initialdir=self.pdf_src_var.get())
        if d:
            self.pdf_src_var.set(d)

    def _pdf_browse_out(self):
        path = filedialog.asksaveasfilename(
            initialdir=os.path.dirname(self.pdf_out_var.get()),
            initialfile=os.path.basename(self.pdf_out_var.get()),
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf")],
        )
        if path:
            self.pdf_out_var.set(path)

    def _pdf_open_folder(self):
        d = os.path.dirname(self.pdf_out_var.get())
        if os.path.isdir(d):
            os.startfile(d)

    def _start_pdf_convert(self):
        src = self.pdf_src_var.get()
        out = self.pdf_out_var.get()
        if not os.path.isdir(src):
            messagebox.showwarning("알림", "이미지 폴더가 존재하지 않습니다.")
            return
        if not out.lower().endswith(".pdf"):
            messagebox.showwarning("알림", "출력 파일명이 .pdf로 끝나야 합니다.")
            return

        exts = [e.strip().lstrip(".").lower() for e in self.pdf_ext_var.get().split(",")]
        sort_by_mtime = self.pdf_sort_var.get() == "수정 시간"

        self.pdf_convert_btn.config(state=tk.DISABLED)
        self.pdf_status_var.set("이미지 수집 중...")
        self.pdf_progress_var.set(0)
        self.pdf_count_var.set("")

        threading.Thread(
            target=self._pdf_worker,
            args=(src, out, exts, sort_by_mtime),
            daemon=True,
        ).start()

    def _pdf_worker(self, src: str, out: str, exts: list, sort_by_mtime: bool):
        try:
            # 이미지 목록 수집
            files = []
            for ext in exts:
                files.extend(_glob.glob(os.path.join(src, f"*.{ext}")))
                files.extend(_glob.glob(os.path.join(src, f"*.{ext.upper()}")))
            files = list(dict.fromkeys(files))  # 중복 제거

            if not files:
                self.root.after(0, messagebox.showwarning, "알림", "이미지 파일을 찾을 수 없습니다.")
                self.root.after(0, self.pdf_convert_btn.config, {"state": tk.NORMAL})
                return

            if sort_by_mtime:
                files.sort(key=os.path.getmtime)
            else:
                files.sort(key=lambda p: os.path.basename(p).lower())

            total = len(files)
            self.root.after(0, self.pdf_count_var.set, f"총 {total}장")

            # 진행률 표시
            for i in range(total):
                self.root.after(0, self.pdf_status_var.set, f"확인 중... ({i + 1}/{total})")
                self.root.after(0, self.pdf_progress_var.set, (i + 1) / total * 50)

            # PDF 저장 (img2pdf: 재인코딩 없이 그대로 삽입)
            import img2pdf
            self.root.after(0, self.pdf_status_var.set, "PDF 저장 중...")
            os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
            with open(out, "wb") as f:
                f.write(img2pdf.convert(files))
            self.root.after(0, self.pdf_progress_var.set, 100)
            self.root.after(
                0, self.pdf_status_var.set,
                f"완료! {total}장 → {os.path.basename(out)}",
            )
            self.root.after(
                0, messagebox.showinfo, "완료",
                f"{total}장의 이미지를 PDF로 저장했습니다.\n\n{out}",
            )
        except Exception as e:
            self.root.after(0, messagebox.showerror, "오류", str(e))
            self.root.after(0, self.pdf_status_var.set, "오류 발생")
        finally:
            self.root.after(0, self.pdf_convert_btn.config, {"state": tk.NORMAL})
            self.root.after(0, self.pdf_progress_var.set, 0)

    # ── 앱 종료 ───────────────────────────────────────────────────────────────

    def _on_close(self):
        self._stop()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    import sys
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_DPI_AWARE
    except Exception:
        ctypes.windll.user32.SetProcessDPIAware()
    if not ctypes.windll.shell32.IsUserAnAdmin():
        if getattr(sys, "frozen", False):
            # PyInstaller exe: sys.executable 자체가 exe
            params = None
        else:
            # python capture.py: 스크립트 경로를 인자로 전달
            params = f'"{os.path.abspath(sys.argv[0])}"'
            if len(sys.argv) > 1:
                params += " " + " ".join(sys.argv[1:])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        sys.exit()
    app = CaptureApp()
    app.run()
