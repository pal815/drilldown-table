# -*- coding: utf-8 -*-
"""
드릴다운 표 생성기 (Drilldown Table) — 비개발자용 데스크톱 앱.

입력양식 엑셀(또는 클립보드)로 ㄱ자(드릴다운/코너헤더) 계층형 표를
Excel/Word/PowerPoint 로 생성한다. 언어모델·파이썬 지식 불필요.

엔진(drilldown_table.py, v14 양식)은 그대로 재사용 — 이 앱은 입력/옵션 UI 래퍼.
GUI는 CustomTkinter(모던, 컴팩트 단일 열: 입력가이드 최상단 + 방향/테마 호버 툴팁) → 없으면 기본 tkinter 폴백.
"""
import os
import sys
import traceback

import excel_io

try:
    import drilldown_table as G
except ImportError:
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir))
    import drilldown_table as G

FORMATS = [("Excel", "xlsx"), ("Word", "docx"), ("PowerPoint", "pptx")]
ORIENT_MAP = {"행 기준": "row", "열 기준": "column", "행+열": "both"}
THEME_MAP = {"컬러": "color", "그레이": "grey", "테두리만": "mono"}
ACCENT, ACCENT_HOVER = "#185FA5", "#0C447C"
KFONT = "맑은 고딕"


# ── 헤드리스 핵심 (GUI 없이 테스트 가능) ────────────────────────────────
def run_generate(in_path, orient, theme, formats, out_dir=None, com="auto"):
    """입력 엑셀 → 모델 → 선택 포맷별 파일 생성. 생성된 경로 리스트 반환.
    com='auto': openpyxl 먼저, DRM/보호로 실패하면 Excel COM 자동 폴백. True=강제 COM."""
    model = excel_io.read_model(in_path, orient, com=com)
    out_dir = out_dir or os.path.dirname(os.path.abspath(in_path))
    base = os.path.splitext(os.path.basename(in_path))[0]
    if base.lower() in ("입력양식", "template", "input"):
        base = "드릴다운표"
    made = []
    for ext in formats:
        out = os.path.join(out_dir, f"{base}_{orient}_{theme}.{ext}")
        G.generate(model, out, orient, theme)
        made.append(out)
    return made


def _default_outdir():
    try:  # 진짜 Desktop(OneDrive 리디렉션 포함)을 레지스트리에서
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as k:
            d = os.path.expandvars(winreg.QueryValueEx(k, "Desktop")[0])
            if os.path.isdir(d):
                return d
    except Exception:
        pass
    for sub in ("Desktop", "바탕 화면", "Documents", "문서"):
        d = os.path.join(os.path.expanduser("~"), sub)
        if os.path.isdir(d):
            return d
    return os.path.expanduser("~")


def _asset_dir():
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "assets")


def run_generate_text(text, orient, theme, formats, out_dir=None):
    """클립보드/붙여넣기 TSV → 모델 → 파일 생성(기본 저장 위치: 바탕화면)."""
    model = excel_io.read_model_from_text(text, orient)
    out_dir = out_dir or _default_outdir()
    os.makedirs(out_dir, exist_ok=True)
    made = []
    for ext in formats:
        out = os.path.join(out_dir, f"드릴다운표_{orient}_{theme}.{ext}")
        G.generate(model, out, orient, theme)
        made.append(out)
    return made


# ── 모던 GUI (CustomTkinter, 컴팩트: 입력가이드 최상단 + 방향/테마 호버 툴팁) ──
def launch_gui():
    try:
        import customtkinter as ctk
    except ImportError:
        return launch_gui_basic()
    from tkinter import filedialog, messagebox
    from PIL import Image

    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = ctk.CTk()
    app.title("드릴다운 표 생성기")
    app.geometry("470x600")
    app.minsize(450, 520)
    state = {"source": None}

    def f(size, weight="normal"):
        return ctk.CTkFont(family=KFONT, size=size, weight=weight)

    # 이미지 툴팁: '단일 공유 창'을 앱 창 바깥(오른쪽, 없으면 왼쪽)에 띄워 버튼을 가리지 않게.
    # 위젯을 벗어나면 즉시 withdraw(빠르게 사라짐), 다른 버튼으로 옮기면 같은 창을 갱신(중복 없음).
    _tip = {"win": None, "lbl": None}

    def tip_hide(_=None):
        if _tip["win"] is not None:
            _tip["win"].withdraw()

    def tip_show(get_img, near):
        img = get_img()
        if img is None:
            return
        if _tip["win"] is None:
            w = ctk.CTkToplevel(app)
            w.overrideredirect(True)
            w.attributes("-topmost", True)
            _tip["win"] = w
            _tip["lbl"] = ctk.CTkLabel(w, text="", corner_radius=6)
            _tip["lbl"].pack(padx=2, pady=2)
        _tip["lbl"].configure(image=img)
        _tip["win"].deiconify()
        _tip["win"].update_idletasks()
        tw = _tip["win"].winfo_width()
        x = app.winfo_rootx() + app.winfo_width() + 8
        if x + tw > app.winfo_screenwidth():
            x = max(0, app.winfo_rootx() - tw - 8)
        _tip["win"].geometry(f"+{x}+{near.winfo_rooty()}")

    def attach_tip(widget, get_img):
        widget.bind("<Enter>", lambda e: tip_show(get_img, widget), add="+")
        widget.bind("<Leave>", tip_hide, add="+")

    root = ctk.CTkFrame(app, fg_color="transparent")
    root.pack(fill="both", expand=True, padx=10, pady=8)

    # 헤더
    head = ctk.CTkFrame(root, fg_color="transparent")
    head.pack(fill="x", pady=(0, 8))
    ctk.CTkLabel(head, text="ㄱ", width=34, height=34, corner_radius=8,
                 fg_color=ACCENT, text_color="#fff", font=f(17, "bold")).pack(side="left")
    ht = ctk.CTkFrame(head, fg_color="transparent")
    ht.pack(side="left", padx=10)
    ctk.CTkLabel(ht, text="ㄱ자 계층형 표 만들기", font=f(16, "bold"), anchor="w").pack(anchor="w")
    ctk.CTkLabel(ht, text="Excel · Word · PowerPoint 표를 한 번에",
                 font=f(11), text_color=("gray45", "gray55"), anchor="w").pack(anchor="w")

    # 팝업 뷰어 (이미지 + 캡션 + 닫기)
    def show_image_popup(title, png_path, caption):
        if not os.path.exists(png_path):
            return
        win = ctk.CTkToplevel(app)
        win.title(title)
        win.attributes("-topmost", True)
        im = Image.open(png_path)
        ratio = min(640 / im.width, 520 / im.height)
        img = ctk.CTkImage(light_image=im, dark_image=im,
                           size=(int(im.width * ratio), int(im.height * ratio)))
        lab = ctk.CTkLabel(win, image=img, text="")
        lab.image = img
        lab.pack(padx=16, pady=(16, 8))
        ctk.CTkLabel(win, text=caption, font=f(12), wraplength=600, justify="left",
                     text_color=ACCENT).pack(padx=16, pady=(0, 10))
        ctk.CTkButton(win, text="닫기", width=120, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=win.destroy).pack(pady=(0, 16))
        win.update_idletasks()
        win.geometry(f"+{app.winfo_rootx() + 30}+{app.winfo_rooty() + 30}")
        win.lift()
        win.focus()

    def show_input_example():
        orient = ORIENT_MAP[seg_orient.get()]
        if orient == "both":
            png = os.path.join(_asset_dir(), "previews", "guide_both.png")
            cap = ("행+열(2D 교차표) 입력 형식입니다. 맨 위 제목·단위·코너 라벨은 선택.\n"
                   "이 표 영역만 엑셀에 채우거나 복사하세요.")
        else:
            png = os.path.join(_asset_dir(), "previews", "guide_multi.png")
            cap = ("왼쪽 라벨 열(대분류·중분류…)을 원하는 단계만큼, 그 오른쪽에 값(매출 등)을 적으세요.\n"
                   "단계 수는 자동 인식되고, 합계 줄은 생략하면 자동 합산됩니다('계'를 적어도 됨).\n"
                   "예시는 사업부>지역>분기(3단계) — 가지마다 깊이가 달라도 됩니다. 제목·단위는 선택.")
        show_image_popup("입력 형식 예시", png, cap)

    _r = {}

    def result_img_for(orient, theme):
        p = os.path.join(_asset_dir(), "previews", f"{orient}_{theme}.png")
        if not os.path.exists(p):
            return None
        im = Image.open(p)
        ratio = min(360 / im.width, 300 / im.height)
        _r["i"] = ctk.CTkImage(light_image=im, dark_image=im,
                               size=(int(im.width * ratio), int(im.height * ratio)))
        return _r["i"]

    # 1) 입력  (+ ⓘ 입력 형식 예시 팝업)
    irow = ctk.CTkFrame(root, fg_color="transparent")
    irow.pack(fill="x", pady=(2, 4))
    ctk.CTkLabel(irow, text="1   입력", font=f(13, "bold")).pack(side="left")
    ctk.CTkButton(irow, text="ⓘ 입력 형식 예시", height=26, font=f(11),
                  fg_color="transparent", border_width=1, text_color=ACCENT, border_color=ACCENT,
                  hover_color=("gray92", "gray22"), command=lambda: show_input_example()).pack(side="right")
    brow = ctk.CTkFrame(root, fg_color="transparent")
    brow.pack(fill="x")
    ctk.CTkButton(brow, text="파일 선택", height=34, font=f(13), fg_color=ACCENT, hover_color=ACCENT_HOVER,
                  command=lambda: pick_file()).pack(side="left", expand=True, fill="x", padx=(0, 3))
    ctk.CTkButton(brow, text="클립보드 붙여넣기", height=34, font=f(13), fg_color="transparent", border_width=1,
                  text_color=("gray20", "gray85"), border_color=("gray70", "gray40"),
                  hover_color=("gray90", "gray25"), command=lambda: paste_clip()).pack(side="left", expand=True, fill="x", padx=(3, 0))
    brow2 = ctk.CTkFrame(root, fg_color="transparent")
    brow2.pack(fill="x", pady=(5, 0))
    ctk.CTkButton(brow2, text="빈 양식 받기", height=28, font=f(11), fg_color="transparent", border_width=1,
                  text_color=("gray35", "gray70"), border_color=("gray75", "gray35"),
                  hover_color=("gray92", "gray22"), width=110, command=lambda: make_tpl()).pack(side="left")
    com_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(brow2, text="DRM·보호 파일", variable=com_var, font=f(11), fg_color=ACCENT,
                    hover_color=ACCENT_HOVER, checkbox_width=17, checkbox_height=17).pack(side="left", padx=10)
    file_lbl = ctk.CTkLabel(root, text="아직 선택 안 됨", font=f(11), text_color=("gray45", "gray55"),
                            anchor="w", wraplength=420, justify="left")
    file_lbl.pack(fill="x", pady=(4, 8))

    # 2) 방향  (+결과 예시 호버 툴팁)
    orow = ctk.CTkFrame(root, fg_color="transparent")
    orow.pack(fill="x", pady=(2, 4))
    ctk.CTkLabel(orow, text="2   방향", font=f(13, "bold")).pack(side="left")
    ctk.CTkLabel(orow, text="   각 버튼에 마우스를 올리면 결과 예시", font=f(11),
                 text_color=("gray45", "gray55")).pack(side="left")
    seg_orient = ctk.CTkSegmentedButton(root, values=list(ORIENT_MAP.keys()), font=f(12),
                                        selected_color=ACCENT, selected_hover_color=ACCENT_HOVER)
    seg_orient.set("행 기준")
    seg_orient.pack(fill="x")

    # 3) 테마  (+결과 예시 호버 툴팁)
    trow = ctk.CTkFrame(root, fg_color="transparent")
    trow.pack(fill="x", pady=(12, 4))
    ctk.CTkLabel(trow, text="3   테마", font=f(13, "bold")).pack(side="left")
    ctk.CTkLabel(trow, text="   각 버튼에 마우스를 올리면 결과 예시", font=f(11),
                 text_color=("gray45", "gray55")).pack(side="left")
    seg_theme = ctk.CTkSegmentedButton(root, values=list(THEME_MAP.keys()), font=f(12),
                                       selected_color=ACCENT, selected_hover_color=ACCENT_HOVER)
    seg_theme.set("컬러")
    seg_theme.pack(fill="x")

    # 4) 포맷
    ctk.CTkLabel(root, text="4   출력 포맷 (복수 선택)", font=f(13, "bold"), anchor="w").pack(fill="x", pady=(12, 4))
    frow = ctk.CTkFrame(root, fg_color="transparent")
    frow.pack(fill="x")
    fmt_vars = {}
    for label, ext in FORMATS:
        v = ctk.BooleanVar(value=(ext == "xlsx"))
        fmt_vars[ext] = v
        ctk.CTkCheckBox(frow, text=label, variable=v, font=f(12), fg_color=ACCENT,
                        hover_color=ACCENT_HOVER, checkbox_width=18, checkbox_height=18).pack(side="left", padx=(0, 16))

    status = ctk.CTkLabel(root, text="", font=f(11), text_color=ACCENT, wraplength=420, justify="left")
    status.pack(fill="x", pady=(12, 6))
    ctk.CTkButton(root, text="표 생성", height=44, font=f(15, "bold"), fg_color=ACCENT,
                  hover_color=ACCENT_HOVER, command=lambda: go()).pack(fill="x", pady=(0, 4))

    # ── 콜백 (위젯 뒤에 정의; 버튼은 lambda 로 지연 호출) ──
    def pick_file(_=None):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        state["source"] = ("file", p)
        file_lbl.configure(text="✓ " + os.path.basename(p), text_color="#1a7f37")

    def make_tpl():
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="입력양식.xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        excel_io.make_template(p)
        state["source"] = ("file", p)
        file_lbl.configure(text="✓ " + os.path.basename(p) + " (예시 채워짐)", text_color="#1a7f37")
        messagebox.showinfo("입력양식 생성", "예시가 채워진 입력양식을 만들었습니다.\n그대로 [표 생성] 하거나, 열어서 데이터를 바꿔 저장하세요.")

    def paste_clip():
        try:
            txt = app.clipboard_get()
        except Exception:
            txt = ""
        if not txt.strip():
            messagebox.showwarning("클립보드 비어있음", "엑셀에서 표 영역을 먼저 복사(Ctrl+C)한 뒤 눌러주세요.")
            return
        state["source"] = ("text", txt)
        n = len([ln for ln in txt.splitlines() if ln.strip()])
        file_lbl.configure(text=f"✓ 클립보드 {n}줄 불러옴 (바탕화면에 저장)", text_color="#1a7f37")

    def go():
        if not state["source"]:
            messagebox.showwarning("입력 필요", "파일을 선택하거나, 엑셀에서 복사 후 [클립보드 붙여넣기] 하세요.")
            return
        fmts = [e for e, v in fmt_vars.items() if v.get()]
        if not fmts:
            messagebox.showwarning("포맷 필요", "출력 포맷을 하나 이상 선택하세요.")
            return
        orient, theme = ORIENT_MAP[seg_orient.get()], THEME_MAP[seg_theme.get()]
        try:
            status.configure(text="생성 중…", text_color=ACCENT)
            app.update_idletasks()
            kind, data = state["source"]
            if kind == "text":
                made = run_generate_text(data, orient, theme, fmts)
            else:
                made = run_generate(data, orient, theme, fmts,
                                    com=(True if com_var.get() else "auto"))
            status.configure(text="완료 · " + ", ".join(os.path.basename(m) for m in made),
                             text_color="#1a7f37")
            if messagebox.askyesno("완료", "표를 만들었습니다. 생성 폴더를 열까요?"):
                os.startfile(os.path.dirname(made[0]))
        except Exception as e:
            status.configure(text="오류: " + str(e), text_color="#cf222e")
            messagebox.showerror("오류", f"{e}\n\n{traceback.format_exc()}")

    def bind_seg_tips(seg, kind):
        for val, btn in getattr(seg, "_buttons_dict", {}).items():
            if kind == "orient":
                fn = (lambda v: lambda: result_img_for(ORIENT_MAP[v], THEME_MAP[seg_theme.get()]))(val)
            else:
                fn = (lambda v: lambda: result_img_for(ORIENT_MAP[seg_orient.get()], THEME_MAP[v]))(val)
            attach_tip(btn, fn)

    bind_seg_tips(seg_orient, "orient")
    bind_seg_tips(seg_theme, "theme")
    app.mainloop()


# ── 폴백 GUI (CustomTkinter 미설치 시 기본 tkinter) ─────────────────────
def launch_gui_basic():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    root = tk.Tk()
    root.title("드릴다운 표 생성기 (ㄱ자/코너헤더)")
    root.geometry("520x560")
    pad = {"padx": 14, "pady": 6}
    state = {"in_path": None}
    tk.Label(root, text="드릴다운(ㄱ자) 계층형 표 생성기", font=(KFONT, 14, "bold")).pack(**pad)
    f1 = ttk.LabelFrame(root, text="1. 입력양식"); f1.pack(fill="x", **pad)

    def make_tpl():
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="입력양식.xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            excel_io.make_template(p); state["in_path"] = p
            lbl_file.config(text="선택: " + os.path.basename(p), fg="#1a7f37")

    def pick_file():
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            state["in_path"] = p; lbl_file.config(text="선택: " + os.path.basename(p), fg="#1a1a1a")

    ttk.Button(f1, text="입력양식 만들기…", command=make_tpl).pack(side="left", padx=8, pady=8)
    ttk.Button(f1, text="입력 파일 선택…", command=pick_file).pack(side="left", padx=8, pady=8)
    lbl_file = tk.Label(f1, text="(없음)", fg="#777"); lbl_file.pack(fill="x", padx=10, pady=(0, 8))
    f2 = ttk.LabelFrame(root, text="2. 방향"); f2.pack(fill="x", **pad)
    v_orient = tk.StringVar(value="row")
    for label, val in ORIENT_MAP.items():
        ttk.Radiobutton(f2, text=label, value=val, variable=v_orient).pack(anchor="w", padx=12)
    f3 = ttk.LabelFrame(root, text="3. 테마"); f3.pack(fill="x", **pad)
    v_theme = tk.StringVar(value="color")
    for label, val in THEME_MAP.items():
        ttk.Radiobutton(f3, text=label, value=val, variable=v_theme).pack(side="left", padx=8)
    f4 = ttk.LabelFrame(root, text="4. 출력 포맷"); f4.pack(fill="x", **pad)
    fmt_vars = {}
    for label, ext in FORMATS:
        bv = tk.BooleanVar(value=(ext == "xlsx")); fmt_vars[ext] = bv
        ttk.Checkbutton(f4, text=label, variable=bv).pack(side="left", padx=8)
    status = tk.Label(root, text="", fg="#1f6feb", wraplength=480); status.pack(fill="x", **pad)

    def go():
        if not state["in_path"]:
            messagebox.showwarning("입력 필요", "입력양식을 선택하세요."); return
        fmts = [e for e, v in fmt_vars.items() if v.get()]
        if not fmts:
            messagebox.showwarning("포맷 필요", "포맷을 선택하세요."); return
        try:
            made = run_generate(state["in_path"], v_orient.get(), v_theme.get(), fmts)
            status.config(text="완료: " + ", ".join(os.path.basename(m) for m in made), fg="#1a7f37")
            if messagebox.askyesno("완료", "생성 폴더를 열까요?"):
                os.startfile(os.path.dirname(made[0]))
        except Exception as e:
            messagebox.showerror("오류", str(e))

    ttk.Button(root, text="표 생성", command=go).pack(pady=10, ipadx=30, ipady=6)
    root.mainloop()


def _cli(args):
    """app IN.xlsx [orient] [theme] [fmt1,fmt2,...] [out_dir] — GUI 없이 생성."""
    in_path = args[0]
    orient = args[1] if len(args) > 1 else "row"
    theme = args[2] if len(args) > 2 else "color"
    fmts = args[3].split(",") if len(args) > 3 else ["xlsx"]
    out_dir = args[4] if len(args) > 4 else None
    for m in run_generate(in_path, orient, theme, fmts, out_dir):
        print("saved", m)


if __name__ == "__main__":
    a = sys.argv[1:]
    if a and a[0] == "--make-template":
        excel_io.make_template(a[1] if len(a) > 1 else "입력양식.xlsx")
        print("template saved")
    elif a and a[0] == "--paste":
        import tkinter as _tk
        _r = _tk.Tk(); _r.withdraw(); _txt = _r.clipboard_get(); _r.destroy()
        o = a[1] if len(a) > 1 else "row"
        th = a[2] if len(a) > 2 else "color"
        fm = a[3].split(",") if len(a) > 3 else ["xlsx"]
        for m in run_generate_text(_txt, o, th, fm):
            print("saved", m)
    elif a:
        _cli(a)
    else:
        launch_gui()
