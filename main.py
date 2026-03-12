"""
main.py
Smart Spice LIB 파일 편집기 - 메인 GUI 애플리케이션
Python 표준 라이브러리 Tkinter 기반
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from collections import OrderedDict

from data_model import LibFile, LibBlock, ModelEntry, ParamEntry, DirectiveEntry
from lib_parser import parse_lib
from lib_writer import save_lib, write_lib
from excel_exporter import export_lib_to_excel

# ─────────────────── 색상 & 폰트 상수 ────────────────────
BG_DARK    = "#1e1e2e"
BG_PANEL   = "#2a2a3d"
BG_HEADER  = "#313151"
BG_ROW_ODD = "#252538"
BG_ROW_EVN = "#2a2a3d"
FG_MAIN    = "#cdd6f4"
FG_DIM     = "#6e7399"
FG_ACCENT  = "#89b4fa"
FG_GREEN   = "#a6e3a1"
FG_RED     = "#f38ba8"
FG_YELLOW  = "#f9e2af"
FG_PURPLE  = "#cba6f7"
SEL_BG     = "#45475a"
BORDER     = "#585b70"
FONT_BODY  = ("Consolas", 11)
FONT_BOLD  = ("Consolas", 11, "bold")
FONT_TITLE = ("Segoe UI", 13, "bold")
FONT_SMALL = ("Consolas", 10)


# ──────────────────── 인라인 셀 편집기 ───────────────────
class InlineCellEditor:
    """
    Tkinter Treeview 안에서 Excel처럼 표의 셀을 직접 클릭해서 수정할 수 있게 해주는 헬퍼 클래스입니다.
    선택된 셀 위치에 임시 `tk.Entry` 텍스트 박스 위젯을 겹쳐 올린 뒤, 
    유저가 값 입력을 마치고 Enter/Tab을 누르면 원본 노드의 데이터를 갱신(commit)합니다.
    """

    def __init__(self, tree: ttk.Treeview, on_commit):
        self.tree = tree
        self.on_commit = on_commit  # 값이 확정되었을 때 호출할 콜백 (item_id, col_index, new_value) -> None
        self._entry = None
        self._item = None
        self._col = None

    def start_edit(self, item: str, col_index: int):
        """지정된 셀에 Entry 위젯을 띄웁니다."""
        self.cancel()
        col_id = f"#{col_index + 1}"
        bbox = self.tree.bbox(item, col_id)
        if not bbox:
            return
        x, y, width, height = bbox
        value = self.tree.set(item, col_id)

        self._item = item
        self._col = col_index

        self._entry = tk.Entry(
            self.tree,
            font=FONT_BODY,
            background=BG_HEADER,
            foreground=FG_MAIN,
            insertbackground=FG_MAIN,
            relief="flat",
            highlightthickness=1,
            highlightcolor=FG_ACCENT,
            highlightbackground=BORDER,
        )
        self._entry.insert(0, value)
        self._entry.place(x=x, y=y, width=width, height=height)
        self._entry.focus_set()
        self._entry.select_range(0, tk.END)
        self._entry.bind("<Return>", self._commit)
        self._entry.bind("<Escape>", lambda e: self.cancel())
        self._entry.bind("<Tab>", self._commit)

    def _commit(self, event=None):
        if self._entry is None:
            return
        new_value = self._entry.get()
        item, col = self._item, self._col
        self.cancel()
        self.on_commit(item, col, new_value)

    def cancel(self):
        if self._entry:
            self._entry.destroy()
            self._entry = None
        self._item = None
        self._col = None


# ─────────────────── 메인 애플리케이션 ───────────────────
class LibEditorApp(tk.Tk):
    """
    Tkinter로 작성된 Smart Spice .lib 에디터 메인 창 클래스입니다.
    파일을 읽고 저장하는 전체 라이프사이클과 화면 분할, 트리뷰 기반 탐색 UI,
    단축키 이벤트 처리 등을 관장합니다.
    """

    def __init__(self):
        super().__init__()
        self.title("Smart Spice LIB Editor")
        self.geometry("1280x780")
        self.minsize(900, 600)
        self.configure(bg=BG_DARK)

        self.lib_file: LibFile = None
        self._current_node = None   # 현재 선택된 트리 노드 (model / params)
        self._node_map = {}          # tree item id → (LibBlock | ModelEntry | 'global_params' | LibBlock[for params])
        self._param_items = []       # 현재 파라미터 테이블 행 (iid 리스트)

        self._setup_styles()
        self._build_toolbar()
        self._build_main_layout()

    # ── 스타일 설정 ───────────────────────────────────────
    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".",
            background=BG_DARK,
            foreground=FG_MAIN,
            fieldbackground=BG_PANEL,
            troughcolor=BG_PANEL,
            bordercolor=BORDER,
            darkcolor=BG_DARK,
            lightcolor=BG_PANEL,
            font=FONT_BODY,
        )
        style.configure("Toolbar.TFrame", background=BG_HEADER)
        style.configure("Toolbar.TButton",
            background=BG_HEADER, foreground=FG_MAIN,
            relief="flat", padding=(10, 5), font=FONT_BOLD,
            borderwidth=0,
        )
        style.map("Toolbar.TButton",
            background=[("active", SEL_BG), ("pressed", BORDER)],
            foreground=[("active", FG_ACCENT)],
        )
        style.configure("Accent.TButton",
            background=FG_ACCENT, foreground=BG_DARK,
            relief="flat", padding=(8, 4), font=FONT_BOLD,
        )
        style.map("Accent.TButton",
            background=[("active", "#74a8e8"), ("pressed", "#5e8fcf")],
        )
        style.configure("Danger.TButton",
            background="#45192a", foreground=FG_RED,
            relief="flat", padding=(8, 4), font=FONT_BOLD,
        )
        style.map("Danger.TButton",
            background=[("active", "#6b2035"), ("pressed", "#8b2a45")],
        )
        style.configure("Tree.Treeview",
            background=BG_PANEL, foreground=FG_MAIN,
            fieldbackground=BG_PANEL, rowheight=26,
            borderwidth=0,
        )
        style.configure("Tree.Treeview.Heading",
            background=BG_HEADER, foreground=FG_ACCENT,
            relief="flat", font=FONT_BOLD,
        )
        style.map("Tree.Treeview",
            background=[("selected", FG_ACCENT)],
            foreground=[("selected", BG_DARK)],
        )
        style.configure("Param.Treeview",
            background=BG_PANEL, foreground=FG_MAIN,
            fieldbackground=BG_PANEL, rowheight=28,
            borderwidth=0,
        )
        style.configure("Param.Treeview.Heading",
            background=BG_HEADER, foreground=FG_ACCENT,
            relief="flat", font=FONT_BOLD,
        )
        style.map("Param.Treeview",
            background=[("selected", SEL_BG)],
            foreground=[("selected", FG_ACCENT)],
        )
        style.configure("TScrollbar",
            background=BG_PANEL, troughcolor=BG_DARK,
            arrowcolor=FG_DIM, borderwidth=0, width=8,
        )
        style.configure("TLabel",
            background=BG_DARK, foreground=FG_MAIN,
            font=FONT_BODY,
        )
        style.configure("Header.TLabel",
            background=BG_HEADER, foreground=FG_ACCENT,
            font=FONT_TITLE, padding=(12, 6),
        )
        style.configure("TEntry",
            fieldbackground=BG_PANEL, foreground=FG_MAIN,
            insertcolor=FG_MAIN, borderwidth=1,
            relief="flat",
        )

    # ── 툴바 ─────────────────────────────────────────────
    def _build_toolbar(self):
        tb = ttk.Frame(self, style="Toolbar.TFrame")
        tb.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(tb, text="📂  파일 열기",  style="Toolbar.TButton",
                   command=self._open_file).pack(side=tk.LEFT, padx=2, pady=4)
        ttk.Button(tb, text="💾  저장",       style="Toolbar.TButton",
                   command=self._save_file).pack(side=tk.LEFT, padx=2, pady=4)
        ttk.Button(tb, text="💾  다른 이름으로 저장", style="Toolbar.TButton",
                   command=self._save_as_file).pack(side=tk.LEFT, padx=2, pady=4)
        ttk.Button(tb, text="👁  내용 미리보기", style="Toolbar.TButton",
                   command=self._preview).pack(side=tk.LEFT, padx=2, pady=4)
        ttk.Button(tb, text="📊  Excel 내보내기", style="Toolbar.TButton",
                   command=self._export_excel).pack(side=tk.LEFT, padx=2, pady=4)

        self._filepath_var = tk.StringVar(value="— 파일을 열어주세요 —")
        ttk.Label(tb, textvariable=self._filepath_var,
                  style="Toolbar.TButton", foreground=FG_DIM).pack(
                  side=tk.LEFT, padx=20)

    # ── 메인 레이아웃 (좌: 트리 / 우: 편집) ──────────────
    def _build_main_layout(self):
        pw = tk.PanedWindow(self, orient=tk.HORIZONTAL,
                            bg=BORDER, sashwidth=4,
                            sashrelief="flat")
        pw.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        # ── 좌측: 트리 뷰 ──
        left = tk.Frame(pw, bg=BG_PANEL, width=280)
        left.pack_propagate(False)
        pw.add(left, minsize=200)

        lbl = tk.Label(left, text="📋  라이브러리 구조",
                       bg=BG_HEADER, fg=FG_ACCENT,
                       font=FONT_TITLE, anchor="w", padx=12, pady=8)
        lbl.pack(fill=tk.X)

        tree_frame = tk.Frame(left, bg=BG_PANEL)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(tree_frame, style="Tree.Treeview",
                                  yscrollcommand=vsb.set, show="tree headings",
                                  selectmode="browse")
        self.tree.heading("#0", text="구조")
        self.tree.column("#0", width=250, minwidth=150)
        self.tree.pack(fill=tk.BOTH, expand=True)
        vsb.config(command=self.tree.yview)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # 뷰 전환 리스트
        ttk.Button(left, text="🔄  파라미터 중심 뷰 열기", style="Toolbar.TButton",
                   command=self._open_param_view).pack(fill=tk.X, side=tk.BOTTOM, padx=4, pady=4)

        # ── 우측: 편집 패널 ──
        right = tk.Frame(pw, bg=BG_DARK)
        pw.add(right, minsize=400)

        # 상단 정보 헤더
        self._info_var = tk.StringVar(value="← 좌측 트리에서 항목을 선택하세요")
        info_lbl = tk.Label(right, textvariable=self._info_var,
                             bg=BG_HEADER, fg=FG_ACCENT,
                             font=FONT_TITLE, anchor="w", padx=14, pady=8)
        info_lbl.pack(fill=tk.X)

        # 파라미터 테이블 영역
        table_frame = tk.Frame(right, bg=BG_DARK)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        tvsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        tvsb.pack(side=tk.RIGHT, fill=tk.Y)
        thsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        thsb.pack(side=tk.BOTTOM, fill=tk.X)

        self.param_tree = ttk.Treeview(
            table_frame,
            style="Param.Treeview",
            columns=("name", "value"),
            show="headings",
            selectmode="browse",
            yscrollcommand=tvsb.set,
            xscrollcommand=thsb.set,
        )
        self.param_tree.heading("name",  text="Parameter 명",  anchor="w")
        self.param_tree.heading("value", text="Parameter 값",  anchor="w")
        self.param_tree.column("name",  width=280, minwidth=120, anchor="w")
        self.param_tree.column("value", width=400, minwidth=120, anchor="w")
        self.param_tree.pack(fill=tk.BOTH, expand=True)
        tvsb.config(command=self.param_tree.yview)
        thsb.config(command=self.param_tree.xview)

        # 색상 태그
        self.param_tree.tag_configure("odd",  background=BG_ROW_ODD)
        self.param_tree.tag_configure("even", background=BG_ROW_EVN)
        self.param_tree.tag_configure("var",  foreground=FG_YELLOW)
        self.param_tree.tag_configure("expr", foreground=FG_PURPLE)
        self.param_tree.tag_configure("num",  foreground=FG_GREEN)

        # 더블클릭 → 인라인 편집
        self.param_tree.bind("<Double-1>", self._on_param_dblclick)
        self._cell_editor = InlineCellEditor(self.param_tree, self._on_cell_commit)

        # 하단 버튼 바
        btn_bar = tk.Frame(right, bg=BG_PANEL, pady=6)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Button(btn_bar, text="＋  파라미터 추가", style="Accent.TButton",
                   command=self._add_param).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_bar, text="－  삭제", style="Danger.TButton",
                   command=self._delete_param).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_bar, text="📋  일괄 수정", style="Toolbar.TButton",
                   command=self._batch_edit_param).pack(side=tk.LEFT, padx=4)

        # 변수(PARAM) 전용 하단 섹션
        self._var_frame = tk.Frame(right, bg=BG_DARK)
        self._var_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=8, pady=(0, 4))

    # ── 파일 열기 ─────────────────────────────────────────
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="LIB 파일 선택",
            filetypes=[("LIB 파일", "*.lib *.LIB"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        try:
            self.lib_file = parse_lib(path)
            self._filepath_var.set(path)
            self._rebuild_tree()
            self._info_var.set("← 좌측 트리에서 MODEL이나 PARAMS를 선택하세요")
            self._clear_param_table()
        except Exception as e:
            messagebox.showerror("파싱 오류", f"파일을 읽는 중 오류가 발생했습니다:\n{e}")

    # ── 저장 ──────────────────────────────────────────────
    def _save_file(self):
        if not self.lib_file:
            messagebox.showwarning("경고", "열린 파일이 없습니다.")
            return
        try:
            path = save_lib(self.lib_file)
            messagebox.showinfo("저장 완료", f"저장 완료:\n{path}")
        except Exception as e:
            messagebox.showerror("저장 오류", str(e))

    def _save_as_file(self):
        if not self.lib_file:
            messagebox.showwarning("경고", "열린 파일이 없습니다.")
            return
        path = filedialog.asksaveasfilename(
            title="다른 이름으로 저장",
            defaultextension=".lib",
            filetypes=[("LIB 파일", "*.lib"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        try:
            save_lib(self.lib_file, filepath=path)
            self.lib_file.filepath = path
            self._filepath_var.set(path)
            messagebox.showinfo("저장 완료", f"저장 완료:\n{path}")
        except Exception as e:
            messagebox.showerror("저장 오류", str(e))

    # ── 미리보기 ──────────────────────────────────────────
    def _preview(self):
        if not self.lib_file:
            messagebox.showwarning("경고", "열린 파일이 없습니다.")
            return
        text = write_lib(self.lib_file)
        win = tk.Toplevel(self)
        win.title("📄 LIB 파일 미리보기")
        win.geometry("900x650")
        win.configure(bg=BG_DARK)

        top_lbl = tk.Label(win, text="LIB 파일 미리보기 (읽기 전용)",
                           bg=BG_HEADER, fg=FG_ACCENT,
                           font=FONT_TITLE, anchor="w", padx=12, pady=6)
        top_lbl.pack(fill=tk.X)

        frame = tk.Frame(win, bg=BG_DARK)
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        sb_y = ttk.Scrollbar(frame); sb_y.pack(side=tk.RIGHT, fill=tk.Y)
        sb_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
        sb_x.pack(side=tk.BOTTOM, fill=tk.X)
        txt = tk.Text(frame,
                      bg=BG_PANEL, fg=FG_MAIN,
                      font=FONT_BODY, wrap=tk.NONE,
                      insertbackground=FG_MAIN,
                      yscrollcommand=sb_y.set,
                      xscrollcommand=sb_x.set,
                      relief="flat")
        txt.pack(fill=tk.BOTH, expand=True)
        sb_y.config(command=txt.yview)
        sb_x.config(command=txt.xview)
        txt.insert("1.0", text)
        txt.configure(state="disabled")

    # ── 트리 빌드 (GUI 갱신) ──────────────────────────────
    def _rebuild_tree(self):
        """
        메모리에 파싱된 `self.lib_file` 객체의 최신 데이터를 바탕으로
        좌측 사이드바 트리뷰(Treeview)를 초기화하고 화면에 다시 그려줍니다.
        전역 파라미터 -> 전역 설정 지시어 -> 각 LIB 블록 (파라미터 -> 지시어 -> 모델 순)으로 렌더링합니다.
        """
        self.tree.delete(*self.tree.get_children())
        self._node_map.clear()  # 트리 노드 ID와 실제 백엔드 데이터 객체 매핑 테이블 초기화

        if not self.lib_file:
            return

        # 전역 PARAMS 노드
        if self.lib_file.global_params:
            node = self.tree.insert("", "end",
                                    text="🔧 전역 PARAMS",
                                    open=True, tags=("params",))
            self._node_map[node] = ("global_params", None)
            for pe in self.lib_file.global_params:
                child = self.tree.insert(node, "end",
                                         text=f"  {pe.name} = {pe.value}",
                                         tags=("param_var",))
                self._node_map[child] = ("global_param_var", pe)

        # 전역 기타 directive
        if self.lib_file.global_directives:
            node = self.tree.insert("", "end",
                                    text="📄 전역 설정 (.directives)",
                                    open=True, tags=("directives",))
            self._node_map[node] = ("global_directives", None)
            for de in self.lib_file.global_directives:
                child = self.tree.insert(node, "end",
                                         text=f"  {de.raw_text}",
                                         tags=("directive_var",))
                self._node_map[child] = ("global_directive_var", de)

        # LIB 블록들
        for lb in self.lib_file.lib_blocks:
            lib_node = self.tree.insert("", "end",
                                         text=f"📁 LIB: {lb.name}",
                                         open=True, tags=("lib",))
            self._node_map[lib_node] = ("lib", lb)

            # LIB 내 PARAMS
            if lb.params:
                p_node = self.tree.insert(lib_node, "end",
                                          text="  🔧 PARAMS",
                                          tags=("params",))
                self._node_map[p_node] = ("lib_params", lb)

            # LIB 내 기타 directive
            if lb.directives:
                d_node = self.tree.insert(lib_node, "end",
                                          text="  📄 설정 (.directives)",
                                          tags=("directives",))
                self._node_map[d_node] = ("lib_directives", lb)

            # MODEL들
            for model in lb.models:
                m_node = self.tree.insert(
                    lib_node, "end",
                    text=f"  ⚙️ {model.name}  [{model.model_type}]",
                    tags=("model",),
                )
                self._node_map[m_node] = ("model", model, lb)

        # 트리 태그 색상
        self.tree.tag_configure("lib",           foreground=FG_ACCENT)
        self.tree.tag_configure("model",         foreground=FG_GREEN)
        self.tree.tag_configure("params",        foreground=FG_YELLOW)
        self.tree.tag_configure("param_var",     foreground=FG_DIM)
        self.tree.tag_configure("directives",    foreground=FG_PURPLE)
        self.tree.tag_configure("directive_var", foreground=FG_DIM)

    # ── 트리 더블클릭 (이름 변경) ─────────────────────────
    def _on_tree_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        info = self._node_map.get(iid)
        if not info:
            return
            
        kind = info[0]
        if kind == "lib":
            lb = info[1]
            new_name = simpledialog.askstring("LIB 이름 변경", "새 LIB 이름을 입력하세요:", initialvalue=lb.name, parent=self)
            if new_name and new_name.strip():
                lb.name = new_name.strip()
                self._rebuild_tree()
                
        elif kind == "model":
            model = info[1]
            dlg = ParamAddDialog(self, title="MODEL 속성 변경")
            dlg._name_var.set(model.name)
            dlg._val_var.set(model.model_type)
            # 재활용 다이얼로그 라벨 변경
            # dialog는 tk.Toplevel 이므로 show_model_params처럼 변경하려면 별도 클래스가 낫지만 임시로
            if dlg.result:
                new_name, new_type = dlg.result
                model.name = new_name
                model.model_type = new_type
                self._rebuild_tree()
                self._on_tree_select()

    # ── 트리 선택 이벤트 ──────────────────────────────────
    def _on_tree_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        info = self._node_map.get(iid)
        if not info:
            return
        kind = info[0]

        if kind == "model":
            _, model, lb = info
            self._current_node = ("model", model, lb)
            self._info_var.set(f"⚙️  MODEL: {model.name}   타입: {model.model_type}")
            self._show_model_params(model)

        elif kind == "global_params":
            self._current_node = ("global_params", None)
            self._info_var.set("🔧  전역 PARAMS  (.PARAM 변수)")
            self._show_param_list(self.lib_file.global_params)

        elif kind == "lib_params":
            lb = info[1]
            self._current_node = ("lib_params", lb)
            self._info_var.set(f"🔧  LIB [{lb.name}] – PARAMS  (.PARAM 변수)")
            self._show_param_list(lb.params)

        elif kind == "lib":
            lb = info[1]
            self._current_node = ("lib", lb)
            self._info_var.set(f"📁  LIB: {lb.name}   (MODEL {len(lb.models)}개 / PARAM {len(lb.params)}개)")
            self._clear_param_table()

        elif kind == "global_directives":
            self._current_node = ("global_directives", None)
            self._info_var.set("📄  전역 설정 (.directives)")
            self._show_directive_list(self.lib_file.global_directives)

        elif kind == "lib_directives":
            lb = info[1]
            self._current_node = ("lib_directives", lb)
            self._info_var.set(f"📄  LIB [{lb.name}] – 설정 (.directives)")
            self._show_directive_list(lb.directives)

        else:
            self._clear_param_table()

    # ── 파라미터 테이블 표시 (MODEL용) ────────────────────
    def _show_model_params(self, model: ModelEntry):
        self._clear_param_table()
        for i, (name, value) in enumerate(model.params.items()):
            tag = self._value_tag(value, i)
            iid = self.param_tree.insert("", "end",
                                          values=(name, value),
                                          tags=(tag,))
            self._param_items.append(iid)

    # ── 파라미터 목록 표시 (PARAM 변수용) ─────────────────
    def _show_param_list(self, param_list: list):
        self._clear_param_table()
        self.param_tree.heading("name",  text="변수명")
        self.param_tree.heading("value", text="변수 값")
        for i, pe in enumerate(param_list):
            tag = self._value_tag(pe.value, i)
            iid = self.param_tree.insert("", "end",
                                          values=(pe.name, pe.value),
                                          tags=(tag,))
            self._param_items.append(iid)

    # ── 지시어 목록 표시 (Directive 변수용) ─────────────────
    def _show_directive_list(self, directive_list: list):
        self._clear_param_table()
        self.param_tree.heading("name",  text="키워드")
        self.param_tree.heading("value", text="전체 지시어 (원문)")
        for i, de in enumerate(directive_list):
            base = "odd" if i % 2 == 0 else "even"
            iid = self.param_tree.insert("", "end",
                                          values=(de.keyword, de.raw_text),
                                          tags=(base,))
            self._param_items.append(iid)

    def _value_tag(self, value: str, row_idx: int) -> str:
        base = "odd" if row_idx % 2 == 0 else "even"
        if "{" in value and "}" in value:
            # 수식/변수 참조
            stripped = value.strip("{} ")
            if any(op in stripped for op in ["+", "-", "*", "/"]):
                return "expr"
            return "var"
        try:
            float(value)
            return "num"
        except ValueError:
            pass
        return base

    def _clear_param_table(self):
        self.param_tree.heading("name",  text="Parameter 명")
        self.param_tree.heading("value", text="Parameter 값")
        for iid in self.param_tree.get_children():
            self.param_tree.delete(iid)
        self._param_items = []

    # ── 인라인 편집 ───────────────────────────────────────
    def _on_param_dblclick(self, event):
        region = self.param_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = self.param_tree.identify_column(event.x)
        col_index = int(col_id.replace("#", "")) - 1  # 0-based
        item = self.param_tree.identify_row(event.y)
        if not item:
            return
        self._cell_editor.start_edit(item, col_index)

    def _on_cell_commit(self, item_id: str, col_index: int, new_value: str):
        """셀 편집 완료 시 데이터 모델 업데이트"""
        col_id = f"#{col_index + 1}"
        old_values = self.param_tree.item(item_id, "values")
        if not old_values:
            return
        old_name, old_val = old_values[0], old_values[1]

        if col_index == 0:
            new_name = new_value.strip()
            if not new_name:
                return
            self._rename_param_key(old_name, new_name)
            self.param_tree.set(item_id, "#1", new_name)
        elif col_index == 1:
            self._update_param_value(old_name, new_value.strip())
            self.param_tree.set(item_id, "#2", new_value.strip())
            # 태그 재적용
            row_idx = self.param_tree.index(item_id)
            tag = self._value_tag(new_value.strip(), row_idx)
            self.param_tree.item(item_id, tags=(tag,))

    def _rename_param_key(self, old_name: str, new_name: str):
        node = self._current_node
        if not node:
            return
        kind = node[0]
        if kind == "model":
            model: ModelEntry = node[1]
            if old_name in model.params:
                keys = list(model.params.keys())
                idx = keys.index(old_name)
                items = list(model.params.items())
                items[idx] = (new_name, items[idx][1])
                model.params = OrderedDict(items)
        elif kind in ("global_params", "lib_params"):
            param_list = (self.lib_file.global_params
                          if kind == "global_params"
                          else node[1].params)
            for pe in param_list:
                if pe.name == old_name:
                    pe.name = new_name
                    break
        elif kind in ("global_directives", "lib_directives"):
            directive_list = (self.lib_file.global_directives
                              if kind == "global_directives"
                              else node[1].directives)
            for de in directive_list:
                if de.keyword == old_name:
                    de.keyword = new_name
                    break
                    
    def _update_param_value(self, param_name: str, new_value: str):
        node = self._current_node
        if not node:
            return
        kind = node[0]
        if kind == "model":
            model: ModelEntry = node[1]
            if param_name in model.params:
                model.params[param_name] = new_value
        elif kind in ("global_params", "lib_params"):
            param_list = (self.lib_file.global_params
                          if kind == "global_params"
                          else node[1].params)
            for pe in param_list:
                if pe.name == param_name:
                    pe.value = new_value
                    break
        elif kind in ("global_directives", "lib_directives"):
            directive_list = (self.lib_file.global_directives
                              if kind == "global_directives"
                              else node[1].directives)
            for de in directive_list:
                if de.keyword == param_name:
                    de.raw_text = new_value
                    break

    # ── 파라미터 추가 ─────────────────────────────────────
    def _add_param(self):
        node = self._current_node
        if not node:
            messagebox.showwarning("선택 필요", "먼저 좌측 트리에서 MODEL 또는 PARAMS를 선택하세요.")
            return

        kind = node[0]

        if kind == "model":
            model: ModelEntry = node[1]
            dlg = ParamAddDialog(self, title="파라미터 추가")
            if dlg.result:
                pname, pval = dlg.result
                model.params[pname] = pval
                row_idx = len(self._param_items)
                tag = self._value_tag(pval, row_idx)
                iid = self.param_tree.insert("", "end",
                                              values=(pname, pval),
                                              tags=(tag,))
                self._param_items.append(iid)
                self.param_tree.see(iid)

        elif kind in ("global_params", "lib_params"):
            param_list = (self.lib_file.global_params
                          if kind == "global_params"
                          else node[1].params)
            dlg = ParamAddDialog(self, title="변수(PARAM) 추가")
            if dlg.result:
                pname, pval = dlg.result
                param_list.append(ParamEntry(name=pname, value=pval))
                row_idx = len(self._param_items)
                tag = self._value_tag(pval, row_idx)
                iid = self.param_tree.insert("", "end",
                                              values=(pname, pval),
                                              tags=(tag,))
                self._param_items.append(iid)
                self.param_tree.see(iid)
                # 트리 업데이트
                self._rebuild_tree()

        elif kind in ("global_directives", "lib_directives"):
            directive_list = (self.lib_file.global_directives
                              if kind == "global_directives"
                              else node[1].directives)
            dlg = ParamAddDialog(self, title="지시어(Directive) 추가")
            if dlg.result:
                keyw, text = dlg.result
                directive_list.append(DirectiveEntry(keyword=keyw, raw_text=text))
                row_idx = len(self._param_items)
                base = "odd" if row_idx % 2 == 0 else "even"
                iid = self.param_tree.insert("", "end",
                                              values=(keyw, text),
                                              tags=(base,))
                self._param_items.append(iid)
                self.param_tree.see(iid)
                self._rebuild_tree()

        elif kind == "lib":
            # LIB 블록에 PARAMS를 추가
            lb: LibBlock = node[1]
            dlg = ParamAddDialog(self, title="변수(PARAM) 추가")
            if dlg.result:
                pname, pval = dlg.result
                lb.params.append(ParamEntry(name=pname, value=pval))
                self._rebuild_tree()

    # ── 파라미터 삭제 ─────────────────────────────────────
    def _delete_param(self):
        sel = self.param_tree.selection()
        if not sel:
            messagebox.showwarning("선택 필요", "삭제할 파라미터를 선택하세요.")
            return
        item_id = sel[0]
        values = self.param_tree.item(item_id, "values")
        if not values:
            return
        param_name = values[0]

        if not messagebox.askyesno("삭제 확인",
                                    f"'{param_name}' 파라미터를 삭제하시겠습니까?"):
            return

        node = self._current_node
        if not node:
            return
        kind = node[0]

        if kind == "model":
            model: ModelEntry = node[1]
            if param_name in model.params:
                del model.params[param_name]

        elif kind in ("global_params", "lib_params"):
            param_list = (self.lib_file.global_params
                          if kind == "global_params"
                          else node[1].params)
            for i, pe in enumerate(param_list):
                if pe.name == param_name:
                    param_list.pop(i)
                    break
            self._rebuild_tree()

        elif kind in ("global_directives", "lib_directives"):
            directive_list = (self.lib_file.global_directives
                              if kind == "global_directives"
                              else node[1].directives)
            for i, de in enumerate(directive_list):
                if de.keyword == param_name:
                    directive_list.pop(i)
                    break
            self._rebuild_tree()

        self.param_tree.delete(item_id)
        if item_id in self._param_items:
            self._param_items.remove(item_id)

    # ── 일괄 수정 / 엑셀 내보내기 ──────────────────────────
    def _export_excel(self):
        if not self.lib_file:
            messagebox.showwarning("경고", "열린 파일이 없습니다.")
            return
            
        default_name = os.path.splitext(os.path.basename(self.lib_file.filepath))[0] + "_export.xlsx"
        path = filedialog.asksaveasfilename(
            title="Excel로 내보내기",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if not path:
            return
            
        try:
            export_lib_to_excel(self.lib_file, path)
            messagebox.showinfo("내보내기 완료", f"Excel 생성 완료:\n{path}")
        except Exception as e:
            messagebox.showerror("오류", f"Excel 내보내기 실패:\n{e}")

    def _batch_edit_param(self):
        """
        일괄 수정(Batch Edit) 다이얼로그를 띄워 특정 이름을 가진 파라미터 값 전체를 변경합니다.
        범위(전체 또는 현재 선택된 LIB 블록 내)를 선택할 수 있으며,
        순회하면서 일치하는 파라미터가 있을 경우 지정한 새 값으로 모두 교체합니다.
        """
        if not self.lib_file:
            return
            
        dlg = BatchEditDialog(self, self.lib_file, getattr(self, '_current_node', None))
        if dlg.result:
            p_name, p_value, scope = dlg.result
            count = 0
            
            blocks = []
            if scope == "all":
                blocks = self.lib_file.lib_blocks
            elif scope == "lib" and self._current_node and self._current_node[0] in ("lib", "model", "lib_params", "lib_directives"):
                # find which lib is selected
                lb = self._current_node[1] if self._current_node[0] != "model" else self._current_node[2]
                blocks = [lb]
            else:
                blocks = self.lib_file.lib_blocks
                
            for lb in blocks:
                for model in lb.models:
                    if p_name in model.params:
                        model.params[p_name] = p_value
                        count += 1
                        
            messagebox.showinfo("적용 완료", f"총 {count}개의 모델에서 '{p_name}' 값이 변경되었습니다.")
            
            # 뷰 갱신
            if self._current_node:
                self._on_tree_select()

    def _open_param_view(self):
        if not self.lib_file:
            messagebox.showwarning("경고", "열린 파일이 없습니다.")
            return
        ParameterViewWindow(self, self.lib_file)


# ─────────────────────── 다이얼로그 ──────────────────────
class ParamAddDialog(tk.Toplevel):
    """파라미터 (또는 변수) 추가 다이얼로그"""

    def __init__(self, parent, title="파라미터 추가"):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.resizable(False, False)
        self.configure(bg=BG_DARK)
        self.grab_set()
        self._build(title)
        self.transient(parent)
        self.wait_window()

    def _build(self, title):
        pad = dict(padx=16, pady=6)

        tk.Label(self, text=title, bg=BG_HEADER, fg=FG_ACCENT,
                 font=FONT_TITLE, anchor="w", padx=14, pady=8
                 ).pack(fill=tk.X, **{})

        frm = tk.Frame(self, bg=BG_DARK, padx=20, pady=16)
        frm.pack(fill=tk.BOTH)

        tk.Label(frm, text="이름:", bg=BG_DARK, fg=FG_MAIN, font=FONT_BODY).grid(
            row=0, column=0, sticky="w", pady=6)
        self._name_var = tk.StringVar()
        tk.Entry(frm, textvariable=self._name_var,
                 bg=BG_PANEL, fg=FG_MAIN, insertbackground=FG_MAIN,
                 font=FONT_BODY, relief="flat",
                 highlightthickness=1, highlightcolor=FG_ACCENT,
                 highlightbackground=BORDER, width=28
                 ).grid(row=0, column=1, pady=6, padx=(8, 0))

        tk.Label(frm, text="값:", bg=BG_DARK, fg=FG_MAIN, font=FONT_BODY).grid(
            row=1, column=0, sticky="w", pady=6)
        self._val_var = tk.StringVar()
        tk.Entry(frm, textvariable=self._val_var,
                 bg=BG_PANEL, fg=FG_MAIN, insertbackground=FG_MAIN,
                 font=FONT_BODY, relief="flat",
                 highlightthickness=1, highlightcolor=FG_ACCENT,
                 highlightbackground=BORDER, width=28
                 ).grid(row=1, column=1, pady=6, padx=(8, 0))

        tk.Label(frm,
                 text="※ 수식은 중괄호로: {var_name}  또는  {var_name * 1.1}",
                 bg=BG_DARK, fg=FG_DIM, font=FONT_SMALL
                 ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 6))

        btn_frm = tk.Frame(self, bg=BG_DARK, pady=10)
        btn_frm.pack()
        ttk.Button(btn_frm, text="확인", style="Accent.TButton",
                   command=self._ok).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frm, text="취소", style="Toolbar.TButton",
                   command=self.destroy).pack(side=tk.LEFT, padx=8)

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())

    def _ok(self):
        name = self._name_var.get().strip()
        val  = self._val_var.get().strip()
        if not name:
            messagebox.showwarning("입력 오류", "이름을 입력하세요.", parent=self)
            return
        self.result = (name, val)
        self.destroy()


class BatchEditDialog(tk.Toplevel):
    """일괄 파라미터 수정 다이얼로그"""

    def __init__(self, parent, lib_file: LibFile, current_node):
        super().__init__(parent)
        self.title("일괄 파라미터 수정")
        self.result = None
        self.resizable(False, False)
        self.configure(bg=BG_DARK)
        self.grab_set()
        self.transient(parent)
        self._build(lib_file, current_node)
        self.wait_window()

    def _build(self, lib_file: LibFile, current_node):
        tk.Label(self, text="일괄 파라미터 수정",
                 bg=BG_HEADER, fg=FG_ACCENT,
                 font=FONT_TITLE, anchor="w", padx=14, pady=8
                 ).pack(fill=tk.X)

        frm = tk.Frame(self, bg=BG_DARK, padx=20, pady=16)
        frm.pack(fill=tk.BOTH)

        # 파라미터 목록 수집 (전체)
        all_params = set()
        for lb in lib_file.lib_blocks:
            for model in lb.models:
                all_params.update(model.params.keys())
        p_names = sorted(list(all_params))

        tk.Label(frm, text="대상 파라미터:", bg=BG_DARK, fg=FG_MAIN,
                 font=FONT_BODY).grid(row=0, column=0, sticky="w", pady=6)
        self._p_name = tk.StringVar()
        var_combo = ttk.Combobox(frm, textvariable=self._p_name,
                                 values=p_names, width=26, font=FONT_BODY)
        var_combo.grid(row=0, column=1, pady=6, padx=(8, 0))

        tk.Label(frm, text="새 파라미터 값:", bg=BG_DARK, fg=FG_MAIN,
                 font=FONT_BODY).grid(row=1, column=0, sticky="w", pady=6)
        self._p_val = tk.StringVar()
        tk.Entry(frm, textvariable=self._p_val,
                 bg=BG_PANEL, fg=FG_MAIN, insertbackground=FG_MAIN,
                 font=FONT_BODY, relief="flat",
                 highlightthickness=1, highlightcolor=FG_ACCENT,
                 highlightbackground=BORDER, width=28
                 ).grid(row=1, column=1, pady=6, padx=(8, 0))

        tk.Label(frm, text="적용 범위:",
                 bg=BG_DARK, fg=FG_MAIN, font=FONT_BODY
                 ).grid(row=2, column=0, sticky="w", pady=6)
                 
        self._scope_var = tk.StringVar(value="all")
        rb_frm = tk.Frame(frm, bg=BG_DARK)
        rb_frm.grid(row=2, column=1, sticky="w", pady=6, padx=(8,0))
        tk.Radiobutton(rb_frm, text="모든 LIB", variable=self._scope_var, value="all",
                       bg=BG_DARK, fg=FG_MAIN, selectcolor=BG_PANEL, font=FONT_BODY).pack(side=tk.LEFT)
        tk.Radiobutton(rb_frm, text="현재 보고있는 LIB", variable=self._scope_var, value="lib",
                       bg=BG_DARK, fg=FG_MAIN, selectcolor=BG_PANEL, font=FONT_BODY).pack(side=tk.LEFT)

        btn_frm = tk.Frame(self, bg=BG_DARK, pady=10)
        btn_frm.pack()
        ttk.Button(btn_frm, text="적용", style="Accent.TButton",
                   command=self._ok).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frm, text="취소", style="Toolbar.TButton",
                   command=self.destroy).pack(side=tk.LEFT, padx=8)

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())

    def _ok(self):
        p_name = self._p_name.get().strip()
        p_val  = self._p_val.get().strip()
        scope  = self._scope_var.get()
        if not p_name:
            messagebox.showwarning("입력 오류", "이름을 입력하세요.", parent=self)
            return
        self.result = (p_name, p_val, scope)
        self.destroy()

class ParameterViewWindow(tk.Toplevel):
    """선택한 파라미터에 대해 모델별 값을 보여주는 창"""

    def __init__(self, parent, lib_file: LibFile):
        super().__init__(parent)
        self.title("파라미터 중심 뷰")
        self.geometry("900x600")
        self.configure(bg=BG_DARK)
        self.lib_file = lib_file

        self._build()

    def _build(self):
        pw = tk.PanedWindow(self, orient=tk.HORIZONTAL, bg=BORDER, sashwidth=4, sashrelief="flat")
        pw.pack(fill=tk.BOTH, expand=True)

        # ── 좌측: 파라미터 목록 ──
        left = tk.Frame(pw, bg=BG_PANEL, width=250)
        pw.add(left, minsize=150)
        
        tk.Label(left, text="파라미터 목록", bg=BG_HEADER, fg=FG_ACCENT, font=FONT_TITLE, anchor="w", padx=12, pady=8).pack(fill=tk.X)

        tframe = tk.Frame(left, bg=BG_PANEL)
        tframe.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        vsb1 = ttk.Scrollbar(tframe, orient=tk.VERTICAL)
        vsb1.pack(side=tk.RIGHT, fill=tk.Y)

        self.p_tree = ttk.Treeview(tframe, style="Tree.Treeview", show="tree", yscrollcommand=vsb1.set)
        self.p_tree.pack(fill=tk.BOTH, expand=True)
        vsb1.config(command=self.p_tree.yview)
        self.p_tree.bind("<<TreeviewSelect>>", self._on_p_select)

        # ── 우측: 모델별 값 목록 ──
        right = tk.Frame(pw, bg=BG_DARK)
        pw.add(right, minsize=400)
        
        self._title_var = tk.StringVar(value="← 파라미터를 선택하세요")
        tk.Label(right, textvariable=self._title_var, bg=BG_HEADER, fg=FG_ACCENT, font=FONT_TITLE, anchor="w", padx=14, pady=8).pack(fill=tk.X)

        rframe = tk.Frame(right, bg=BG_DARK)
        rframe.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        vsb2 = ttk.Scrollbar(rframe, orient=tk.VERTICAL)
        vsb2.pack(side=tk.RIGHT, fill=tk.Y)

        self.v_tree = ttk.Treeview(rframe, style="Param.Treeview", columns=("lib", "model", "val"), show="headings", yscrollcommand=vsb2.set)
        self.v_tree.heading("lib", text="LIB명", anchor="w")
        self.v_tree.heading("model", text="MODEL명", anchor="w")
        self.v_tree.heading("val", text="값", anchor="w")
        self.v_tree.column("lib", width=200, anchor="w")
        self.v_tree.column("model", width=200, anchor="w")
        self.v_tree.column("val", width=200, anchor="w")
        
        self.v_tree.tag_configure("odd",  background=BG_ROW_ODD)
        self.v_tree.tag_configure("even", background=BG_ROW_EVN)
        
        self.v_tree.pack(fill=tk.BOTH, expand=True)
        vsb2.config(command=self.v_tree.yview)

        self._populate_params()

    def _populate_params(self):
        all_params = set()
        for lb in self.lib_file.lib_blocks:
            for model in lb.models:
                all_params.update(model.params.keys())
                
        for p in sorted(list(all_params)):
            self.p_tree.insert("", "end", text=f"  {p}", iid=p)

    def _on_p_select(self, event=None):
        sel = self.p_tree.selection()
        if not sel: return
        p_name = sel[0]
        
        for iid in self.v_tree.get_children():
            self.v_tree.delete(iid)
            
        self._title_var.set(f"파라미터: {p_name}")
        
        row_idx = 0
        for lb in self.lib_file.lib_blocks:
            for model in lb.models:
                if p_name in model.params:
                    val = model.params[p_name]
                    tag = "odd" if row_idx % 2 == 0 else "even"
                    self.v_tree.insert("", "end", values=(lb.name, model.name, val), tags=(tag,))
                    row_idx += 1


# ─────────────────────── 진입점 ──────────────────────────
if __name__ == "__main__":
    app = LibEditorApp()
    app.mainloop()
