"""
Bingo Grid Generator  v7
========================
Run with: python bingo_generator.py
Requires: pip install Pillow

PIPELINE:
  1. Build a number matrix (integers 0..pool_size-1) — no images yet
  2. Verify the matrix fully (row + column uniqueness)
  3. If verification fails, throw it away and rebuild — loop until clean
  4. Show a progress bar window while rendering images
  5. Save PNGs to the chosen folder
"""

import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import threading
import random
import os
from pathlib import Path

try:
    from PIL import Image, ImageTk, ImageDraw
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# ─── UI palette ───────────────────────────────────────────────────────────────
PREVIEW_W = 340
PREVIEW_H = 440
ACCENT    = "#b5651d"
ACCENT_LT = "#d4956a"
BG_DARK   = "#1a1008"
BG_MID    = "#2b1e10"
BG_PANEL  = "#3a2a18"
BG_CARD   = "#4a3520"
TEXT_MAIN = "#f5e6d0"
TEXT_DIM  = "#a08060"
BORDER    = "#6b4c2a"

FONT_HEAD  = ("Georgia", 13, "bold")
FONT_BODY  = ("Georgia", 10)
FONT_SMALL = ("Georgia", 9)
FONT_BIG   = ("Georgia", 22, "bold")
FONT_MONO  = ("Courier", 9)


# ═══════════════════════════════════════════════════════════════════════════════
#  STAGE 1 — NUMBER MATRIX
# ═══════════════════════════════════════════════════════════════════════════════

def build_matrix(pool_size, cells, n_grids, max_attempts=200):
    """
    Build and VERIFY an (n_grids x cells) integer matrix where:
      • Each row   contains `cells` distinct integers from 0..pool_size-1
      • Each column contains distinct integers across all rows
      (i.e. no index repeats in the same slot across different grids)

    Uses recursive backtracking to fill one row at a time.
    Retries up to max_attempts times from scratch if a dead-end is hit.

    Returns (matrix, attempts_taken) or raises RuntimeError if all attempts fail.
    """
    for attempt in range(1, max_attempts + 1):
        matrix = _try_build(pool_size, cells, n_grids)
        if matrix is not None:
            ok, errors = verify_matrix(matrix, cells, n_grids)
            if ok:
                return matrix, attempt
            # Verification found a bug in the builder — retry
        # else: backtracking failed — retry with fresh shuffle

    raise RuntimeError(
        f"Could not build a valid matrix after {max_attempts} attempts.\n"
        f"Try reducing the number of grids or adding more images to the pool.")


def _try_build(pool_size, cells, n_grids):
    """
    One attempt to build the full matrix using slot-aware backtracking.
    Returns the matrix on success, None on failure.
    """
    # col_used[s] = set of indices already placed in column s
    col_used = [set() for _ in range(cells)]
    matrix   = []

    for g in range(n_grids):
        row = _fill_row(pool_size, cells, col_used)
        if row is None:
            return None   # dead-end, caller should retry
        for s, idx in enumerate(row):
            col_used[s].add(idx)
        matrix.append(row)

    return matrix


def _fill_row(pool_size, cells, col_used):
    """
    Fill one row of `cells` distinct indices, respecting col_used.
    Uses randomised backtracking.
    Returns a list of `cells` ints, or None if no solution found.
    """
    # Shuffle the slot order so we don't always fill left-to-right
    slots = list(range(cells))
    random.shuffle(slots)

    row = [None] * cells

    def backtrack(depth):
        if depth == cells:
            return True
        s        = slots[depth]
        used_row = {row[slots[i]] for i in range(depth) if row[slots[i]] is not None}
        forbidden = used_row | col_used[s]
        candidates = [i for i in range(pool_size) if i not in forbidden]
        random.shuffle(candidates)
        for idx in candidates:
            row[s] = idx
            if backtrack(depth + 1):
                return True
            row[s] = None
        return False

    if backtrack(0):
        return row
    return None


def verify_matrix(matrix, cells, n_grids):
    """
    Full cross-referenced verification.
    Returns (True, []) if clean, or (False, [error_strings]) if not.
    """
    errors = []

    # Row check: no index repeats within a single grid
    for g, row in enumerate(matrix):
        seen = {}
        for s, idx in enumerate(row):
            if idx in seen:
                errors.append(
                    f"Grid {g+1}: index {idx} appears in slots {seen[idx]} and {s}")
            else:
                seen[idx] = s

    # Column check: no index repeats in the same slot across grids
    for s in range(cells):
        seen = {}
        for g, row in enumerate(matrix):
            idx = row[s]
            if idx in seen:
                errors.append(
                    f"Slot {s}: index {idx} appears in grids {seen[idx]+1} and {g+1}")
            else:
                seen[idx] = g

    return (len(errors) == 0), errors


# ═══════════════════════════════════════════════════════════════════════════════
#  STAGE 2 — PROGRESS BAR WINDOW
# ═══════════════════════════════════════════════════════════════════════════════

class ProgressWindow(tk.Toplevel):
    """
    Modal progress window shown during generation.
    Has three phases:
      Phase 1 — Building & verifying matrix  (indeterminate spinner)
      Phase 2 — Rendering images             (determinate bar)
      Phase 3 — Done / error
    """
    def __init__(self, parent, n_grids):
        super().__init__(parent)
        self.title("Generating…")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)
        self.grab_set()               # modal
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # block close

        self._n_grids  = n_grids
        self._canceled = False

        # ── Widgets ────────────────────────────────────────────────────────
        tk.Label(self, text="🌸  Generating Bingo Grids", font=FONT_HEAD,
                 bg=BG_DARK, fg=ACCENT).pack(padx=24, pady=(18, 4))

        self._phase_lbl = tk.Label(self, text="Phase 1 / 2 — Building matrix…",
                                   font=FONT_SMALL, bg=BG_DARK, fg=TEXT_DIM)
        self._phase_lbl.pack(padx=24, pady=(0, 6))

        self._detail_lbl = tk.Label(self, text="", font=FONT_MONO,
                                    bg=BG_DARK, fg=TEXT_MAIN)
        self._detail_lbl.pack(padx=24)

        # Outer frame for the bar
        bar_frame = tk.Frame(self, bg=BORDER, padx=1, pady=1)
        bar_frame.pack(padx=24, pady=10, fill="x")
        bar_bg = tk.Frame(bar_frame, bg=BG_MID, height=18)
        bar_bg.pack(fill="x")
        self._bar_fill = tk.Frame(bar_bg, bg=ACCENT, height=18, width=0)
        self._bar_fill.place(x=0, y=0, relheight=1.0)
        self._bar_bg   = bar_bg
        self._bar_w    = 340   # will update after pack

        self._pct_lbl = tk.Label(self, text="0 / 0", font=FONT_SMALL,
                                 bg=BG_DARK, fg=TEXT_DIM)
        self._pct_lbl.pack(padx=24, pady=(0, 16))

        self.update_idletasks()
        self._bar_w = bar_bg.winfo_width() or 340

        # Centre on parent
        self.update_idletasks()
        pw, ph = self.winfo_width(), self.winfo_height()
        px = parent.winfo_x() + (parent.winfo_width()  - pw) // 2
        py = parent.winfo_y() + (parent.winfo_height() - ph) // 2
        self.geometry(f"+{px}+{py}")

    # ── Public update methods (called from worker thread via after()) ──────

    def set_phase_matrix(self, attempt):
        self._phase_lbl.config(text="Phase 1 / 2 — Building & verifying matrix…")
        self._detail_lbl.config(text=f"Attempt #{attempt}…")
        self._set_bar_indeterminate()

    def set_phase_render(self, current, total):
        self._phase_lbl.config(text="Phase 2 / 2 — Rendering grids…")
        self._detail_lbl.config(text=f"Grid {current} of {total}")
        self._pct_lbl.config(text=f"{current} / {total}")
        pct = current / total if total else 0
        self._bar_w = self._bar_bg.winfo_width() or self._bar_w
        self._bar_fill.place(x=0, y=0, relheight=1.0,
                             width=int(self._bar_w * pct))

    def set_done(self, message):
        self._phase_lbl.config(text="✓  Done!", fg=ACCENT_LT)
        self._detail_lbl.config(text=message)
        self._bar_fill.place(x=0, y=0, relheight=1.0, width=self._bar_w)
        self._pct_lbl.config(text="Complete")
        self.after(1200, self.destroy)

    def set_error(self, message):
        self._phase_lbl.config(text="✗  Error", fg="#cc4444")
        self._detail_lbl.config(text=message, fg="#cc4444")
        # Add a close button
        btn_close = tk.Label(self, text="Close", font=FONT_BODY,
                             bg="#8b2020", fg=TEXT_MAIN, padx=12, pady=5,
                             cursor="hand2")
        btn_close.pack(pady=(4, 16))
        btn_close.bind("<Button-1>", lambda e: self.destroy())

    def _set_bar_indeterminate(self):
        # Animate a sliding block to indicate "working"
        self._ind_pos = getattr(self, "_ind_pos", 0)
        self._bar_w   = self._bar_bg.winfo_width() or self._bar_w
        block = max(40, self._bar_w // 4)
        x     = self._ind_pos % (self._bar_w + block) - block
        self._bar_fill.place(x=x, y=0, relheight=1.0, width=block)
        self._ind_pos += 8
        if self.winfo_exists():
            self._ind_job = self.after(30, self._set_bar_indeterminate)


# ═══════════════════════════════════════════════════════════════════════════════
#  Reusable widget helpers
# ═══════════════════════════════════════════════════════════════════════════════

def btn(parent, text, cmd, style="accent"):
    colors = {
        "accent": (ACCENT,    BG_DARK,   ACCENT_LT),
        "dim":    (BG_PANEL,  TEXT_DIM,  BG_CARD),
        "danger": ("#8b2020", TEXT_MAIN, "#a03030"),
    }
    bg, fg, hover = colors.get(style, colors["accent"])
    b = tk.Label(parent, text=text, font=FONT_BODY, bg=bg, fg=fg,
                 padx=12, pady=5, cursor="hand2", relief="flat")
    b.bind("<Button-1>", lambda e: cmd())
    b.bind("<Enter>",    lambda e: b.config(bg=hover))
    b.bind("<Leave>",    lambda e: b.config(bg=bg))
    return b

def sep(parent):
    return tk.Frame(parent, bg=BORDER, height=1)

def pil_to_tk(img, max_w, max_h):
    img = img.copy()
    img.thumbnail((max_w, max_h), Image.LANCZOS)
    return ImageTk.PhotoImage(img)

def hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def _fit_to_cell(img, cell_w_px, cell_h_px, target_inches=1.4, dpi=96):
    """
    Resize `img` so its longest side is exactly `target_inches` at `dpi`,
    then clamp it so it never exceeds the cell bounds.

    This makes every image the same physical size regardless of its original
    pixel dimensions — a tall image and a wide image both have their longest
    side = 1.4 inches.  The result is centred in the cell by the caller.

    dpi=96 is the standard screen/export DPI used by most design tools.
    Change to 150 or 300 if you need print-quality output.
    """
    target_px = int(target_inches * dpi)          # 1.4 * 96 = 134 px at 96 dpi

    iw, ih = img.size
    if iw == 0 or ih == 0:
        return img

    # Scale so the longest side == target_px
    scale = target_px / max(iw, ih)
    new_w = max(1, round(iw * scale))
    new_h = max(1, round(ih * scale))

    # Safety clamp: never exceed the available cell area
    if new_w > cell_w_px or new_h > cell_h_px:
        clamp = min(cell_w_px / new_w, cell_h_px / new_h)
        new_w = max(1, round(new_w * clamp))
        new_h = max(1, round(new_h * clamp))

    return img.resize((new_w, new_h), Image.LANCZOS)


# ═══════════════════════════════════════════════════════════════════════════════
#  Settings Window
# ═══════════════════════════════════════════════════════════════════════════════

class SettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.title("Settings")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)
        self._build()

    def _build(self):
        pad = dict(padx=14, pady=6)
        tk.Label(self, text="⚙  Settings", font=FONT_HEAD,
                 bg=BG_DARK, fg=ACCENT).pack(**pad, anchor="w")
        sep(self).pack(fill="x", padx=14)

        tk.Label(self, text="Grid template image (optional PNG overlay):",
                 font=FONT_SMALL, bg=BG_DARK, fg=TEXT_DIM).pack(**pad, anchor="w")
        row = tk.Frame(self, bg=BG_DARK)
        row.pack(fill="x", padx=14)
        self._grid_var = tk.StringVar(value=self.app.grid_template_path or "")
        tk.Entry(row, textvariable=self._grid_var, font=FONT_MONO,
                 bg=BG_MID, fg=TEXT_MAIN, insertbackground=ACCENT,
                 relief="flat", width=30).pack(side="left", padx=(0, 6))
        btn(row, "Browse", self._pick_grid).pack(side="left")
        sep(self).pack(fill="x", padx=14, pady=6)

        tk.Label(self, text="Layout  (all values in pixels):",
                 font=FONT_SMALL, bg=BG_DARK, fg=TEXT_DIM).pack(**pad, anchor="w")
        self._vars = {}
        fields = [
            ("Columns",                     "cols"),
            ("Rows  (excl. title bar)",     "rows"),
            ("Title bar height (px)",       "title_h"),
            ("Cell width (px)",             "cell_w"),
            ("Cell height (px)",            "cell_h"),
            ("Left / right margin (px)",    "margin_x"),
            ("Top margin after title (px)", "margin_y"),
            ("Cell image padding (px)",     "cell_pad"),
        ]
        for label, key in fields:
            v = tk.StringVar(value=str(getattr(self.app, key)))
            self._vars[key] = v
            r = tk.Frame(self, bg=BG_DARK)
            r.pack(fill="x", padx=14, pady=2)
            tk.Label(r, text=label, font=FONT_SMALL, bg=BG_DARK,
                     fg=TEXT_MAIN, width=30, anchor="w").pack(side="left")
            tk.Entry(r, textvariable=v, font=FONT_MONO, bg=BG_MID, fg=TEXT_MAIN,
                     insertbackground=ACCENT, relief="flat", width=8).pack(side="left")
        sep(self).pack(fill="x", padx=14, pady=6)

        tk.Label(self, text="Grid lines:", font=FONT_SMALL,
                 bg=BG_DARK, fg=TEXT_DIM).pack(**pad, anchor="w")
        lw_row = tk.Frame(self, bg=BG_DARK)
        lw_row.pack(fill="x", padx=14, pady=2)
        tk.Label(lw_row, text="Line width (px)", font=FONT_SMALL, bg=BG_DARK,
                 fg=TEXT_MAIN, width=30, anchor="w").pack(side="left")
        self._lw_var = tk.IntVar(value=self.app.grid_line_width)
        tk.Spinbox(lw_row, from_=1, to=20, textvariable=self._lw_var,
                   font=FONT_MONO, bg=BG_MID, fg=TEXT_MAIN, width=6,
                   insertbackground=ACCENT, relief="flat",
                   buttonbackground=BG_CARD).pack(side="left")

        lc_row = tk.Frame(self, bg=BG_DARK)
        lc_row.pack(fill="x", padx=14, pady=4)
        tk.Label(lc_row, text="Line color", font=FONT_SMALL, bg=BG_DARK,
                 fg=TEXT_MAIN, width=30, anchor="w").pack(side="left")
        self._lc_var = tk.StringVar(value=self.app.grid_line_color)
        self._lc_swatch = tk.Label(lc_row, bg=self.app.grid_line_color,
                                    width=4, relief="solid", cursor="hand2")
        self._lc_swatch.pack(side="left", padx=(0, 6))
        self._lc_swatch.bind("<Button-1>", lambda e: self._pick_line_color())
        tk.Label(lc_row, textvariable=self._lc_var, font=FONT_MONO,
                 bg=BG_DARK, fg=TEXT_DIM).pack(side="left")
        sep(self).pack(fill="x", padx=14, pady=6)

        tk.Label(self, text="Board background color:", font=FONT_SMALL,
                 bg=BG_DARK, fg=TEXT_DIM).pack(**pad, anchor="w")
        bc_row = tk.Frame(self, bg=BG_DARK)
        bc_row.pack(fill="x", padx=14, pady=4)
        tk.Label(bc_row, text="Background color", font=FONT_SMALL, bg=BG_DARK,
                 fg=TEXT_MAIN, width=30, anchor="w").pack(side="left")
        self._bc_var = tk.StringVar(value=self.app.board_bg_color)
        self._bc_swatch = tk.Label(bc_row, bg=self.app.board_bg_color,
                                    width=4, relief="solid", cursor="hand2")
        self._bc_swatch.pack(side="left", padx=(0, 6))
        self._bc_swatch.bind("<Button-1>", lambda e: self._pick_bg_color())
        tk.Label(bc_row, textvariable=self._bc_var, font=FONT_MONO,
                 bg=BG_DARK, fg=TEXT_DIM).pack(side="left")
        sep(self).pack(fill="x", padx=14, pady=6)

        tk.Label(self, text="Background image alpha  (0 = hidden  ·  255 = solid):",
                 font=FONT_SMALL, bg=BG_DARK, fg=TEXT_DIM).pack(**pad, anchor="w")
        alpha_row = tk.Frame(self, bg=BG_DARK)
        alpha_row.pack(fill="x", padx=14, pady=4)
        self._alpha_var = tk.IntVar(value=self.app.bg_alpha)
        tk.Scale(alpha_row, from_=0, to=255, orient="horizontal",
                 variable=self._alpha_var, length=220,
                 bg=BG_DARK, fg=TEXT_MAIN, troughcolor=BG_MID,
                 highlightthickness=0, relief="flat").pack(side="left")
        tk.Label(alpha_row, textvariable=self._alpha_var, font=FONT_MONO,
                 bg=BG_DARK, fg=ACCENT, width=4).pack(side="left", padx=6)
        sep(self).pack(fill="x", padx=14, pady=8)

        bf = tk.Frame(self, bg=BG_DARK)
        bf.pack(pady=10)
        btn(bf, "Save & Close", self._save).pack(side="left", padx=6)
        btn(bf, "Cancel", self.destroy, style="dim").pack(side="left", padx=6)

    def _pick_grid(self):
        p = filedialog.askopenfilename(title="Select grid template overlay",
                                       filetypes=[("PNG", "*.png"), ("All", "*.*")])
        if p: self._grid_var.set(p)

    def _pick_line_color(self):
        result = colorchooser.askcolor(color=self._lc_var.get(),
                                       title="Grid line color", parent=self)
        if result and result[1]:
            self._lc_var.set(result[1])
            self._lc_swatch.config(bg=result[1])

    def _pick_bg_color(self):
        result = colorchooser.askcolor(color=self._bc_var.get(),
                                       title="Board background color", parent=self)
        if result and result[1]:
            self._bc_var.set(result[1])
            self._bc_swatch.config(bg=result[1])

    def _save(self):
        p = self._grid_var.get().strip()
        if p and not os.path.isfile(p):
            messagebox.showerror("Error", "Grid template file not found.", parent=self)
            return
        self.app.grid_template_path = p or None
        for key, v in self._vars.items():
            try: setattr(self.app, key, int(v.get()))
            except ValueError: pass
        self.app.grid_line_width = int(self._lw_var.get())
        self.app.grid_line_color = self._lc_var.get()
        self.app.board_bg_color  = self._bc_var.get()
        self.app.bg_alpha        = int(self._alpha_var.get())
        self.app.refresh_preview()
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  Main Application
# ═══════════════════════════════════════════════════════════════════════════════

class BingoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bingo Grid Generator")
        self.configure(bg=BG_DARK)
        self.resizable(True, True)

        self.grid_template_path = None
        self.title_image_path   = None
        self.bg_image_path      = None
        self.item_paths         = []

        self.cols     = 4
        self.rows     = 4
        self.title_h  = 160
        self.cell_w   = 180
        self.cell_h   = 180
        self.margin_x = 0
        self.margin_y = 0
        self.cell_pad = 8

        self.grid_line_width = 3
        self.grid_line_color = "#000000"
        self.board_bg_color  = "#ffffff"
        self.bg_alpha        = 200

        self._preview_tk = None
        self._last_grid  = None

        self._build_ui()
        self.refresh_preview()

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self):
        top = tk.Frame(self, bg=BG_DARK, pady=8)
        top.pack(fill="x", padx=16)
        tk.Label(top, text="🌸  Bingo Grid Generator", font=FONT_BIG,
                 bg=BG_DARK, fg=ACCENT).pack(side="left")
        btn(top, "⚙  Settings", self._open_settings, "dim").pack(side="right", padx=4)
        sep(self).pack(fill="x", padx=16)

        body = tk.Frame(self, bg=BG_DARK)
        body.pack(fill="both", expand=True, padx=16, pady=10)

        left = tk.Frame(body, bg=BG_PANEL, padx=12, pady=12)
        left.pack(side="left", fill="y", padx=(0, 10))
        self._build_left(left)

        center = tk.Frame(body, bg=BG_MID)
        center.pack(side="left", fill="both", expand=True, padx=(0, 10))
        self._build_center(center)

        right = tk.Frame(body, bg=BG_PANEL, padx=12, pady=12)
        right.pack(side="right", fill="y")
        self._build_right(right)

    def _build_left(self, parent):
        tk.Label(parent, text="ACTIONS", font=FONT_SMALL,
                 bg=BG_PANEL, fg=TEXT_DIM).pack(anchor="w", pady=(0, 8))
        btn(parent, "🎲  Generate Grids", self._generate).pack(fill="x", pady=4)
        btn(parent, "🖼  Change Background", self._pick_bg, "dim").pack(fill="x", pady=4)
        sep(parent).pack(fill="x", pady=10)

        tk.Label(parent, text="TITLE IMAGE", font=FONT_SMALL,
                 bg=BG_PANEL, fg=TEXT_DIM).pack(anchor="w", pady=(0, 4))
        tk.Label(parent, text="Transparent PNG placed\nin the top title bar:",
                 font=FONT_SMALL, bg=BG_PANEL, fg=TEXT_MAIN, justify="left").pack(anchor="w")
        btn(parent, "📝  Select Title Image", self._pick_title, "dim").pack(fill="x", pady=(6, 2))
        self._title_lbl = tk.Label(parent, text="None selected", font=FONT_SMALL,
                                   bg=BG_PANEL, fg=TEXT_DIM, wraplength=160, justify="left")
        self._title_lbl.pack(anchor="w")
        sep(parent).pack(fill="x", pady=10)

        tk.Label(parent, text="GRIDS TO GENERATE", font=FONT_SMALL,
                 bg=BG_PANEL, fg=TEXT_DIM).pack(anchor="w", pady=(0, 4))
        spin_row = tk.Frame(parent, bg=BG_PANEL)
        spin_row.pack(anchor="w")
        btn(spin_row, "−", self._dec_count, "dim").pack(side="left")
        self._count_var = tk.IntVar(value=1)
        tk.Label(spin_row, textvariable=self._count_var, font=FONT_HEAD,
                 bg=BG_PANEL, fg=TEXT_MAIN, width=4).pack(side="left")
        btn(spin_row, "+", self._inc_count, "dim").pack(side="left")
        sep(parent).pack(fill="x", pady=10)

        self._pool_info = tk.Label(parent, text="0 images in pool\n0 cells per grid",
                                   font=FONT_SMALL, bg=BG_PANEL, fg=TEXT_MAIN, justify="left")
        self._pool_info.pack(anchor="w")

    def _build_center(self, parent):
        tk.Label(parent, text="PREVIEW", font=FONT_SMALL,
                 bg=BG_MID, fg=TEXT_DIM).pack(pady=(8, 0))
        self._canvas = tk.Canvas(parent, width=PREVIEW_W, height=PREVIEW_H,
                                 bg=BG_DARK, highlightthickness=0)
        self._canvas.pack(padx=10, pady=8)

    def _build_right(self, parent):
        tk.Label(parent, text="IMAGE POOL", font=FONT_SMALL,
                 bg=BG_PANEL, fg=TEXT_DIM).pack(anchor="w", pady=(0, 8))
        tk.Label(parent, text="Images that fill the grid cells:",
                 font=FONT_SMALL, bg=BG_PANEL, fg=TEXT_MAIN, justify="left").pack(anchor="w")
        br = tk.Frame(parent, bg=BG_PANEL)
        br.pack(fill="x", pady=(6, 2))
        btn(br, "+ Add", self._add_images).pack(side="left")
        btn(br, "✕ Remove", self._remove_selected, "danger").pack(side="left", padx=(4, 0))
        frame = tk.Frame(parent, bg=BG_DARK)
        frame.pack(fill="both", expand=True, pady=(4, 0))
        sb = tk.Scrollbar(frame, orient="vertical", bg=BG_MID)
        self._listbox = tk.Listbox(frame, font=FONT_SMALL, bg=BG_DARK, fg=TEXT_MAIN,
                                   selectbackground=ACCENT, selectforeground=BG_DARK,
                                   relief="flat", bd=0, width=24, height=20,
                                   yscrollcommand=sb.set)
        sb.config(command=self._listbox.yview)
        self._listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        sep(parent).pack(fill="x", pady=8)
        btn(parent, "Clear All", self._clear_pool, "danger").pack(fill="x")

    # ── Actions ────────────────────────────────────────────────────────────
    def _open_settings(self): SettingsWindow(self)

    def _pick_bg(self):
        p = filedialog.askopenfilename(
            title="Select background image",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.webp"), ("All", "*.*")])
        if p:
            self.bg_image_path = p
            self.refresh_preview()

    def _pick_title(self):
        p = filedialog.askopenfilename(
            title="Select title image (transparent PNG)",
            filetypes=[("PNG", "*.png"), ("All", "*.*")])
        if p:
            self.title_image_path = p
            self._title_lbl.config(text=Path(p).name)
            self.refresh_preview()

    def _add_images(self):
        paths = filedialog.askopenfilenames(
            title="Select cell images",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.webp"), ("All", "*.*")])
        for p in paths:
            if p not in self.item_paths:
                self.item_paths.append(p)
                self._listbox.insert("end", Path(p).name)
        self._update_pool_info()
        self.refresh_preview()

    def _remove_selected(self):
        for i in reversed(list(self._listbox.curselection())):
            self._listbox.delete(i)
            self.item_paths.pop(i)
        self._update_pool_info()
        self.refresh_preview()

    def _clear_pool(self):
        if messagebox.askyesno("Clear pool", "Remove all images from the pool?"):
            self.item_paths.clear()
            self._listbox.delete(0, "end")
            self._update_pool_info()
            self.refresh_preview()

    def _dec_count(self):
        v = self._count_var.get()
        if v > 1: self._count_var.set(v - 1)

    def _inc_count(self):
        self._count_var.set(self._count_var.get() + 1)

    def _update_pool_info(self):
        n     = len(self.item_paths)
        cells = self.cols * self.rows
        self._pool_info.config(
            text=f"{n} image{'s' if n != 1 else ''} in pool\n{cells} cells per grid")

    # ── Grid rendering ─────────────────────────────────────────────────────
    def _draw_grid_lines(self, draw, out_w, out_h):
        lw    = self.grid_line_width
        color = self.grid_line_color
        x0    = self.margin_x
        y0    = self.title_h + self.margin_y
        x1    = out_w - self.margin_x
        y1    = out_h - self.margin_y
        draw.rectangle([0, 0, out_w-1, out_h-1], outline=color, width=lw)
        draw.line([(0, self.title_h), (out_w, self.title_h)], fill=color, width=lw)
        draw.rectangle([x0, y0, x1, y1], outline=color, width=lw)
        for c in range(1, self.cols):
            lx = x0 + c * self.cell_w
            draw.line([(lx, y0), (lx, y1)], fill=color, width=lw)
        for r in range(1, self.rows):
            ly = y0 + r * self.cell_h
            draw.line([(x0, ly), (x1, ly)], fill=color, width=lw)

    def _compose_grid(self, index_row, pool):
        """Render one grid card from a row of pool indices."""
        out_w = self.margin_x * 2 + self.cols * self.cell_w
        out_h = self.title_h + self.margin_y + self.rows * self.cell_h + self.margin_y
        r, g, b = hex_to_rgb(self.board_bg_color)
        card = Image.new("RGBA", (out_w, out_h), (r, g, b, 255))

        if self.bg_image_path and os.path.isfile(self.bg_image_path):
            bg_img = Image.open(self.bg_image_path).convert("RGBA")
            bg_img = bg_img.resize((out_w, out_h), Image.LANCZOS)
            rc, gc, bc, ac = bg_img.split()
            ac = ac.point(lambda x: min(x, self.bg_alpha))
            bg_img.putalpha(ac)
            card.paste(bg_img, (0, 0), bg_img)

        if self.grid_template_path and os.path.isfile(self.grid_template_path):
            overlay = Image.open(self.grid_template_path).convert("RGBA")
            overlay = overlay.resize((out_w, out_h), Image.LANCZOS)
            card.paste(overlay, (0, 0), overlay)

        for slot, idx in enumerate(index_row):
            path = pool[idx]
            ri   = slot // self.cols
            ci   = slot  % self.cols
            x    = self.margin_x + ci * self.cell_w + self.cell_pad
            y    = self.title_h  + self.margin_y + ri * self.cell_h + self.cell_pad
            cw   = self.cell_w - self.cell_pad * 2
            ch   = self.cell_h - self.cell_pad * 2
            cell = Image.open(path).convert("RGBA")
            cell = _fit_to_cell(cell, cw, ch)
            cx   = x + (cw - cell.width)  // 2
            cy   = y + (ch - cell.height) // 2
            card.paste(cell, (cx, cy), cell)

        draw = ImageDraw.Draw(card)
        self._draw_grid_lines(draw, out_w, out_h)

        if self.title_image_path and os.path.isfile(self.title_image_path):
            title_img = Image.open(self.title_image_path).convert("RGBA")
            title_img.thumbnail((out_w - 20, max(1, self.title_h - 10)), Image.LANCZOS)
            tx = (out_w - title_img.width)  // 2
            ty = (self.title_h - title_img.height) // 2
            card.paste(title_img, (tx, ty), title_img)

        return card

    # ── Generate pipeline ──────────────────────────────────────────────────
    def _generate(self):
        if not PIL_AVAILABLE:
            messagebox.showerror("Missing library", "Run: pip install Pillow")
            return

        cells = self.cols * self.rows
        n     = self._count_var.get()

        if not self.item_paths:
            messagebox.showerror("Error", "Add images to the pool first.")
            return

        if len(self.item_paths) < cells:
            go = messagebox.askyesno(
                "Not enough images",
                f"You have {len(self.item_paths)} images but each grid needs {cells}.\n"
                f"Images will repeat within grids. Generate anyway?")
            if not go: return

        out_dir = filedialog.askdirectory(title="Select output folder")
        if not out_dir: return

        # Snapshot pool order (indices refer to this list)
        pool = self.item_paths[:]

        # Open progress window and run worker in background thread
        progress = ProgressWindow(self, n)
        self._run_worker(progress, pool, cells, n, out_dir)

    def _run_worker(self, progress, pool, cells, n_grids, out_dir):
        """Launch the generation worker in a background thread."""

        def worker():
            try:
                # ── Phase 1: build & verify matrix ────────────────────────
                attempt = [0]

                def on_attempt(a):
                    attempt[0] = a
                    if progress.winfo_exists():
                        progress.after(0, lambda: progress.set_phase_matrix(a))

                matrix, total_attempts = _build_matrix_with_callback(
                    len(pool), cells, n_grids, on_attempt)

                print(f"✓ Matrix verified after {total_attempts} attempt(s).")

                # ── Phase 2: render images ─────────────────────────────────
                last_card = None
                for i, index_row in enumerate(matrix):
                    if not progress.winfo_exists():
                        return   # window was closed, abort
                    # Update progress bar on main thread
                    progress.after(0, lambda i=i: progress.set_phase_render(i+1, n_grids))
                    card  = self._compose_grid(index_row, pool)
                    fname = os.path.join(out_dir, f"bingo_grid_{i+1:03d}.png")
                    card.save(fname, "PNG")
                    last_card = card

                # ── Done ──────────────────────────────────────────────────
                if last_card:
                    self._last_grid = last_card

                def finish():
                    if progress.winfo_exists():
                        progress.set_done(f"Saved {n_grids} grid(s) to {out_dir}")
                    self.refresh_preview(use_last=True)

                self.after(0, finish)

            except Exception as e:
                def show_err(msg=str(e)):
                    if progress.winfo_exists():
                        progress.set_error(msg)
                self.after(0, show_err)

        threading.Thread(target=worker, daemon=True).start()

    # ── Preview ────────────────────────────────────────────────────────────
    def refresh_preview(self, use_last=False):
        if not PIL_AVAILABLE:
            self._canvas.delete("all")
            self._canvas.create_text(PREVIEW_W//2, PREVIEW_H//2,
                                     text="Pillow not installed\npip install Pillow",
                                     fill=TEXT_DIM, font=FONT_BODY, justify="center")
            return
        self._canvas.delete("all")
        img = self._last_grid if (use_last and self._last_grid) else self._build_preview_image()
        self._preview_tk = pil_to_tk(img, PREVIEW_W, PREVIEW_H)
        self._canvas.config(width=self._preview_tk.width(), height=self._preview_tk.height())
        self._canvas.create_image(0, 0, anchor="nw", image=self._preview_tk)
        self._update_pool_info()

    def _build_preview_image(self):
        cells = self.cols * self.rows
        out_w = max(1, self.margin_x * 2 + self.cols * self.cell_w)
        out_h = max(1, self.title_h + self.margin_y + self.rows * self.cell_h + self.margin_y)
        r, g, b = hex_to_rgb(self.board_bg_color)
        card = Image.new("RGBA", (out_w, out_h), (r, g, b, 255))

        if self.bg_image_path and os.path.isfile(self.bg_image_path):
            bg_img = Image.open(self.bg_image_path).convert("RGBA")
            bg_img = bg_img.resize((out_w, out_h), Image.LANCZOS)
            rc, gc, bc, ac = bg_img.split()
            ac = ac.point(lambda x: min(x, self.bg_alpha))
            bg_img.putalpha(ac)
            card.paste(bg_img, (0, 0), bg_img)

        if self.grid_template_path and os.path.isfile(self.grid_template_path):
            overlay = Image.open(self.grid_template_path).convert("RGBA")
            overlay = overlay.resize((out_w, out_h), Image.LANCZOS)
            card.paste(overlay, (0, 0), overlay)

        if self.item_paths:
            pool   = self.item_paths[:]
            sample = (pool * (cells // len(pool) + 1))[:cells]
            random.shuffle(sample)
            for idx, path in enumerate(sample):
                ri = idx // self.cols
                ci = idx  % self.cols
                x  = self.margin_x + ci * self.cell_w + self.cell_pad
                y  = self.title_h  + self.margin_y + ri * self.cell_h + self.cell_pad
                cw = self.cell_w - self.cell_pad * 2
                ch = self.cell_h - self.cell_pad * 2
                try:
                    cell = Image.open(path).convert("RGBA")
                    cell = _fit_to_cell(cell, cw, ch)
                    cx = x + (cw - cell.width)  // 2
                    cy = y + (ch - cell.height) // 2
                    card.paste(cell, (cx, cy), cell)
                except Exception:
                    pass

        draw = ImageDraw.Draw(card)
        self._draw_grid_lines(draw, out_w, out_h)

        if self.title_image_path and os.path.isfile(self.title_image_path):
            try:
                title_img = Image.open(self.title_image_path).convert("RGBA")
                title_img.thumbnail((out_w - 20, max(1, self.title_h - 10)), Image.LANCZOS)
                tx = (out_w - title_img.width)  // 2
                ty = (self.title_h - title_img.height) // 2
                card.paste(title_img, (tx, ty), title_img)
            except Exception:
                pass

        return card


# ═══════════════════════════════════════════════════════════════════════════════
#  Matrix builder with callback (called from worker thread)
# ═══════════════════════════════════════════════════════════════════════════════

def _build_matrix_with_callback(pool_size, cells, n_grids, on_attempt, max_attempts=500):
    """
    Build + verify the integer matrix. Calls on_attempt(n) each retry.
    Returns (matrix, attempt_count).
    """
    for attempt in range(1, max_attempts + 1):
        on_attempt(attempt)
        matrix = _try_build(pool_size, cells, n_grids)
        if matrix is None:
            continue
        ok, errors = verify_matrix(matrix, cells, n_grids)
        if ok:
            return matrix, attempt
        # Verification failed — print why and retry
        print(f"  Attempt {attempt} failed verification ({len(errors)} errors). Retrying…")

    raise RuntimeError(
        f"Could not build a valid matrix after {max_attempts} attempts.\n"
        f"Pool size: {pool_size}, cells: {cells}, grids: {n_grids}.\n"
        f"Try reducing the number of grids or adding more images.")


# ═══════════════════════════════════════════════════════════════════════════════
#  Entry
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    if not PIL_AVAILABLE:
        print("ERROR: Pillow not installed. Run: pip install Pillow")
    app = BingoApp()
    app.mainloop()

if __name__ == "__main__":
    main()
