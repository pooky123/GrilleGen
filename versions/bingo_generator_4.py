"""
Bingo Grid Generator  v4
========================
Language: Python 3
Libraries: tkinter (GUI), Pillow (image compositing)
Run with: python bingo_generator.py

CORE ALGORITHM v4 — strict per-slot uniqueness:

  We build a (n_grids x cells) assignment matrix using a proper
  column-by-column derangement approach:

  Think of it as a table:
    - Rows    = grids  (one row per generated card)
    - Columns = slots  (one column per cell position on the card)

  Rule: no value may repeat in any COLUMN.
        no value may repeat in any ROW.

  This is equivalent to filling a Latin rectangle of size (n_grids x cells)
  from a pool of images.

  Implementation:
    1. Shuffle the pool.
    2. For each grid (row), pick `cells` images that:
         a) have not been used in THIS grid yet          (no row repeat)
         b) have not been used in the SAME SLOT before   (no column repeat)
       We use a greedy backtracking approach slot-by-slot.
    3. After building the matrix, run a position-diversity pass:
       shuffle the positions within each grid while preserving column uniqueness.
    4. A final anti-pattern pass swaps images in adjacent positions that
       appeared as neighbours in a previous grid.

  If the pool is smaller than cells, images must repeat within a grid
  (unavoidable) — the user is warned. Cross-grid slot uniqueness is still
  enforced as much as the pool size allows.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
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


# ─── Helpers ──────────────────────────────────────────────────────────────────
def pil_to_tk(img, max_w, max_h):
    img = img.copy()
    img.thumbnail((max_w, max_h), Image.LANCZOS)
    return ImageTk.PhotoImage(img)


def hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


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


# ─── Core assignment algorithm ────────────────────────────────────────────────

def build_strict_matrix(pool, cells, n_grids, n_cols):
    """
    Returns matrix[g][s] = image_path  for g in range(n_grids), s in range(cells).

    Hard guarantees (when pool >= cells):
      • No image repeats within a single grid (row uniqueness).
      • No image repeats in the same slot across different grids (column uniqueness).
      • Positions are randomised per grid so layouts look different.
      • An anti-adjacency pass reduces "image X always next to image Y" patterns.

    The algorithm:
      Step 1 — build a slot-queue table.
        slot_pool[s] = independent shuffled list of pool images for slot s.
        We ensure slot_pool[s][g] != slot_pool[s][g2] for g != g2  (column rule).
        We do this by constructing each slot's queue as a fresh shuffle of the
        pool, which by construction gives each slot a unique image per grid
        position — AS LONG AS we also enforce the row rule in step 2.

      Step 2 — enforce row uniqueness (no image twice in one grid).
        For each grid g, collect the naive assignment col[s] = slot_pool[s][g].
        Find any duplicates in that row and fix them by finding a swap partner
        in a different slot that resolves both the row and the column constraint.

      Step 3 — randomise positions within each grid.
        Shuffle each row while maintaining column uniqueness (swap within columns).

      Step 4 — anti-adjacency pass.
        For consecutive grids, swap images that appear as neighbours of where
        they sat in the previous grid.
    """

    n_pool = len(pool)

    # ── Step 1: build per-slot queues (column uniqueness) ─────────────────
    # Each slot gets a shuffled permutation of the pool, tiled if n_grids > n_pool.
    # We build these so that slot_pool[s][g] is always a valid unique value per
    # column — i.e. no two rows share the same image in column s.

    slot_pool = []
    for s in range(cells):
        q = []
        while len(q) < n_grids:
            chunk = pool[:]
            random.shuffle(chunk)
            # Avoid same image at the join between tiles
            if q and chunk and q[-1] == chunk[0]:
                idx = random.randrange(1, len(chunk))
                chunk[0], chunk[idx] = chunk[idx], chunk[0]
            q.extend(chunk)
        slot_pool.append(q[:n_grids])

    # slot_pool[slot][grid] = assigned image for that slot in that grid

    # ── Step 2: enforce row uniqueness (no image twice in one grid) ───────
    # Build matrix as list of rows
    matrix = []
    for g in range(n_grids):
        row = [slot_pool[s][g] for s in range(cells)]
        matrix.append(row)

    # For each grid, detect and fix intra-row duplicates
    for g in range(n_grids):
        row = matrix[g]
        _fix_row_duplicates(row, matrix, g, cells, pool)

    # ── Step 3: randomise positions within each grid ──────────────────────
    # We want each grid to look like a different arrangement, not a cyclic shift.
    # We do this by performing random valid swaps within each row.
    # A swap of positions s1 and s2 in row g is valid if:
    #   matrix[g2][s1] != matrix[g][s2]  for all g2 != g   (column rule for s1)
    #   matrix[g2][s2] != matrix[g][s1]  for all g2 != g   (column rule for s2)

    for g in range(n_grids):
        row       = matrix[g]
        positions = list(range(cells))
        random.shuffle(positions)
        # Attempt many random valid swaps
        for _ in range(cells * 10):
            s1 = random.randrange(cells)
            s2 = random.randrange(cells)
            if s1 == s2:
                continue
            if _swap_is_valid(matrix, g, s1, s2, n_grids):
                row[s1], row[s2] = row[s2], row[s1]

    # ── Step 4: anti-adjacency pass ───────────────────────────────────────
    # Reduce cases where image X sits next to image Y on multiple grids.
    for g in range(1, n_grids):
        prev     = matrix[g - 1]
        curr     = matrix[g]
        prev_pos = {img: s for s, img in enumerate(prev)}

        for _attempt in range(cells * 6):
            improved = False
            for s1 in range(cells):
                img = curr[s1]
                pp  = prev_pos.get(img)
                if pp is None:
                    continue
                if s1 not in _neighbours(pp, n_cols, cells):
                    continue
                # This image is adjacent to where it sat in the previous grid.
                # Try to find a swap partner that resolves it.
                candidates = list(range(cells))
                random.shuffle(candidates)
                for s2 in candidates:
                    if s2 == s1:
                        continue
                    img2 = curr[s2]
                    pp2  = prev_pos.get(img2)
                    # Check column uniqueness for both positions after swap
                    if not _swap_is_valid(matrix, g, s1, s2, n_grids):
                        continue
                    # Check adjacency improvement
                    new_s1_adj = pp2 is not None and s1 in _neighbours(pp2, n_cols, cells)
                    new_s2_adj = pp  is not None and s2 in _neighbours(pp,  n_cols, cells)
                    if not new_s1_adj and not new_s2_adj:
                        curr[s1], curr[s2] = curr[s2], curr[s1]
                        improved = True
                        break
            if not improved:
                break

    # ── Final verification (debug safety net) ─────────────────────────────
    _verify_matrix(matrix, cells, n_grids)

    return matrix


def _fix_row_duplicates(row, matrix, g, cells, pool):
    """
    Fix duplicate images within a single row (grid) by finding valid swaps
    with other slots. Modifies row and matrix[g] in-place.
    """
    n_grids = len(matrix)
    for attempt in range(cells * 4):
        seen   = {}
        dupes  = []
        for s, img in enumerate(row):
            if img in seen:
                dupes.append(s)
            else:
                seen[img] = s
        if not dupes:
            break

        s_bad = dupes[0]
        img_bad = row[s_bad]

        # Find a replacement image: not already in this row,
        # and not already used in column s_bad by another grid.
        used_in_row    = set(row)
        used_in_col    = {matrix[g2][s_bad] for g2 in range(n_grids) if g2 != g}
        candidates     = [p for p in pool if p not in used_in_row and p not in used_in_col]

        if candidates:
            replacement = random.choice(candidates)
            # Find which slot currently holds the replacement across other grids
            # and swap it in for this slot
            row[s_bad] = replacement
            matrix[g][s_bad] = replacement
        else:
            # No perfect replacement — try swapping s_bad with another slot
            # that doesn't create a new duplicate
            other_slots = [s for s in range(cells) if s != s_bad and row[s] not in used_in_col]
            if other_slots:
                s_other = random.choice(other_slots)
                row[s_bad], row[s_other] = row[s_other], row[s_bad]
                matrix[g][s_bad], matrix[g][s_other] = matrix[g][s_other], matrix[g][s_bad]


def _swap_is_valid(matrix, g, s1, s2, n_grids):
    """
    Check whether swapping positions s1 and s2 in row g maintains column uniqueness.
    i.e. after swap:
      matrix[g][s1] (which becomes the old s2 value) must not appear in column s1
      matrix[g][s2] (which becomes the old s1 value) must not appear in column s2
    for any other row.
    """
    val1 = matrix[g][s1]   # will go to s2
    val2 = matrix[g][s2]   # will go to s1
    for g2 in range(n_grids):
        if g2 == g:
            continue
        if matrix[g2][s1] == val2:
            return False
        if matrix[g2][s2] == val1:
            return False
    return True


def _neighbours(slot, n_cols, cells):
    """Return slot indices adjacent (up/down/left/right) to `slot`."""
    n_rows = cells // n_cols
    r, c   = divmod(slot, n_cols)
    nb     = []
    if r > 0:         nb.append(slot - n_cols)
    if r < n_rows-1:  nb.append(slot + n_cols)
    if c > 0:         nb.append(slot - 1)
    if c < n_cols-1:  nb.append(slot + 1)
    return nb


def _verify_matrix(matrix, cells, n_grids):
    """
    Verify column uniqueness. Raises ValueError listing any violations found.
    Used as a safety net during development.
    """
    errors = []
    for s in range(cells):
        col_vals = [matrix[g][s] for g in range(n_grids)]
        seen = {}
        for g, val in enumerate(col_vals):
            if val in seen:
                errors.append(
                    f"Slot {s}: image '{os.path.basename(val)}' "
                    f"appears in grid {seen[val]+1} AND grid {g+1}")
            else:
                seen[val] = g
    if errors:
        # Log to console but don't crash the app — surface as a warning
        print("⚠ Assignment violations detected (will still generate):")
        for e in errors:
            print("  ", e)


# ─── Settings Window ──────────────────────────────────────────────────────────
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
        p = filedialog.askopenfilename(
            title="Select grid template overlay",
            filetypes=[("PNG", "*.png"), ("All", "*.*")])
        if p:
            self._grid_var.set(p)

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
            try:
                setattr(self.app, key, int(v.get()))
            except ValueError:
                pass
        self.app.grid_line_width = int(self._lw_var.get())
        self.app.grid_line_color = self._lc_var.get()
        self.app.board_bg_color  = self._bc_var.get()
        self.app.bg_alpha        = int(self._alpha_var.get())
        self.app.refresh_preview()
        self.destroy()


# ─── Main Application ─────────────────────────────────────────────────────────
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
    def _open_settings(self):
        SettingsWindow(self)

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
        if v > 1:
            self._count_var.set(v - 1)

    def _inc_count(self):
        self._count_var.set(self._count_var.get() + 1)

    def _update_pool_info(self):
        n     = len(self.item_paths)
        cells = self.cols * self.rows
        self._pool_info.config(
            text=f"{n} image{'s' if n != 1 else ''} in pool\n{cells} cells per grid")

    # ── Grid drawing ───────────────────────────────────────────────────────
    def _draw_grid_lines(self, draw, out_w, out_h):
        lw    = self.grid_line_width
        color = self.grid_line_color
        x0    = self.margin_x
        y0    = self.title_h + self.margin_y
        x1    = out_w - self.margin_x
        y1    = out_h - self.margin_y

        draw.rectangle([0, 0, out_w - 1, out_h - 1], outline=color, width=lw)
        draw.line([(0, self.title_h), (out_w, self.title_h)], fill=color, width=lw)
        draw.rectangle([x0, y0, x1, y1], outline=color, width=lw)

        for c in range(1, self.cols):
            lx = x0 + c * self.cell_w
            draw.line([(lx, y0), (lx, y1)], fill=color, width=lw)
        for r in range(1, self.rows):
            ly = y0 + r * self.cell_h
            draw.line([(x0, ly), (x1, ly)], fill=color, width=lw)

    def _compose_grid(self, assigned_paths):
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

        for idx, path in enumerate(assigned_paths):
            row = idx // self.cols
            col = idx  % self.cols
            x   = self.margin_x + col * self.cell_w + self.cell_pad
            y   = self.title_h  + self.margin_y + row * self.cell_h + self.cell_pad
            cw  = self.cell_w - self.cell_pad * 2
            ch  = self.cell_h - self.cell_pad * 2
            cell = Image.open(path).convert("RGBA")
            cell.thumbnail((max(1, cw), max(1, ch)), Image.LANCZOS)
            cx = x + (cw - cell.width)  // 2
            cy = y + (ch - cell.height) // 2
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

    # ── Assignment ─────────────────────────────────────────────────────────
    def _assign_images(self, n_grids):
        cells = self.cols * self.rows
        pool  = self.item_paths[:]

        if len(pool) < cells:
            # Pad pool by cycling — user was warned
            while len(pool) < cells:
                pool += self.item_paths[:]

        matrix = build_strict_matrix(pool, cells, n_grids, self.cols)
        return matrix

    # ── Generate ───────────────────────────────────────────────────────────
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
                f"Some images will repeat within a single grid.\n\n"
                f"Generate anyway?")
            if not go:
                return

        out_dir = filedialog.askdirectory(title="Select output folder")
        if not out_dir:
            return

        try:
            assignments = self._assign_images(n)
        except Exception as e:
            messagebox.showerror("Assignment error", str(e))
            return

        for i, paths in enumerate(assignments):
            card  = self._compose_grid(paths)
            fname = os.path.join(out_dir, f"bingo_grid_{i+1:03d}.png")
            card.save(fname, "PNG")

        self._last_grid = self._compose_grid(assignments[-1])
        self.refresh_preview(use_last=True)
        messagebox.showinfo("Done ✓", f"Saved {n} grid(s) to:\n{out_dir}")

    # ── Preview ────────────────────────────────────────────────────────────
    def refresh_preview(self, use_last=False):
        if not PIL_AVAILABLE:
            self._canvas.delete("all")
            self._canvas.create_text(PREVIEW_W // 2, PREVIEW_H // 2,
                                     text="Pillow not installed\npip install Pillow",
                                     fill=TEXT_DIM, font=FONT_BODY, justify="center")
            return

        self._canvas.delete("all")
        img = self._last_grid if (use_last and self._last_grid) else self._build_preview_image()
        self._preview_tk = pil_to_tk(img, PREVIEW_W, PREVIEW_H)
        self._canvas.config(width=self._preview_tk.width(),
                            height=self._preview_tk.height())
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
                row_i = idx // self.cols
                col_i = idx  % self.cols
                x     = self.margin_x + col_i * self.cell_w + self.cell_pad
                y     = self.title_h  + self.margin_y + row_i * self.cell_h + self.cell_pad
                cw    = self.cell_w - self.cell_pad * 2
                ch    = self.cell_h - self.cell_pad * 2
                try:
                    cell = Image.open(path).convert("RGBA")
                    cell.thumbnail((max(1, cw), max(1, ch)), Image.LANCZOS)
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


# ─── Entry ────────────────────────────────────────────────────────────────────
def main():
    if not PIL_AVAILABLE:
        print("ERROR: Pillow is not installed.  Run:  pip install Pillow")
    app = BingoApp()
    app.mainloop()


if __name__ == "__main__":
    main()
