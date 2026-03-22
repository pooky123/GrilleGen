"""
Bingo Grid Generator
====================
Language: Python 3
Libraries: tkinter (GUI), Pillow (image compositing), tkinter.filedialog, random
Run with: python bingo_generator.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import random
import math
import os
from pathlib import Path

try:
    from PIL import Image, ImageTk, ImageDraw
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# ─── Constants ───────────────────────────────────────────────────────────────
PREVIEW_W = 340
PREVIEW_H = 420
ACCENT     = "#b5651d"
ACCENT_LT  = "#d4956a"
BG_DARK    = "#1a1008"
BG_MID     = "#2b1e10"
BG_PANEL   = "#3a2a18"
BG_CARD    = "#4a3520"
TEXT_MAIN  = "#f5e6d0"
TEXT_DIM   = "#a08060"
BORDER     = "#6b4c2a"

FONT_HEAD  = ("Georgia", 13, "bold")
FONT_BODY  = ("Georgia", 10)
FONT_SMALL = ("Georgia", 9)
FONT_BIG   = ("Georgia", 22, "bold")
FONT_MONO  = ("Courier", 9)


# ─── Helpers ──────────────────────────────────────────────────────────────────
def pil_to_tk(img: "Image.Image", max_w: int, max_h: int) -> "ImageTk.PhotoImage":
    img = img.copy()
    img.thumbnail((max_w, max_h), Image.LANCZOS)
    return ImageTk.PhotoImage(img)


def make_placeholder(w: int, h: int, text: str, bg: str = "#2b1e10") -> "Image.Image":
    img = Image.new("RGBA", (w, h), bg + "ff")
    draw = ImageDraw.Draw(img)
    draw.rectangle([2, 2, w - 3, h - 3], outline="#6b4c2a", width=2)
    # simple centered text
    draw.text((w // 2, h // 2), text, fill="#a08060", anchor="mm")
    return img


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
        pad = dict(padx=14, pady=8)

        tk.Label(self, text="⚙  Settings", font=FONT_HEAD,
                 bg=BG_DARK, fg=ACCENT).pack(**pad, anchor="w")

        sep(self).pack(fill="x", padx=14)

        # Grid template
        tk.Label(self, text="Grid template image (PNG with transparent bg):",
                 font=FONT_SMALL, bg=BG_DARK, fg=TEXT_DIM).pack(**pad, anchor="w")
        row = tk.Frame(self, bg=BG_DARK)
        row.pack(fill="x", padx=14)
        self._grid_var = tk.StringVar(value=self.app.grid_template_path or "")
        tk.Entry(row, textvariable=self._grid_var, font=FONT_MONO,
                 bg=BG_MID, fg=TEXT_MAIN, insertbackground=ACCENT,
                 relief="flat", width=32).pack(side="left", padx=(0, 6))
        btn(row, "Browse", self._pick_grid).pack(side="left")

        sep(self).pack(fill="x", padx=14, pady=6)

        # Cell layout config
        tk.Label(self, text="Grid cell layout  (pixels from top-left of template):",
                 font=FONT_SMALL, bg=BG_DARK, fg=TEXT_DIM).pack(**pad, anchor="w")

        fields = [
            ("Columns",        "cols",        "4"),
            ("Rows (excl. title)", "rows",   "4"),
            ("Title row height (px)", "title_h", "80"),
            ("Cell width (px)", "cell_w",    "180"),
            ("Cell height (px)","cell_h",    "180"),
            ("Left margin (px)","margin_x",  "5"),
            ("Top margin after title (px)","margin_y", "5"),
            ("Cell padding (px)","cell_pad",  "8"),
        ]
        self._vars = {}
        for label, key, default in fields:
            cur = str(getattr(self.app, key, default))
            v = tk.StringVar(value=cur)
            self._vars[key] = v
            r = tk.Frame(self, bg=BG_DARK)
            r.pack(fill="x", padx=14, pady=2)
            tk.Label(r, text=label, font=FONT_SMALL, bg=BG_DARK,
                     fg=TEXT_MAIN, width=28, anchor="w").pack(side="left")
            tk.Entry(r, textvariable=v, font=FONT_MONO, bg=BG_MID, fg=TEXT_MAIN,
                     insertbackground=ACCENT, relief="flat", width=8).pack(side="left")

        sep(self).pack(fill="x", padx=14, pady=6)

        bf = tk.Frame(self, bg=BG_DARK)
        bf.pack(pady=10)
        btn(bf, "Save & Close", self._save).pack(side="left", padx=6)
        btn(bf, "Cancel", self.destroy, style="dim").pack(side="left", padx=6)

    def _pick_grid(self):
        p = filedialog.askopenfilename(
            title="Select grid template",
            filetypes=[("PNG Images", "*.png"), ("All files", "*.*")])
        if p:
            self._grid_var.set(p)

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
        self.app.refresh_preview()
        self.destroy()


# ─── Reusable widget helpers ───────────────────────────────────────────────────
def btn(parent, text, cmd, style="accent"):
    colors = {
        "accent": (ACCENT,   BG_DARK, ACCENT_LT),
        "dim":    (BG_PANEL, TEXT_DIM, BG_CARD),
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


# ─── Main Application ─────────────────────────────────────────────────────────
class BingoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bingo Grid Generator")
        self.configure(bg=BG_DARK)
        self.resizable(True, True)

        # ── State ──────────────────────────────────────────────────────────
        self.grid_template_path: str | None = None
        self.title_image_path:   str | None = None
        self.bg_image_path:      str | None = None
        self.item_paths: list[str] = []          # pool of cell images

        # Layout defaults (match the sample bingo card)
        self.cols     = 4
        self.rows     = 4
        self.title_h  = 80
        self.cell_w   = 180
        self.cell_h   = 180
        self.margin_x = 5
        self.margin_y = 5
        self.cell_pad = 8

        self._preview_tk  = None   # keep reference to avoid GC
        self._last_grid   = None   # last PIL image

        self._build_ui()
        self.refresh_preview()

    # ── UI construction ────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Top bar ────────────────────────────────────────────────────────
        top = tk.Frame(self, bg=BG_DARK, pady=8)
        top.pack(fill="x", padx=16)
        tk.Label(top, text="🌸  Bingo Grid Generator", font=FONT_BIG,
                 bg=BG_DARK, fg=ACCENT).pack(side="left")
        btn(top, "⚙  Settings", self._open_settings, "dim").pack(side="right", padx=4)

        sep(self).pack(fill="x", padx=16)

        # ── Main body ──────────────────────────────────────────────────────
        body = tk.Frame(self, bg=BG_DARK)
        body.pack(fill="both", expand=True, padx=16, pady=10)

        # LEFT panel
        left = tk.Frame(body, bg=BG_PANEL, bd=0, relief="flat", padx=12, pady=12)
        left.pack(side="left", fill="y", padx=(0, 10))
        self._build_left(left)

        # CENTER preview
        center = tk.Frame(body, bg=BG_MID, bd=0)
        center.pack(side="left", fill="both", expand=True, padx=(0, 10))
        self._build_center(center)

        # RIGHT panel
        right = tk.Frame(body, bg=BG_PANEL, bd=0, padx=12, pady=12)
        right.pack(side="right", fill="y")
        self._build_right(right)

    def _build_left(self, parent):
        tk.Label(parent, text="ACTIONS", font=FONT_SMALL, bg=BG_PANEL,
                 fg=TEXT_DIM).pack(anchor="w", pady=(0, 8))

        btn(parent, "🎲  Generate Grids", self._generate).pack(fill="x", pady=4)
        btn(parent, "🖼  Change Background", self._pick_bg, "dim").pack(fill="x", pady=4)

        sep(parent).pack(fill="x", pady=10)

        tk.Label(parent, text="TITLE IMAGE", font=FONT_SMALL, bg=BG_PANEL,
                 fg=TEXT_DIM).pack(anchor="w", pady=(0, 6))
        tk.Label(parent, text="Transparent PNG placed in\nthe top title bar:",
                 font=FONT_SMALL, bg=BG_PANEL, fg=TEXT_MAIN,
                 justify="left").pack(anchor="w")
        btn(parent, "📝  Select Title Image", self._pick_title, "dim").pack(fill="x", pady=(6, 2))

        self._title_lbl = tk.Label(parent, text="None selected", font=FONT_SMALL,
                                   bg=BG_PANEL, fg=TEXT_DIM, wraplength=160, justify="left")
        self._title_lbl.pack(anchor="w")

        sep(parent).pack(fill="x", pady=10)

        # Count spinner
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

        tk.Label(parent, text="POOL INFO", font=FONT_SMALL, bg=BG_PANEL,
                 fg=TEXT_DIM).pack(anchor="w")
        self._pool_info = tk.Label(parent, text="0 images in pool\n0 cells per grid",
                                   font=FONT_SMALL, bg=BG_PANEL, fg=TEXT_MAIN,
                                   justify="left")
        self._pool_info.pack(anchor="w", pady=(4, 0))

    def _build_center(self, parent):
        tk.Label(parent, text="PREVIEW", font=FONT_SMALL, bg=BG_MID,
                 fg=TEXT_DIM).pack(pady=(8, 0))

        self._canvas = tk.Canvas(parent, width=PREVIEW_W, height=PREVIEW_H,
                                 bg=BG_DARK, highlightthickness=0)
        self._canvas.pack(padx=10, pady=8)

    def _build_right(self, parent):
        tk.Label(parent, text="IMAGE POOL", font=FONT_SMALL, bg=BG_PANEL,
                 fg=TEXT_DIM).pack(anchor="w", pady=(0, 8))
        tk.Label(parent, text="All images available to\nfill grid cells:",
                 font=FONT_SMALL, bg=BG_PANEL, fg=TEXT_MAIN,
                 justify="left").pack(anchor="w")

        btn_row = tk.Frame(parent, bg=BG_PANEL)
        btn_row.pack(fill="x", pady=(6, 2))
        btn(btn_row, "+ Add Images", self._add_images).pack(side="left")
        btn(btn_row, "✕ Remove", self._remove_selected, "danger").pack(side="left", padx=(4, 0))

        frame = tk.Frame(parent, bg=BG_DARK)
        frame.pack(fill="both", expand=True, pady=(4, 0))

        scrollbar = tk.Scrollbar(frame, orient="vertical", bg=BG_MID)
        self._listbox = tk.Listbox(frame, font=FONT_SMALL, bg=BG_DARK, fg=TEXT_MAIN,
                                   selectbackground=ACCENT, selectforeground=BG_DARK,
                                   relief="flat", bd=0, width=24, height=20,
                                   yscrollcommand=scrollbar.set)
        scrollbar.config(command=self._listbox.yview)
        self._listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

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
            title="Select title text image (transparent PNG)",
            filetypes=[("PNG Images", "*.png"), ("All", "*.*")])
        if p:
            self.title_image_path = p
            name = Path(p).name
            self._title_lbl.config(text=name)
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
        sel = list(self._listbox.curselection())
        for i in reversed(sel):
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
        n = len(self.item_paths)
        cells = self.cols * self.rows
        self._pool_info.config(
            text=f"{n} image{'s' if n != 1 else ''} in pool\n{cells} cells per grid")

    # ── Grid composition ───────────────────────────────────────────────────
    def _compose_grid(self, assigned: list["Image.Image"]) -> "Image.Image":
        """Build one bingo card PIL image."""
        # --- background ---
        if self.bg_image_path and os.path.isfile(self.bg_image_path):
            bg = Image.open(self.bg_image_path).convert("RGBA")
        else:
            # soft pastel blue-sky default
            bg = Image.new("RGBA", (800, 960), "#cce8f0")

        # Figure out output size from grid params
        out_w = self.margin_x * 2 + self.cols * self.cell_w
        out_h = self.title_h + self.margin_y + self.rows * self.cell_h + self.margin_y
        bg = bg.resize((out_w, out_h), Image.LANCZOS)

        # --- grid template overlay ---
        if self.grid_template_path and os.path.isfile(self.grid_template_path):
            grid_img = Image.open(self.grid_template_path).convert("RGBA")
            grid_img = grid_img.resize((out_w, out_h), Image.LANCZOS)
            bg.paste(grid_img, (0, 0), grid_img)

        # --- paste cell images ---
        for idx, img in enumerate(assigned):
            row = idx // self.cols
            col = idx  % self.cols
            x = self.margin_x + col * self.cell_w + self.cell_pad
            y = self.title_h + self.margin_y + row * self.cell_h + self.cell_pad
            cw = self.cell_w - self.cell_pad * 2
            ch = self.cell_h - self.cell_pad * 2
            cell = img.convert("RGBA").copy()
            cell.thumbnail((cw, ch), Image.LANCZOS)
            # center in cell slot
            cx = x + (cw - cell.width)  // 2
            cy = y + (ch - cell.height) // 2
            bg.paste(cell, (cx, cy), cell)

        # --- title image overlay ---
        if self.title_image_path and os.path.isfile(self.title_image_path):
            title_img = Image.open(self.title_image_path).convert("RGBA")
            tw = out_w - 20
            title_img.thumbnail((tw, self.title_h - 10), Image.LANCZOS)
            tx = (out_w - title_img.width) // 2
            ty = (self.title_h - title_img.height) // 2
            bg.paste(title_img, (tx, ty), title_img)

        return bg

    def _validate_pool(self, n_grids: int) -> bool:
        """Check we have enough images so no image repeats across grids in the same slot."""
        need = self.cols * self.rows  # cells per grid
        have = len(self.item_paths)
        if have == 0:
            messagebox.showerror("Error", "Add at least some images to the pool first.")
            return False
        # Each grid needs `need` unique images, and no image can appear in the same
        # slot position in two grids.  Simplest rule: total pool >= need * n_grids
        # (strict: each image used at most once across all grids in the same slot).
        # We implement: shuffle the whole pool, assign sequentially per grid.
        # That means we need pool_size >= need * n_grids for zero repeats anywhere.
        # If pool is smaller we allow repeats across DIFFERENT slots but never same slot.
        # We warn if pool < need.
        if have < need:
            messagebox.showwarning(
                "Small pool",
                f"You have {have} images but each grid has {need} cells.\n"
                "Some images will repeat within a single grid.\n"
                "Add more images for best results.")
        return True

    def _assign_images(self, n_grids: int) -> list[list["Image.Image"]]:
        """
        Return n_grids lists, each with (cols*rows) PIL images.
        Guarantee: for position p, the image at position p in grid A
        never appears at position p in grid B.
        Strategy: build a per-slot shuffled rotation.
        """
        cells = self.cols * self.rows
        pool = self.item_paths[:]

        # Build per-slot queues: each slot gets a independently shuffled copy
        # of the pool so the same image never lands in the same slot twice.
        slot_queues = []
        for _ in range(cells):
            q = pool[:]
            random.shuffle(q)
            slot_queues.append(q)

        assignments = []
        for _ in range(n_grids):
            grid_imgs = []
            for slot in range(cells):
                path = slot_queues[slot].pop(0)
                # reload the queue when exhausted
                if not slot_queues[slot]:
                    slot_queues[slot] = pool[:]
                    random.shuffle(slot_queues[slot])
                grid_imgs.append(Image.open(path).convert("RGBA"))
            assignments.append(grid_imgs)

        return assignments

    # ── Generate ───────────────────────────────────────────────────────────
    def _generate(self):
        if not PIL_AVAILABLE:
            messagebox.showerror("Missing library",
                                 "Pillow is required.\nRun: pip install Pillow")
            return

        n = self._count_var.get()
        if not self._validate_pool(n):
            return

        out_dir = filedialog.askdirectory(title="Select output folder")
        if not out_dir:
            return

        try:
            assignments = self._assign_images(n)
        except Exception as e:
            messagebox.showerror("Error", f"Could not assign images:\n{e}")
            return

        saved = []
        for i, imgs in enumerate(assignments):
            card = self._compose_grid(imgs)
            fname = os.path.join(out_dir, f"bingo_grid_{i+1:03d}.png")
            card.save(fname, "PNG")
            saved.append(fname)

        # Show last card as preview
        self._last_grid = self._compose_grid(assignments[-1])
        self.refresh_preview(use_last=True)

        messagebox.showinfo("Done", f"Saved {n} grid(s) to:\n{out_dir}")

    # ── Preview ────────────────────────────────────────────────────────────
    def refresh_preview(self, use_last: bool = False):
        if not PIL_AVAILABLE:
            self._canvas.delete("all")
            self._canvas.create_text(PREVIEW_W // 2, PREVIEW_H // 2,
                                     text="Pillow not installed\npip install Pillow",
                                     fill=TEXT_DIM, font=FONT_BODY, justify="center")
            return

        self._canvas.delete("all")

        if use_last and self._last_grid:
            img = self._last_grid
        else:
            # Build a dummy preview
            img = self._build_preview_image()

        self._preview_tk = pil_to_tk(img, PREVIEW_W, PREVIEW_H)
        self._canvas.config(width=self._preview_tk.width(),
                            height=self._preview_tk.height())
        self._canvas.create_image(0, 0, anchor="nw", image=self._preview_tk)
        self._update_pool_info()

    def _build_preview_image(self) -> "Image.Image":
        """Compose a lightweight preview using current settings."""
        cells = self.cols * self.rows
        # background
        if self.bg_image_path and os.path.isfile(self.bg_image_path):
            bg = Image.open(self.bg_image_path).convert("RGBA")
        else:
            bg = Image.new("RGBA", (720, 900), "#cce8f0")

        # scale to preview
        bg.thumbnail((PREVIEW_W * 2, PREVIEW_H * 2), Image.LANCZOS)
        out_w, out_h = bg.size

        # grid template
        if self.grid_template_path and os.path.isfile(self.grid_template_path):
            grid_img = Image.open(self.grid_template_path).convert("RGBA")
            grid_img = grid_img.resize((out_w, out_h), Image.LANCZOS)
            bg.paste(grid_img, (0, 0), grid_img)

        # sample images
        if self.item_paths:
            sample = random.sample(self.item_paths,
                                   min(cells, len(self.item_paths)))
            while len(sample) < cells:
                sample.append(random.choice(self.item_paths))
            scale = out_w / (self.margin_x * 2 + self.cols * self.cell_w)
            cw = int((self.cell_w - self.cell_pad * 2) * scale)
            ch = int((self.cell_h - self.cell_pad * 2) * scale)
            for idx, path in enumerate(sample):
                r = idx // self.cols
                c = idx  % self.cols
                x = int((self.margin_x + c * self.cell_w + self.cell_pad) * scale)
                y = int((self.title_h + self.margin_y + r * self.cell_h + self.cell_pad) * scale)
                cell = Image.open(path).convert("RGBA")
                cell.thumbnail((cw, ch), Image.LANCZOS)
                cx = x + (cw - cell.width)  // 2
                cy = y + (ch - cell.height) // 2
                bg.paste(cell, (cx, cy), cell)

        # title
        if self.title_image_path and os.path.isfile(self.title_image_path):
            title_img = Image.open(self.title_image_path).convert("RGBA")
            scale = out_w / (self.margin_x * 2 + self.cols * self.cell_w)
            th = int(self.title_h * scale)
            title_img.thumbnail((out_w - 20, th - 4), Image.LANCZOS)
            tx = (out_w - title_img.width) // 2
            ty = max(0, (th - title_img.height) // 2)
            bg.paste(title_img, (tx, ty), title_img)

        return bg


# ─── Entry ────────────────────────────────────────────────────────────────────
def main():
    if not PIL_AVAILABLE:
        import sys
        print("ERROR: Pillow is not installed.")
        print("Install it with:  pip install Pillow")
        # Still launch the app so the user sees the error in the GUI
    app = BingoApp()
    app.mainloop()


if __name__ == "__main__":
    main()
