"""
Splash Screen — OSEAudit
A2Z Projetos em parceria com 2S Engenharia e Geotecnologia
"""

import os
import sys
import tkinter as tk
from tkinter import ttk

NOME_PROGRAMA = "OSEAudit"
VERSAO        = "1.1"
SUBTITULO     = "Comparação e Auditoria de Documentos OSE"

_BG   = "#0D1117"
_SRF  = "#161B22"
_CARD = "#1C2128"
_RED  = "#DA3633"
_RED2 = "#3D0A0A"
_REDL = "#FF7B72"
_BDR  = "#30363D"
_WHT  = "#E6EDF3"
_GRY  = "#8B949E"
_DIM  = "#484F58"


def _res(*parts):
    base = (os.path.dirname(sys.executable)
            if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def _load_img(path, max_w, max_h, bg_hex):
    try:
        from PIL import Image, ImageTk
        img = Image.open(path).convert("RGBA")
        img.thumbnail((max_w, max_h), Image.LANCZOS)
        r, g, b = (int(bg_hex.lstrip("#")[i:i+2], 16) for i in (0, 2, 4))
        canvas = Image.new("RGBA", img.size, (r, g, b, 255))
        canvas.paste(img, mask=img.split()[3])
        return ImageTk.PhotoImage(canvas.convert("RGB"))
    except Exception:
        return None


class SplashScreen:
    W = 540
    H = 310

    def __init__(self, parent, on_close):
        self.on_close = on_close
        self._closed  = False

        self.top = tk.Toplevel(parent)
        self.top.overrideredirect(True)
        self.top.configure(bg=_BG)
        self.top.wm_attributes("-topmost", True)
        self.top.wm_attributes("-alpha", 0.0)
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)

        sw = self.top.winfo_screenwidth()
        sh = self.top.winfo_screenheight()
        x  = (sw - self.W) // 2
        y  = (sh - self.H) // 2
        self.top.geometry(f"{self.W}x{self.H}+{x}+{y}")

        self._build()
        self._fade(0.0, 1.0, 220,
                   lambda: self.top.after(700, self._begin_close))

    # ────────────────────────────────────────────────────────── build
    def _build(self):
        # Top accent bar
        tk.Frame(self.top, bg=_RED, height=3).pack(fill="x")

        body = tk.Frame(self.top, bg=_BG, padx=32, pady=18)
        body.pack(fill="both", expand=True)

        # ── Logos row ──────────────────────────────────────────────
        self._img_a2z = _load_img(_res("assets", "logo_a2z.png"), 148, 58, _BG)
        self._img_2s  = _load_img(_res("assets", "logo_2s.png"),  140, 58, _BG)

        logos = tk.Frame(body, bg=_BG)
        logos.pack(fill="x")

        if self._img_a2z:
            tk.Label(logos, image=self._img_a2z, bg=_BG).pack(side="left")
        else:
            tk.Label(logos, text="A2Z PROJETOS",
                     font=("Segoe UI", 12, "bold"),
                     fg=_RED, bg=_BG).pack(side="left")

        mid = tk.Frame(logos, bg=_BG)
        mid.pack(side="left", expand=True, fill="x")
        tk.Label(mid, text="×", font=("Segoe UI", 16),
                 fg=_DIM, bg=_BG).pack(expand=True)

        if self._img_2s:
            tk.Label(logos, image=self._img_2s, bg=_BG).pack(side="right")
        else:
            tk.Label(logos, text="2S ENGENHARIA",
                     font=("Segoe UI", 11, "bold"),
                     fg=_GRY, bg=_BG).pack(side="right")

        # ── Divider ────────────────────────────────────────────────
        tk.Frame(body, bg=_BDR, height=1).pack(fill="x", pady=(14, 0))

        # ── Title block ────────────────────────────────────────────
        ctr = tk.Frame(body, bg=_BG)
        ctr.pack(fill="x", pady=(16, 0))

        tk.Label(ctr, text=NOME_PROGRAMA,
                 font=("Segoe UI", 26, "bold"),
                 fg=_WHT, bg=_BG).pack()

        tk.Label(ctr, text=SUBTITULO,
                 font=("Segoe UI", 9),
                 fg=_GRY, bg=_BG).pack(pady=(3, 0))

        # Version pill
        pill_wrap = tk.Frame(ctr, bg=_BG)
        pill_wrap.pack(pady=(9, 0))
        pill = tk.Frame(pill_wrap, bg=_RED2, padx=12, pady=3)
        pill.pack()
        tk.Label(pill, text=f"v{VERSAO}",
                 font=("Segoe UI", 8, "bold"),
                 fg=_REDL, bg=_RED2).pack()

        # ── Credits ────────────────────────────────────────────────
        tk.Frame(body, bg=_CARD, height=1).pack(fill="x", pady=(14, 0))

        credits = tk.Frame(body, bg=_BG)
        credits.pack(fill="x", pady=(8, 0))
        tk.Label(credits, text="Desenvolvido por A2Z Projetos",
                 font=("Segoe UI", 8, "bold"),
                 fg=_GRY, bg=_BG).pack()
        tk.Label(credits, text="em parceria com 2S Engenharia e Geotecnologia",
                 font=("Segoe UI", 8),
                 fg=_DIM, bg=_BG).pack(pady=(1, 0))

        # ── Progress bar ───────────────────────────────────────────
        bottom = tk.Frame(self.top, bg=_BG)
        bottom.pack(fill="x", side="bottom")

        s = ttk.Style(self.top)
        s.theme_use("clam")
        s.configure("Spl.Horizontal.TProgressbar",
                    troughcolor=_CARD, background=_RED, thickness=3)
        self._pb = ttk.Progressbar(bottom, style="Spl.Horizontal.TProgressbar",
                                   mode="indeterminate")
        self._pb.pack(fill="x")
        self._pb.start(10)

        # Bottom accent bar
        tk.Frame(self.top, bg=_RED2, height=3).pack(fill="x", side="bottom")

    # ────────────────────────────────────────────────────── animation
    def _fade(self, start, end, ms, callback=None):
        steps = 20
        delay = max(1, ms // steps)
        delta = (end - start) / steps

        def step(alpha, n):
            try:
                self.top.wm_attributes("-alpha", max(0.0, min(1.0, alpha)))
                if n > 0:
                    self.top.after(delay, step, alpha + delta, n - 1)
                elif callback:
                    callback()
            except tk.TclError:
                pass

        step(start, steps)

    def _begin_close(self):
        try:
            self._pb.stop()
        except Exception:
            pass
        self._fade(1.0, 0.0, 160, self._close)

    def _close(self):
        if self._closed:
            return
        self._closed = True
        try:
            self.top.destroy()
        except Exception:
            pass
        if self.on_close:
            self.on_close()
