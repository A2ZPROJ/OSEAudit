"""
Splash Screen — OSEAudit
Desenvolvido por A2Z Projetos
"""

import os
import sys
import tkinter as tk
from tkinter import ttk

NOME_PROGRAMA = "OSEAudit"
VERSAO        = "1.2"
SUBTITULO     = "COMPARAÇÃO E AUDITORIA DE OSEs"
FABRICANTE    = "A2Z PROJETOS"
FABRICANTE_SUB = "Soluções em Engenharia"

# Paleta CodePro-style (navy escuro + vermelho no lugar do azul)
_BG    = "#060C18"     # fundo principal — navy muito escuro
_SRF   = "#0D1626"     # superfície elevada
_CARD  = "#111E30"     # card/input
_RED   = "#DA3633"     # accent principal
_RED2  = "#A0201E"     # accent escuro
_BLU   = "#DA3633"     # reuse red as accent (brand)
_BDR   = "rgba(255,255,255,.08)"
_WHT   = "#F0F6FC"
_GRY   = "#B0BEC9"
_DIM   = "#4A5568"
_DIM2  = "#2D3748"

# Rgb tuples para cálculos
_BG_RGB   = (6,  12, 24)
_SRF_RGB  = (13, 22, 38)
_RED_RGB  = (218, 54, 51)


def _res(*parts):
    base = (os.path.dirname(sys.executable)
            if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def _load_img(path, max_w, max_h, bg_rgb=_BG_RGB):
    try:
        from PIL import Image, ImageTk
        img = Image.open(path).convert("RGBA")
        img.thumbnail((max_w, max_h), Image.LANCZOS)
        canvas = Image.new("RGBA", img.size, bg_rgb + (255,))
        canvas.paste(img, mask=img.split()[3])
        return ImageTk.PhotoImage(canvas.convert("RGB"))
    except Exception:
        return None


class SplashScreen:
    W = 480
    H = 300

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
        self.top.geometry(f"{self.W}x{self.H}+{(sw-self.W)//2}+{(sh-self.H)//2}")

        self._build()
        self._fade(0.0, 1.0, 200, lambda: self.top.after(800, self._begin_close))

    # ──────────────────────────────────────────────────────── build
    def _build(self):
        # ── Brand header (estilo CodePro) ──────────────────────────
        hdr = tk.Frame(self.top, bg=_SRF)
        hdr.pack(fill="x")

        # Linha divisória sutil abaixo do header
        hdr_inner = tk.Frame(hdr, bg=_SRF, padx=18, pady=12)
        hdr_inner.pack(fill="x")

        # Logo A2Z (esquerda)
        self._img_a2z = _load_img(
            _res("assets", "logo_a2z.png"), 110, 40, _SRF_RGB)

        logo_frame = tk.Frame(hdr_inner, bg=_SRF)
        logo_frame.pack(side="left")

        if self._img_a2z:
            tk.Label(logo_frame, image=self._img_a2z,
                     bg=_SRF).pack(side="left", padx=(0, 10))
        else:
            tk.Label(logo_frame, text=FABRICANTE,
                     font=("Segoe UI", 11, "bold"),
                     fg=_WHT, bg=_SRF).pack(side="left", padx=(0, 10))

        brand_txt = tk.Frame(hdr_inner, bg=_SRF)
        brand_txt.pack(side="left")
        tk.Label(brand_txt, text=FABRICANTE,
                 font=("Segoe UI", 10, "bold"),
                 fg=_WHT, bg=_SRF).pack(anchor="w")
        tk.Label(brand_txt, text=FABRICANTE_SUB,
                 font=("Segoe UI", 8),
                 fg=_DIM, bg=_SRF).pack(anchor="w")

        tk.Frame(self.top, bg="#1A2640", height=1).pack(fill="x")

        # ── Product hero ───────────────────────────────────────────
        hero = tk.Frame(self.top, bg=_BG, padx=18, pady=14)
        hero.pack(fill="x")

        # Ícone do produto (quadrado arredondado vermelho com lupa)
        icon_wrap = tk.Frame(hero, bg=_RED2,
                              width=48, height=48,
                              highlightbackground=_RED,
                              highlightthickness=1)
        icon_wrap.pack(side="left", padx=(0, 14))
        icon_wrap.pack_propagate(False)
        tk.Label(icon_wrap, text="⌕",
                 font=("Segoe UI", 22), fg=_WHT, bg=_RED2).pack(expand=True)

        info = tk.Frame(hero, bg=_BG)
        info.pack(side="left", fill="x", expand=True)

        name_row = tk.Frame(info, bg=_BG)
        name_row.pack(anchor="w")
        tk.Label(name_row, text="OSE",
                 font=("Segoe UI", 20, "bold"),
                 fg=_WHT, bg=_BG).pack(side="left")
        tk.Label(name_row, text="Audit",
                 font=("Segoe UI", 20),
                 fg=_RED, bg=_BG).pack(side="left")

        tk.Label(info, text=SUBTITULO,
                 font=("Segoe UI", 8, "bold"),
                 fg=_RED, bg=_BG).pack(anchor="w", pady=(2, 0))

        # Version pill
        pill = tk.Frame(info, bg="#1A0D0D",
                         highlightbackground="#5A1A18",
                         highlightthickness=1,
                         padx=9, pady=2)
        pill.pack(anchor="w", pady=(6, 0))
        tk.Label(pill, text=f"v{VERSAO}",
                 font=("Consolas", 9, "bold"),
                 fg=_RED, bg="#1A0D0D").pack()

        tk.Frame(self.top, bg="#1A2640", height=1).pack(fill="x")

        # ── Corpo (crédito + info) ─────────────────────────────────
        body = tk.Frame(self.top, bg=_BG, padx=18, pady=14)
        body.pack(fill="both", expand=True)

        credit_row = tk.Frame(body, bg=_BG)
        credit_row.pack(fill="x")

        tk.Label(credit_row,
                 text="Desenvolvido por  ",
                 font=("Segoe UI", 9),
                 fg=_DIM, bg=_BG).pack(side="left")
        tk.Label(credit_row,
                 text=FABRICANTE,
                 font=("Segoe UI", 9, "bold"),
                 fg=_GRY, bg=_BG).pack(side="left")

        tk.Label(body,
                 text="em parceria com 2S Engenharia e Geotecnologia",
                 font=("Segoe UI", 8),
                 fg=_DIM, bg=_BG).pack(anchor="w", pady=(2, 0))

        # ── Progress bar ───────────────────────────────────────────
        bottom = tk.Frame(self.top, bg=_BG)
        bottom.pack(fill="x", side="bottom", pady=(0, 0))

        s = ttk.Style(self.top)
        s.theme_use("clam")
        s.configure("Spl.Horizontal.TProgressbar",
                    troughcolor=_CARD, background=_RED, thickness=3)
        self._pb = ttk.Progressbar(
            bottom, style="Spl.Horizontal.TProgressbar",
            mode="indeterminate")
        self._pb.pack(fill="x")
        self._pb.start(10)

        # Bottom accent
        tk.Frame(self.top, bg=_RED2, height=3).pack(fill="x", side="bottom")

    # ──────────────────────────────────────────────── animation
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
