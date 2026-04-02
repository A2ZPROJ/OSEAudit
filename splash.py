"""
Splash Screen — OSEAudit
Desenvolvido por A2Z Projetos
"""

import os
import sys
import tkinter as tk
from tkinter import ttk

NOME_PROGRAMA  = "OSEAudit"
VERSAO         = "1.5"
SUBTITULO      = "Comparação e Auditoria de OSEs"
FABRICANTE     = "A2Z Projetos"
PARCEIRO       = "2S Engenharia e Geotecnologia"

_BG   = "#0D1117"
_SRF  = "#161B22"
_CARD = "#1C2128"
_RED  = "#DA3633"
_RED2 = "#3D0A0A"
_WHT  = "#F0F6FC"
_GRY  = "#8B949E"
_DIM  = "#30363D"

_BG_RGB  = (13, 17, 23)
_SRF_RGB = (22, 27, 34)


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
    W = 500
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
        self.top.geometry(f"{self.W}x{self.H}+{(sw-self.W)//2}+{(sh-self.H)//2}")

        self._build()
        self._fade(0.0, 1.0, 220, lambda: self.top.after(1200, self._begin_close))

    def _build(self):
        # ── Borda / outline da janela
        self.top.configure(highlightbackground=_DIM, highlightthickness=1)

        # ── Faixa vermelha no topo
        tk.Frame(self.top, bg=_RED, height=3).pack(fill="x")

        # ── Área central — produto
        center = tk.Frame(self.top, bg=_BG)
        center.pack(fill="both", expand=True, padx=40, pady=0)

        # Espaçamento superior
        tk.Frame(center, bg=_BG, height=30).pack()

        # Ícone + nome
        name_row = tk.Frame(center, bg=_BG)
        name_row.pack()

        icon_wrap = tk.Frame(name_row, bg=_RED2,
                             width=44, height=44,
                             highlightbackground=_RED,
                             highlightthickness=1)
        icon_wrap.pack(side="left", padx=(0, 14))
        icon_wrap.pack_propagate(False)
        tk.Label(icon_wrap, text="⌕",
                 font=("Segoe UI", 20), fg=_WHT, bg=_RED2).pack(expand=True)

        name_txt = tk.Frame(name_row, bg=_BG)
        name_txt.pack(side="left")
        title_row = tk.Frame(name_txt, bg=_BG)
        title_row.pack(anchor="w")
        tk.Label(title_row, text="OSE",
                 font=("Segoe UI", 26, "bold"),
                 fg=_WHT, bg=_BG).pack(side="left")
        tk.Label(title_row, text="Audit",
                 font=("Segoe UI", 26),
                 fg=_RED, bg=_BG).pack(side="left")

        tk.Label(name_txt, text=SUBTITULO,
                 font=("Segoe UI", 9),
                 fg=_GRY, bg=_BG).pack(anchor="w", pady=(2, 0))

        # Version pill
        pill = tk.Frame(name_txt, bg=_RED2,
                        padx=8, pady=1)
        pill.pack(anchor="w", pady=(6, 0))
        tk.Label(pill, text=f"v{VERSAO}",
                 font=("Consolas", 8, "bold"),
                 fg=_RED, bg=_RED2).pack()

        # Separador
        tk.Frame(center, bg=_DIM, height=1).pack(fill="x", pady=(20, 12))

        # ── Logos das empresas lado a lado
        logos_row = tk.Frame(center, bg=_BG)
        logos_row.pack()

        self._img_a2z = _load_img(_res("assets", "logo_a2z.png"), 90, 36, _BG_RGB)
        self._img_2s  = _load_img(_res("assets", "logo_2s.png"),  120, 36, _BG_RGB)

        # Bloco A2Z
        bloco_a2z = tk.Frame(logos_row, bg=_BG)
        bloco_a2z.pack(side="left", padx=(0, 30))
        if self._img_a2z:
            tk.Label(bloco_a2z, image=self._img_a2z,
                     bg=_BG).pack()
        else:
            tk.Label(bloco_a2z, text="A2Z PROJETOS",
                     font=("Segoe UI", 9, "bold"),
                     fg=_WHT, bg=_BG).pack()
        tk.Label(bloco_a2z, text="Desenvolvedor",
                 font=("Segoe UI", 7), fg=_GRY, bg=_BG).pack()

        # Separador vertical
        tk.Frame(logos_row, bg=_DIM, width=1).pack(side="left", fill="y", padx=(0, 30))

        # Bloco 2S
        bloco_2s = tk.Frame(logos_row, bg=_BG)
        bloco_2s.pack(side="left")
        if self._img_2s:
            tk.Label(bloco_2s, image=self._img_2s,
                     bg=_BG).pack()
        else:
            tk.Label(bloco_2s, text="2S ENGENHARIA",
                     font=("Segoe UI", 9, "bold"),
                     fg=_WHT, bg=_BG).pack()
        tk.Label(bloco_2s, text="Parceiro",
                 font=("Segoe UI", 7), fg=_GRY, bg=_BG).pack()

        # ── Barra de progresso + faixa inferior
        bottom = tk.Frame(self.top, bg=_BG)
        bottom.pack(fill="x", side="bottom")

        s = ttk.Style(self.top)
        s.theme_use("clam")
        s.configure("Spl.Horizontal.TProgressbar",
                    troughcolor=_CARD, background=_RED, thickness=3)
        self._pb = ttk.Progressbar(
            bottom, style="Spl.Horizontal.TProgressbar",
            mode="indeterminate")
        self._pb.pack(fill="x")
        self._pb.start(8)

        tk.Frame(self.top, bg=_RED2, height=3).pack(fill="x", side="bottom")

    # ─────────────────────────────────────────── animação
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
        self._fade(1.0, 0.0, 180, self._close)

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
