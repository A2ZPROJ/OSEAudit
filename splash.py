"""
Splash Screen — OSEAudit v1.7
A2Z Projetos em parceria com 2S Engenharia e Geotecnologia

CodePro-style dark splash with integrated GitHub update check.
"""

import os
import sys
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk
import urllib.error
import urllib.request
import json

VERSAO    = "1.7"
SUBTITULO = "COMPARAÇÃO E AUDITORIA DE OSEs"

# ── Palette ────────────────────────────────────────────────────────────────
_BG    = "#060D1B"   # Deep navy
_SRF   = "#0D1424"   # Surface
_CARD  = "#1C2128"   # Card
_BDR   = "#1E293B"   # Border
_RED   = "#DA3633"   # Accent red
_RED2  = "#3D0A0A"   # Deep red bg
_WHT   = "#F0F6FC"   # Primary text
_GRY   = "#8B949E"   # Muted text
_DIM   = "#30363D"   # Separator
_GRN   = "#3FB950"   # Green (ready)
_BLU   = "#3B82F6"   # Blue (downloading)

_BG_RGB  = (6, 13, 27)
_SRF_RGB = (13, 20, 36)

W = 480
H = 360


# ── Resource helper (frozen-safe) ──────────────────────────────────────────
def _res(*parts):
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
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
    """
    Splash screen with integrated GitHub update check.

    Parameters
    ----------
    parent       : tk.Tk — the hidden main window
    on_ready     : callable — invoked when no update found / after min display
    on_update    : callable — invoked after installer launched (destroy app)
    versao       : str — current version string, e.g. "1.7"
    github_repo  : str — "owner/repo" for GitHub Releases API
    """

    def __init__(self, parent, on_ready, on_update, versao, github_repo):
        self._parent      = parent
        self.on_ready     = on_ready
        self.on_update    = on_update
        self._versao      = versao
        self._github_repo = github_repo
        self._closed      = False
        self._ready_armed = False   # True once min 2s elapsed
        self._update_done = False   # True once update check finished (no update)
        self._start_ms    = None    # used for 2s minimum timing

        self.top = tk.Toplevel(parent)
        self.top.overrideredirect(True)
        self.top.configure(bg=_BG, highlightbackground=_DIM, highlightthickness=1)
        self.top.wm_attributes("-topmost", True)
        self.top.wm_attributes("-alpha", 0.0)
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)

        sw = self.top.winfo_screenwidth()
        sh = self.top.winfo_screenheight()
        self.top.geometry(f"{W}x{H}+{(sw - W) // 2}+{(sh - H) // 2}")

        self._build()
        self._fade(0.0, 1.0, 220, self._after_fade_in)

    # ──────────────────────────────────────────────────────────────── Build
    def _build(self):
        # ── Top accent bar (3px red)
        tk.Frame(self.top, bg=_RED, height=3).pack(fill="x")

        # ── Brand header (~60px): A2Z logo + tagline
        hdr = tk.Frame(self.top, bg=_BG, padx=24, pady=10)
        hdr.pack(fill="x")

        self._img_a2z = _load_img(_res("assets", "logo_a2z.png"), 80, 32, _BG_RGB)
        if self._img_a2z:
            tk.Label(hdr, image=self._img_a2z, bg=_BG).pack(side="left", padx=(0, 12))
        else:
            tk.Label(hdr, text="A2Z", font=("Segoe UI", 12, "bold"),
                     fg=_WHT, bg=_BG).pack(side="left", padx=(0, 12))

        brand_txt = tk.Frame(hdr, bg=_BG)
        brand_txt.pack(side="left", fill="y", justify="left")
        tk.Label(brand_txt, text="A2Z PROJETOS",
                 font=("Segoe UI", 10, "bold"), fg=_WHT, bg=_BG,
                 anchor="w").pack(anchor="w")
        tk.Label(brand_txt, text="Soluções em Engenharia",
                 font=("Segoe UI", 8), fg=_GRY, bg=_BG,
                 anchor="w").pack(anchor="w")

        # Header bottom border
        tk.Frame(self.top, bg=_BDR, height=1).pack(fill="x")

        # ── Product hero section (~120px)
        hero = tk.Frame(self.top, bg=_BG, padx=30, pady=22)
        hero.pack(fill="x")

        # Icon box (44x44)
        icon_wrap = tk.Frame(hero, bg=_RED2, width=44, height=44,
                             highlightbackground=_RED, highlightthickness=1)
        icon_wrap.pack(side="left", padx=(0, 16))
        icon_wrap.pack_propagate(False)
        tk.Label(icon_wrap, text="⌕", font=("Segoe UI", 20),
                 fg=_WHT, bg=_RED2).pack(expand=True)

        # Product text
        prod = tk.Frame(hero, bg=_BG)
        prod.pack(side="left", fill="y")

        title_row = tk.Frame(prod, bg=_BG)
        title_row.pack(anchor="w")
        tk.Label(title_row, text="OSE", font=("Segoe UI", 22, "bold"),
                 fg=_WHT, bg=_BG).pack(side="left")
        tk.Label(title_row, text="Audit", font=("Segoe UI", 22),
                 fg=_RED, bg=_BG).pack(side="left")

        tk.Label(prod, text=SUBTITULO,
                 font=("Segoe UI", 7, "bold"), fg=_RED, bg=_BG,
                 anchor="w").pack(anchor="w", pady=(3, 0))

        # Version pill
        pill = tk.Frame(prod, bg=_RED2, padx=8, pady=1)
        pill.pack(anchor="w", pady=(6, 0))
        tk.Label(pill, text=f"v{self._versao}",
                 font=("Consolas", 8, "bold"), fg=_RED, bg=_RED2).pack()

        # ── Status section (~60px)
        status_area = tk.Frame(self.top, bg=_BG, padx=30, pady=8)
        status_area.pack(fill="x")

        status_row = tk.Frame(status_area, bg=_BG)
        status_row.pack(anchor="w")

        self._dot = tk.Label(status_row, text="●", font=("Segoe UI", 7),
                             fg=_GRY, bg=_BG)
        self._dot.pack(side="left", padx=(0, 6))

        self._status_lbl = tk.Label(status_row, text="Verificando atualizações…",
                                    font=("Segoe UI", 9), fg=_GRY, bg=_BG)
        self._status_lbl.pack(side="left")

        # ── Progress bar (4px) — hidden until download
        pb_frame = tk.Frame(self.top, bg=_BG, padx=30)
        pb_frame.pack(fill="x")

        pb_bg = tk.Frame(pb_frame, bg=_BDR, height=4)
        pb_bg.pack(fill="x")
        pb_bg.pack_propagate(False)
        self._pb_fill = tk.Frame(pb_bg, bg=_RED, height=4)
        self._pb_fill.place(x=0, y=0, height=4, relwidth=0.0)

        # ── Spacer
        tk.Frame(self.top, bg=_BG, height=6).pack()

        # ── Footer separator + logos
        tk.Frame(self.top, bg=_BDR, height=1).pack(fill="x", padx=0)

        footer = tk.Frame(self.top, bg=_BG, padx=24, pady=10)
        footer.pack(fill="x", side="bottom")

        self._img_2s = _load_img(_res("assets", "logo_2s.png"), 80, 28, _BG_RGB)
        if self._img_2s:
            tk.Label(footer, image=self._img_2s, bg=_BG).pack(side="left", padx=(0, 10))
        else:
            tk.Label(footer, text="2S", font=("Segoe UI", 9, "bold"),
                     fg=_GRY, bg=_BG).pack(side="left", padx=(0, 10))

        tk.Label(footer, text="em parceria com 2S Engenharia e Geotecnologia",
                 font=("Segoe UI", 8), fg=_GRY, bg=_BG).pack(side="left")

    # ──────────────────────────────────────────── after fade-in
    def _after_fade_in(self):
        self._start_ms = self.top.tk.call("clock", "milliseconds")
        # Arm 2s timer
        self.top.after(2000, self._min_elapsed)
        # Start update check in background
        threading.Thread(target=self._check_updates, daemon=True).start()

    def _min_elapsed(self):
        """Called after the 2-second minimum display time."""
        self._ready_armed = True
        if self._update_done:
            # Update check already returned (no update) — close now
            self._begin_close()

    # ──────────────────────────────────────────── update check (thread)
    def _check_updates(self):
        try:
            url = f"https://api.github.com/repos/{self._github_repo}/releases/latest"
            req = urllib.request.Request(url, headers={"User-Agent": "OSEAudit"})
            with urllib.request.urlopen(req, timeout=4) as r:
                data = json.loads(r.read())
            tag = data.get("tag_name", "").lstrip("v")
            if tag and tag != self._versao:
                asset = next(
                    (a for a in data.get("assets", [])
                     if a["name"].lower().endswith(".exe")),
                    None)
                if asset:
                    self.top.after(0, lambda: self._begin_download(
                        tag, asset["browser_download_url"]))
                    return
            # No update (or tag matches)
            self.top.after(0, self._no_update)
        except Exception:
            self.top.after(0, self._no_update)

    def _no_update(self):
        """No update found — show green 'Tudo atualizado', then close."""
        self._set_status("Tudo atualizado", _GRN)
        self._update_done = True
        if self._ready_armed:
            # Min time already elapsed — wait 0.5s then close
            self.top.after(500, self._begin_close)
        # else: _min_elapsed will trigger _begin_close

    # ──────────────────────────────────────────── download flow
    def _begin_download(self, nova_versao, url):
        self._set_status(f"Baixando v{nova_versao}… 0%", _BLU)
        self._pulse_dot(_BLU)
        threading.Thread(target=self._download,
                         args=(nova_versao, url), daemon=True).start()

    def _download(self, nova_versao, url):
        try:
            tmp = tempfile.mktemp(suffix="_OSEAudit_Setup.exe")

            def _hook(count, block_size, total_size):
                if total_size > 0:
                    pct = min(100, int(count * block_size * 100 / total_size))
                    self.top.after(0, self._set_download_progress, nova_versao, pct)

            urllib.request.urlretrieve(url, tmp, reporthook=_hook)
            self.top.after(0, self._do_install, tmp)
        except Exception:
            # Download failed — proceed normally
            self.top.after(0, self._no_update)

    def _set_download_progress(self, nova_versao, pct):
        self._set_status(f"Baixando v{nova_versao}… {pct}%", _BLU)
        self._pb_fill.place(x=0, y=0, height=4, relwidth=min(1.0, pct / 100))

    def _do_install(self, tmp_path):
        self._set_status("Instalando…", _BLU)
        self._pb_fill.place(x=0, y=0, height=4, relwidth=1.0)
        try:
            subprocess.Popen([tmp_path, "/VERYSILENT", "/NORESTART"], shell=False)
        except Exception:
            self._no_update()
            return
        self._fade(1.0, 0.0, 180, self.on_update)

    # ──────────────────────────────────────────── dot pulse
    def _pulse_dot(self, color, state=True):
        if self._closed:
            return
        try:
            self._dot.config(fg=color if state else _GRY)
            self.top.after(500, self._pulse_dot, color, not state)
        except tk.TclError:
            pass

    # ──────────────────────────────────────────── status helper
    def _set_status(self, text, color=None):
        try:
            self._status_lbl.config(text=text)
            if color:
                self._dot.config(fg=color)
                self._status_lbl.config(fg=color)
        except tk.TclError:
            pass

    # ──────────────────────────────────────────── close / fade
    def _begin_close(self):
        self._fade(1.0, 0.0, 180, self._close)

    def _close(self):
        if self._closed:
            return
        self._closed = True
        try:
            self.top.destroy()
        except Exception:
            pass
        if self.on_ready:
            self.on_ready()

    # ──────────────────────────────────────────── fade animation
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
