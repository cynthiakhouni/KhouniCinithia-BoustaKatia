"""
WT-Desal Lab — Main Application.
Glassmorphism navbar with centered tabs, light "Frost" theme.
"""

import customtkinter as ctk
from theme import Colors, Fonts, Spacing
from pages.source_page import SourcePage
from pages.desal_page import DesalPage
from pages.econ_page import EconPage
from pages.info_page import InfoPage


class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("WT-Desal Lab")
        self.geometry("1360x820")
        self.minsize(1000, 650)
        self.configure(fg_color=Colors.BG_ROOT)
        self._center()

        self._pages: dict[str, ctk.CTkFrame] = {}
        self._btns: dict[str, ctk.CTkButton] = {}
        self._bars: dict[str, ctk.CTkFrame] = {}
        self._current: str = ""

        self._build_topbar()

        self._content = ctk.CTkFrame(self, fg_color=Colors.BG_PRIMARY, corner_radius=0)
        self._content.pack(fill="both", expand=True)
        self._navigate("Source")

    def _build_topbar(self):
        # ── Outer wrapper with padding ──────────────────────────────────────
        wrapper = ctk.CTkFrame(self, fg_color=Colors.BG_ROOT, height=68, corner_radius=0)
        wrapper.pack(fill="x")
        wrapper.pack_propagate(False)

        # ── Glass bar (white frosted panel, floating) ───────────────────────
        glass = ctk.CTkFrame(
            wrapper,
            fg_color=Colors.BG_TOPBAR,
            corner_radius=Spacing.RADIUS_LG,
            border_width=1,
            border_color=Colors.BORDER_SUBTLE,
        )
        glass.pack(fill="x", padx=Spacing.PAD_MD, pady=Spacing.PAD_SM)

        inner = ctk.CTkFrame(glass, fg_color="transparent", height=44)
        inner.pack(fill="x", padx=Spacing.PAD_LG)
        inner.pack_propagate(False)

        # ── LEFT: Brand ─────────────────────────────────────────────────────
        brand = ctk.CTkFrame(inner, fg_color="transparent")
        brand.pack(side="left", pady=6)

        box = ctk.CTkFrame(brand, width=30, height=30,
                           fg_color=Colors.ACCENT, corner_radius=8)
        box.pack(side="left", padx=(0, 8))
        box.pack_propagate(False)
        ctk.CTkLabel(box, text="W", font=(Fonts.FAMILY, 13, "bold"),
                     text_color=Colors.TEXT_ON_ACCENT
                     ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(brand, text="WT-Desal", font=Fonts.H3,
                     text_color=Colors.TEXT_DARK).pack(side="left")

        # ── CENTER: Nav tabs (absolutely centered) ──────────────────────────
        nav_group = ctk.CTkFrame(inner, fg_color=Colors.BG_ELEVATED,
                                 corner_radius=Spacing.RADIUS_MD)
        nav_group.place(relx=0.5, rely=0.5, anchor="center")

        nav_inner = ctk.CTkFrame(nav_group, fg_color="transparent")
        nav_inner.pack(padx=3, pady=3)

        for name in ["Source", "Desalination", "Economics", "Info"]:
            col = ctk.CTkFrame(nav_inner, fg_color="transparent")
            col.pack(side="left", padx=1)

            btn = ctk.CTkButton(
                col, text=name,
                command=lambda n=name: self._navigate(n),
                font=Fonts.BODY,
                text_color=Colors.TEXT_SECONDARY,
                fg_color="transparent",
                hover_color=Colors.BG_HOVER,
                height=30, corner_radius=Spacing.RADIUS_SM,
                width=len(name) * 9 + 28,
            )
            btn.pack()

            ind = ctk.CTkFrame(col, height=2, fg_color="transparent", corner_radius=1)
            ind.pack(fill="x", padx=10)

            self._btns[name] = btn
            self._bars[name] = ind

        # ── RIGHT: Status ───────────────────────────────────────────────────
        right = ctk.CTkFrame(inner, fg_color="transparent")
        right.pack(side="right", pady=6)

        status = ctk.CTkFrame(right, fg_color=Colors.BG_ELEVATED,
                              corner_radius=10)
        status.pack(side="left")
        si = ctk.CTkFrame(status, fg_color="transparent")
        si.pack(padx=10, pady=3)

        dot = ctk.CTkFrame(si, width=6, height=6,
                           fg_color=Colors.SUCCESS, corner_radius=3)
        dot.pack(side="left", padx=(0, 5))
        dot.pack_propagate(False)

        ctk.CTkLabel(si, text="Online", font=Fonts.MICRO,
                     text_color=Colors.TEXT_SECONDARY).pack(side="left")

        ctk.CTkLabel(right, text="  v3.0", font=Fonts.MICRO,
                     text_color=Colors.TEXT_MUTED).pack(side="left")

    # ─── NAV ────────────────────────────────────────────────────────────────

    def _navigate(self, page: str):
        if page == self._current:
            return
        for w in self._content.winfo_children():
            w.pack_forget()
        self._current = page

        for name, btn in self._btns.items():
            active = name == page
            btn.configure(
                font=Fonts.BODY_BOLD if active else Fonts.BODY,
                text_color=Colors.ACCENT if active else Colors.TEXT_SECONDARY,
                fg_color=Colors.BG_TOPBAR if active else "transparent",
            )
            self._bars[name].configure(
                fg_color=Colors.ACCENT if active else "transparent"
            )

        if page not in self._pages:
            cls = {"Source": SourcePage, "Desalination": DesalPage, "Economics": EconPage, "Info": InfoPage}[page]
            self._pages[page] = cls(self._content)
        self._pages[page].pack(fill="both", expand=True)

    def _center(self):
        self.update_idletasks()
        w, h = 1360, 820
        x = (self.winfo_screenwidth() - w) // 2
        y = max(0, (self.winfo_screenheight() - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
