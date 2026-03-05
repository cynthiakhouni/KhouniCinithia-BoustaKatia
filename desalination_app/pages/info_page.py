"""
WT-Desal Lab — Info Page. Light "Frost" theme.
Application information and credits.
"""

import customtkinter as ctk
from theme import Colors, Fonts, Spacing


class InfoPage(ctk.CTkFrame):

    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._build()

    def _build(self):
        # Header
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=Spacing.PAD_XL, pady=(Spacing.PAD_LG, 0))

        title_row = ctk.CTkFrame(header, fg_color="transparent")
        title_row.pack(fill="x")

        bar = ctk.CTkFrame(title_row, width=4, height=24,
                           fg_color=Colors.ACCENT, corner_radius=2)
        bar.pack(side="left", padx=(0, Spacing.PAD_SM))
        bar.pack_propagate(False)

        ctk.CTkLabel(title_row, text="About", font=Fonts.H1,
                     text_color=Colors.TEXT_DARK).pack(side="left")

        badge = ctk.CTkFrame(title_row, fg_color=Colors.ACCENT_BG,
                             corner_radius=Spacing.RADIUS_SM)
        badge.pack(side="left", padx=(Spacing.PAD_SM, 0))
        ctk.CTkLabel(badge, text="  Info  ", font=Fonts.MICRO,
                     text_color=Colors.ACCENT).pack(padx=4, pady=2)

        ctk.CTkLabel(header, text="Application information and project credits.",
                     font=Fonts.CAPTION, text_color=Colors.TEXT_SECONDARY
                     ).pack(anchor="w", pady=(Spacing.PAD_XS, 0))

        # Main content card
        card = ctk.CTkFrame(self, fg_color=Colors.BG_CARD,
                            corner_radius=Spacing.RADIUS_XL,
                            border_width=1, border_color=Colors.BORDER_SUBTLE)
        card.pack(fill="both", expand=True,
                  padx=Spacing.PAD_XL, pady=Spacing.PAD_LG)

        ctk.CTkFrame(card, height=3, fg_color=Colors.ACCENT,
                     corner_radius=0).pack(fill="x")

        # Scrollable content
        scroll_frame = ctk.CTkScrollableFrame(card, fg_color="transparent",
                                               scrollbar_button_color=Colors.BORDER_DEFAULT)
        scroll_frame.pack(fill="both", expand=True, padx=Spacing.PAD_LG, pady=Spacing.PAD_LG)

        # App Icon and Name
        center = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        center.pack(pady=Spacing.PAD_LG)

        icon_bg = ctk.CTkFrame(center, width=80, height=80,
                               fg_color=Colors.ACCENT_BG, corner_radius=Spacing.RADIUS_XL)
        icon_bg.pack()
        icon_bg.pack_propagate(False)
        ctk.CTkLabel(icon_bg, text="🌊", font=(Fonts.FAMILY, 36),
                     text_color=Colors.ACCENT
                     ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(center, text="Eolien Desalination App", font=Fonts.H1,
                     text_color=Colors.TEXT_DARK).pack(pady=(Spacing.PAD_MD, Spacing.PAD_XS))

        # Info sections
        info_container = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        info_container.pack(fill="x", padx=Spacing.PAD_XL, pady=Spacing.PAD_LG)

        # Section helper
        def add_section(parent, title, content, icon=""):
            section = ctk.CTkFrame(parent, fg_color=Colors.BG_ELEVATED,
                                   corner_radius=Spacing.RADIUS_MD)
            section.pack(fill="x", pady=Spacing.PAD_SM)

            inner = ctk.CTkFrame(section, fg_color="transparent")
            inner.pack(fill="x", padx=Spacing.PAD_MD, pady=Spacing.PAD_MD)

            header_row = ctk.CTkFrame(inner, fg_color="transparent")
            header_row.pack(fill="x")

            if icon:
                ctk.CTkLabel(header_row, text=icon, font=(Fonts.FAMILY, 16),
                            text_color=Colors.ACCENT).pack(side="left", padx=(0, Spacing.PAD_SM))

            ctk.CTkLabel(header_row, text=title, font=Fonts.BODY_BOLD,
                        text_color=Colors.TEXT_DARK).pack(side="left")

            ctk.CTkLabel(inner, text=content, font=Fonts.BODY,
                        text_color=Colors.TEXT_SECONDARY,
                        wraplength=600, justify="left").pack(anchor="w", pady=(Spacing.PAD_XS, 0))

        # Developers
        add_section(info_container, "Developers",
                   "Khouni Cinithia / Bousta Katia", "👨‍💻")

        # Supervisor
        add_section(info_container, "Supervisor",
                   "Kirati Sidahmed Khodja", "👨‍🏫")

        # Institution
        add_section(info_container, "Institution",
                   "USTHB (Université des Sciences et de la Technologie Houari Boumediene)", "🏛️")

        # Project
        add_section(info_container, "Project",
                   "Final Year Project (PFE) 2026", "🎓")

        # Footer
        footer = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        footer.pack(pady=Spacing.PAD_LG)

        ctk.CTkLabel(footer, text="© 2026 Eolien Desalination App. All rights reserved.",
                     font=Fonts.CAPTION, text_color=Colors.TEXT_MUTED).pack()
