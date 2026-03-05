"""
WT-Desal Lab — Desalination Page. Light "Frost" theme.
"""

import customtkinter as ctk
from theme import Colors, Fonts, Spacing


class DesalPage(ctk.CTkFrame):

    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._build()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=Spacing.PAD_XL, pady=(Spacing.PAD_LG, 0))

        title_row = ctk.CTkFrame(header, fg_color="transparent")
        title_row.pack(fill="x")

        bar = ctk.CTkFrame(title_row, width=4, height=24,
                           fg_color=Colors.ACCENT, corner_radius=2)
        bar.pack(side="left", padx=(0, Spacing.PAD_SM))
        bar.pack_propagate(False)

        ctk.CTkLabel(title_row, text="Desalination", font=Fonts.H1,
                     text_color=Colors.TEXT_DARK).pack(side="left")

        badge = ctk.CTkFrame(title_row, fg_color=Colors.ACCENT_BG,
                             corner_radius=Spacing.RADIUS_SM)
        badge.pack(side="left", padx=(Spacing.PAD_SM, 0))
        ctk.CTkLabel(badge, text="  RO Module  ", font=Fonts.MICRO,
                     text_color=Colors.ACCENT).pack(padx=4, pady=2)

        ctk.CTkLabel(header, text="Reverse Osmosis simulation powered by wind energy.",
                     font=Fonts.CAPTION, text_color=Colors.TEXT_SECONDARY
                     ).pack(anchor="w", pady=(Spacing.PAD_XS, 0))

        # Hero card
        card = ctk.CTkFrame(self, fg_color=Colors.BG_CARD,
                            corner_radius=Spacing.RADIUS_XL,
                            border_width=1, border_color=Colors.BORDER_SUBTLE)
        card.pack(fill="both", expand=True,
                  padx=Spacing.PAD_XL, pady=Spacing.PAD_LG)

        ctk.CTkFrame(card, height=3, fg_color=Colors.ACCENT,
                     corner_radius=0).pack(fill="x")

        center = ctk.CTkFrame(card, fg_color="transparent")
        center.place(relx=0.5, rely=0.42, anchor="center")

        icon_bg = ctk.CTkFrame(center, width=70, height=70,
                               fg_color=Colors.ACCENT_BG, corner_radius=Spacing.RADIUS_XL)
        icon_bg.pack()
        icon_bg.pack_propagate(False)
        ctk.CTkLabel(icon_bg, text="💧", font=Fonts.ICON_LG,
                     text_color=Colors.ACCENT
                     ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(center, text="Desalination Module", font=Fonts.H1,
                     text_color=Colors.TEXT_DARK).pack(pady=(Spacing.PAD_MD, Spacing.PAD_XS))

        pill = ctk.CTkFrame(center, fg_color=Colors.ACCENT_SOFT, corner_radius=10)
        pill.pack(pady=(0, Spacing.PAD_LG))
        ctk.CTkLabel(pill, text="  COMING SOON  ",
                     font=(Fonts.FAMILY, 9, "bold"),
                     text_color=Colors.ACCENT).pack(padx=Spacing.PAD_MD, pady=3)

        ctk.CTkLabel(center,
                     text="Simulate water production from wind-powered\nreverse osmosis systems with full membrane modeling.",
                     font=Fonts.BODY, text_color=Colors.TEXT_SECONDARY,
                     justify="center").pack(pady=(0, Spacing.PAD_XL))

        features = ctk.CTkFrame(center, fg_color="transparent")
        features.pack()
        for feat in ["RO Membrane Simulation", "Water Salinity Modeling",
                     "Recovery Rate Optimization", "Specific Energy Consumption"]:
            row = ctk.CTkFrame(features, fg_color="transparent")
            row.pack(fill="x", pady=2)
            dot = ctk.CTkFrame(row, width=6, height=6,
                               fg_color=Colors.ACCENT, corner_radius=3)
            dot.pack(side="left", padx=(0, Spacing.PAD_SM), pady=6)
            dot.pack_propagate(False)
            ctk.CTkLabel(row, text=feat, font=Fonts.BODY,
                         text_color=Colors.TEXT_SECONDARY).pack(side="left")
