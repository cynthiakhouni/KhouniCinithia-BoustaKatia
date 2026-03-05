"""
WT-Desal Lab — Design System.
"Frost" — Light white / grey / blue palette.
"""


class Colors:
    # ── Backgrounds ─────────────────────────────────────────────────────────
    BG_ROOT        = "#f0f4f8"       # Light blue-grey root
    BG_PRIMARY     = "#f0f4f8"       # Main content
    BG_TOPBAR      = "#ffffff"       # Navbar (white glass)
    BG_CARD        = "#ffffff"       # Cards
    BG_ELEVATED    = "#f7f9fc"       # Inputs / subtle
    BG_HOVER       = "#e8edf4"       # Hover state
    BG_ACTIVE      = "#dce3ee"       # Active / pressed

    # ── Borders ─────────────────────────────────────────────────────────────
    BORDER_SUBTLE  = "#e2e8f0"       # Soft borders
    BORDER_DEFAULT = "#cbd5e1"       # Visible
    BORDER_BRIGHT  = "#94a3b8"       # Focus

    # ── Accent — Soft Blue ──────────────────────────────────────────────────
    ACCENT         = "#3b82f6"       # Primary blue
    ACCENT_HOVER   = "#60a5fa"       # Lighter blue
    ACCENT_DIM     = "#2563eb"       # Deeper blue
    ACCENT_BG      = "#eff6ff"       # Blue-tinted bg
    ACCENT_SOFT    = "#dbeafe"       # Soft blue bg

    # ── Text ────────────────────────────────────────────────────────────────
    TEXT_DARK      = "#0f172a"        # Headings
    TEXT_PRIMARY   = "#1e293b"        # Body text
    TEXT_SECONDARY = "#64748b"        # Labels
    TEXT_MUTED     = "#94a3b8"        # Hints
    TEXT_FAINT     = "#cbd5e1"        # Ultra-subtle
    TEXT_ON_ACCENT = "#ffffff"        # On blue buttons

    # ── Status ──────────────────────────────────────────────────────────────
    SUCCESS        = "#22c55e"
    ERROR          = "#ef4444"


class Fonts:
    FAMILY     = "Segoe UI"

    H1         = (FAMILY, 20, "bold")
    H2         = (FAMILY, 15, "bold")
    H3         = (FAMILY, 12, "bold")
    BODY       = (FAMILY, 12)
    BODY_BOLD  = (FAMILY, 12, "bold")
    CAPTION    = (FAMILY, 10)
    TINY       = (FAMILY, 9)
    MICRO      = (FAMILY, 8)
    BIG_NUM    = (FAMILY, 28, "bold")
    ICON_LG    = (FAMILY, 36)
    ICON_MD    = (FAMILY, 20)


class Spacing:
    RADIUS_XL  = 18
    RADIUS_LG  = 12
    RADIUS_MD  = 8
    RADIUS_SM  = 6
    PAD_XL     = 28
    PAD_LG     = 20
    PAD_MD     = 14
    PAD_SM     = 8
    PAD_XS     = 4
