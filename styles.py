"""
Centralized styling module for Aghosh Care System.
Defines color palette, fonts, ttk theme configuration,
and helper functions for consistent professional UI.
"""
import tkinter as tk
from tkinter import ttk
import os
import sys


# ============================================================
# COLOR PALETTE
# ============================================================
COLORS = {
    "primary":        "#005BA1",
    "primary_light":  "#007ACC",
    "primary_dark":   "#004578",
    "accent":         "#0098FF",
    "accent_hover":   "#007ACC",
    "success":        "#28A745",
    "warning":        "#FFC107",
    "danger":         "#DC3545",
    "white":          "#FFFFFF",
    "bg":             "#F3F3F3",
    "bg_card":        "#FFFFFF",
    "bg_header":      "#005BA1",
    "bg_section":     "#E8E8E8",
    "border":         "#CCCEDB",
    "text_primary":   "#1E1E1E",
    "text_secondary": "#616161",
    "text_light":     "#FFFFFF",
    "text_muted":     "#A0A0A0",
    "separator":      "#CCCEDB",
    "input_bg":       "#FFFFFF",
    "input_border":   "#CCCEDB",
    "row_alt":        "#F5F5F5",
    "row_hover":      "#E4E6F1",
}

# ============================================================
# FONT DEFINITIONS
# ============================================================
FONTS = {
    "title":          ("Segoe UI", 20, "bold"),
    "subtitle":       ("Segoe UI", 16, "bold"),
    "heading":        ("Segoe UI", 14, "bold"),
    "subheading":     ("Segoe UI", 12, "bold"),
    "body":           ("Segoe UI", 11),
    "body_bold":      ("Segoe UI", 11, "bold"),
    "small":          ("Segoe UI", 10),
    "small_bold":     ("Segoe UI", 10, "bold"),
    "button":         ("Segoe UI", 11, "bold"),
    "entry":          ("Segoe UI", 11),
    "menu":           ("Segoe UI", 10),
    "footer":         ("Segoe UI", 9),
    "section_header": ("Segoe UI", 13, "bold"),
}

ORG_SUFFIX = " - Al-Khidmat Khawateen Trust Pakistan"


def apply_theme(root_window):
    """
    Apply the professional ttk theme to the root window.
    Call this once after creating the root Tk() instance.
    """
    style = ttk.Style(root_window)

    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure(".", font=FONTS["body"], background=COLORS["bg"])

    style.configure("TFrame", background=COLORS["bg"])
    style.configure("Card.TFrame", background=COLORS["bg_card"], relief="flat")
    style.configure("Header.TFrame", background=COLORS["bg_header"])
    style.configure("Section.TFrame", background=COLORS["bg_section"])

    style.configure("TLabel",
                     background=COLORS["bg"],
                     foreground=COLORS["text_primary"],
                     font=FONTS["body"])
    style.configure("Header.TLabel",
                     background=COLORS["bg_header"],
                     foreground=COLORS["text_light"],
                     font=FONTS["title"])
    style.configure("SubHeader.TLabel",
                     background=COLORS["bg_header"],
                     foreground=COLORS["text_light"],
                     font=FONTS["subtitle"])
    style.configure("Section.TLabel",
                     background=COLORS["bg"],
                     foreground=COLORS["primary"],
                     font=FONTS["section_header"])
    style.configure("Card.TLabel",
                     background=COLORS["bg_card"],
                     foreground=COLORS["text_primary"],
                     font=FONTS["body"])
    style.configure("Footer.TLabel",
                     background=COLORS["bg_section"],
                     foreground=COLORS["text_secondary"],
                     font=FONTS["footer"])
    style.configure("Muted.TLabel",
                     background=COLORS["bg"],
                     foreground=COLORS["text_secondary"],
                     font=FONTS["small"])

    style.configure("TButton",
                     font=FONTS["button"],
                     padding=(16, 8),
                     background=COLORS["accent"],
                     foreground=COLORS["white"])
    style.map("TButton",
              background=[("active", COLORS["accent_hover"]),
                          ("pressed", COLORS["primary_dark"])],
              foreground=[("active", COLORS["white"])])

    style.configure("Primary.TButton",
                     font=FONTS["button"],
                     padding=(20, 10),
                     background=COLORS["primary"],
                     foreground=COLORS["white"])
    style.map("Primary.TButton",
              background=[("active", COLORS["primary_light"]),
                          ("pressed", COLORS["primary_dark"])],
              foreground=[("active", COLORS["white"])])

    style.configure("Success.TButton",
                     font=FONTS["button"],
                     padding=(16, 8),
                     background=COLORS["success"],
                     foreground=COLORS["white"])
    style.map("Success.TButton",
              background=[("active", "#218838")],
              foreground=[("active", COLORS["white"])])

    style.configure("Danger.TButton",
                     font=FONTS["button"],
                     padding=(16, 8),
                     background=COLORS["danger"],
                     foreground=COLORS["white"])
    style.map("Danger.TButton",
              background=[("active", "#C82333")],
              foreground=[("active", COLORS["white"])])

    style.configure("Outline.TButton",
                     font=FONTS["button"],
                     padding=(16, 8),
                     background=COLORS["bg_card"],
                     foreground=COLORS["primary"],
                     bordercolor=COLORS["primary"],
                     borderwidth=2)
    style.map("Outline.TButton",
              background=[("active", COLORS["bg_section"])],
              foreground=[("active", COLORS["primary"])])

    style.configure("Nav.TButton",
                     font=FONTS["button"],
                     padding=(24, 14),
                     background=COLORS["accent"],
                     foreground=COLORS["white"],
                     anchor="center")
    style.map("Nav.TButton",
              background=[("active", COLORS["accent_hover"]),
                          ("pressed", COLORS["primary"])],
              foreground=[("active", COLORS["white"])])

    style.configure("TEntry",
                     font=FONTS["entry"],
                     padding=6,
                     fieldbackground=COLORS["input_bg"],
                     bordercolor=COLORS["input_border"],
                     lightcolor=COLORS["input_border"],
                     darkcolor=COLORS["input_border"])
    style.map("TEntry",
              bordercolor=[("focus", COLORS["accent"])],
              lightcolor=[("focus", COLORS["accent"])])

    style.configure("TCombobox",
                     font=FONTS["entry"],
                     padding=6,
                     fieldbackground=COLORS["input_bg"],
                     bordercolor=COLORS["input_border"])
    style.map("TCombobox",
              bordercolor=[("focus", COLORS["accent"])],
              fieldbackground=[("readonly", COLORS["input_bg"])])

    style.configure("Treeview",
                     font=FONTS["body"],
                     rowheight=28,
                     background=COLORS["white"],
                     foreground=COLORS["text_primary"],
                     fieldbackground=COLORS["white"])
    style.configure("Treeview.Heading",
                     font=FONTS["body_bold"],
                     background=COLORS["primary"],
                     foreground=COLORS["white"],
                     padding=6)
    style.map("Treeview",
              background=[("selected", COLORS["accent"])],
              foreground=[("selected", COLORS["white"])])

    style.configure("TSeparator", background=COLORS["separator"])

    style.configure("TLabelframe",
                     background=COLORS["bg"],
                     foreground=COLORS["primary"],
                     font=FONTS["subheading"])
    style.configure("TLabelframe.Label",
                     background=COLORS["bg"],
                     foreground=COLORS["primary"],
                     font=FONTS["subheading"])

    style.configure("TCheckbutton",
                     font=FONTS["body"],
                     background=COLORS["bg"],
                     foreground=COLORS["text_primary"])

    style.configure("Horizontal.TProgressbar",
                     background=COLORS["accent"],
                     troughcolor=COLORS["bg_section"],
                     bordercolor=COLORS["border"],
                     lightcolor=COLORS["accent"],
                     darkcolor=COLORS["primary"])

    root_window.configure(bg=COLORS["bg"])

    return style


def center_window(window, width, height):
    """Center a window on screen with specified dimensions."""
    window.update_idletasks()
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    x = (screen_w // 2) - (width // 2)
    y = (screen_h // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")
    window.minsize(width, height)


def setup_modal_window(window, parent, title, width, height):
    """
    Configure a Toplevel window as a professional modal dialog.
    Sets title, size, centers it, and makes it modal relative to parent.
    """
    window.title(title + ORG_SUFFIX)
    window.configure(bg=COLORS["bg"])
    center_window(window, width, height)
    if parent:
        window.transient(parent)
    window.grab_set()
    window.focus_force()
    window.bind('<Escape>', lambda e: window.destroy())


def setup_fullscreen_window(window, parent, title):
    """
    Configure a Toplevel window to be maximized/fullscreen with professional styling.
    """
    window.title(title + ORG_SUFFIX)
    window.configure(bg=COLORS["bg"])
    window.state('zoomed')
    if parent:
        window.transient(parent)
    window.grab_set()
    window.focus_force()
    window.bind('<Escape>', lambda e: window.destroy())


def create_window_header(parent, title_text, subtitle_text=None):
    """
    Create a styled header bar at the top of a window.
    Returns the header frame.
    """
    header = tk.Frame(parent, bg=COLORS["bg_header"], height=60)
    header.pack(fill="x", side="top")
    header.pack_propagate(False)

    tk.Label(header,
             text=title_text,
             font=FONTS["subtitle"],
             bg=COLORS["bg_header"],
             fg=COLORS["text_light"]).pack(side="left", padx=20, pady=10)

    if subtitle_text:
        tk.Label(header,
                 text=subtitle_text,
                 font=FONTS["small"],
                 bg=COLORS["bg_header"],
                 fg=COLORS["text_muted"]).pack(side="left", padx=10, pady=10)

    return header


def create_section_header(parent, text, row=None, column=0, columnspan=6, pady=(20, 8)):
    """
    Create a styled section header with a separator line.
    Works with both pack and grid layouts.
    Returns the frame containing the header.
    """
    section_frame = tk.Frame(parent, bg=COLORS["bg"])

    lbl = tk.Label(section_frame,
                   text=text,
                   font=FONTS["section_header"],
                   fg=COLORS["primary"],
                   bg=COLORS["bg"],
                   anchor="w")
    lbl.pack(fill="x", padx=5)

    sep = ttk.Separator(section_frame, orient="horizontal")
    sep.pack(fill="x", padx=5, pady=(2, 0))

    if row is not None:
        section_frame.grid(row=row, column=column, columnspan=columnspan,
                          sticky="ew", pady=pady, padx=5)
    else:
        section_frame.pack(fill="x", padx=10, pady=pady)

    return section_frame


def create_form_frame(parent, padding=20):
    """Create a styled frame for form fields."""
    frame = tk.Frame(parent, bg=COLORS["bg_card"],
                     bd=1, relief="solid",
                     highlightbackground=COLORS["border"],
                     highlightthickness=1)
    frame.pack(fill="both", expand=True, padx=padding, pady=10)
    return frame


def create_button_bar(parent, buttons, padding=15):
    """
    Create a styled button bar.
    buttons: list of tuples (text, command, style)
    style can be: 'primary', 'success', 'danger', 'outline', or None for default
    """
    bar = tk.Frame(parent, bg=COLORS["bg"])
    bar.pack(pady=padding, fill="x")

    style_map = {
        'primary': 'Primary.TButton',
        'success': 'Success.TButton',
        'danger':  'Danger.TButton',
        'outline': 'Outline.TButton',
        None:      'TButton',
    }

    for i, (text, command, btn_style) in enumerate(buttons):
        ttk.Button(bar,
                   text=text,
                   command=command,
                   style=style_map.get(btn_style, 'TButton'),
                   width=18).pack(side="left", padx=8, pady=5)

    bar.pack_configure(anchor="center")
    return bar


def styled_label(parent, text, font_key="body", bg=None, fg=None, **kwargs):
    """Create a styled tk.Label with theme defaults."""
    return tk.Label(parent,
                    text=text,
                    font=FONTS.get(font_key, FONTS["body"]),
                    bg=bg or COLORS["bg"],
                    fg=fg or COLORS["text_primary"],
                    **kwargs)


def styled_entry(parent, width=30, font_key="entry", **kwargs):
    """Create a styled tk.Entry with theme defaults."""
    entry = tk.Entry(parent,
                     font=FONTS.get(font_key, FONTS["entry"]),
                     width=width,
                     bg=COLORS["input_bg"],
                     relief="solid",
                     bd=1,
                     highlightthickness=1,
                     highlightcolor=COLORS["accent"],
                     highlightbackground=COLORS["input_border"],
                     **kwargs)
    return entry


def create_scrollable_frame(parent, bg_color=None):
    """
    Create a scrollable frame with canvas and scrollbar.
    Returns (main_frame, canvas, scroll_frame) tuple.
    The scroll_frame is where widgets should be placed.
    """
    bg = bg_color or COLORS["bg"]

    main_frame = tk.Frame(parent, bg=bg)
    main_frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(main_frame, bg=bg, highlightthickness=0)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg=bg)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel(event):
        try:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    return main_frame, canvas, scroll_frame


def show_loading_window(parent_title="Loading..."):
    """Show a small centered loading window with progress bar. Returns (window, progress_var, status_label) tuple."""
    loading = tk.Tk() if parent_title == "Loading..." else tk.Toplevel()
    loading.overrideredirect(True)
    loading.configure(bg=COLORS["primary"])

    w, h = 350, 100
    x = (loading.winfo_screenwidth() - w) // 2
    y = (loading.winfo_screenheight() - h) // 2
    loading.geometry(f"{w}x{h}+{x}+{y}")

    tk.Label(loading, text=parent_title, font=FONTS["body_bold"],
             bg=COLORS["primary"], fg=COLORS["text_light"]).pack(pady=(15, 5))

    progress_var = tk.IntVar(value=0)
    progress = ttk.Progressbar(loading, variable=progress_var, maximum=100, length=300,
                                style="Horizontal.TProgressbar")
    progress.pack(pady=5)

    status_label = tk.Label(loading, text="Please wait...", font=FONTS["small"],
                            bg=COLORS["primary"], fg=COLORS["text_muted"])
    status_label.pack()

    loading.update()
    return loading, progress_var, status_label


def setup_scrollable_window(window, parent, title, width, height):
    """Set up a window with a scrollable content area. Returns the inner frame for content."""
    window.title(title + ORG_SUFFIX)
    window.configure(bg=COLORS["bg"])
    center_window(window, width, height)
    if parent:
        window.transient(parent)
    window.grab_set()

    create_window_header(window, title)

    container = tk.Frame(window, bg=COLORS["bg"])
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container, bg=COLORS["bg"], highlightthickness=0)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    inner_frame = tk.Frame(canvas, bg=COLORS["bg"])

    inner_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=inner_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel(event):
        try:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    return inner_frame
