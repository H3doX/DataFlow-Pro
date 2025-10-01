"""
UI Improvements and styling for the Automation GUI
"""

import tkinter as tk
from tkinter import ttk
import tkinter.font as tkFont

# Color schemes for themes - FIXED with proper contrast ratios
THEMES = {
    "light": {
        # Main colors
        "bg": "#f5f5f5",               # Light gray background
        "fg": "#212121",               # Almost black text
        "heading_fg": "#000000",       # Black for headings

        # Component backgrounds
        "frame_bg": "#ffffff",         # White for frames
        "entry_bg": "#ffffff",         # White for input fields
        "button_bg": "#e0e0e0",        # Light gray for buttons

        # Component foregrounds
        "entry_fg": "#212121",         # Dark text in entries
        "button_fg": "#212121",        # Dark text on buttons
        "label_fg": "#424242",         # Dark gray for labels

        # Interactive states
        "select_bg": "#1976d2",        # Blue for selection
        "select_fg": "#ffffff",        # White text on selection
        "button_hover_bg": "#d0d0d0",  # Darker gray on hover
        "button_active_bg": "#1976d2", # Blue when active
        "button_active_fg": "#ffffff", # White text when active

        # Semantic colors
        "success": "#2e7d32",          # Dark green
        "success_fg": "#ffffff",       # White text
        "warning": "#f57c00",          # Orange
        "warning_fg": "#ffffff",       # White text
        "error": "#d32f2f",            # Red
        "error_fg": "#ffffff",         # White text
        "info": "#0288d1",             # Light blue
        "info_fg": "#ffffff",          # White text

        # Treeview specific
        "tree_bg": "#ffffff",          # White background
        "tree_fg": "#212121",          # Dark text
        "tree_selected_bg": "#1976d2", # Blue selection
        "tree_selected_fg": "#ffffff", # White text when selected
        "tree_heading_bg": "#f5f5f5",  # Light gray heading
        "tree_heading_fg": "#212121",  # Dark heading text

        # Notebook (tabs)
        "tab_bg": "#e0e0e0",          # Gray for inactive tabs
        "tab_fg": "#424242",          # Dark gray text
        "tab_selected_bg": "#ffffff", # White for active tab
        "tab_selected_fg": "#212121", # Dark text for active tab

        # Progressbar
        "progress_bg": "#e0e0e0",      # Light gray trough
        "progress_fg": "#1976d2",      # Blue progress

        # Menu
        "menu_bg": "#ffffff",          # White menu background
        "menu_fg": "#212121",          # Dark menu text
        "menu_select_bg": "#e3f2fd",  # Light blue selection
        "menu_select_fg": "#212121",  # Dark text on selection
    },

    "dark": {
        # Main colors
        "bg": "#121212",               # Very dark background
        "fg": "#e0e0e0",               # Light gray text
        "heading_fg": "#ffffff",       # White for headings

        # Component backgrounds
        "frame_bg": "#1e1e1e",         # Dark gray for frames
        "entry_bg": "#2d2d2d",         # Darker for input fields
        "button_bg": "#3a3a3a",        # Gray for buttons

        # Component foregrounds
        "entry_fg": "#e0e0e0",         # Light text in entries
        "button_fg": "#e0e0e0",        # Light text on buttons
        "label_fg": "#b0b0b0",         # Medium gray for labels

        # Interactive states
        "select_bg": "#1e88e5",        # Bright blue for selection
        "select_fg": "#ffffff",        # White text on selection
        "button_hover_bg": "#4a4a4a",  # Lighter gray on hover
        "button_active_bg": "#1e88e5", # Blue when active
        "button_active_fg": "#ffffff", # White text when active

        # Semantic colors
        "success": "#43a047",          # Green
        "success_fg": "#ffffff",       # White text
        "warning": "#fb8c00",          # Orange
        "warning_fg": "#000000",       # Black text for contrast
        "error": "#e53935",            # Bright red
        "error_fg": "#ffffff",         # White text
        "info": "#29b6f6",             # Light blue
        "info_fg": "#000000",          # Black text for contrast

        # Treeview specific
        "tree_bg": "#2d2d2d",          # Dark background
        "tree_fg": "#e0e0e0",          # Light text
        "tree_selected_bg": "#1e88e5", # Blue selection
        "tree_selected_fg": "#ffffff", # White text when selected
        "tree_heading_bg": "#3a3a3a",  # Gray heading
        "tree_heading_fg": "#e0e0e0",  # Light heading text

        # Notebook (tabs)
        "tab_bg": "#2d2d2d",          # Dark gray for inactive tabs
        "tab_fg": "#9e9e9e",          # Medium gray text
        "tab_selected_bg": "#3a3a3a", # Lighter gray for active tab
        "tab_selected_fg": "#ffffff", # White text for active tab

        # Progressbar
        "progress_bg": "#3a3a3a",      # Dark gray trough
        "progress_fg": "#1e88e5",      # Blue progress

        # Menu
        "menu_bg": "#2d2d2d",          # Dark menu background
        "menu_fg": "#e0e0e0",          # Light menu text
        "menu_select_bg": "#1e88e5",  # Blue selection
        "menu_select_fg": "#ffffff",  # White text on selection
    }
}

# Icons (Unicode characters for cross-platform compatibility)
ICONS = {
    "file": "📁",
    "excel": "📊",
    "save": "💾",
    "load": "📂",
    "play": "▶️",
    "stop": "⏹️",
    "add": "➕",
    "remove": "➖",
    "up": "⬆️",
    "down": "⬇️",
    "settings": "⚙️",
    "info": "ℹ️",
    "warning": "⚠️",
    "error": "❌",
    "success": "✅",
    "click": "👆",
    "keyboard": "⌨️",
    "timer": "⏱️",
    "mouse": "🖱️",
    "test": "🧪",
    "language": "🌐",
    "theme": "🎨",
    "capture": "📍",
    "mapping": "🔗"
}

def setup_styles(root, theme="light"):
    """Setup ttk styles for modern look with proper contrast"""
    style = ttk.Style(root)
    colors = THEMES[theme]

    # Configure root window
    root.configure(bg=colors["bg"])

    # Configure style theme
    style.theme_use('clam')  # Use clam for better customization

    # Global font configuration
    default_font = ("Segoe UI", 10)
    heading_font = ("Segoe UI", 11, "bold")
    small_font = ("Segoe UI", 9)

    # Configure all base styles
    style.configure(".",
                   background=colors["bg"],
                   foreground=colors["fg"],
                   font=default_font)

    # Frame styles
    style.configure("TFrame",
                   background=colors["frame_bg"],
                   relief="flat",
                   borderwidth=0)

    style.configure("Card.TFrame",
                   background=colors["frame_bg"],
                   relief="solid",
                   borderwidth=1,
                   bordercolor=colors["button_bg"])

    # Label styles
    style.configure("TLabel",
                   background=colors["frame_bg"],
                   foreground=colors["fg"],
                   font=default_font)

    style.configure("Heading.TLabel",
                   background=colors["frame_bg"],
                   foreground=colors["heading_fg"],
                   font=heading_font)

    style.configure("Body.TLabel",
                   background=colors["frame_bg"],
                   foreground=colors["label_fg"],
                   font=small_font)

    # Button styles
    style.configure("TButton",
                   font=default_font,
                   borderwidth=1,
                   focuscolor="none",
                   relief="raised",
                   padding=(10, 6))

    style.map("TButton",
              background=[("pressed", colors["button_active_bg"]),
                         ("active", colors["button_hover_bg"]),
                         ("!disabled", colors["button_bg"]),
                         ("disabled", colors["button_bg"])],
              foreground=[("pressed", colors["button_active_fg"]),
                         ("active", colors["button_fg"]),
                         ("!disabled", colors["button_fg"]),
                         ("disabled", colors["label_fg"])])

    # Primary button style
    style.configure("Primary.TButton",
                   font=("Segoe UI", 10, "bold"),
                   borderwidth=0,
                   focuscolor="none",
                   padding=(12, 8))

    style.map("Primary.TButton",
              background=[("pressed", colors["select_bg"]),
                         ("active", colors["select_bg"]),
                         ("!disabled", colors["select_bg"]),
                         ("disabled", colors["button_bg"])],
              foreground=[("pressed", colors["select_fg"]),
                         ("active", colors["select_fg"]),
                         ("!disabled", colors["select_fg"]),
                         ("disabled", colors["label_fg"])])

    # Secondary button style
    style.configure("Secondary.TButton",
                   font=default_font,
                   borderwidth=1,
                   focuscolor="none",
                   padding=(10, 6))

    style.map("Secondary.TButton",
              background=[("pressed", colors["button_active_bg"]),
                         ("active", colors["button_hover_bg"]),
                         ("!disabled", colors["button_bg"]),
                         ("disabled", colors["button_bg"])],
              foreground=[("pressed", colors["button_active_fg"]),
                         ("active", colors["button_fg"]),
                         ("!disabled", colors["button_fg"]),
                         ("disabled", colors["label_fg"])])

    # Success button style
    style.configure("Success.TButton",
                   font=("Segoe UI", 10, "bold"),
                   borderwidth=0,
                   focuscolor="none",
                   padding=(12, 8))

    style.map("Success.TButton",
              background=[("pressed", colors["success"]),
                         ("active", colors["success"]),
                         ("!disabled", colors["success"]),
                         ("disabled", colors["button_bg"])],
              foreground=[("pressed", colors["success_fg"]),
                         ("active", colors["success_fg"]),
                         ("!disabled", colors["success_fg"]),
                         ("disabled", colors["label_fg"])])

    # Danger button style
    style.configure("Danger.TButton",
                   font=default_font,
                   borderwidth=0,
                   focuscolor="none",
                   padding=(10, 6))

    style.map("Danger.TButton",
              background=[("pressed", colors["error"]),
                         ("active", colors["error"]),
                         ("!disabled", colors["error"]),
                         ("disabled", colors["button_bg"])],
              foreground=[("pressed", colors["error_fg"]),
                         ("active", colors["error_fg"]),
                         ("!disabled", colors["error_fg"]),
                         ("disabled", colors["label_fg"])])

    # Entry style
    style.configure("TEntry",
                   fieldbackground=colors["entry_bg"],
                   background=colors["entry_bg"],
                   foreground=colors["entry_fg"],
                   insertcolor=colors["entry_fg"],
                   borderwidth=1,
                   relief="solid",
                   padding=5)

    style.map("TEntry",
              fieldbackground=[("focus", colors["entry_bg"]),
                              ("!focus", colors["entry_bg"])],
              foreground=[("focus", colors["entry_fg"]),
                         ("!focus", colors["entry_fg"])])

    # Combobox style
    style.configure("TCombobox",
                   fieldbackground=colors["entry_bg"],
                   background=colors["button_bg"],
                   foreground=colors["entry_fg"],
                   borderwidth=1,
                   relief="solid",
                   padding=5,
                   arrowcolor=colors["button_fg"])

    style.map("TCombobox",
              fieldbackground=[("focus", colors["entry_bg"]),
                              ("!focus", colors["entry_bg"])],
              foreground=[("focus", colors["entry_fg"]),
                         ("!focus", colors["entry_fg"])])

    # Treeview style
    style.configure("Treeview",
                   background=colors["tree_bg"],
                   foreground=colors["tree_fg"],
                   fieldbackground=colors["tree_bg"],
                   borderwidth=0,
                   font=default_font)

    style.configure("Treeview.Heading",
                   background=colors["tree_heading_bg"],
                   foreground=colors["tree_heading_fg"],
                   font=heading_font,
                   borderwidth=1,
                   relief="raised")

    style.map("Treeview",
              background=[("selected", colors["tree_selected_bg"])],
              foreground=[("selected", colors["tree_selected_fg"])])

    style.map("Treeview.Heading",
              background=[("active", colors["button_hover_bg"]),
                         ("!active", colors["tree_heading_bg"])],
              foreground=[("active", colors["tree_heading_fg"]),
                         ("!active", colors["tree_heading_fg"])])

    # Notebook (tabs) style
    style.configure("TNotebook",
                   background=colors["bg"],
                   borderwidth=0,
                   tabmargins=[2, 5, 2, 0])

    style.configure("TNotebook.Tab",
                   padding=[20, 10],
                   background=colors["tab_bg"],
                   foreground=colors["tab_fg"],
                   font=default_font,
                   borderwidth=0)

    style.map("TNotebook.Tab",
              background=[("selected", colors["tab_selected_bg"]),
                         ("!selected", colors["tab_bg"])],
              foreground=[("selected", colors["tab_selected_fg"]),
                         ("!selected", colors["tab_fg"])],
              expand=[("selected", [1, 1, 1, 0])])

    # LabelFrame style
    style.configure("TLabelframe",
                   background=colors["frame_bg"],
                   foreground=colors["fg"],
                   borderwidth=1,
                   relief="solid",
                   bordercolor=colors["button_bg"])

    style.configure("TLabelframe.Label",
                   background=colors["frame_bg"],
                   foreground=colors["heading_fg"],
                   font=heading_font)

    # Checkbutton style
    style.configure("TCheckbutton",
                   background=colors["frame_bg"],
                   foreground=colors["fg"],
                   font=default_font,
                   focuscolor="none")

    style.map("TCheckbutton",
              background=[("active", colors["frame_bg"]),
                         ("!active", colors["frame_bg"])],
              foreground=[("active", colors["fg"]),
                         ("!active", colors["fg"])])

    # Progressbar style
    style.configure("TProgressbar",
                   background=colors["progress_fg"],
                   troughcolor=colors["progress_bg"],
                   borderwidth=0,
                   lightcolor=colors["progress_fg"],
                   darkcolor=colors["progress_fg"])

    # Scrollbar style
    style.configure("TScrollbar",
                   background=colors["button_bg"],
                   troughcolor=colors["frame_bg"],
                   borderwidth=0,
                   arrowcolor=colors["button_fg"],
                   width=12)

    style.map("TScrollbar",
              background=[("active", colors["button_hover_bg"]),
                         ("!active", colors["button_bg"])])

    # Separator style
    style.configure("TSeparator",
                   background=colors["button_bg"])

    # Text widget configuration (not ttk but needed)
    text_config = {
        "bg": colors["entry_bg"],
        "fg": colors["entry_fg"],
        "insertbackground": colors["entry_fg"],
        "selectbackground": colors["select_bg"],
        "selectforeground": colors["select_fg"],
        "font": default_font
    }

    return style, text_config

def create_tooltip(widget, text, delay=500):
    """Create a tooltip for a widget"""
    tooltip = None

    def on_enter(event):
        nonlocal tooltip
        if tooltip:
            return

        # Get theme colors
        theme = getattr(widget, '_theme', 'light')
        colors = THEMES[theme]

        tooltip = tk.Toplevel()
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")

        label = tk.Label(tooltip, text=text,
                        background=colors.get("info", "#17a2b8"),
                        foreground=colors.get("info_fg", "#ffffff"),
                        borderwidth=1,
                        relief="solid",
                        font=("Segoe UI", 9),
                        padx=8,
                        pady=4)
        label.pack()

    def on_leave(event):
        nonlocal tooltip
        if tooltip:
            tooltip.destroy()
            tooltip = None

    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)
    widget.bind("<ButtonPress>", on_leave)

def create_icon_button(parent, icon, text, command, style="Secondary.TButton", **kwargs):
    """Create a button with icon and text"""
    button_text = f"{ICONS.get(icon, '')} {text}"
    btn = ttk.Button(parent, text=button_text, command=command, style=style, **kwargs)
    # Store theme reference
    if hasattr(parent, '_theme'):
        btn._theme = parent._theme
    return btn

def create_status_bar(parent, theme="light"):
    """Create a status bar at the bottom of the window"""
    colors = THEMES[theme]

    status_frame = tk.Frame(parent, relief="sunken", borderwidth=1,
                           bg=colors["frame_bg"], height=25)
    status_frame.pack(side=tk.BOTTOM, fill=tk.X)
    status_frame.pack_propagate(False)

    status_label = tk.Label(status_frame, text="Ready",
                           font=("Segoe UI", 9),
                           bg=colors["frame_bg"],
                           fg=colors["label_fg"],
                           anchor="w")
    status_label.pack(side=tk.LEFT, padx=5, fill=tk.BOTH)

    return status_label

def setup_keyboard_shortcuts(root, shortcuts):
    """Setup keyboard shortcuts"""
    for key_combo, callback in shortcuts.items():
        root.bind(key_combo, lambda e, cb=callback: cb())

def show_notification(parent, message, type="info", duration=3000):
    """Show a temporary notification"""
    # Get current theme
    theme = getattr(parent, '_current_theme', 'light')
    colors = THEMES[theme]

    notif = tk.Toplevel(parent)
    notif.wm_overrideredirect(True)

    # Position at top-right of parent window
    parent.update_idletasks()
    x = parent.winfo_x() + parent.winfo_width() - 300
    y = parent.winfo_y() + 50
    notif.wm_geometry(f"280x60+{x}+{y}")

    # Style based on type
    bg_colors = {
        "info": colors["info"],
        "success": colors["success"],
        "warning": colors["warning"],
        "error": colors["error"]
    }

    fg_colors = {
        "info": colors["info_fg"],
        "success": colors["success_fg"],
        "warning": colors["warning_fg"],
        "error": colors["error_fg"]
    }

    icons_map = {
        "info": ICONS["info"],
        "success": ICONS["success"],
        "warning": ICONS["warning"],
        "error": ICONS["error"]
    }

    bg_color = bg_colors.get(type, colors["info"])
    fg_color = fg_colors.get(type, colors["info_fg"])

    notif.config(bg=bg_color)

    # Icon and message
    frame = tk.Frame(notif, bg=bg_color)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    icon_label = tk.Label(frame, text=icons_map.get(type, ""),
                         font=("Segoe UI", 16),
                         bg=bg_color,
                         fg=fg_color)
    icon_label.pack(side=tk.LEFT, padx=(0, 10))

    msg_label = tk.Label(frame, text=message,
                        font=("Segoe UI", 10),
                        bg=bg_color,
                        fg=fg_color,
                        wraplength=200,
                        justify=tk.LEFT)
    msg_label.pack(side=tk.LEFT)

    # Auto-destroy after duration
    notif.after(duration, notif.destroy)

    # Fade in effect
    notif.attributes('-alpha', 0.0)
    for i in range(10):
        notif.after(i * 30, lambda alpha=i/10: notif.attributes('-alpha', alpha))

    return notif