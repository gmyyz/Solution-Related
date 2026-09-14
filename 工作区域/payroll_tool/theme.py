"""theme: extracted payroll responsibilities; business rules preserved."""
import tkinter as tk
from tkinter import ttk


COLOR_PRIMARY = '#FFD700'


COLOR_SUCCESS = '#FFD700'


COLOR_DANGER = '#FF4444'


COLOR_WARNING = '#FFA500'


COLOR_BG = '#1A1A1A'


COLOR_CARD = '#2A2A2A'


COLOR_BORDER = '#3D3D3D'


COLOR_TEXT_MAIN = '#FFFFFF'


COLOR_TEXT_SUB = '#AAAAAA'


COLOR_HEADER_BG = '#111111'


def _apply_style(root):
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TFrame', background=COLOR_BG)
    style.configure(
        'Horizontal.TProgressbar',
        background=COLOR_PRIMARY,
        troughcolor='#333333',
        borderwidth=0,
        thickness=8,
    )


def _btn(parent, text, bg, command, padx=20, pady=8):
    if bg in (COLOR_PRIMARY, COLOR_SUCCESS):
        fg_color = '#000000'
        active_bg = '#E6C200'
    elif bg == COLOR_DANGER:
        fg_color = '#FFFFFF'
        active_bg = '#CC2222'
    else:
        fg_color = '#FFFFFF'
        active_bg = '#555555'
    return tk.Button(
        parent,
        text=text,
        bg=bg,
        fg=fg_color,
        font=('微软雅黑', 11, 'bold'),
        relief='flat',
        padx=padx,
        pady=pady,
        cursor='hand2',
        activebackground=active_bg,
        activeforeground=fg_color,
        command=command,
    )


def _circle_label(parent, text, size=26, bg_color=None, fg_color='#000000', parent_bg=None):
    bg_color = bg_color or COLOR_PRIMARY
    parent_bg = parent_bg or COLOR_BG
    canvas = tk.Canvas(parent, width=size, height=size, bg=parent_bg, highlightthickness=0)
    canvas.create_oval(1, 1, size - 1, size - 1, fill=bg_color, outline='')
    canvas.create_text(
        size // 2,
        size // 2,
        text=text,
        font=('微软雅黑', max(size // 3, 8), 'bold'),
        fill=fg_color,
        anchor='center',
    )
    return canvas
