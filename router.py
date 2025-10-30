# -*- coding: utf-8 -*-
"""
Created on Thu Oct 16 13:06:42 2025

@author: bbruna
"""

from __future__ import annotations

import tkinter as tk
from typing import List, Optional

class Router:
    """
    Vista pequeña, cada "view" es un frame embedded en el parent.
    """
    def __init__(self, parent: tk.Widget):
        self.parent = parent
        self._stack: List[tk.Frame] = []

    def push(self, view: tk.Frame) -> None:
        if self._stack:
            self._stack[-1].pack_forget()
        self._stack.append(view)
        view.pack(fill="both", expand=True)

    def pop(self) -> None:
        if not self._stack:
            return
        top = self._stack.pop()
        try:
            top.destroy()
        except Exception:
            pass
        if self._stack:
            self._stack[-1].pack(fill="both", expand=True)

    def home(self) -> None:
        while len(self._stack) > 1:
            self.pop()

    def current(self) -> Optional[tk.Frame]:
        return self._stack[-1] if self._stack else None
