"""
main_app.py
===========
Interfaz gráfica (CustomTkinter) del
«Sistema de Formulación Presupuestal y Operativa POA - DPDI»
(Proceso Gestión Estratégica GD-DO-01 / Insumo PP-PR-02).

Pestañas:
  0. Inicio: datos de la unidad, generador de códigos por área y «Correr Actividades de Funcionamiento».
  1. Parámetros y catálogos NIIF (solo Administrador DPDI).
  2. Iniciativas del Plan Estratégico (maestro-detalle).
  3. Actividades de funcionamiento (Ingresos / Egresos).
  4. Otras actividades operativas (maestro-detalle).
  5. Actividades propias (maestro-detalle + punto de equilibrio).
  6. Actividades de inversión - CAPEX (maestro-detalle).
  7. Ficha Resumen y auditoría de directrices (12 secciones).

Ejecutar:  python main_app.py
"""
from __future__ import annotations

import os
import sys
import tkinter as tk
import traceback
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Optional

import customtkinter as ctk

sys.path.insert(0, str(Path(__file__).resolve().parent))

from catalogs_seed import BASES_CALCULO, MACROGRUPOS  # noqa: E402
from controllers import (DETAIL_TABLE, PRORRATEO_12, POAController, calc_detalle, calc_fegr, calc_fing,  # noqa: E402
                         code_display, control_meta, distribuir, fmt_mark, fmt_money, fmt_pct, s, to_float)
from database import MCOLS, MESES, Database  # noqa: E402
from exporter import (export_bd_consolidada, export_libro_oficial, export_pdf,  # noqa: E402
                      export_plantilla_parametros)

APP_TITLE = "Sistema de Formulación Presupuestal y Operativa POA - DPDI"
ROLE_ADMIN = "Administrador DPDI"
ROLE_UNIT = "Responsable de Unidad Orgánica"
COLOR_OK = "#1E8449"
COLOR_BAD = "#C0392B"
COLOR_WARN = "#B9770E"
OK_WORDS = {"OK", "Cumple", "Dentro del techo", "Autofinanciada", "Superávit", "Alcanza la meta"}
BAD_WORDS = {"Revisar", "Excede techo", "Déficit", "Sin techo", "Bajo la meta", "Sin meta"}

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


def status_color(text: str) -> str:
    if text in OK_WORDS:
        return COLOR_OK
    if text in BAD_WORDS:
        return COLOR_BAD
    return "gray30"


def fmt_value(value: Any, fmt: str) -> str:
    if fmt == "money":
        return fmt_money(value)
    if fmt == "pct":
        return fmt_pct(value)
    if fmt == "int":
        v = to_float(value)
        return "" if v is None else f"{v:,.0f}"
    if fmt == "mark":
        return fmt_mark(value)
    if fmt == "num":
        v = to_float(value)
        return "" if v is None else (f"{v:,.0f}" if v.is_integer() else f"{v:,.4f}".rstrip("0").rstrip("."))
    return s(value).replace("\n", " ⏎ ")


def guarded(fn: Callable) -> Callable:
    """Muestra los errores de negocio en un cuadro de diálogo en lugar de cerrar la aplicación."""

    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except ValueError as exc:
            messagebox.showwarning("Validación", str(exc))
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Error", f"{type(exc).__name__}: {exc}")
    wrapper.__name__ = fn.__name__
    return wrapper


# =============================================================================
# Widgets reutilizables
# =============================================================================
class DataTable(ctk.CTkFrame):
    """Treeview con barras de desplazamiento y formato por columna."""

    def __init__(self, master, columns: list[tuple], height: int = 10, on_select: Optional[Callable] = None,
                 status_cols: tuple[str, ...] = ()):
        super().__init__(master, fg_color="transparent")
        self.columns = columns
        self.status_cols = status_cols
        keys = [c[0] for c in columns]
        self.tree = ttk.Treeview(self, columns=keys, show="headings", height=height, selectmode="browse")
        for key, header, width, *rest in columns:
            anchor = rest[1] if len(rest) > 1 else ("e" if (rest and rest[0] in ("money", "pct", "int", "num"))
                                                    else "w")
            self.tree.heading(key, text=header)
            self.tree.column(key, width=width, minwidth=40, anchor=anchor, stretch=False)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure("bad", foreground=COLOR_BAD)
        self.tree.tag_configure("ok", foreground="black")
        self.tree.tag_configure("placeholder", foreground="gray55")
        self.tree.tag_configure("total", background="#D9D9D9", font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("odd", background="#F7F9FC")
        if on_select:
            self.tree.bind("<<TreeviewSelect>>", lambda e: on_select(self.selected_id()))
        self._rows: dict[str, dict] = {}

    def set_rows(self, rows: list[dict], id_key: str = "id", tag_fn: Optional[Callable] = None):
        sel = self.selected_id()
        self.tree.delete(*self.tree.get_children())
        self._rows.clear()
        for i, r in enumerate(rows):
            vals = []
            for key, header, width, *rest in self.columns:
                fmt = rest[0] if rest else "text"
                vals.append(fmt_value(r.get(key), fmt))
            iid = str(r.get(id_key, i))
            tags = [tag_fn(r)] if tag_fn and tag_fn(r) else []
            if i % 2:
                tags.append("odd")
            self.tree.insert("", "end", iid=iid, values=vals, tags=tags)
            self._rows[iid] = r
        if sel and str(sel) in self._rows:
            self.tree.selection_set(str(sel))

    def selected_id(self) -> Optional[int]:
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except ValueError:
            return None

    def selected_row(self) -> Optional[dict]:
        sel = self.tree.selection()
        return self._rows.get(sel[0]) if sel else None

    def select(self, iid: Any):
        if iid is not None and str(iid) in self._rows:
            self.tree.selection_set(str(iid))
            self.tree.see(str(iid))


class SearchDialog(ctk.CTkToplevel):
    """Buscador modal sobre una lista de valores (catálogo NIIF de 250+ partidas)."""

    def __init__(self, master, title: str, values: list[str], on_pick: Callable[[str], None], initial: str = "",
                 describe: Optional[Callable[[str], str]] = None):
        super().__init__(master)
        self.title(title)
        self.geometry("680x560" if describe else "620x480")
        self.values = values
        self.on_pick = on_pick
        self.describe = describe
        self.entry = ctk.CTkEntry(self, placeholder_text="Escriba para filtrar (código, partida o categoría)…")
        self.entry.pack(fill="x", padx=10, pady=(10, 6))
        self.entry.insert(0, initial)
        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=10, pady=4)
        self.lb = tk.Listbox(frame, font=("Segoe UI", 10), activestyle="dotbox")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.info = ctk.CTkLabel(self, text="")
        self.info.pack(pady=2)
        self.desc = None
        if describe:
            ctk.CTkLabel(self, text="Descripción orientativa", font=ctk.CTkFont(weight="bold")).pack(anchor="w",
                                                                                                  padx=10)
            self.desc = ctk.CTkTextbox(self, height=90, wrap="word")
            self.desc.pack(fill="x", padx=10, pady=(0, 6))
            self.lb.bind("<<ListboxSelect>>", lambda e: self._show_desc())
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(bar, text="Seleccionar", command=self._pick).pack(side="right")
        ctk.CTkButton(bar, text="Cancelar", fg_color="gray60", command=self.destroy).pack(side="right", padx=6)
        self.entry.bind("<KeyRelease>", lambda e: self._filter())
        self.entry.bind("<Return>", lambda e: self._pick())
        self.lb.bind("<Double-Button-1>", lambda e: self._pick())
        self.lb.bind("<Return>", lambda e: self._pick())
        self._filter()
        self.after(50, self._focus)

    def _focus(self):
        try:
            self.grab_set()
        except tk.TclError:
            pass
        self.entry.focus_set()

    def _filter(self):
        q = self.entry.get().strip().upper()
        words = q.split()
        self.lb.delete(0, "end")
        n = 0
        for v in self.values:
            up = v.upper()
            if all(w in up for w in words):
                self.lb.insert("end", v)
                n += 1
        self.info.configure(text=f"{n} de {len(self.values)} opciones")
        if n:
            self.lb.selection_set(0)
        self._show_desc()

    def _show_desc(self):
        if not self.desc:
            return
        sel = self.lb.curselection()
        text = self.describe(self.lb.get(sel[0])) if sel else ""
        self.desc.configure(state="normal")
        self.desc.delete("1.0", "end")
        self.desc.insert("1.0", text)
        self.desc.configure(state="disabled")

    def _pick(self):
        sel = self.lb.curselection()
        if not sel:
            return
        value = self.lb.get(sel[0])
        self.destroy()
        self.on_pick(value)


class SearchCombo(ctk.CTkFrame):
    """Combo con buscador: entrada + botón que abre el buscador del catálogo."""

    def __init__(self, master, values_fn: Callable[[], list[str]], width: int = 260,
                 on_change: Optional[Callable] = None, display_fn: Optional[Callable[[str], str]] = None):
        super().__init__(master, fg_color="transparent")
        self.values_fn = values_fn
        self.on_change = on_change
        self.display_fn = display_fn
        self.entry = ctk.CTkEntry(self, width=width - 36)
        self.entry.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(self, text="🔍", width=32, command=self.open).grid(row=0, column=1, padx=(4, 0))
        self.grid_columnconfigure(0, weight=1)
        self.entry.bind("<Return>", lambda e: self.open())
        self.entry.bind("<FocusOut>", lambda e: on_change and on_change())

    def open(self):
        vals = self.values_fn()
        shown = [self.display_fn(v) for v in vals] if self.display_fn else vals
        mapping = dict(zip(shown, vals))
        SearchDialog(self.winfo_toplevel(), "Buscar en catálogo", shown,
                     lambda v: self.set(mapping.get(v, v), fire=True), self.entry.get())

    def get(self) -> str:
        return self.entry.get().strip()

    def set(self, value: Any, fire: bool = False):
        self.entry.delete(0, "end")
        self.entry.insert(0, s(value))
        if fire and self.on_change:
            self.on_change()


class PartidaCombo(ctk.CTkFrame):
    """Combo «Partida» (Partidas GG) + buscador con ventana de descripción orientativa."""

    def __init__(self, master, values_fn: Callable[[], list[str]], describe: Optional[Callable[[str], str]],
                 width: int = 300, on_change: Optional[Callable] = None):
        super().__init__(master, fg_color="transparent")
        self.values_fn, self.describe, self.on_change = values_fn, describe, on_change
        self.combo = ctk.CTkComboBox(self, values=[], width=width - 36, state="readonly",
                                     command=lambda v: on_change and on_change())
        self.combo.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(self, text="🔍", width=32, command=self.open).grid(row=0, column=1, padx=(4, 0))
        self.grid_columnconfigure(0, weight=1)
        self.combo.set("")

    def reload(self):
        self.combo.configure(values=self.values_fn())

    def configure_values(self, values: list[str]):
        self.combo.configure(values=values)

    def open(self):
        SearchDialog(self.winfo_toplevel(), "Buscar Partida", self.values_fn(), lambda v: self.set(v, fire=True),
                     "", describe=self.describe)

    def get(self) -> str:
        return self.combo.get().strip()

    def set(self, value: Any, fire: bool = False):
        self.combo.set(s(value))
        if fire and self.on_change:
            self.on_change()


class Field:
    def __init__(self, key: str, label: str, kind: str = "entry", values: Any = None, width: int = 200,
                 span: int = 1, height: int = 60, anchor: str = "e", describe: Optional[Callable] = None):
        self.key, self.label, self.kind, self.values = key, label, kind, values
        self.width, self.span, self.height, self.anchor = width, span, height, anchor
        self.describe = describe


class Form(ctk.CTkFrame):
    """Formulario declarativo: entry | combo | search | text | calc | check."""

    def __init__(self, master, fields: list[Field], columns: int = 2, on_change: Optional[Callable] = None,
                 label_width: int = 120):
        super().__init__(master, fg_color="transparent")
        self.fields = {f.key: f for f in fields}
        self.widgets: dict[str, Any] = {}
        self.on_change = on_change
        r = c = 0
        for f in fields:
            if c + f.span > columns:
                r, c = r + 1, 0
            lab = ctk.CTkLabel(self, text=f.label, anchor="w", width=label_width, wraplength=label_width,
                               justify="left", font=ctk.CTkFont(size=12))
            lab.grid(row=r, column=c * 2, sticky="nw", padx=(4, 4), pady=3)
            w = self._make(f)
            w.grid(row=r, column=c * 2 + 1, columnspan=f.span * 2 - 1, sticky="ew", padx=(0, 10), pady=3)
            self.widgets[f.key] = w
            c += f.span
            if c >= columns:
                r, c = r + 1, 0
        for i in range(columns):
            self.grid_columnconfigure(i * 2 + 1, weight=1)

    def _changed(self, *_):
        if self.on_change:
            self.on_change()

    def _values(self, f: Field) -> list[str]:
        v = f.values() if callable(f.values) else (f.values or [])
        return [s(x) for x in v]

    def _make(self, f: Field):
        if f.kind == "combo":
            w = ctk.CTkComboBox(self, values=self._values(f), width=f.width, command=self._changed, state="readonly")
            w.set("")
        elif f.kind == "combo_free":
            w = ctk.CTkComboBox(self, values=self._values(f), width=f.width, command=self._changed)
            w.set("")
            w.bind("<KeyRelease>", self._changed)
        elif f.kind == "search":
            w = SearchCombo(self, f.values, width=f.width, on_change=self._changed)
        elif f.kind == "partida":
            w = PartidaCombo(self, lambda f=f: self._values(f), f.describe, width=f.width, on_change=self._changed)
            w.reload()
        elif f.kind == "text":
            w = ctk.CTkTextbox(self, width=f.width, height=f.height, wrap="word")
            w.bind("<KeyRelease>", self._changed)
        elif f.kind == "calc":
            w = ctk.CTkLabel(self, text="", width=f.width, anchor=f.anchor, fg_color=("gray88", "gray25"),
                             corner_radius=6, padx=8, wraplength=max(120, f.width - 10) if f.anchor == "w" else 0)
        elif f.kind == "check":
            w = ctk.CTkCheckBox(self, text="", command=self._changed)
        else:
            w = ctk.CTkEntry(self, width=f.width)
            w.bind("<KeyRelease>", self._changed)
        return w

    def reload_values(self):
        for k, f in self.fields.items():
            if f.kind in ("combo", "combo_free"):
                self.widgets[k].configure(values=self._values(f))
            elif f.kind == "partida":
                self.widgets[k].reload()

    def set_combo_values(self, key: str, values: list[str]):
        w = self.widgets[key]
        if isinstance(w, PartidaCombo):
            w.configure_values([s(v) for v in values])
        else:
            w.configure(values=[s(v) for v in values])

    def get(self) -> dict:
        out = {}
        for k, f in self.fields.items():
            w = self.widgets[k]
            if f.kind == "calc":
                continue
            if f.kind == "text":
                out[k] = w.get("1.0", "end-1c").strip()
            elif f.kind == "check":
                out[k] = 1 if w.get() else 0
            else:
                out[k] = w.get().strip() if isinstance(w.get(), str) else w.get()
        return out

    def value(self, key: str) -> str:
        return s(self.get().get(key))

    def set(self, data: dict):
        for k, f in self.fields.items():
            if k not in data:
                continue
            self.set_value(k, data[k])

    def set_value(self, k: str, v: Any):
        f, w = self.fields[k], self.widgets[k]
        if f.kind == "calc":
            text = s(v) if not isinstance(v, tuple) else v[0]
            color = status_color(text) if not isinstance(v, tuple) else v[1]
            w.configure(text=text, text_color=color)
        elif f.kind == "text":
            w.delete("1.0", "end")
            w.insert("1.0", s(v))
        elif f.kind == "check":
            w.select() if v and str(v) not in ("0", "False") else w.deselect()
        elif f.kind in ("combo", "combo_free"):
            w.set(s(v))
        elif f.kind in ("search", "partida"):
            w.set(v)
        else:
            w.delete(0, "end")
            w.insert(0, s(v) if not isinstance(v, float) else fmt_value(v, "num").replace(",", ""))

    def clear(self):
        for k, f in self.fields.items():
            self.set_value(k, "" if f.kind != "check" else 0)


class MonthGrid(ctk.CTkFrame):
    """Distribución mensual Ene–Dic con selección de meses y asistentes de prorrateo."""

    def __init__(self, master, mode: str = "amount", on_change: Optional[Callable] = None,
                 total_fn: Optional[Callable[[], float]] = None, show_calc: bool = False, title: str = ""):
        super().__init__(master, fg_color=("gray95", "gray17"), corner_radius=8)
        self.mode, self.on_change, self.total_fn = mode, on_change, total_fn
        self.entries: list[ctk.CTkEntry] = []
        self.checks: list[ctk.CTkCheckBox] = []
        self.calcs: list[ctk.CTkLabel] = []
        ctk.CTkLabel(self, text=title or ("Distribución mensual del monto (S/)" if mode == "amount"
                                          else "Meses de pago/cobro: «x» o factor (ej. 0.08333333)"),
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=6, sticky="w", padx=8, pady=(6, 2))
        for i, m in enumerate(MESES):
            blk = 1 + (i // 6) * (3 if show_calc else 2)
            colm = i % 6
            cb = ctk.CTkCheckBox(self, text=m, width=52, checkbox_width=16, checkbox_height=16)
            cb.grid(row=blk, column=colm, padx=3, pady=(4, 0), sticky="w")
            e = ctk.CTkEntry(self, width=80, justify="right")
            e.grid(row=blk + 1, column=colm, padx=3, pady=2)
            e.bind("<KeyRelease>", lambda ev: self._changed())
            self.checks.append(cb)
            self.entries.append(e)
            if show_calc:
                lab = ctk.CTkLabel(self, text="", width=80, anchor="e", text_color="gray25",
                                   font=ctk.CTkFont(size=11))
                lab.grid(row=blk + 2, column=colm, padx=3)
                self.calcs.append(lab)
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=10, column=0, columnspan=6, sticky="w", padx=4, pady=(6, 0))
        bar2 = ctk.CTkFrame(self, fg_color="transparent")
        bar2.grid(row=11, column=0, columnspan=6, sticky="w", padx=4, pady=(2, 6))
        if mode == "amount":
            ctk.CTkButton(bar, text="Distribuir equitativamente en meses seleccionados", width=10,
                          command=self.distribute_selected).pack(side="left", padx=2)
            ctk.CTkButton(bar, text="Prorrateo anual 1/12", width=10, command=self.prorate_all).pack(side="left", padx=2)
        else:
            ctk.CTkButton(bar, text="Prorrateo anual 1/12 (0.08333333)", width=10,
                          command=self.prorate_all).pack(side="left", padx=2)
            ctk.CTkButton(bar, text="Marcar «x» en seleccionados", width=10,
                          command=self.mark_selected).pack(side="left", padx=2)
            ctk.CTkButton(bar2, text="Factor 1/n en seleccionados", width=10,
                          command=self.distribute_selected).pack(side="left", padx=2)
        ctk.CTkButton(bar2, text="Seleccionar todos", width=10, fg_color="gray55",
                      command=lambda: [c.select() for c in self.checks]).pack(side="left", padx=2)
        ctk.CTkButton(bar2, text="Quitar selección", width=10, fg_color="gray55",
                      command=lambda: [c.deselect() for c in self.checks]).pack(side="left", padx=2)
        ctk.CTkButton(bar2, text="Limpiar montos", width=10, fg_color="gray55", command=self.clear).pack(side="left",
                                                                                                    padx=2)

    def _changed(self):
        if self.on_change:
            self.on_change()

    def selected(self) -> list[int]:
        return [i + 1 for i, c in enumerate(self.checks) if c.get()]

    def get(self) -> list[str]:
        return [e.get().strip() for e in self.entries]

    def set(self, values: list[Any]):
        for e, v in zip(self.entries, values):
            e.delete(0, "end")
            if self.mode == "amount":
                f = to_float(v)
                e.insert(0, "" if f is None else f"{f:.2f}")
            else:
                e.insert(0, fmt_mark(v))

    def set_calc(self, values: list[float]):
        for lab, v in zip(self.calcs, values):
            lab.configure(text=fmt_money(v) if v else "")

    def clear(self):
        self.set([None] * 12)
        for c in self.checks:
            c.deselect()
        self._changed()

    def prorate_all(self):
        if self.mode == "amount":
            total = self.total_fn() if self.total_fn else 0.0
            self.set(distribuir(total, range(1, 13)))
        else:
            self.set([PRORRATEO_12] * 12)
        self._changed()

    def distribute_selected(self):
        sel = self.selected()
        if not sel:
            messagebox.showinfo("Distribución", "Marque primero los meses (casillas) donde distribuir.")
            return
        if self.mode == "amount":
            total = self.total_fn() if self.total_fn else 0.0
            self.set(distribuir(total, sel))
        else:
            self.set([1.0 / len(sel) if (i + 1) in sel else None for i in range(12)])
        self._changed()

    def mark_selected(self):
        sel = self.selected()
        if not sel:
            messagebox.showinfo("Meses", "Marque primero los meses (casillas).")
            return
        self.set(["x" if (i + 1) in sel else None for i in range(12)])
        self._changed()


def button_bar(master, buttons: list[tuple[str, Callable]], side: str = "left") -> ctk.CTkFrame:
    bar = ctk.CTkFrame(master, fg_color="transparent")
    for item in buttons:
        text, cmd = item[0], item[1]
        color = item[2] if len(item) > 2 else None
        b = ctk.CTkButton(bar, text=text, command=cmd, width=10, **({"fg_color": color} if color else {}))
        b.pack(side=side, padx=3, pady=3)
    return bar


# =============================================================================
# Pantalla de acceso
# =============================================================================
class LoginFrame(ctk.CTkFrame):
    def __init__(self, master, app: "App"):
        super().__init__(master)
        self.app = app
        card = ctk.CTkFrame(self, corner_radius=14)
        card.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(card, text="Sistema de Formulación Presupuestal\ny Operativa POA - DPDI",
                     font=ctk.CTkFont(size=22, weight="bold"), justify="center").pack(padx=40, pady=(28, 4))
        ctk.CTkLabel(card, text="Proceso Gestión Estratégica GD-DO-01 · Insumo PP-PR-02",
                     text_color="gray40").pack(pady=(0, 18))
        self.role = ctk.CTkSegmentedButton(card, values=[ROLE_UNIT, ROLE_ADMIN], command=self._role_changed)
        self.role.set(ROLE_UNIT)
        self.role.pack(padx=30, pady=6, fill="x")
        ctk.CTkLabel(card, text="Unidad orgánica / Centro de costo", anchor="w").pack(fill="x", padx=30, pady=(12, 0))
        self.unit = ctk.CTkComboBox(card, values=[], width=520, state="readonly")
        self.unit.pack(padx=30, pady=4)
        self.pwd_box = ctk.CTkFrame(card, fg_color="transparent", height=1)
        self.pwd_box.pack(fill="x", padx=30)
        self.pwd_label = ctk.CTkLabel(self.pwd_box, text="Clave de administrador", anchor="w")
        self.pwd = ctk.CTkEntry(self.pwd_box, show="•", width=520)
        self.pwd.bind("<Return>", lambda e: self.login())
        ctk.CTkButton(card, text="Ingresar", height=38, command=self.login).pack(padx=30, pady=(18, 8), fill="x")
        self.anio = ctk.CTkLabel(card, text="", text_color="gray40")
        self.anio.pack(pady=(0, 22))
        self.refresh()

    def refresh(self):
        units = self.app.c.unidades()
        self._units = {self.app.c.unidad_display(u): u["id"] for u in units}
        self.unit.configure(values=list(self._units))
        if self._units and not self.unit.get():
            self.unit.set(next(iter(self._units)))
        p = self.app.c.params
        self.anio.configure(text=f"Año de formulación: {p.anio} · Fecha límite: {s(p.raw.get('fecha_limite'))}")

    def _role_changed(self, value):
        if value == ROLE_ADMIN:
            self.pwd_label.pack(fill="x", pady=(10, 0))
            self.pwd.pack(pady=4)
            self.pwd.focus_set()
        else:
            self.pwd_label.pack_forget()
            self.pwd.pack_forget()

    def login(self):
        role = self.role.get()
        uid = self._units.get(self.unit.get())
        if role == ROLE_ADMIN:
            if self.pwd.get() != s(self.app.db.parametros().get("clave_admin") or "DPDI"):
                messagebox.showerror("Acceso", "Clave de administrador incorrecta.")
                return
        elif uid is None:
            messagebox.showwarning("Acceso", "Seleccione su unidad orgánica.")
            return
        self.pwd.delete(0, "end")
        self.app.start_session(role, uid)


# =============================================================================
# Pestaña 0 — Inicio, códigos y funcionamiento
# =============================================================================
class InicioTab(ctk.CTkScrollableFrame):
    def __init__(self, master, app: "App"):
        super().__init__(master, fg_color="transparent")
        self.app, self.c = app, app.c
        self.grid_columnconfigure((0, 1, 2), weight=1)

        # --- Datos de la unidad y conformidad
        box = ctk.CTkFrame(self)
        box.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ctk.CTkLabel(box, text="Datos de la unidad (Ficha_Resumen)", font=ctk.CTkFont(size=15, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 2))
        self.form = Form(box, [
            Field("nombre_unidad", "Unidad orgánica", "calc", width=260, anchor="w"),
            Field("centro_costo", "Centro de costo", "calc", width=260, anchor="w"),
            Field("responsable", "Responsable de la unidad"), Field("cargo", "Cargo"),
            Field("correo", "Correo electrónico"), Field("fecha_presentacion", "Fecha de presentación"),
            Field("version_ficha", "Versión de la ficha"),
            Field("elaborado_nombre", "Elaborado por (nombre)"), Field("elaborado_cargo", "Elaborado por (cargo)"),
            Field("elaborado_fecha", "Fecha (elaborado)"),
            Field("revisado_nombre", "Revisado por DPDI (nombre)"), Field("revisado_cargo", "Revisado por (cargo)"),
            Field("revisado_fecha", "Fecha (revisado)"),
        ], columns=1, label_width=150)
        self.form.pack(fill="x", padx=6)
        button_bar(box, [("Guardar datos de la unidad", self.save_unidad),
                         ("Hoy en fechas", self.fill_today, "gray55")]).pack(anchor="w", padx=6, pady=6)

        # --- Generador de códigos
        gen = ctk.CTkFrame(self)
        gen.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)
        ctk.CTkLabel(gen, text="Generador automático de códigos", font=ctk.CTkFont(size=15, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 2))
        ctk.CTkLabel(gen, text="Indique la cantidad de actividades aprobadas por tipo. Los códigos se crean "
                               "correlativos y se prefijan con el código de área en la base consolidada.",
                     wraplength=380, justify="left", text_color="gray35").pack(anchor="w", padx=10)
        grid = ctk.CTkFrame(gen, fg_color="transparent")
        grid.pack(anchor="w", padx=10, pady=6)
        self.qty: dict[str, ctk.CTkEntry] = {}
        for i, (k, lab) in enumerate((("INI", "Iniciativas (INI-NN)"), ("FI", "Funcionamiento ingresos (FI-NNN)"),
                                      ("FE", "Funcionamiento egresos (FE-NNN)"), ("OO", "Otras operativas (OO-NNN)"),
                                      ("AP", "Actividades propias (AP-NNN)"), ("INV", "Inversiones (INV-NNN)"))):
            ctk.CTkLabel(grid, text=lab, anchor="w", width=230).grid(row=i, column=0, sticky="w", pady=2)
            e = ctk.CTkEntry(grid, width=70, justify="right")
            e.insert(0, "0")
            e.grid(row=i, column=1, pady=2)
            self.qty[k] = e
        button_bar(gen, [("Vista previa", self.preview_codes, "gray55"), ("Generar códigos", self.generate_codes)]).pack(
            anchor="w", padx=6)
        self.codes_box = ctk.CTkTextbox(gen, height=200, width=380)
        self.codes_box.pack(fill="both", expand=True, padx=10, pady=8)

        # --- Correr actividades de funcionamiento
        run = ctk.CTkFrame(self)
        run.grid(row=0, column=2, sticky="nsew", padx=6, pady=6)
        ctk.CTkLabel(run, text="Correr Actividades de Funcionamiento",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=10, pady=(8, 2))
        ctk.CTkLabel(run, text="Importa la plantilla base de ingresos y egresos fijos (planillas, CTS, depreciación, "
                               "servicios básicos, provisiones) prorrateando 1/12 (0.08333333) en Ene–Dic.",
                     wraplength=420, justify="left", text_color="gray35").pack(anchor="w", padx=10)
        self.tpl_table = DataTable(run, [("tipo_flujo", "Tipo", 70), ("partida", "Partida / rubro", 210),
                                         ("base_calculo", "Base", 140), ("_monto", "Monto anual (S/)", 110, "money")],
                                   height=9)
        self.tpl_table.pack(fill="both", expand=True, padx=8, pady=6)
        self.replace_chk = ctk.CTkCheckBox(run, text="Reemplazar las filas generadas previamente por la plantilla")
        self.replace_chk.select()
        self.replace_chk.pack(anchor="w", padx=10)
        ctk.CTkButton(run, text="▶ Correr Actividades de Funcionamiento", height=36,
                      command=self.run_funcionamiento).pack(fill="x", padx=10, pady=8)
        self.run_info = ctk.CTkLabel(run, text="", justify="left", wraplength=420)
        self.run_info.pack(anchor="w", padx=10, pady=(0, 8))

        # --- Consumo de techos
        tech = ctk.CTkFrame(self)
        tech.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
        ctk.CTkLabel(tech, text="Consumo de techos presupuestales", font=ctk.CTkFont(size=15, weight="bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", padx=10, pady=(8, 4))
        self.tech_rows = {}
        for i, (k, lab) in enumerate((("INI", "Iniciativas con impacto estratégico"),
                                      ("FE", "Actividades de funcionamiento"), ("OO", "Otras actividades operativas"),
                                      ("INV", "Actividades de inversión"), ("AP", "Actividades propias (autofinanciadas)"))):
            ctk.CTkLabel(tech, text=lab, anchor="w", width=260).grid(row=i + 1, column=0, sticky="w", padx=10, pady=3)
            bar = ctk.CTkProgressBar(tech, width=360)
            bar.grid(row=i + 1, column=1, padx=8)
            info = ctk.CTkLabel(tech, text="", anchor="w", width=420)
            info.grid(row=i + 1, column=2, sticky="w")
            st = ctk.CTkLabel(tech, text="", anchor="w", width=150, font=ctk.CTkFont(weight="bold"))
            st.grid(row=i + 1, column=3, sticky="w")
            self.tech_rows[k] = (bar, info, st)

    def refresh(self, preview: bool = True):
        uid = self.app.uid
        if uid is None:
            return
        u = self.c.unidad(uid)
        self.form.clear()
        self.form.set(u)
        if preview:
            self.preview_codes()
        tpl = self.c.plantilla_aplicable(uid)
        for r in tpl:
            try:
                r["_monto"] = self.c.monto_plantilla(r, u)
            except ValueError:
                r["_monto"] = None
        self.tpl_table.set_rows(tpl)
        for k, (bar, info, st) in self.tech_rows.items():
            cons = self.c.consumo_techo(uid, k)
            if k == "AP":
                bar.set(1.0 if cons["ingresos"] >= cons["egresos"] else 0.0)
                info.configure(text=f"Ingresos S/ {fmt_money(cons['ingresos'])} · Costos S/ {fmt_money(cons['egresos'])}")
            else:
                uso = cons["uso"] or 0.0
                bar.set(min(1.0, uso))
                bar.configure(progress_color=COLOR_BAD if uso > 1 else ("#2E86C1"))
                info.configure(text=f"Egresos S/ {fmt_money(cons['egresos'])} de S/ {fmt_money(cons['techo'])} "
                                    f"({fmt_pct(cons['uso'])}) · saldo S/ {fmt_money(cons['saldo'])}")
            st.configure(text=cons["estado"], text_color=status_color(cons["estado"]))

    @guarded
    def save_unidad(self):
        data = self.form.get()
        self.app.db.update("unidades_organicas", self.app.uid, data)
        self.app.db.log(self.app.role, "Datos de unidad", self.c.unidad(self.app.uid)["codigo_area"])
        messagebox.showinfo("Unidad", "Datos de la unidad guardados.")
        self.app.update_header()

    def fill_today(self):
        today = date.today().isoformat()
        for k in ("fecha_presentacion", "elaborado_fecha"):
            self.form.set_value(k, today)

    def _quantities(self) -> dict[str, int]:
        out = {}
        for k, e in self.qty.items():
            v = to_float(e.get(), 0) or 0
            if v < 0 or v > 500:
                raise ValueError("Las cantidades deben estar entre 0 y 500.")
            out[k] = int(v)
        return out

    @guarded
    def preview_codes(self):
        uid = self.app.uid
        area = self.c.unidad(uid)["codigo_area"]
        q = self._quantities()
        lines = [f"Unidad: {area} · Año {self.c.anio}", ""]
        for k, n in q.items():
            existing = len(self.c._existing_codes(k, uid))
            if n:
                first = self.c.next_code(k, uid)
                last = self.c.next_code(k, uid, n - 1)
                lines.append(f"{k}: {n} nuevos → {first} … {last}  (BD: {area}-{first} … {area}-{last})")
            else:
                lines.append(f"{k}: {existing} registrados; siguiente código {self.c.next_code(k, uid)}")
        self.codes_box.delete("1.0", "end")
        self.codes_box.insert("1.0", "\n".join(lines))

    @guarded
    def generate_codes(self):
        q = self._quantities()
        if not sum(q.values()):
            raise ValueError("Indique al menos una cantidad mayor que cero.")
        if not messagebox.askyesno("Generar códigos", f"Se crearán {sum(q.values())} registros con códigos "
                                                       "correlativos. ¿Continuar?"):
            return
        creados = self.c.generar_codigos(self.app.uid, q)
        self.app.db.log(self.app.role, "Generar códigos", f"{len(creados)} códigos")
        self.app.mark_dirty(except_tab=self)
        self.codes_box.delete("1.0", "end")
        self.codes_box.insert("1.0", "Códigos generados:\n" + "\n".join(f"{k:<4} {c:<9} → {bd}" for k, c, bd in creados))
        for e in self.qty.values():
            e.delete(0, "end")
            e.insert(0, "0")

    @guarded
    def run_funcionamiento(self):
        rep = self.c.correr_funcionamiento(self.app.uid, reemplazar=bool(self.replace_chk.get()))
        self.app.db.log(self.app.role, "Correr funcionamiento", str(rep))
        self.run_info.configure(text=f"✔ Plantilla ejecutada: {rep['FE']} egresos (S/ {fmt_money(rep['monto_FE'])}) y "
                                     f"{rep['FI']} ingresos (S/ {fmt_money(rep['monto_FI'])}) prorrateados 1/12. "
                                     f"Filas reemplazadas: {rep['eliminadas']}. Complete los montos en cero en la "
                                     "pestaña Funcionamiento.", text_color=COLOR_OK)
        self.app.mark_dirty(except_tab=self)
        self.refresh(preview=False)


# =============================================================================
# Pestaña 1 — Parámetros y catálogos (Administrador DPDI)
# =============================================================================
class CrudPanel(ctk.CTkFrame):
    """Mantenedor genérico tabla + formulario."""

    def __init__(self, master, app: "App", table: str, columns: list[tuple], fields: list[Field],
                 where: Callable[[], tuple[str, tuple]] = lambda: ("", ()), order: str = "id",
                 save_fn: Optional[Callable[[dict, Optional[int]], int]] = None,
                 delete_fn: Optional[Callable[[int], None]] = None, search_cols: tuple[str, ...] = (),
                 defaults: Optional[Callable[[], dict]] = None, extra_buttons: list = (), form_columns: int = 1,
                 transform: Optional[Callable[[dict], dict]] = None, rows_fn: Optional[Callable[[], list]] = None,
                 after_load: Optional[Callable[[dict, Form], None]] = None):
        super().__init__(master, fg_color="transparent")
        self.app, self.table, self.where, self.order = app, table, where, order
        self.save_fn, self.delete_fn, self.search_cols = save_fn, delete_fn, search_cols
        self.defaults, self.transform, self.rows_fn, self.after_load = defaults, transform, rows_fn, after_load
        self.current: Optional[int] = None
        left = ctk.CTkScrollableFrame(self, width=430)
        left.pack(side="left", fill="y", padx=(0, 6))
        self.form = Form(left, fields, columns=form_columns, label_width=130)
        self.form.pack(fill="x")
        button_bar(left, [("Nuevo", self.new, "gray55"), ("Guardar", self.save), ("Eliminar", self.delete, COLOR_BAD)]
                   + list(extra_buttons)).pack(anchor="w", pady=6)
        right = ctk.CTkFrame(self, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True)
        if search_cols:
            self.search = ctk.CTkEntry(right, placeholder_text="Buscar…")
            self.search.pack(fill="x", pady=(0, 4))
            self.search.bind("<KeyRelease>", lambda e: self.refresh())
        self.grid = DataTable(right, columns, height=20, on_select=self.load)
        self.grid.pack(fill="both", expand=True)
        self.count = ctk.CTkLabel(right, text="", text_color="gray40")
        self.count.pack(anchor="e")

    def refresh(self):
        if self.rows_fn:
            rows = self.rows_fn()
        else:
            w, p = self.where()
            rows = self.app.db.all(self.table, w, p, self.order)
        if self.search_cols and self.search.get().strip():
            q = self.search.get().strip().upper()
            rows = [r for r in rows if any(q in s(r.get(c)).upper() for c in self.search_cols)]
        self.grid.set_rows(rows)
        self.count.configure(text=f"{len(rows)} registros")
        self.form.reload_values()

    def load(self, rid):
        if rid is None:
            return
        row = self.app.db.get(self.table, rid)
        if row:
            self.current = rid
            self.form.clear()
            self.form.set(row)
            if self.after_load:
                self.after_load(row, self.form)

    def new(self):
        self.current = None
        self.form.clear()
        if self.defaults:
            self.form.set(self.defaults())

    @guarded
    def save(self):
        data = self.form.get()
        if self.transform:
            data = self.transform(data)
        if self.save_fn:
            rid = self.save_fn(data, self.current)
        else:
            rid = self.app.db.upsert(self.table, data, self.current)
        self.current = rid
        self.app.db.log(self.app.role, f"Guardar {self.table}", str(rid))
        self.refresh()
        self.grid.select(rid)
        self.app.mark_dirty()

    @guarded
    def delete(self):
        if not self.current:
            return
        if not messagebox.askyesno("Eliminar", "¿Eliminar el registro seleccionado?"):
            return
        if self.delete_fn:
            self.delete_fn(self.current)
        else:
            self.app.db.delete(self.table, self.current)
        self.new()
        self.refresh()
        self.app.mark_dirty()


class ParametrosTab(ctk.CTkFrame):
    def __init__(self, master, app: "App"):
        super().__init__(master, fg_color="transparent")
        self.app, self.c, db = app, app.c, app.db
        bar = ctk.CTkFrame(self, fg_color=("#E8EEF7", "gray20"), corner_radius=8)
        bar.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(bar, text="Actualización masiva:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=8)
        ctk.CTkButton(bar, text="⬇ Descargar plantilla de Parámetros y Catálogos (.xlsx)", width=10, fg_color="#1E8449",
                      command=self.download_template).pack(side="left", padx=4, pady=5)
        ctk.CTkButton(bar, text="⬆ Importar plantilla de Parámetros y Catálogos (.xlsx)", width=10, fg_color="#2E86C1",
                      command=self.import_template).pack(side="left", padx=4, pady=5)
        tv = ctk.CTkTabview(self)
        tv.pack(fill="both", expand=True)
        self.panels: list = []

        # --- Datos generales
        t = tv.add("Datos generales y supuestos")
        sc = ctk.CTkScrollableFrame(t)
        sc.pack(fill="both", expand=True)
        self.pform = Form(sc, [
            Field("institucion", "Institución", width=320), Field("codigo_formato", "Código del formato"),
            Field("version_formato", "Versión"), Field("fecha_aprobacion", "Fecha de aprobación"),
            Field("anio_formulacion", "Año de formulación"), Field("fecha_emision", "Fecha de emisión (AAAA-MM-DD)"),
            Field("fecha_limite", "Fecha límite (AAAA-MM-DD)"), Field("inflacion_proyectada", "Inflación proyectada (%)"),
            Field("tipo_cambio", "Tipo de cambio (S/ por US$)"), Field("valor_uit", "Valor UIT (S/)"),
            Field("rmv", "Remuneración mínima vital (S/)"), Field("reajuste_remunerativo_pct", "Reajuste remunerativo (%)"),
            Field("mes_inicio_reajuste", "Mes de reajuste (1–12)", "combo", [str(i) for i in range(1, 13)]),
            Field("tasa_morosidad_pct", "Tasa de morosidad (%)"),
            Field("estudiantes_pregrado", "Estudiantes pregrado"), Field("estudiantes_posgrado", "Estudiantes posgrado"),
            Field("pension_ref_pregrado", "Pensión ref. pregrado (S/)"),
            Field("pension_ref_posgrado", "Pensión ref. posgrado (S/)"),
            Field("cuotas_periodo", "N° cuotas por periodo"), Field("director_dpdi", "Director(a) DPDI"),
            Field("cargo_dpdi", "Cargo DPDI", width=320), Field("clave_admin", "Clave de administrador"),
        ], columns=2, label_width=190)
        self.pform.pack(fill="x", padx=6, pady=6)
        button_bar(sc, [("Guardar parámetros", self.save_params),
                        ("Importar parámetros desde Ficha Excel (.xlsx)", self.import_params, "gray40")]).pack(
            anchor="w", padx=6)
        ctk.CTkLabel(sc, text="Porcentajes: escriba 2.5, 2.5% o 0.025 (todos equivalen a 2.5 %).",
                     text_color="gray40").pack(anchor="w", padx=10, pady=4)

        # --- Unidades y techos
        t = tv.add("Unidades y techos")
        fields_u = [
            Field("codigo_area", "Código de área"), Field("macrogrupo", "Macrogrupo", "combo", MACROGRUPOS),
            Field("nombre_unidad", "Nombre de la unidad"), Field("centro_costo", "Centro de costo"),
            Field("responsable", "Responsable"), Field("cargo", "Cargo"), Field("correo", "Correo"),
            Field("genera_ingresos", "¿Genera ingresos?", "check"),
            Field("techo_iniciativas", "Techo iniciativas (S/)"), Field("techo_funcionamiento", "Techo funcionamiento (S/)"),
            Field("techo_otras_operativas", "Techo otras operativas (S/)"), Field("techo_inversion", "Techo inversión (S/)"),
            Field("techo_planilla_docente", "Techo planilla docente (S/)"),
            Field("techo_planilla_adm", "Techo planilla adm. (S/)"), Field("techo_marketing", "Techo marketing (S/)"),
            Field("meta_ingresos_func", "Meta ingresos funcionamiento (S/)"),
            Field("meta_resultado_otras", "Meta ingresos otras operativas (S/)"),
            Field("meta_resultado_propias", "Meta resultado act. propias (S/)"),
        ]
        p = CrudPanel(t, app, "unidades_organicas",
                      [("codigo_area", "Área", 80), ("macrogrupo", "Macrogrupo", 170), ("nombre_unidad", "Unidad", 300),
                       ("techo_iniciativas", "Techo INI", 100, "money"), ("techo_funcionamiento", "Techo FUN", 110, "money"),
                       ("techo_otras_operativas", "Techo OO", 100, "money"), ("techo_inversion", "Techo INV", 100, "money"),
                       ("techo_planilla_docente", "Planilla doc.", 100, "money"),
                       ("techo_planilla_adm", "Planilla adm.", 100, "money"), ("techo_marketing", "Marketing", 100, "money"),
                       ("meta_ingresos_func", "Meta ingresos", 110, "money")],
                      fields_u, order="macrogrupo, nombre_unidad", save_fn=lambda d, i: self.c.save_unidad(d, i),
                      delete_fn=self.c.delete_unidad, search_cols=("codigo_area", "nombre_unidad", "macrogrupo"),
                      defaults=lambda: {"macrogrupo": MACROGRUPOS[0]})
        p.pack(fill="both", expand=True)
        self.panels.append(p)

        # --- Base consolidada Cuentas NIIF + Partidas GG
        t = tv.add("Cuentas NIIF + Partidas GG")
        ctk.CTkLabel(t, text="Base única: cada cuenta contable NIIF lleva la «Partida» (GG) que el usuario elige en el "
                             "detalle y su descripción orientativa. Las cuentas sin Partida GG solo se usan en "
                             "Funcionamiento.", text_color="gray35", wraplength=1100, justify="left").pack(anchor="w",
                                                                                                          padx=6)
        p = CrudPanel(t, app, "catalogo_niif_partidas",
                      [("tipo_flujo", "Tipo", 75), ("partida_gg", "Partida (GG)", 230),
                       ("descripcion", "Descripción orientativa", 300), ("codigo_niif", "Código NIIF", 95),
                       ("partida_rubro", "Cuenta contable NIIF", 260), ("categoria", "Categoría", 190),
                       ("naturaleza", "Naturaleza", 180), ("resultado_funcion", "Resultado por función", 170),
                       ("activo", "Activo", 50, "int")],
                      [Field("tipo_flujo", "Tipo", "combo", ["Egresos", "Ingresos", "Inversion"]),
                       Field("partida_gg", "Partida (GG)", "combo_free",
                             lambda: sorted(set(self.c.partidas_gg("Egresos") + self.c.partidas_gg("Ingresos")
                                                + self.c.partidas_gg("Inversion"))), width=260),
                       Field("descripcion", "Descripción orientativa", "text", height=80),
                       Field("codigo_niif", "Código NIIF"), Field("partida_rubro", "Cuenta contable NIIF"),
                       Field("categoria", "Categoría", "combo_free",
                             lambda: [r["categoria"] for r in db.query(
                                 "SELECT DISTINCT categoria FROM catalogo_niif_partidas ORDER BY categoria")]),
                       Field("naturaleza", "Naturaleza"), Field("resultado_funcion", "Resultado por función"),
                       Field("activo", "Activo", "check")],
                      order="tipo_flujo, partida_gg='', partida_gg COLLATE NOCASE, partida_rubro",
                      save_fn=self.c.save_partida_niif, delete_fn=self.c.delete_partida_niif,
                      search_cols=("codigo_niif", "partida_rubro", "categoria", "tipo_flujo", "partida_gg", "descripcion"),
                      defaults=lambda: {"tipo_flujo": "Egresos", "activo": 1},
                      extra_buttons=[("Carga masiva NIIF (.xlsx/.csv)", self.carga_masiva, "gray40")])
        p.pack(fill="both", expand=True)
        self.panels.append(p)

        # --- Catálogos estratégicos
        t = tv.add("Catálogos estratégicos")
        top = ctk.CTkFrame(t, fg_color="transparent")
        top.pack(fill="x")
        ctk.CTkLabel(top, text="Catálogo:").pack(side="left", padx=6)
        self.cat_names = {"Ejes estratégicos (EJ)": "EJE", "Objetivos estratégicos PEI (OE)": "OE",
                          "Objetivos anuales POA (OBJ) — Eje → OBJ → OE": "OBJ", "Orígenes de iniciativa": "ORIGEN",
                          "Niveles de prioridad": "PRIORIDAD", "Sustentos de egresos": "SUSTENTO_EGR",
                          "Sustentos de ingresos": "SUSTENTO_ING", "Categorías otras operativas (OO)": "CAT_OO",
                          "Tipos de actividad propia": "TIPO_AP", "Tipos de inversión": "TIPO_INV",
                          "Clasificación de actividades": "TIPO_ACTIVIDAD"}
        self.cat_sel = ctk.CTkComboBox(top, values=list(self.cat_names), width=360, state="readonly",
                                       command=lambda v: self.cat_panel.refresh())
        self.cat_sel.set("Ejes estratégicos (EJ)")
        self.cat_sel.pack(side="left")
        ctk.CTkLabel(top, text="En los objetivos anuales registre el Eje al que pertenecen y el OE del PEI al que "
                               "contribuyen · Extra = indicador / alcance / criterio",
                     text_color="gray40").pack(side="left", padx=10)
        self.cat_panel = CrudPanel(
            t, app, "catalogos_estrategicos",
            [("codigo", "Código", 110), ("descripcion", "Descripción", 360), ("eje", "Eje", 70),
             ("padre", "OE del PEI", 80), ("extra", "Extra", 360), ("orden", "Orden", 60, "int")],
            [Field("codigo", "Código"), Field("descripcion", "Descripción", "text", height=70),
             Field("eje", "Eje estratégico (solo OBJ)", "combo_free", lambda: [""] + self.c.catalogo_values("EJE")),
             Field("padre", "OE del PEI (solo OBJ)", "combo_free", lambda: [""] + self.c.catalogo_values("OE")),
             Field("extra", "Extra", "text", height=70), Field("orden", "Orden")],
            where=lambda: ("tipo_catalogo = ?", (self.cat_names[self.cat_sel.get()],)), order="orden, codigo",
            save_fn=self._save_catalogo, search_cols=("codigo", "descripcion", "padre", "eje"))
        self.cat_panel.pack(fill="both", expand=True)
        self.panels.append(self.cat_panel)

        # --- Costos de referencia
        t = tv.add("Costos de referencia")
        p = CrudPanel(t, app, "costos_referencia",
                      [("concepto", "Concepto", 260), ("unidad_medida", "Unidad de medida", 160),
                       ("costo_unitario", "Costo unitario (S/)", 140, "money"), ("fuente", "Fuente / vigencia", 300)],
                      [Field("concepto", "Concepto"), Field("unidad_medida", "Unidad de medida"),
                       Field("costo_unitario", "Costo unitario (S/)"), Field("fuente", "Fuente / vigencia")],
                      transform=lambda d: dict(d, costo_unitario=to_float(d.get("costo_unitario"))))
        p.pack(fill="both", expand=True)
        self.panels.append(p)

        # --- Plantilla de funcionamiento
        t = tv.add("Plantilla de funcionamiento")
        self.unit_map: dict[str, Optional[int]] = {}
        p = CrudPanel(t, app, "plantilla_funcionamiento",
                      [("_unidad", "Ámbito", 110), ("macrogrupo", "Macrogrupo", 150), ("tipo_flujo", "Tipo", 80),
                       ("partida", "Partida / rubro", 260), ("base_calculo", "Base de cálculo", 160),
                       ("porcentaje", "%", 80, "pct"), ("monto_fijo", "Monto fijo (S/)", 120, "money"),
                       ("aplica_ajuste", "Reajuste/Morosidad", 120), ("sustento", "Sustento", 160),
                       ("activo", "Activo", 60, "int")],
                      [Field("_unidad", "Ámbito (global o unidad)", "combo", self._unit_choices),
                       Field("macrogrupo", "Macrogrupo", "combo", ["*"] + MACROGRUPOS),
                       Field("solo_genera_ingresos", "Solo unidades que generan ingresos", "check"),
                       Field("tipo_flujo", "Tipo", "combo", ["Egresos", "Ingresos"]),
                       Field("partida", "Partida / rubro", "search",
                             lambda: self.c.partidas("Egresos") + self.c.partidas("Ingresos"), width=300),
                       Field("base_calculo", "Base de cálculo", "combo", BASES_CALCULO),
                       Field("porcentaje", "Porcentaje de la base"), Field("monto_fijo", "Monto fijo anual (S/)"),
                       Field("aplica_ajuste", "Aplica reajuste / morosidad", "combo", ["Sí", "No"]),
                       Field("sustento", "Sustento", "combo",
                             lambda: self.c.catalogo_values("SUSTENTO_EGR") + self.c.catalogo_values("SUSTENTO_ING")),
                       Field("activo", "Activo", "check"), Field("orden", "Orden")],
                      rows_fn=self._plantilla_rows, save_fn=self._save_plantilla,
                      defaults=lambda: {"_unidad": "GLOBAL", "macrogrupo": "*", "tipo_flujo": "Egresos",
                                        "base_calculo": "MONTO_FIJO", "porcentaje": "1", "aplica_ajuste": "No",
                                        "activo": 1},
                      search_cols=("partida", "macrogrupo", "base_calculo", "_unidad"),
                      after_load=self._plantilla_scope)
        p.pack(fill="both", expand=True)
        self.panels.append(p)
        self.plantilla_panel = p

    def _save_catalogo(self, d: dict, rid: Optional[int]) -> int:
        tipo = self.cat_names[self.cat_sel.get()]
        data = dict(d, orden=int(to_float(d.get("orden"), 0) or 0))
        if tipo == "OBJ":
            if not data.get("eje") or not data.get("padre"):
                raise ValueError("Cada objetivo anual del POA debe tener su Eje estratégico y su OE del PEI.")
            data["eje"], data["padre"] = code_display(data["eje"]), code_display(data["padre"])
        else:
            data["eje"], data["padre"] = "", ""
        return self.c.save_catalogo(tipo, data, rid)

    @guarded
    def download_template(self):
        path = filedialog.asksaveasfilename(title="Descargar plantilla de Parámetros y Catálogos",
                                            defaultextension=".xlsx",
                                            initialfile=f"Plantilla_Parametros_Catalogos_POA_{self.c.anio}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        out = export_plantilla_parametros(self.c, path)
        self.app.db.log(self.app.role, "Descargar plantilla de parámetros", Path(path).name)
        if messagebox.askyesno("Plantilla", f"Plantilla generada:\n{out}\n\nEdítela y luego use «Importar plantilla»."
                                            "\n¿Abrir el archivo?"):
            self.app._open(out)

    @guarded
    def import_template(self):
        path = filedialog.askopenfilename(title="Importar plantilla de Parámetros y Catálogos",
                                          filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        if not messagebox.askyesno("Importar plantilla", "Se actualizarán o agregarán los parámetros, unidades, "
                                                         "cuentas, catálogos, costos y la plantilla de funcionamiento "
                                                         "del archivo (no se elimina nada). ¿Continuar?"):
            return
        rep = self.c.import_plantilla_parametros(path)
        self.app.db.log(self.app.role, "Importar plantilla de parámetros", Path(path).name)
        avisos = rep.pop("avisos")
        msg = "\n".join(f"{k}: {v} registros" for k, v in rep.items())
        if avisos:
            msg += "\n\nAvisos:\n" + "\n".join(f"• {a}" for a in avisos[:15])
        messagebox.showinfo("Plantilla importada", msg)
        self.refresh()
        self.app.mark_dirty(except_tab=self)
        self.app.update_header()

    def _unit_choices(self) -> list[str]:
        self.unit_map = {"GLOBAL": None}
        for u in self.c.unidades():
            self.unit_map[u["codigo_area"]] = u["id"]
        return list(self.unit_map)

    def _plantilla_scope(self, row: dict, form: Form) -> None:
        u = self.app.db.get("unidades_organicas", row["unidad_id"]) if row["unidad_id"] else None
        form.set_value("_unidad", u["codigo_area"] if u else "GLOBAL")

    def _plantilla_rows(self) -> list[dict]:
        areas = {u["id"]: u["codigo_area"] for u in self.c.unidades()}
        rows = self.app.db.all("plantilla_funcionamiento", order="unidad_id IS NOT NULL, unidad_id, orden, id")
        for r in rows:
            r["_unidad"] = areas.get(r["unidad_id"], "GLOBAL")
        return rows

    def _save_plantilla(self, d: dict, rid: Optional[int]) -> int:
        self._unit_choices()
        if not d.get("partida"):
            raise ValueError("Seleccione la partida.")
        data = dict(d)
        data["unidad_id"] = self.unit_map.get(d.pop("_unidad", "GLOBAL") or "GLOBAL")
        data.pop("_unidad", None)
        data["porcentaje"] = to_float(d.get("porcentaje"), 1.0)
        data["monto_fijo"] = to_float(d.get("monto_fijo"), 0.0)
        data["orden"] = int(to_float(d.get("orden"), 0) or 0)
        return self.app.db.upsert("plantilla_funcionamiento", data, rid)

    def refresh(self):
        p = self.app.db.parametros()
        self.pform.clear()
        self.pform.set(p)
        for k in ("inflacion_proyectada", "reajuste_remunerativo_pct", "tasa_morosidad_pct"):
            v = to_float(p.get(k))
            self.pform.set_value(k, "" if v is None else f"{v * 100:g}%")
        for pnl in self.panels:
            pnl.refresh()

    @guarded
    def save_params(self):
        from controllers import parse_pct
        d = self.pform.get()
        for k in ("inflacion_proyectada", "reajuste_remunerativo_pct", "tasa_morosidad_pct"):
            d[k] = parse_pct(d.get(k))
        for k in ("tipo_cambio", "valor_uit", "rmv", "estudiantes_pregrado", "estudiantes_posgrado",
                  "pension_ref_pregrado", "pension_ref_posgrado", "cuotas_periodo"):
            d[k] = to_float(d.get(k))
        d["mes_inicio_reajuste"] = int(to_float(d.get("mes_inicio_reajuste"), 1) or 1)
        if not d.get("tipo_cambio"):
            raise ValueError("El tipo de cambio es obligatorio.")
        if not d.get("clave_admin"):
            d.pop("clave_admin")
        self.app.db.save_parametros(d)
        self.app.db.log(self.app.role, "Parámetros", "guardados")
        messagebox.showinfo("Parámetros", "Parámetros guardados. Los cálculos se actualizan en todas las pestañas.")
        self.app.mark_dirty()
        self.app.update_header()

    @guarded
    def import_params(self):
        path = filedialog.askopenfilename(title="Ficha_Formulacion_POA (.xlsx)", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        rep = self.c.import_parametros(path)
        messagebox.showinfo("Parámetros importados", "\n".join(f"{k}: {v}" for k, v in rep.items()))
        self.refresh()
        self.app.mark_dirty()

    @guarded
    def carga_masiva(self):
        path = filedialog.askopenfilename(title="Índice de cuentas NIIF", filetypes=[("Excel / CSV", "*.xlsx *.csv")])
        if not path:
            return
        rep = self.c.carga_masiva_niif(path)
        messagebox.showinfo("Carga masiva NIIF", "\n".join(f"{k}: {v}" for k, v in rep.items()))
        self.refresh()


# =============================================================================
# Pestañas maestro-detalle (Iniciativas, Otras operativas, Act. propias, Inversión)
# =============================================================================
DETAIL_COLUMNS = [("tipo", "Tipo", 75), ("partida", "Partida", 220), ("categoria", "Categoría", 150),
                  ("descripcion", "Descripción", 220), ("cantidad", "Cant.", 60, "num"), ("moneda", "Mon.", 45),
                  ("_cu", "C.U. (S/)", 95, "money"), ("_total", "Total (S/)", 105, "money"),
                  *[(f"_m{i}", m, 82, "money") for i, m in enumerate(MESES)],
                  ("_prog", "Programado (S/)", 105, "money"), ("_control", "Control", 70), ("sustento", "Sustento", 150)]


class MasterDetailTab(ctk.CTkFrame):
    CONFIG = {
        "INI": dict(title="Iniciativas con impacto estratégico", sheet="Iniciativas + Detalle_Iniciativas",
                    tipos=["Egresos", "Ingresos"], gg=True, techo="INI"),
        "OO": dict(title="Otras actividades operativas", sheet="Otras_Operat + Detalle_Otras_Operativas",
                   tipos=["Egresos", "Ingresos"], gg=False, techo="OO"),
        "AP": dict(title="Actividades propias", sheet="Act_Prop + Detalle_Act_Prop", tipos=["Ingresos", "Egresos"],
                   gg=False, techo="AP"),
        "INV": dict(title="Actividades de inversión (CAPEX)", sheet="Inversion + Detalle_Inversiones",
                    tipos=["Inversion"], gg=False, techo="INV"),
    }

    def __init__(self, master, app: "App", kind: str):
        super().__init__(master, fg_color="transparent")
        self.app, self.c, self.kind = app, app.c, kind
        self.cfg = self.CONFIG[kind]
        self.mid: Optional[int] = None
        self.did: Optional[int] = None
        paned = tk.PanedWindow(self, orient="vertical", sashwidth=7, sashrelief="raised", bg="#C8D2E0", bd=0)
        paned.pack(fill="both", expand=True)
        top = ctk.CTkFrame(paned)
        bottom = ctk.CTkFrame(paned)
        paned.add(top, minsize=230, height=380)
        paned.add(bottom, minsize=260)

        # ---------------- maestro
        head = ctk.CTkFrame(top, fg_color="transparent")
        head.pack(fill="x", padx=6, pady=(4, 0))
        ctk.CTkLabel(head, text=f"{self.cfg['title']} — maestro ({self.cfg['sheet'].split(' + ')[0]})",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        self.summary = ctk.CTkLabel(head, text="", font=ctk.CTkFont(size=12, weight="bold"))
        self.summary.pack(side="right")
        body = ctk.CTkFrame(top, fg_color="transparent")
        body.pack(fill="both", expand=True)
        left = ctk.CTkScrollableFrame(body, width=560)
        left.pack(side="left", fill="y", padx=4, pady=4)
        self.mform = Form(left, self._master_fields(), columns=2, on_change=self._master_changed, label_width=105)
        self.mform.pack(fill="x")
        mbtns = [("Nuevo", self.new_master, "gray55"), ("Guardar", self.save_master),
                 ("Duplicar", self.duplicate_master, "gray40"), ("Eliminar", self.delete_master, COLOR_BAD)]
        if kind == "AP":
            mbtns.append(("Generar ingreso desde tarifa", self.ap_ingreso, "#117A65"))
        button_bar(left, mbtns).pack(anchor="w", pady=4)
        self.mtable = DataTable(body, self._master_columns(), height=8, on_select=self.load_master)
        self.mtable.pack(side="left", fill="both", expand=True, padx=4, pady=4)

        # ---------------- detalle
        head2 = ctk.CTkFrame(bottom, fg_color="transparent")
        head2.pack(fill="x", padx=6, pady=(4, 0))
        self.det_title = ctk.CTkLabel(head2, text="Detalle — seleccione un registro maestro",
                                      font=ctk.CTkFont(size=14, weight="bold"))
        self.det_title.pack(side="left")
        self.det_info = ctk.CTkLabel(head2, text="")
        self.det_info.pack(side="right")
        body2 = ctk.CTkFrame(bottom, fg_color="transparent")
        body2.pack(fill="both", expand=True)
        left2 = ctk.CTkScrollableFrame(body2, width=560)
        left2.pack(side="left", fill="y", padx=4, pady=4)
        self.dform = Form(left2, self._detail_fields(), columns=2, on_change=self._detail_changed, label_width=105)
        self.dform.pack(fill="x")
        self.months = MonthGrid(left2, "amount", on_change=self._detail_changed, total_fn=self._detail_total)
        self.months.pack(fill="x", pady=4)
        self.dcalc = Form(left2, [Field("_prog", "Programado (S/)", "calc", width=140),
                                  Field("_control", "Control de distribución", "calc", width=140)], columns=2,
                          label_width=105)
        self.dcalc.pack(fill="x")
        button_bar(left2, [("Nuevo ítem", self.new_detail, "gray55"), ("Guardar ítem", self.save_detail),
                           ("Duplicar ítem", self.duplicate_detail, "gray40"),
                           ("Eliminar ítem", self.delete_detail, COLOR_BAD)]).pack(anchor="w", pady=4)
        self.dtable = DataTable(body2, DETAIL_COLUMNS, height=8, on_select=self.load_detail)
        self.dtable.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        self.new_detail()

    # ---------------------------------------------------- definición de campos
    def _master_fields(self) -> list[Field]:
        c = self.c
        k = self.kind
        if k == "INI":
            return [
                Field("codigo", "Código"),
                Field("eje", "1) Eje estratégico", "combo", c.ejes_display),
                Field("obj_poa", "2) Objetivo anual del POA (del eje)", "combo",
                      lambda: c.objs_por_eje(self.mform.value("eje")) if hasattr(self, "mform") else [],
                      span=2, width=420),
                Field("_oe", "3) OE del PEI (automático)", "calc", span=2, width=420, anchor="w"),
                Field("nombre", "Nombre de la iniciativa", span=2, width=420),
                Field("descripcion", "Descripción y justificación", "text", span=2, width=420),
                Field("origen", "Origen", "combo", lambda: c.catalogo_values("ORIGEN")),
                Field("directriz", "Directriz que atiende"),
                Field("indicador", "Indicador"),
                Field("unidad_medida", "Unidad de medida", "combo", c.unidades_medida()),
                Field("linea_base", "Línea base"), Field("meta_anual", "Meta anual"), Field("t1", "T1"),
                Field("t2", "T2"), Field("t3", "T3"),
                Field("_control_meta", "Control de meta", "calc", width=120, anchor="center"),
                Field("responsable", "Responsable"),
                Field("prioridad", "Prioridad (1–4)", "combo", lambda: c.catalogo_values("PRIORIDAD", "both")),
                Field("observaciones", "Observaciones DPDI", "text", span=2, width=420, height=45),
            ]
        if k == "OO":
            return [
                Field("codigo", "Código"),
                Field("categoria", "Categoría", "combo", lambda: c.catalogo_values("CAT_OO", "both")),
                Field("nombre", "Nombre de la actividad", span=2, width=420),
                Field("descripcion", "Descripción y justificación", "text", span=2, width=420),
                Field("responsable", "Responsable"),
                Field("observaciones", "Observaciones", "text", span=2, width=420, height=45),
            ]
        if k == "AP":
            return [
                Field("codigo", "Código"), Field("tipo", "Tipo", "combo", lambda: c.catalogo_values("TIPO_AP")),
                Field("nombre", "Nombre de la actividad", span=2, width=420),
                Field("publico_objetivo", "Público objetivo"),
                Field("mes_ejecucion", "Mes de ejecución", "combo", MESES),
                Field("participantes", "Participantes proyectados"), Field("tarifa", "Tarifa por participante (S/)"),
                Field("becas_pct", "Becas y descuentos (%)"), Field("responsable", "Responsable"),
                Field("sustento", "Sustento", "combo", lambda: c.catalogo_values("SUSTENTO_ING")),
                Field("_tarifa_neta", "Tarifa neta (S/)", "calc", width=120),
                Field("observaciones", "Observaciones", "text", span=2, width=420, height=45),
            ]
        return [
            Field("codigo", "Código"),
            Field("tipo_inversion", "Tipo de inversión", "combo", lambda: c.catalogo_values("TIPO_INV")),
            Field("proyecto", "Proyecto de inversión", span=2, width=420),
            Field("prioridad", "Prioridad (1–4)", "combo", lambda: c.catalogo_values("PRIORIDAD", "both")),
            Field("ficha_sustento", "Ficha de sustento adjunta", "combo", ["Sí", "No"]),
            Field("observaciones", "Observaciones", "text", span=2, width=420, height=45),
        ]

    def _master_columns(self) -> list[tuple]:
        money = [("_ingresos", "Ingresos Total (S/)", 120, "money"), ("_egresos", "Costo / Gasto (S/)", 120, "money"),
                 ("_resultado", "Resultado (S/)", 115, "money"), ("_margen", "Margen", 70, "pct")]
        if self.kind == "INI":
            return [("codigo", "Código", 70), ("eje", "Eje estratégico", 160), ("nombre", "Nombre", 240),
                    ("obj_poa", "OBJ POA", 70), ("oe", "OE", 50), ("unidad_medida", "U. medida", 90),
                    ("meta_anual", "Meta", 60, "num"),
                    ("_control_meta", "Control meta", 90), ("responsable", "Responsable", 150),
                    ("prioridad", "Prior.", 50, "int"), *money]
        if self.kind == "OO":
            return [("codigo", "Código", 70), ("categoria", "Categoría", 200), ("nombre", "Nombre", 260),
                    ("responsable", "Responsable", 160), *money]
        if self.kind == "AP":
            return [("codigo", "Código", 70), ("tipo", "Tipo", 90), ("nombre", "Nombre", 250),
                    ("mes_ejecucion", "Mes", 50), ("_participantes", "Particip.", 70, "num"),
                    ("_tarifa_neta", "Tarifa neta", 90, "money"), *money,
                    ("_punto_equilibrio", "P. equilibrio", 90, "int"), ("_estado", "Estado", 110)]
        return [("codigo", "Código", 75), ("tipo_inversion", "Tipo de inversión", 190), ("proyecto", "Proyecto", 330),
                ("_egresos", "Programado (S/)", 120, "money"), ("prioridad", "Prior.", 55, "int"),
                ("ficha_sustento", "Ficha", 55), ("_n_detalle", "Ítems", 50, "int")]

    def _detail_fields(self) -> list[Field]:
        c = self.c
        tipos = self.cfg["tipos"]
        f = [Field("tipo", "Tipo", "combo", tipos),
             Field("partida", "Partida", "partida", lambda: c.partidas_gg(self._tipo_detalle()), width=300,
                   describe=lambda v: c.descripcion_partida(v, self._tipo_detalle())),
             Field("_desc_partida", "Descripción orientativa", "calc", span=2, width=420, anchor="w"),
             Field("_categoria", "Categoría", "calc", width=200, anchor="w"),
              Field("descripcion", "Descripción del bien o servicio", "text", span=2, width=420,
                    height=70 if self.kind == "INV" else 45),
              Field("cantidad", "Cantidad"), Field("moneda", "Moneda", "combo", ["S/", "US$"]),
              Field("costo_unit_moneda", "Costo unitario (moneda)"),
              Field("_cu", "Costo unitario (S/)", "calc", width=140), Field("_total", "Total (S/)", "calc", width=140),
              Field("sustento", "Sustento", "combo", lambda: self._sustentos()),
              Field("observaciones", "Observaciones", span=2, width=420)]
        return f

    def _tipo_detalle(self) -> str:
        if hasattr(self, "dform"):
            return self.dform.value("tipo") or self.cfg["tipos"][0]
        return self.cfg["tipos"][0]

    def _sustentos(self) -> list[str]:
        tipo = self.dform.value("tipo") if hasattr(self, "dform") else ""
        return self.c.catalogo_values("SUSTENTO_ING" if tipo == "Ingresos" else "SUSTENTO_EGR")

    # --------------------------------------------------------------- maestro
    def refresh(self):
        if self.app.uid is None:
            return
        self.mform.reload_values()
        self.dform.reload_values()
        rows = self.c.list_masters(self.kind, self.app.uid)
        self.mtable.set_rows(rows, tag_fn=self._master_tag)
        cons = self.c.consumo_techo(self.app.uid, self.cfg["techo"])
        ing = sum(r["_ingresos"] for r in rows)
        egr = sum(r["_egresos"] for r in rows)
        res = ing - egr
        txt = (f"Ingresos S/ {fmt_money(ing)}   ·   Costo S/ {fmt_money(egr)}   ·   Resultado S/ {fmt_money(res)}   ·   "
               f"Margen {fmt_pct(res / ing) if ing else '—'}")
        if self.kind != "AP":
            txt += f"   ·   Techo S/ {fmt_money(cons['techo'])} ({fmt_pct(cons['uso'])}) → {cons['estado']}"
        else:
            txt += f"   ·   {cons['estado'] or 'Sin movimientos'}"
        self.summary.configure(text=txt, text_color=status_color(cons["estado"]))
        if self.mid and str(self.mid) in self.mtable._rows:
            self.mtable.select(self.mid)
            self._load_details()
        else:
            self.mid = None
            self.mform.clear()
            self._load_details()

    def _master_tag(self, r: dict) -> str:
        name = r.get("nombre") if self.kind != "INV" else r.get("proyecto")
        if not s(name):
            return "placeholder"
        if self.kind == "INI" and r.get("_control_meta") != "OK":
            return "bad"
        if self.kind == "AP" and r.get("_estado") == "Déficit":
            return "bad"
        if self.kind == "INV" and r["_egresos"] > 0 and r.get("ficha_sustento") != "Sí":
            return "bad"
        return ""

    def load_master(self, mid):
        if mid is None:
            return
        m = self.c.get_master(self.kind, mid)
        if not m:
            return
        self.mid = mid
        self.mform.clear()
        data = dict(m)
        if self.kind == "INI":
            data["eje"] = self.c.eje_display(m["eje"])
            objs = self.c.objs_por_eje(data["eje"])
            self.mform.set_combo_values("obj_poa", objs)
            data["obj_poa"] = next((v for v in objs if code_display(v) == m["obj_poa"]), m["obj_poa"])
            data["prioridad"] = next((v for v in self.c.catalogo_values("PRIORIDAD", "both")
                                      if code_display(v) == s(m["prioridad"])), s(m["prioridad"]))
        elif self.kind == "OO":
            data["categoria"] = next((v for v in self.c.catalogo_values("CAT_OO", "both")
                                      if v.split(" · ", 1)[-1] == m["categoria"]), m["categoria"])
        elif self.kind == "AP":
            data["becas_pct"] = "" if m["becas_pct"] is None else f"{m['becas_pct'] * 100:g}%"
        elif self.kind == "INV":
            data["prioridad"] = next((v for v in self.c.catalogo_values("PRIORIDAD", "both")
                                      if code_display(v) == s(m["prioridad"])), s(m["prioridad"]))
        self.mform.set(data)
        self._master_changed()
        self._load_details()

    def _master_changed(self):
        if self.kind == "INI":
            d = self.mform.get()
            # Eje → Objetivos anuales del POA de ese eje → OE del PEI automático
            objs = self.c.objs_por_eje(d.get("eje", ""))
            self.mform.set_combo_values("obj_poa", objs)
            if d.get("obj_poa") and code_display(d["obj_poa"]) not in [code_display(o) for o in objs]:
                self.mform.set_value("obj_poa", "")
                d["obj_poa"] = ""
            if d.get("obj_poa"):
                oe = self.c.oe_de_obj(d["obj_poa"])
                self.mform.set_value("_oe", (self.c.oe_display(oe) if oe else "(objetivo sin OE vinculado)",
                                             "gray20" if oe else COLOR_WARN))
            elif d.get("eje") and not objs:
                self.mform.set_value("_oe", ("El eje no tiene objetivos anuales vinculados (ver Catálogos)", COLOR_WARN))
            else:
                self.mform.set_value("_oe", ("Seleccione el eje y luego el objetivo anual", "gray45"))
            self.mform.set_value("_control_meta", control_meta(dict(d, nombre=d.get("nombre") or "x")))
        elif self.kind == "AP":
            from controllers import parse_pct
            d = self.mform.get()
            t = to_float(d.get("tarifa"))
            b = parse_pct(d.get("becas_pct")) or 0.0
            self.mform.set_value("_tarifa_neta", fmt_money(t * (1 - b)) if t is not None else "")

    def new_master(self):
        self.mid = None
        self.mtable.tree.selection_remove(self.mtable.tree.selection())
        self.mform.clear()
        self.mform.set_value("codigo", self.c.next_code(self.kind, self.app.uid))
        if self.kind == "INV":
            self.mform.set_value("ficha_sustento", "No")
        self._load_details()

    @guarded
    def save_master(self):
        data = self.mform.get()
        if self.kind == "INI" and not data.get("nombre"):
            raise ValueError("Registre el nombre de la iniciativa.")
        self.mid = self.c.save_master(self.kind, self.app.uid, data, self.mid)
        self.app.db.log(self.app.role, f"Guardar {self.kind}", data.get("codigo", ""))
        self.refresh()
        self.app.mark_dirty(except_tab=self)

    @guarded
    def duplicate_master(self):
        if not self.mid:
            return
        src = self.c.get_master(self.kind, self.mid)
        data = {k: v for k, v in src.items() if not k.startswith("_") and k not in ("id", "codigo")}
        new_id = self.c.save_master(self.kind, self.app.uid, data)
        for dd in self.c.list_detalle(self.kind, self.mid):
            self.c.save_detalle(self.kind, new_id, {k: v for k, v in dd.items() if not k.startswith("_") and k != "id"})
        self.mid = new_id
        self.refresh()
        self.app.mark_dirty(except_tab=self)

    @guarded
    def delete_master(self):
        if not self.mid:
            return
        if not messagebox.askyesno("Eliminar", "¿Eliminar el registro y todo su detalle?"):
            return
        self.c.delete_master(self.kind, self.mid)
        self.mid = None
        self.refresh()
        self.app.mark_dirty(except_tab=self)

    @guarded
    def ap_ingreso(self):
        if not self.mid:
            raise ValueError("Guarde primero la actividad propia.")
        self.save_master()
        self.c.crear_ingreso_desde_tarifa(self.mid)
        self.refresh()

    # --------------------------------------------------------------- detalle
    def _load_details(self):
        if not self.mid:
            self.dtable.set_rows([])
            self.det_title.configure(text="Detalle — guarde o seleccione un registro maestro")
            self.det_info.configure(text="")
            return
        m = self.c.get_master(self.kind, self.mid)
        name = m.get("nombre") or m.get("proyecto") or ""
        self.det_title.configure(text=f"Detalle de {m['codigo']} · {name[:70]}")
        rows = self.c.list_detalle(self.kind, self.mid)
        for r in rows:
            cc = r["_c"]
            r.update(_cu=cc["costo_unit_soles"], _total=cc["total"], _prog=cc["programado"], _control=cc["control"])
            for i, v in enumerate(cc["meses"]):
                r[f"_m{i}"] = v or None
        self.dtable.set_rows(rows, tag_fn=lambda r: "bad" if r["_control"] == "Revisar" or (
            r["_total"] > 0 and not r["sustento"]) else "")
        extra = ""
        if self.kind == "AP":
            pe = m.get("_punto_equilibrio")
            extra = (f" · Punto de equilibrio: {pe if pe is not None else '—'} participantes · "
                     f"Estado: {m.get('_estado') or '—'}")
        self.det_info.configure(text=f"{len(rows)} ítems · Ingresos S/ {fmt_money(m['_ingresos'])} · "
                                     f"Egresos S/ {fmt_money(m['_egresos'])} · Resultado S/ {fmt_money(m['_resultado'])}"
                                     + extra,
                                text_color=COLOR_BAD if self.kind == "AP" and m.get("_estado") == "Déficit" else "gray20")

    def _detail_row_from_form(self) -> dict:
        d = self.dform.get()
        for i, v in enumerate(self.months.get()):
            d[MCOLS[i]] = v
        return d

    def _detail_total(self) -> float:
        return calc_detalle(self._detail_row_from_form(), self.c.params)["total"]

    def _detail_changed(self):
        d = self._detail_row_from_form()
        tipo = d.get("tipo") or self.cfg["tipos"][0]
        cc = calc_detalle(d, self.c.params)
        opciones = self.c.partidas_gg(tipo)
        self.dform.set_combo_values("partida", opciones)
        if d.get("partida") and d["partida"] not in opciones:
            nuevo = self.c.a_partida_gg(d["partida"], tipo)
            self.dform.set_value("partida", nuevo if nuevo in opciones else "")
            d["partida"] = self.dform.value("partida")
        desc = self.c.descripcion_partida(d.get("partida", ""), tipo) if d.get("partida") else ""
        self.dform.set_value("_desc_partida", (desc or "Seleccione la Partida (🔍 para buscar con descripción)",
                                               "gray20" if desc else "gray45"))
        cat = self.c.categoria_de(d.get("partida", ""), tipo)
        self.dform.set_value("_categoria", (cat or ("(partida fuera del catálogo)" if d.get("partida") else ""),
                                            "gray20" if cat else COLOR_WARN))
        self.dform.set_value("_cu", ("" if cc["costo_unit_soles"] is None else fmt_money(cc["costo_unit_soles"]),
                                     "gray20"))
        self.dform.set_value("_total", (fmt_money(cc["total"]), "gray20"))
        self.dcalc.set_value("_prog", (fmt_money(cc["programado"]), "gray20"))
        self.dcalc.set_value("_control", cc["control"])
        self.dform.set_combo_values("sustento", self._sustentos())

    def load_detail(self, did):
        if did is None:
            return
        row = self.app.db.get(DETAIL_TABLE[self.kind], did)
        if not row:
            return
        self.did = did
        self.dform.clear()
        self.dform.set(row)
        self.months.set([row[m] for m in MCOLS])
        self._detail_changed()

    def new_detail(self):
        self.did = None
        if hasattr(self, "dtable"):
            self.dtable.tree.selection_remove(self.dtable.tree.selection())
        self.dform.clear()
        self.months.clear()
        self.dform.set({"tipo": self.cfg["tipos"][0], "moneda": "S/"})
        self._detail_changed()

    @guarded
    def save_detail(self):
        if not self.mid:
            raise ValueError("Seleccione o guarde primero el registro maestro.")
        d = self._detail_row_from_form()
        if not d.get("partida"):
            raise ValueError("Seleccione la Partida.")
        self.did = self.c.save_detalle(self.kind, self.mid, d, self.did)
        self.refresh()
        self.dtable.select(self.did)
        self.app.mark_dirty(except_tab=self)

    @guarded
    def duplicate_detail(self):
        if not self.did:
            return
        self.did = None
        self.save_detail()

    @guarded
    def delete_detail(self):
        if not self.did:
            return
        if not messagebox.askyesno("Eliminar", "¿Eliminar el ítem seleccionado?"):
            return
        self.c.delete_detalle(self.kind, self.did)
        self.new_detail()
        self.refresh()
        self.app.mark_dirty(except_tab=self)


# =============================================================================
# Pestaña 3 — Funcionamiento (Ingresos / Egresos)
# =============================================================================
class FlatTab(ctk.CTkFrame):
    def __init__(self, master, app: "App", kind: str):
        super().__init__(master, fg_color="transparent")
        self.app, self.c, self.kind = app, app.c, kind
        self.rid: Optional[int] = None
        fi = kind == "FI"
        left = ctk.CTkScrollableFrame(self, width=600)
        left.pack(side="left", fill="y", padx=4, pady=4)
        tipo = "Ingresos" if fi else "Egresos"
        fields = [Field("codigo", "Código"),
                  Field("partida", "Partida / rubro", "search", lambda: self.c.partidas(tipo), width=300),
                  Field("_categoria", "Categoría", "calc", width=200, anchor="w"),
                  Field("cantidad", "Cantidad proyectada" if fi else "Cantidad"),
                  Field("moneda", "Moneda", "combo", ["S/", "US$"])]
        if fi:
            fields += [Field("tarifa_moneda", "Tarifa por cobro (moneda)"),
                       Field("_soles", "Tarifa por cobro (S/)", "calc", width=140),
                       Field("becas_pct", "Becas y descuentos (%)"),
                       Field("aplica_morosidad", "¿Aplica morosidad?", "combo", ["Sí", "No"])]
        else:
            fields += [Field("costo_unit_moneda", "Costo unitario por pago (moneda)"),
                       Field("_soles", "Costo unitario por pago (S/)", "calc", width=140),
                       Field("_coste", "Coste total (S/)", "calc", width=140),
                       Field("aplica_reajuste", "¿Aplica reajuste?", "combo", ["Sí", "No"])]
        fields += [Field("sustento", "Sustento", "combo",
                         lambda: self.c.catalogo_values("SUSTENTO_ING" if fi else "SUSTENTO_EGR")),
                   Field("observaciones", "Observaciones", span=2, width=440)]
        self.form = Form(left, fields, columns=2, on_change=self._changed, label_width=115)
        self.form.pack(fill="x")
        self.months = MonthGrid(left, "mark", on_change=self._changed, show_calc=True,
                                title=("Meses de cobro: «x» o factor (ej. 0.08333333) · fila inferior = ingreso mensual neto (S/)"
                                       if fi else "Meses de pago: «x» o factor (ej. 0.08333333) · fila inferior = monto mensual (S/)"))
        self.months.pack(fill="x", pady=4)
        if fi:
            calc = [Field("_bruto", "Ingreso bruto anual (S/)", "calc", width=150),
                    Field("_desc", "Becas, descuentos y morosidad (S/)", "calc", width=150),
                    Field("_neto", "Ingreso neto anual (S/)", "calc", width=150),
                    Field("_n", "N° de meses de cobro", "calc", width=150)]
        else:
            calc = [Field("_total", "Total anual (S/)", "calc", width=150),
                    Field("_n", "N° de pagos", "calc", width=150),
                    Field("_control", "Control de distribución", "calc", width=150),
                    Field("_reaj", "Reajuste vigente", "calc", width=150)]
        self.calc = Form(left, calc, columns=2, label_width=115)
        self.calc.pack(fill="x")
        button_bar(left, [("Nuevo", self.new, "gray55"), ("Guardar", self.save), ("Duplicar", self.duplicate, "gray40"),
                          ("Eliminar", self.delete, COLOR_BAD), ("Renumerar códigos", self.renumber, "gray40")]).pack(
            anchor="w", pady=4)
        right = ctk.CTkFrame(self, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        self.info = ctk.CTkLabel(right, text="", font=ctk.CTkFont(size=12, weight="bold"))
        self.info.pack(anchor="w")
        if fi:
            cols = [("codigo", "Código", 70), ("partida", "Partida / rubro", 230), ("categoria", "Categoría", 150),
                    ("cantidad", "Cant.", 60, "num"), ("moneda", "Mon.", 45), ("_soles", "Tarifa (S/)", 100, "money"),
                    ("becas_pct", "Becas", 60, "pct"), ("aplica_morosidad", "Moros.", 55),
                    *[(f"_m{i}", m, 90, "money") for i, m in enumerate(MESES)],
                    ("_bruto", "Bruto anual", 110, "money"), ("_desc", "Desc. y moros.", 100, "money"),
                    ("_neto", "Neto anual", 110, "money"), ("sustento", "Sustento", 160)]
        else:
            cols = [("codigo", "Código", 70), ("partida", "Partida / rubro", 230), ("categoria", "Categoría", 160),
                    ("cantidad", "Cant.", 55, "num"), ("moneda", "Mon.", 45), ("_soles", "C.U. por pago (S/)", 110, "money"),
                    ("_coste", "Coste total", 110, "money"), ("aplica_reajuste", "Reaj.", 50),
                    *[(f"_m{i}", m, 90, "money") for i, m in enumerate(MESES)],
                    ("_total", "Total anual", 115, "money"), ("_n", "Pagos", 50, "int"), ("_control", "Control", 70),
                    ("sustento", "Sustento", 150)]
        self.table = DataTable(right, cols, height=22, on_select=self.load)
        self.table.pack(fill="both", expand=True)

    def refresh(self):
        if self.app.uid is None:
            return
        self.form.reload_values()
        rows = self.c.list_flat(self.kind, self.app.uid)
        for r in rows:
            cc = r["_c"]
            if self.kind == "FI":
                r.update(_soles=cc["tarifa_soles"], _bruto=cc["bruto"], _desc=cc["descuentos"], _neto=cc["neto"])
            else:
                r.update(_soles=cc["costo_unit_soles"], _coste=cc["coste_total"], _total=cc["total"],
                         _n=cc["n_pagos"], _control=cc["control"])
            for i, v in enumerate(cc["meses"]):
                r[f"_m{i}"] = v or None
        self.table.set_rows(rows, tag_fn=self._tag)
        if self.kind == "FI":
            tot = sum(r["_neto"] for r in rows)
            u = self.c.unidad(self.app.uid)
            meta = u["meta_ingresos_func"] or 0
            self.info.configure(text=f"{len(rows)} conceptos · Ingreso neto anual S/ {fmt_money(tot)} · Meta S/ "
                                     f"{fmt_money(meta)} ({fmt_pct(tot / meta) if meta else 'sin meta'}) · Morosidad "
                                     f"{fmt_pct(self.c.params.morosidad)}")
        else:
            cons = self.c.consumo_techo(self.app.uid, "FE")
            p = self.c.params
            self.info.configure(text=f"{len(rows)} conceptos · Total anual S/ {fmt_money(cons['egresos'])} · Techo S/ "
                                     f"{fmt_money(cons['techo'])} ({fmt_pct(cons['uso'])}) → {cons['estado']} · "
                                     f"Reajuste {fmt_pct(p.reajuste)} desde {MESES[p.mes_reajuste - 1]}",
                                text_color=status_color(cons["estado"]))

    def _tag(self, r: dict) -> str:
        if self.kind == "FE" and (r["_control"] == "Revisar" or (r["_total"] > 0 and not r["sustento"])
                                  or (to_float(r["costo_unit_moneda"]) is not None and r["_n"] == 0)):
            return "bad"
        if self.kind == "FI" and ((r["_neto"] > 0 and not r["sustento"]) or
                                  (to_float(r["cantidad"]) is not None and r["_c"]["n_meses"] == 0)):
            return "bad"
        return ""

    def _row(self) -> dict:
        d = self.form.get()
        for i, v in enumerate(self.months.get()):
            d[MCOLS[i]] = v
        if self.kind == "FI":
            from controllers import parse_pct
            d["becas_pct"] = parse_pct(d.get("becas_pct"))
        return d

    def _changed(self):
        d = self._row()
        p = self.c.params
        cat = self.c.categoria_de(d.get("partida", ""), "Ingresos" if self.kind == "FI" else "Egresos")
        self.form.set_value("_categoria", (cat or ("(fuera del catálogo)" if d.get("partida") else ""),
                                           "gray20" if cat else COLOR_WARN))
        if self.kind == "FI":
            cc = calc_fing(d, p)
            self.form.set_value("_soles", ("" if cc["tarifa_soles"] is None else fmt_money(cc["tarifa_soles"]), "gray20"))
            self.calc.set_value("_bruto", (fmt_money(cc["bruto"]), "gray20"))
            self.calc.set_value("_desc", (fmt_money(cc["descuentos"]), "gray20"))
            self.calc.set_value("_neto", (fmt_money(cc["neto"]), "gray20"))
            self.calc.set_value("_n", (str(cc["n_meses"]), "gray20"))
        else:
            cc = calc_fegr(d, p)
            self.form.set_value("_soles", ("" if cc["costo_unit_soles"] is None else fmt_money(cc["costo_unit_soles"]),
                                           "gray20"))
            self.form.set_value("_coste", (fmt_money(cc["coste_total"]), "gray20"))
            self.calc.set_value("_total", (fmt_money(cc["total"]), "gray20"))
            self.calc.set_value("_n", (str(cc["n_pagos"]), "gray20"))
            self.calc.set_value("_control", cc["control"])
            self.calc.set_value("_reaj", (f"{fmt_pct(p.reajuste)} desde {MESES[p.mes_reajuste - 1]}"
                                          if d.get("aplica_reajuste") == "Sí" else "No aplica", "gray20"))
        self.months.set_calc(cc["meses"])

    def load(self, rid):
        if rid is None:
            return
        row = self.app.db.get("funcionamiento_ing" if self.kind == "FI" else "funcionamiento_egr", rid)
        if not row:
            return
        self.rid = rid
        self.form.clear()
        data = dict(row)
        if self.kind == "FI":
            data["becas_pct"] = "" if row["becas_pct"] is None else f"{row['becas_pct'] * 100:g}%"
        self.form.set(data)
        self.months.set([row[m] for m in MCOLS])
        self._changed()

    def new(self):
        self.rid = None
        self.table.tree.selection_remove(self.table.tree.selection())
        self.form.clear()
        self.months.clear()
        self.form.set({"codigo": self.c.next_code(self.kind, self.app.uid), "moneda": "S/",
                       ("aplica_morosidad" if self.kind == "FI" else "aplica_reajuste"): "No"})
        self._changed()

    @guarded
    def save(self):
        d = self._row()
        if not d.get("partida"):
            raise ValueError("Seleccione la Partida / rubro del catálogo NIIF.")
        self.rid = self.c.save_flat(self.kind, self.app.uid, d, self.rid)
        self.refresh()
        self.table.select(self.rid)
        self.app.mark_dirty(except_tab=self)

    @guarded
    def duplicate(self):
        if not self.rid:
            return
        self.rid = None
        self.form.set_value("codigo", self.c.next_code(self.kind, self.app.uid))
        self.save()

    @guarded
    def delete(self):
        if not self.rid or not messagebox.askyesno("Eliminar", "¿Eliminar la fila seleccionada?"):
            return
        self.c.delete_flat(self.kind, self.rid)
        self.new()
        self.refresh()
        self.app.mark_dirty(except_tab=self)

    @guarded
    def renumber(self):
        self.c.renumerar_flat(self.kind, self.app.uid)
        self.refresh()


class FuncionamientoTab(ctk.CTkFrame):
    def __init__(self, master, app: "App"):
        super().__init__(master, fg_color="transparent")
        tv = ctk.CTkTabview(self, command=self.refresh)
        tv.pack(fill="both", expand=True)
        self.tv = tv
        self.ing = FlatTab(tv.add("Ingresos (Funcionamiento_Ing)"), app, "FI")
        self.ing.pack(fill="both", expand=True)
        self.egr = FlatTab(tv.add("Egresos (Funcionamiento_Egr)"), app, "FE")
        self.egr.pack(fill="both", expand=True)
        tv.set("Egresos (Funcionamiento_Egr)")

    def refresh(self):
        (self.ing if self.tv.get().startswith("Ingresos") else self.egr).refresh()
        other = self.egr if self.tv.get().startswith("Ingresos") else self.ing
        other.after(10, other.refresh)


# =============================================================================
# Pestaña 7 — Ficha Resumen
# =============================================================================
class ResumenTab(ctk.CTkFrame):
    def __init__(self, master, app: "App"):
        super().__init__(master, fg_color="transparent")
        self.app, self.c = app, app.c
        buttons = [("⟳ Actualizar", self.refresh), ("Exportar Ficha Resumen (.pdf)", app.export_pdf),
                   ("Exportar Libro Oficial (.xlsx) + BD NIIF", app.export_xlsx)]
        if app.role == ROLE_ADMIN:
            buttons.append(("BD NIIF consolidada de todas las unidades (.xlsx + .csv)", app.export_bd_all, "#117A65"))
        bar = button_bar(self, buttons)
        bar.pack(anchor="w", padx=4)
        self.head = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=13, weight="bold"), justify="left")
        self.head.pack(anchor="w", padx=8)
        self.sc = ctk.CTkScrollableFrame(self)
        self.sc.pack(fill="both", expand=True, padx=4, pady=4)
        self._widgets: list = []
        self.last: Optional[dict] = None

    def refresh(self):
        if self.app.uid is None:
            return
        for w in self._widgets:
            w.destroy()
        self._widgets.clear()
        summ = self.c.build_summary(self.app.uid)
        self.last = summ
        u = summ["unidad"]
        dirs = summ["directrices"]
        n_rev = sum(1 for d in dirs if d["resultado"] == "Revisar")
        self.head.configure(
            text=f"{u['codigo_area']} · {u['nombre_unidad']} · Año {summ['anio']} · Responsable: {u['responsable'] or '—'}"
                 f"   |   Directrices: {11 - n_rev}/11 cumplen · Resultado S/ {fmt_money(summ['s3']['resultado'])}",
            text_color=COLOR_OK if not n_rev else COLOR_BAD)
        for sec in summ["sections"]:
            frame = ctk.CTkFrame(self.sc)
            frame.pack(fill="x", padx=4, pady=6)
            self._widgets.append(frame)
            ctk.CTkLabel(frame, text=sec["title"], font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8,
                                                                                                pady=(6, 2))
            if sec["key"] == "s11":
                self._semaforo(frame, dirs)
                continue
            headers, fmts = sec["headers"], sec["formats"]
            cols = []
            for j, h in enumerate(headers):
                width = 330 if j == 0 else (95 if sec["key"] == "s10" else 150)
                if sec["key"] == "s12":
                    width = 120 if j == 0 else 430
                fmt = fmts[j] if fmts[j] in ("money", "pct", "int") else "text"
                cols.append((f"c{j}", h or " ", width, fmt))
            rows = []
            for i, r in enumerate(sec["rows"]):
                row = {"id": i}
                for j, v in enumerate(r):
                    fmt = fmts[j]
                    if fmt == "money_or_pct":
                        fmt = "pct" if "(%)" in s(r[0]) else "money"
                    if i in sec.get("pct_rows", set()) and j > 0:
                        fmt = "pct"
                    if isinstance(v, str) or v is None:
                        row[f"c{j}"] = v or ""
                    else:
                        row[f"c{j}"] = fmt_value(v, fmt if fmt in ("money", "pct", "int") else "money")
                rows.append(row)
            cols = [(k, h, w, "text", "e" if f in ("money", "pct", "int") else "w") for k, h, w, f in cols]
            dt = DataTable(frame, cols, height=len(rows))
            dt.pack(fill="x", padx=6, pady=(0, 6))
            totals = sec.get("totals", set())

            def tag(r, totals=totals, sec=sec):
                if r["id"] in totals:
                    return "total"
                raw = sec["rows"][r["id"]]
                return "bad" if any(isinstance(v, str) and v in BAD_WORDS for v in raw) else ""
            dt.set_rows(rows, tag_fn=tag)
            if sec.get("note"):
                ctk.CTkLabel(frame, text=sec["note"], text_color="gray40").pack(anchor="w", padx=8, pady=(0, 6))

    def _semaforo(self, frame, dirs):
        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(fill="x", padx=8, pady=4)
        for i, d in enumerate(dirs):
            color = COLOR_OK if d["resultado"] == "Cumple" else COLOR_BAD
            ctk.CTkLabel(grid, text="●", text_color=color, font=ctk.CTkFont(size=20)).grid(row=i, column=0, padx=4)
            ctk.CTkLabel(grid, text=d["codigo"], font=ctk.CTkFont(weight="bold"), width=60).grid(row=i, column=1)
            ctk.CTkLabel(grid, text=d["descripcion"], anchor="w", width=700).grid(row=i, column=2, sticky="w")
            ctk.CTkLabel(grid, text=f"{d['n']} obs.", width=70).grid(row=i, column=3)
            ctk.CTkLabel(grid, text=d["resultado"], text_color=color, font=ctk.CTkFont(weight="bold"), width=80).grid(
                row=i, column=4)
            if d["items"]:
                ctk.CTkButton(grid, text="Ver detalle", width=90,
                              command=lambda d=d: self._detalle(d)).grid(row=i, column=5, padx=4, pady=1)

    def _detalle(self, d: dict):
        top = ctk.CTkToplevel(self)
        top.title(f"{d['codigo']} — {d['descripcion']}")
        top.geometry("820x420")
        tb = ctk.CTkTextbox(top, wrap="word")
        tb.pack(fill="both", expand=True, padx=8, pady=8)
        tb.insert("1.0", "\n".join(f"• {it}" for it in d["items"]))
        tb.configure(state="disabled")


# =============================================================================
# Aplicación principal
# =============================================================================
class App(ctk.CTk):
    def __init__(self, db_path: Optional[str] = None):
        super().__init__()
        self.title(APP_TITLE)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{min(1400, sw - 20)}x{min(860, sh - 70)}+5+5")
        self.minsize(min(1200, sw - 40), min(700, sh - 80))
        self.db = Database(db_path)
        self.c = POAController(self.db)
        self.role: Optional[str] = None
        self.uid: Optional[int] = None
        self.tabs: dict[str, Any] = {}
        self.dirty: set = set()
        self._style()
        self.login = LoginFrame(self, self)
        self.login.pack(fill="both", expand=True)
        self.main: Optional[ctk.CTkFrame] = None
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(0, self._maximize)

    def _maximize(self):
        try:
            self.state("zoomed")
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)
            except tk.TclError:
                pass

    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Treeview", rowheight=22, font=("Segoe UI", 9), background="white", fieldbackground="white")
        st.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#1F3864", foreground="white")
        st.map("Treeview.Heading", background=[("active", "#2E5597")])
        st.map("Treeview", background=[("selected", "#BBD3F2")], foreground=[("selected", "black")])

    # ---------------------------------------------------------------- sesión
    def start_session(self, role: str, uid: Optional[int]):
        self.role = role
        if uid is None:
            units = self.c.unidades()
            uid = units[0]["id"] if units else None
        self.uid = uid
        self.login.pack_forget()
        if self.main is not None:
            self.main.destroy()
        self._build_main()
        self.db.log(role, "Inicio de sesión", self.c.unidad(uid)["codigo_area"] if uid else "")

    def logout(self):
        if self.main is not None:
            self.main.destroy()
            self.main = None
        self.tabs.clear()
        self.role, self.uid = None, None
        self.login.refresh()
        self.login.pack(fill="both", expand=True)

    def _build_main(self):
        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.pack(fill="both", expand=True)
        top = ctk.CTkFrame(self.main, corner_radius=0, fg_color=("#1F3864", "#1F3864"))
        top.pack(fill="x")
        ctk.CTkLabel(top, text="POA - DPDI", text_color="white", font=ctk.CTkFont(size=18, weight="bold")).pack(
            side="left", padx=12, pady=6)
        self.header = ctk.CTkLabel(top, text="", text_color="white", justify="left")
        self.header.pack(side="left", padx=8)
        ctk.CTkButton(top, text="Cerrar sesión", width=10, fg_color="#7F8C8D", command=self.logout).pack(
            side="right", padx=8)
        bar = ctk.CTkFrame(self.main, corner_radius=0, fg_color=("#D6DEEB", "#2B2B2B"))
        bar.pack(fill="x")
        if self.role == ROLE_ADMIN:
            ctk.CTkLabel(bar, text="Unidad en formulación:").pack(side="left", padx=(8, 4), pady=5)
            self.unit_sel = ctk.CTkComboBox(bar, width=340, state="readonly", command=self._unit_changed)
            self.unit_sel.pack(side="left", pady=5)
            self._reload_units()
        ctk.CTkButton(bar, text="Exportar Ficha Resumen Firmada (.pdf)", width=10, fg_color="#A93226",
                      command=self.export_pdf).pack(side="right", padx=4)
        ctk.CTkButton(bar, text="Exportar Libro Oficial POA (.xlsx) y BD NIIF Consolidada", width=10,
                      fg_color="#1E8449", command=self.export_xlsx).pack(side="right", padx=4)
        ctk.CTkButton(bar, text="Importar Ficha Excel (.xlsx)", width=10, fg_color="#2E86C1",
                      command=self.import_ficha).pack(side="right", padx=4, pady=5)

        self.tv = ctk.CTkTabview(self.main, command=self._tab_changed)
        self.tv.pack(fill="both", expand=True, padx=6, pady=4)
        spec = [("0. Inicio y códigos", lambda m: InicioTab(m, self))]
        if self.role == ROLE_ADMIN:
            spec.append(("1. Parámetros y catálogos NIIF", lambda m: ParametrosTab(m, self)))
        spec += [("2. Iniciativas", lambda m: MasterDetailTab(m, self, "INI")),
                 ("3. Funcionamiento", lambda m: FuncionamientoTab(m, self)),
                 ("4. Otras operativas", lambda m: MasterDetailTab(m, self, "OO")),
                 ("5. Actividades propias", lambda m: MasterDetailTab(m, self, "AP")),
                 ("6. Inversión (CAPEX)", lambda m: MasterDetailTab(m, self, "INV")),
                 ("7. Ficha Resumen", lambda m: ResumenTab(m, self))]
        for name, factory in spec:
            frame = self.tv.add(name)
            w = factory(frame)
            w.pack(fill="both", expand=True)
            self.tabs[name] = w
        self.dirty = set(self.tabs)
        self.update_header()
        self.tv.set(spec[0][0])
        self._tab_changed()

    def _reload_units(self):
        units = self.c.unidades()
        self._units = {self.c.unidad_display(u): u["id"] for u in units}
        self.unit_sel.configure(values=list(self._units))
        for k, v in self._units.items():
            if v == self.uid:
                self.unit_sel.set(k)

    def _unit_changed(self, value):
        self.uid = self._units.get(value)
        self.mark_dirty()
        self.update_header()
        self._tab_changed()

    def update_header(self):
        if not self.uid:
            return
        u = self.c.unidad(self.uid)
        p = self.c.params
        self.header.configure(text=f"{self.role}  ·  {u['codigo_area']} · {u['nombre_unidad']}  ·  {u['macrogrupo']}"
                                   f"\nAño {p.anio} · TC {p.tc:.3f} · Inflación {fmt_pct(p.inflacion)} · "
                                   f"Morosidad {fmt_pct(p.morosidad)} · Reajuste {fmt_pct(p.reajuste)} desde "
                                   f"{MESES[p.mes_reajuste - 1]}")
        if self.role == ROLE_ADMIN and hasattr(self, "unit_sel"):
            self._reload_units()

    def mark_dirty(self, except_tab: Any = None):
        self.dirty = {k for k, w in self.tabs.items() if w is not except_tab}
        if except_tab is None:
            self._tab_changed()

    def _tab_changed(self):
        if not self.tabs:
            return
        name = self.tv.get()
        if name in self.dirty:
            self.dirty.discard(name)
            try:
                self.tabs[name].refresh()
            except Exception:  # noqa: BLE001
                traceback.print_exc()

    # ------------------------------------------------------- import / export
    def _default_name(self, prefix: str, ext: str) -> str:
        u = self.c.unidad(self.uid)
        return f"{prefix}_{self.c.anio}_{u['codigo_area']}.{ext}"

    @guarded
    def import_ficha(self):
        path = filedialog.askopenfilename(title="Importar Ficha_Formulacion_POA_*.xlsx",
                                          filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        target = self.uid if self.role == ROLE_UNIT else None
        msg = ("Se reemplazarán los registros del año en curso de la unidad destino con el contenido del archivo.\n"
               + ("La unidad destino es su unidad orgánica." if target else
                  "La unidad destino se identifica por el nombre de la unidad del archivo (o se crea)."))
        if not messagebox.askyesno("Importar ficha", msg + "\n\n¿Continuar?"):
            return
        params = False
        if self.role == ROLE_ADMIN:
            params = messagebox.askyesno("Importar ficha", "¿Actualizar también los parámetros globales y catálogos "
                                                           "con la hoja «Parámetros» del archivo?")
        self.configure(cursor="watch")
        self.update()
        try:
            rep = self.c.import_ficha(path, unidad_id=target, importar_parametros=params)
        finally:
            self.configure(cursor="")
        self.db.log(self.role, "Importar ficha", Path(path).name)
        if self.role == ROLE_ADMIN:
            self.uid = rep["unidad_id"]
            self._reload_units()
        avisos = "\n".join(f"• {a}" for a in rep["avisos"][:12])
        messagebox.showinfo("Importación completada",
                            f"Archivo: {rep['archivo']}\nUnidad: {self.c.unidad(rep['unidad_id'])['codigo_area']}"
                            f"{' (creada)' if rep['creada'] else ''}\n\n"
                            + "\n".join(f"{k}: {v}" for k, v in rep["conteos"].items())
                            + (f"\n\nAvisos:\n{avisos}" if avisos else ""))
        self.mark_dirty()
        self.update_header()

    @guarded
    def export_xlsx(self):
        path = filedialog.asksaveasfilename(title="Exportar Libro Oficial POA", defaultextension=".xlsx",
                                            initialfile=self._default_name("Ficha_Formulacion_POA", "xlsx"),
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.configure(cursor="watch")
        self.update()
        try:
            out = export_libro_oficial(self.c, self.uid, path)
        finally:
            self.configure(cursor="")
        self.db.log(self.role, "Exportar libro", Path(path).name)
        if messagebox.askyesno("Exportación", f"Libro generado:\n{out}\n\n16 pestañas oficiales + BD_CONSOLIDADA_NIIF."
                                              "\n¿Abrir el archivo?"):
            self._open(out)

    @guarded
    def export_pdf(self):
        path = filedialog.asksaveasfilename(title="Exportar Ficha Resumen (PDF)", defaultextension=".pdf",
                                            initialfile=self._default_name("Ficha_Resumen_POA", "pdf"),
                                            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        out = export_pdf(self.c, self.uid, path)
        self.db.log(self.role, "Exportar PDF", Path(path).name)
        if messagebox.askyesno("Exportación", f"Ficha Resumen generada:\n{out}\n¿Abrir el archivo?"):
            self._open(out)

    @guarded
    def export_bd_all(self):
        path = filedialog.asksaveasfilename(title="BD NIIF consolidada (todas las unidades)", defaultextension=".xlsx",
                                            initialfile=f"BD_CONSOLIDADA_NIIF_{self.c.anio}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        uids = [u["id"] for u in self.c.unidades()]
        out = export_bd_consolidada(self.c, uids, path)
        messagebox.showinfo("BD consolidada", "Archivos generados:\n" + "\n".join(str(p) for p in out))

    @staticmethod
    def _open(path: Path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}" >/dev/null 2>&1 &')
        except OSError:
            pass

    def _on_close(self):
        try:
            self.db.close()
        finally:
            self.destroy()


def main():
    db_path = sys.argv[1] if len(sys.argv) > 1 else None
    app = App(db_path)
    app.mainloop()


if __name__ == "__main__":
    main()
