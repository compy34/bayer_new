import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from config import (
    SERVICE_PRICES,
    DEFAULT_CONTRACT_NO,
    DEFAULT_CONTRACT_DATE,
    DEFAULT_TMA_CONTRACT_DATE,
    DEFAULT_EXECUTOR_NAME,
    DB_BRAND_TO_TEMPLATE,
    LIK_BRANDS_TEMPLATE,
    DD_BRANDS_TEMPLATE,
)
from convertor import (
    build_report_dataframe,
    generate_excel_report,
    mhl_brand_counts,
    mhl_unknown_samples,
    get_unknown_skus,
    save_sku_classification,
    get_mhl_lik_sku,
    get_mhl_dd_sku,
    get_all_skus_with_status,
    set_sku_active,
    get_tma_auth_amount,
    set_tma_auth_amount,
)
from Data.Controller import get_mhl_data, get_pharmacy_data, load_requisites_file, get_tma_offtake_data, DataLoadError


# ============================================================
# UI CONSTANTS — clean, professional palette
# ============================================================
BG_COLOR = '#fafafa'
SURFACE = '#ffffff'
PRIMARY = '#1a56a0'
PRIMARY_HOVER = '#154380'
PRIMARY_LIGHT = '#e8eef7'
SUCCESS = '#007b5e'
SUCCESS_HOVER = '#00664d'
WARNING = '#e67e22'
DANGER = '#c0392b'
TEXT_MAIN = '#1a1a2e'
TEXT_SECONDARY = '#4a4a5a'
TEXT_MUTED = '#888899'
BORDER = '#e0e0e8'
BORDER_FOCUS = '#1a56a0'

FONT_TITLE = ('Segoe UI', 18, 'bold')
FONT_SUBTITLE = ('Segoe UI', 11, 'bold')
FONT_BODY = ('Segoe UI', 10)
FONT_SMALL = ('Segoe UI', 9)
FONT_MONO = ('Consolas', 9)
FONT_BUTTON = ('Segoe UI', 10, 'bold')
FONT_BUTTON_SMALL = ('Segoe UI', 9)

PAD_X = 20
PAD_Y = 16
CARD_RADIUS = 8


def _find_requisites_file():
    """Auto-detect requisites file in project root."""
    root = Path(__file__).resolve().parent
    patterns = ['реквізити*', 'requisites*', 'реквизиты*']
    all_matches = []
    for pattern in patterns:
        matches = list(root.glob(f'{pattern}.*'))
        for m in matches:
            if m.suffix.lower() in ('.xlsx', '.xls', '.csv'):
                all_matches.append(m)
    if not all_matches:
        return None
    ext_priority = {'.xlsx': 0, '.xls': 1, '.csv': 2}
    all_matches.sort(key=lambda p: ext_priority.get(p.suffix.lower(), 99))
    return all_matches[0]


class ModernButton(tk.Button):
    """Clean button with hover effects."""
    def __init__(self, master, style='primary', **kwargs):
        self.style_type = style
        colors = {
            'primary': (PRIMARY, PRIMARY_HOVER, 'white'),
            'secondary': (SURFACE, '#f0f0f5', TEXT_MAIN),
            'success': (SUCCESS, SUCCESS_HOVER, 'white'),
            'danger': (DANGER, '#a02d2b', 'white'),
            'outline': (SURFACE, PRIMARY_LIGHT, PRIMARY),
        }
        bg, hover_bg, fg = colors.get(style, colors['primary'])
        super().__init__(
            master,
            bg=bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            font=FONT_BUTTON if 'height' not in kwargs else FONT_BUTTON_SMALL,
            relief='flat',
            bd=0,
            padx=20,
            pady=10,
            cursor='hand2',
            **kwargs
        )
        self.default_bg = bg
        self.hover_bg = hover_bg
        self.bind('<Enter>', lambda e: self.config(bg=hover_bg))
        self.bind('<Leave>', lambda e: self.config(bg=bg))


class Card(tk.Frame):
    """Clean card container."""
    def __init__(self, master, **kwargs):
        super().__init__(master, bg=SURFACE, **kwargs)
        self.configure(highlightbackground=BORDER, highlightthickness=1)


class SettingsDialog(tk.Toplevel):
    """All settings in one modal window with tabs."""
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title('Налаштування')
        self.geometry('700x550')
        self.minsize(650, 500)
        self.configure(bg=BG_COLOR)
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f'+{x}+{y}')

        self._build_ui()

    def _build_ui(self):
        # Notebook for tabs
        style = ttk.Style()
        style.configure('Custom.TNotebook', background=BG_COLOR, borderwidth=0)
        style.configure('Custom.TNotebook.Tab', padding=[20, 10], font=FONT_BODY)

        notebook = ttk.Notebook(self, style='Custom.TNotebook')
        notebook.pack(fill='both', expand=True, padx=PAD_X, pady=PAD_X)

        # --- Tab 1: SKU Management ---
        sku_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(sku_frame, text='  SKU Класифікація  ')
        self._build_sku_tab(sku_frame)

        # --- Tab 2: TMA Authorization ---
        tma_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(tma_frame, text='  ТМА Авторизація  ')
        self._build_tma_tab(tma_frame)

        # --- Tab 3: Contract Defaults ---
        contract_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(contract_frame, text='  Договори  ')
        self._build_contract_tab(contract_frame)

        # --- Tab 4: Service Prices ---
        prices_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(prices_frame, text='  Ціни послуг  ')
        self._build_prices_tab(prices_frame)

        # Close button
        btn_frame = tk.Frame(self, bg=BG_COLOR)
        btn_frame.pack(fill='x', padx=PAD_X, pady=(0, PAD_X))
        ModernButton(btn_frame, text='Закрити', style='secondary', command=self.destroy).pack(side='right')

    def _build_sku_tab(self, parent):
        """SKU classification management."""
        # Toolbar
        toolbar = tk.Frame(parent, bg=BG_COLOR)
        toolbar.pack(fill='x', pady=(0, 8))

        ModernButton(toolbar, text='➕ Додати SKU', style='primary', command=self._add_sku).pack(side='left', padx=(0, 8))
        ModernButton(toolbar, text='💾 Зберегти зміни', style='success', command=self._save_sku_changes).pack(side='left', padx=(0, 8))
        ModernButton(toolbar, text='🔄 Оновити', style='secondary', command=self._refresh_sku_table).pack(side='left')

        # Table
        table_frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        table_frame.pack(fill='both', expand=True)

        columns = ('SKU_UA', 'BRAND', 'CATEGORY', 'ACTIVE')
        self.sku_tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='extended')

        col_config = {
            'SKU_UA': ('SKU / Назва', 220),
            'BRAND': ('Бренд', 150),
            'CATEGORY': ('Категорія', 100),
            'ACTIVE': ('Активний', 80),
        }
        for col, (title, width) in col_config.items():
            self.sku_tree.heading(col, text=title)
            self.sku_tree.column(col, width=width, anchor='w' if col != 'ACTIVE' else 'center')

        # Scrollbars
        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.sku_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient='horizontal', command=self.sku_tree.xview)
        self.sku_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.sku_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Context menu
        self.sku_tree.bind('<Double-1>', self._on_sku_double_click)
        self.sku_tree.bind('<Delete>', lambda e: self._delete_selected_skus())

        self._refresh_sku_table()

    def _refresh_sku_table(self):
        for item in self.sku_tree.get_children():
            self.sku_tree.delete(item)
        df = get_all_skus_with_status()
        for _, row in df.iterrows():
            active_text = '✓' if row['active'] else '✗'
            active_tag = 'active' if row['active'] else 'inactive'
            self.sku_tree.insert('', 'end', values=(
                row['SKU_UA'], row['BRAND'], row['CATEGORY'], active_text
            ), tags=(active_tag,))
        self.sku_tree.tag_configure('active', foreground=SUCCESS)
        self.sku_tree.tag_configure('inactive', foreground=DANGER)

    def _on_sku_double_click(self, event):
        item = self.sku_tree.selection()[0] if self.sku_tree.selection() else None
        if not item:
            return
        values = self.sku_tree.item(item, 'values')
        self._edit_sku_dialog(values)

    def _add_sku(self):
        self._edit_sku_dialog()

    def _edit_sku_dialog(self, values=None):
        """Dialog to add/edit SKU."""
        dialog = tk.Toplevel(self)
        dialog.title('Редагувати SKU' if values else 'Додати SKU')
        dialog.geometry('420x300')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self)
        dialog.grab_set()

        # Center on parent
        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f'+{x}+{y}')

        fields = {}
        labels = [
            ('SKU_UA', 'SKU / Назва товару (обов\'язково)'),
            ('BRAND', 'Бренд'),
            ('CATEGORY', 'Категорія (ЛІК/ДД)'),
        ]
        for i, (key, label) in enumerate(labels):
            tk.Label(dialog, text=label, font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(
                row=i, column=0, sticky='w', padx=PAD_X, pady=(PAD_Y, 4)
            )
            if key == 'CATEGORY':
                cb = ttk.Combobox(dialog, values=['ЛІК', 'ДД'], state='readonly', font=FONT_BODY, width=35)
                cb.grid(row=i, column=0, sticky='ew', padx=PAD_X, pady=(0, 4))
                fields[key] = cb
            else:
                entry = tk.Entry(dialog, font=FONT_BODY, relief='solid', bd=1, highlightthickness=1,
                                highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
                entry.grid(row=i, column=0, sticky='ew', padx=PAD_X, pady=(0, 4))
                fields[key] = entry

        # Active checkbox
        active_var = tk.BooleanVar(value=True)
        tk.Checkbutton(dialog, text='Активний', variable=active_var, font=FONT_BODY, bg=BG_COLOR).grid(
            row=3, column=0, sticky='w', padx=PAD_X, pady=(PAD_Y, 4)
        )

        # Pre-fill if editing
        if values:
            fields['SKU_UA'].insert(0, values[0])
            fields['SKU_UA'].config(state='readonly')
            fields['BRAND'].insert(0, values[1])
            fields['CATEGORY'].set(values[2])
            active_var.set(values[3] == '✓')

        dialog.grid_columnconfigure(0, weight=1)

        def save():
            sku = fields['SKU_UA'].get().strip()
            brand = fields['BRAND'].get().strip()
            category = fields['CATEGORY'].get().strip()
            active = active_var.get()

            if not sku:
                messagebox.showwarning('Помилка', 'SKU не може бути порожнім', parent=dialog)
                return
            if category not in ('ЛІК', 'ДД'):
                messagebox.showwarning('Помилка', 'Категорія: ЛІК або ДД', parent=dialog)
                return

            save_sku_classification(sku, brand, category, active)
            self._refresh_sku_table()
            dialog.destroy()
            self.app.status_var.set(f'SKU {"оновлено" if values else "додано"}: {sku}')

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.grid(row=4, column=0, sticky='ew', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='Скасувати', style='secondary', command=dialog.destroy).pack(side='right', padx=(8, 0))
        ModernButton(btn_frame, text='Зберегти', style='success', command=save).pack(side='right')

    def _delete_selected_skus(self):
        selected = self.sku_tree.selection()
        if not selected:
            return
        if messagebox.askyesno('Підтвердження', f'Видалити {len(selected)} SKU?', parent=self):
            for item in selected:
                values = self.sku_tree.item(item, 'values')
                sku = values[0]
                # Deactivate instead of delete
                save_sku_classification(sku, '', '', False)
            self._refresh_sku_table()

    def _save_sku_changes(self):
        messagebox.showinfo('Збережено', 'Зміни автоматично зберігаються при редагуванні', parent=self)

    def _build_tma_tab(self, parent):
        """TMA Authorization per pharmacy."""
        # Instructions
        tk.Label(
            parent,
            text='Вартість послуг із авторизації (без ПДВ) для кожної мережі та категорії.',
            font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED, wraplength=600, justify='left'
        ).pack(anchor='w', pady=(0, PAD_Y))

        # Table
        table_frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        table_frame.pack(fill='both', expand=True)

        columns = ('PHARMACY', 'LIK', 'DD')
        self.tma_tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        self.tma_tree.heading('PHARMACY', text='Мережа / Клієнт')
        self.tma_tree.heading('LIK', text='ЛІК (грн.)')
        self.tma_tree.heading('DD', text='ДД (грн.)')
        self.tma_tree.column('PHARMACY', width=300, anchor='w')
        self.tma_tree.column('LIK', width=150, anchor='e')
        self.tma_tree.column('DD', width=150, anchor='e')

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.tma_tree.yview)
        self.tma_tree.configure(yscrollcommand=vsb.set)

        self.tma_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tma_tree.bind('<Double-1>', self._edit_tma_auth)

        self._refresh_tma_table()

    def _refresh_tma_table(self):
        for item in self.tma_tree.get_children():
            self.tma_tree.delete(item)
        from convertor import load_tma_auth
        auth_data = load_tma_auth()
        for pharmacy, cats in sorted(auth_data.items()):
            self.tma_tree.insert('', 'end', values=(
                pharmacy,
                f'{cats.get("ЛІК", 0):.2f}',
                f'{cats.get("ДД", 0):.2f}',
            ))

    def _edit_tma_auth(self, event):
        item = self.tma_tree.selection()[0] if self.tma_tree.selection() else None
        if not item:
            return
        values = self.tma_tree.item(item, 'values')
        pharmacy = values[0]
        lik = float(values[1]) if values[1] else 0
        dd = float(values[2]) if values[2] else 0

        dialog = tk.Toplevel(self)
        dialog.title(f'ТМА Авторизація: {pharmacy}')
        dialog.geometry('350x220')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self)
        dialog.grab_set()

        for i, (label, var, val) in enumerate([
            ('ЛІК (грн., без ПДВ)', 'lik', lik),
            ('ДД (грн., без ПДВ)', 'dd', dd),
        ]):
            tk.Label(dialog, text=label, font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(
                row=i, column=0, sticky='w', padx=PAD_X, pady=(PAD_Y, 4)
            )
            entry = tk.Entry(dialog, font=FONT_BODY, relief='solid', bd=1,
                            highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
            entry.insert(0, f'{val:.2f}')
            entry.grid(row=i, column=0, sticky='ew', padx=PAD_X, pady=(0, 4))
            dialog.grid_columnconfigure(0, weight=1)
            if i == 0:
                lik_entry = entry
            else:
                dd_entry = entry

        def save():
            try:
                set_tma_auth_amount(pharmacy, True, float(lik_entry.get() or 0))
                set_tma_auth_amount(pharmacy, False, float(dd_entry.get() or 0))
                self._refresh_tma_table()
                self.app.status_var.set(f'ТМА авторизація оновлена: {pharmacy}')
                dialog.destroy()
            except ValueError:
                messagebox.showwarning('Помилка', 'Введіть коректні числа', parent=dialog)

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.grid(row=2, column=0, sticky='ew', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='Скасувати', style='secondary', command=dialog.destroy).pack(side='right', padx=(8, 0))
        ModernButton(btn_frame, text='Зберегти', style='success', command=save).pack(side='right')

    def _build_contract_tab(self, parent):
        """Contract default values."""
        tk.Label(
            parent,
            text='Значення за замовчуванням (використовуються, якщо в файлі реквізитів немає даних).',
            font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED, wraplength=600, justify='left'
        ).pack(anchor='w', pady=(0, PAD_Y))

        fields = [
            ('DEFAULT_CONTRACT_NO', 'Номер договору', DEFAULT_CONTRACT_NO),
            ('DEFAULT_CONTRACT_DATE', 'Дата договору (напр. 01 березня 2026 року)', DEFAULT_CONTRACT_DATE),
            ('DEFAULT_TMA_CONTRACT_DATE', 'Дата договору для ТМА (напр. 27 лютого 2026 року)', DEFAULT_TMA_CONTRACT_DATE),
            ('DEFAULT_EXECUTOR_NAME', 'Виконавець', DEFAULT_EXECUTOR_NAME),
        ]

        self.contract_entries = {}
        for i, (key, label, default) in enumerate(fields):
            tk.Label(parent, text=label, font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(
                row=i, column=0, sticky='w', padx=PAD_X, pady=(PAD_Y, 4)
            )
            entry = tk.Entry(parent, font=FONT_BODY, relief='solid', bd=1,
                            highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER, width=50)
            entry.insert(0, default)
            entry.grid(row=i, column=0, sticky='ew', padx=PAD_X, pady=(0, 4))
            self.contract_entries[key] = entry

        parent.grid_columnconfigure(0, weight=1)

        # Save button
        btn_frame = tk.Frame(parent, bg=BG_COLOR)
        btn_frame.grid(row=len(fields), column=0, sticky='ew', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='💾 Зберегти як нові дефолти', style='success',
                    command=self._save_contract_defaults).pack(side='right')

    def _save_contract_defaults(self):
        """Note: This only affects current session. For persistent changes, edit config.py."""
        from config import DEFAULT_CONTRACT_NO, DEFAULT_CONTRACT_DATE, DEFAULT_TMA_CONTRACT_DATE, DEFAULT_EXECUTOR_NAME
        # Update runtime config
        import config as cfg
        cfg.DEFAULT_CONTRACT_NO = self.contract_entries['DEFAULT_CONTRACT_NO'].get()
        cfg.DEFAULT_CONTRACT_DATE = self.contract_entries['DEFAULT_CONTRACT_DATE'].get()
        cfg.DEFAULT_TMA_CONTRACT_DATE = self.contract_entries['DEFAULT_TMA_CONTRACT_DATE'].get()
        cfg.DEFAULT_EXECUTOR_NAME = self.contract_entries['DEFAULT_EXECUTOR_NAME'].get()
        messagebox.showinfo('Збережено', 'Зміни застосовано до поточної сесії.\nДля постійних змін відредагуйте config.py', parent=self)

    def _build_prices_tab(self, parent):
        """Service prices per brand."""
        tk.Label(
            parent,
            text='Ціни послуг за брендом (грн.). Використовуються для розрахунку вартості.',
            font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED, wraplength=600, justify='left'
        ).pack(anchor='w', pady=(0, PAD_Y))

        # Table with two columns: Brand | Price
        table_frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        table_frame.pack(fill='both', expand=True)

        columns = ('BRAND', 'PRICE')
        self.price_tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        self.price_tree.heading('BRAND', text='Бренд')
        self.price_tree.heading('PRICE', text='Ціна (грн.)')
        self.price_tree.column('BRAND', width=300, anchor='w')
        self.price_tree.column('PRICE', width=150, anchor='e')

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.price_tree.yview)
        self.price_tree.configure(yscrollcommand=vsb.set)

        self.price_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.price_tree.bind('<Double-1>', self._edit_price)

        self._refresh_price_table()

    def _refresh_price_table(self):
        for item in self.price_tree.get_children():
            self.price_tree.delete(item)
        for brand, price in sorted(SERVICE_PRICES.items()):
            self.price_tree.insert('', 'end', values=(brand, f'{price:.2f}'))

    def _edit_price(self, event):
        item = self.price_tree.selection()[0] if self.price_tree.selection() else None
        if not item:
            return
        values = self.price_tree.item(item, 'values')
        brand = values[0]
        price = float(values[1])

        dialog = tk.Toplevel(self)
        dialog.title(f'Ціна: {brand}')
        dialog.geometry('300x150')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self)
        dialog.grab_set()

        tk.Label(dialog, text=f'Бренд: {brand}', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', padx=PAD_X, pady=(PAD_Y, 4))
        tk.Label(dialog, text='Ціна (грн.):', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', padx=PAD_X, pady=(PAD_Y, 4))

        entry = tk.Entry(dialog, font=FONT_BODY, relief='solid', bd=1,
                        highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        entry.insert(0, f'{price:.2f}')
        entry.pack(fill='x', padx=PAD_X, pady=(0, PAD_Y))

        def save():
            try:
                SERVICE_PRICES[brand] = float(entry.get())
                self._refresh_price_table()
                self.app.status_var.set(f'Ціна оновлена: {brand} = {SERVICE_PRICES[brand]:.2f}')
                dialog.destroy()
            except ValueError:
                messagebox.showwarning('Помилка', 'Введіть коректне число', parent=dialog)

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.pack(fill='x', padx=PAD_X, pady=(0, PAD_Y))
        ModernButton(btn_frame, text='Скасувати', style='secondary', command=dialog.destroy).pack(side='right', padx=(8, 0))
        ModernButton(btn_frame, text='Зберегти', style='success', command=save).pack(side='right')


class ReportApp:
    """Main application window - clean and focused."""

    def __init__(self, root):
        self.root = root
        self.root.title('Генератор маркетингових звітів та актів')
        self.root.geometry('850x650')
        self.root.minsize(750, 550)
        self.root.configure(bg=BG_COLOR)

        # Data
        self.df = None
        self.df_mhl = None
        self.df_tma_offtake = None
        self.df_req_lik = None
        self.df_req_dd = None
        self.loaded_period = None
        self.loading = False

        # UI variables
        self.requisites_lik_var = tk.StringVar(value='')
        self.requisites_dd_var = tk.StringVar(value='')
        self.pharmacy_name_var = tk.StringVar(value='')
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        self.month_var = tk.StringVar(value=f'{datetime.now().month:02d}')

        # Auto-detect requisites
        self._auto_detect_requisites()

        self._build_ui()

        # Update pharmacy combo after UI is built
        self._update_pharmacy_combo()

    def _auto_detect_requisites(self):
        """Auto-detect separate requisites files for ЛИК and ДД."""
        root = Path(__file__).resolve().parent
        # Look for ЛИК requisites
        lik_patterns = ['реквізити*лік*', 'реквізити*лик*', 'requisites*lik*', 'requisites*lic*']
        dd_patterns = ['реквізити*дд*', 'реквізити*dd*', 'requisites*dd*']

        def find_file(patterns):
            for pattern in patterns:
                matches = list(root.glob(f'{pattern}.*'))
                for m in matches:
                    if m.suffix.lower() in ('.xlsx', '.xls', '.csv'):
                        return m
            return None

        # Try to find ЛИК file
        lik_file = find_file(lik_patterns)
        if lik_file:
            try:
                self.df_req_lik = load_requisites_file(lik_file)
                self.requisites_lik_var.set(lik_file.name)
            except Exception:
                pass

        # Try to find ДД file
        dd_file = find_file(dd_patterns)
        if dd_file:
            try:
                self.df_req_dd = load_requisites_file(dd_file)
                self.requisites_dd_var.set(dd_file.name)
            except Exception:
                pass

        # Fallback: if no separate files found, try generic requisites file
        if not lik_file and not dd_file:
            found = _find_requisites_file()
            if found:
                try:
                    df_req = load_requisites_file(found)
                    self.df_req_lik = df_req
                    self.df_req_dd = df_req
                    self.requisites_lik_var.set(found.name + ' (спільний)')
                    self.requisites_dd_var.set(found.name + ' (спільний)')
                except Exception:
                    pass

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = tk.Frame(self.root, bg=BG_COLOR)
        main.grid(row=0, column=0, sticky='nsew', padx=24, pady=24)
        main.columnconfigure(0, weight=1)

        # --- Header ---
        header = tk.Frame(main, bg=BG_COLOR)
        header.grid(row=0, column=0, sticky='ew', pady=(0, 24))
        header.columnconfigure(1, weight=1)

        tk.Label(header, text='Генератор звітів та актів',
                font=FONT_TITLE, bg=BG_COLOR, fg=TEXT_MAIN).grid(row=0, column=0, sticky='w')

        # Settings button (top right)
        ModernButton(header, text='⚙️ Налаштування', style='outline',
                    command=self.open_settings).grid(row=0, column=2, sticky='e')

        # --- Requisites Card (separate for ЛИК and ДД) ---
        req_card = Card(main, padx=20, pady=16)
        req_card.grid(row=1, column=0, sticky='ew', pady=(0, 16))
        req_card.columnconfigure(1, weight=1)

        tk.Label(req_card, text='Файли реквізитів',
                font=FONT_SUBTITLE, bg=SURFACE, fg=TEXT_MAIN).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 8))

        # ЛИК requisites
        tk.Label(req_card, text='ЛІК:', font=FONT_BODY, bg=SURFACE, fg=TEXT_MAIN).grid(row=1, column=0, sticky='w', padx=(0, 10), pady=(0, 4))
        self.req_lik_entry = tk.Entry(req_card, textvariable=self.requisites_lik_var,
                                       font=FONT_BODY, state='readonly',
                                       readonlybackground=SURFACE, relief='solid', bd=1,
                                       highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        self.req_lik_entry.grid(row=1, column=1, sticky='ew', padx=(0, 10), pady=(0, 4))

        ModernButton(req_card, text='Обрати...', style='outline',
                    command=lambda: self.select_requisites_file('lik')).grid(row=1, column=2, sticky='e', pady=(0, 4))

        # ДД requisites
        tk.Label(req_card, text='ДД:', font=FONT_BODY, bg=SURFACE, fg=TEXT_MAIN).grid(row=2, column=0, sticky='w', padx=(0, 10), pady=(0, 4))
        self.req_dd_entry = tk.Entry(req_card, textvariable=self.requisites_dd_var,
                                      font=FONT_BODY, state='readonly',
                                      readonlybackground=SURFACE, relief='solid', bd=1,
                                      highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        self.req_dd_entry.grid(row=2, column=1, sticky='ew', padx=(0, 10), pady=(0, 4))

        ModernButton(req_card, text='Обрати...', style='outline',
                    command=lambda: self.select_requisites_file('dd')).grid(row=2, column=2, sticky='e', pady=(0, 4))

        # --- Pharmacy Selection ---
        pharm_card = Card(main, padx=20, pady=16)
        pharm_card.grid(row=2, column=0, sticky='ew', pady=(0, 16))
        pharm_card.columnconfigure(1, weight=1)

        # Row 0: label
        tk.Label(pharm_card, text='Мережа / Клієнт',
                font=FONT_SUBTITLE, bg=SURFACE, fg=TEXT_MAIN).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 8))

        # Row 1: combobox + TMA auth preview
        self.pharmacy_combo = ttk.Combobox(pharm_card, textvariable=self.pharmacy_name_var,
                                          state='readonly', font=FONT_BODY, width=45)
        self.pharmacy_combo.grid(row=1, column=0, sticky='ew')
        self.pharmacy_combo.bind('<<ComboboxSelected>>', lambda e: self._on_pharmacy_change())

        # TMA auth preview (read-only)
        auth_frame = tk.Frame(pharm_card, bg=SURFACE)
        auth_frame.grid(row=1, column=1, sticky='e', padx=(16, 0))
        self.tma_auth_preview = tk.Label(auth_frame, text='ТМА: ЛІК 0.00 / ДД 0.00',
                                        font=FONT_SMALL, bg=SURFACE, fg=TEXT_MUTED)
        self.tma_auth_preview.pack()

        # --- Period Selection ---
        period_card = Card(main, padx=20, pady=16)
        period_card.grid(row=3, column=0, sticky='ew', pady=(0, 24))
        period_card.columnconfigure(1, weight=1)

        tk.Label(period_card, text='Звітний період',
                font=FONT_SUBTITLE, bg=SURFACE, fg=TEXT_MAIN).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 8))

        years = [str(y) for y in range(2024, datetime.now().year + 2)]
        tk.Label(period_card, text='Рік', font=FONT_BODY, bg=SURFACE, fg=TEXT_MUTED).grid(row=1, column=0, sticky='w')
        year_combo = ttk.Combobox(period_card, textvariable=self.year_var, values=years,
                                 state='readonly', font=FONT_BODY, width=10)
        year_combo.grid(row=2, column=0, sticky='w', padx=(0, 16), pady=(0, 8))

        months = [f'{m:02d} — {n}' for m, n in zip(range(1, 13), [
            'Січень', 'Лютий', 'Березень', 'Квітень', 'Травень', 'Червень',
            'Липень', 'Серпень', 'Вересень', 'Жовтень', 'Листопад', 'Грудень'
        ])]
        tk.Label(period_card, text='Місяць', font=FONT_BODY, bg=SURFACE, fg=TEXT_MUTED).grid(row=1, column=1, sticky='w')
        self.month_combo = ttk.Combobox(period_card, textvariable=self.month_var,
                                       values=months, state='readonly', font=FONT_BODY, width=22)
        self.month_combo.grid(row=2, column=1, sticky='w', pady=(0, 8))

        # --- Action Buttons ---
        actions = tk.Frame(main, bg=BG_COLOR)
        actions.grid(row=4, column=0, sticky='ew', pady=(8, 16))
        actions.columnconfigure(0, weight=1)

        self.load_button = ModernButton(actions, text='🔄 Завантажити дані', style='primary',
                                       command=self.load_data)
        self.load_button.pack(side='left', padx=(0, 10))

        ModernButton(actions, text='📥 Згенерувати Excel', style='success',
                    command=self.download).pack(side='left', padx=(0, 10))

        ModernButton(actions, text='👁 Попередній перегляд', style='secondary',
                    command=self.preview).pack(side='left')

        # --- Status Bar (no progress bar) ---
        self.status_var = tk.StringVar(value='Готово до роботи')
        status_bar = tk.Frame(main, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        status_bar.grid(row=5, column=0, sticky='ew', pady=(16, 0))
        status_bar.columnconfigure(0, weight=1)

        self.status_label = tk.Label(status_bar, textvariable=self.status_var,
                                    font=FONT_MONO, bg=SURFACE, fg=TEXT_MUTED,
                                    anchor='w', padx=12, pady=8)
        self.status_label.grid(row=0, column=0, sticky='ew')

    def open_settings(self):
        SettingsDialog(self.root, self)

    def _on_pharmacy_change(self):
        self._load_tma_auth()

    def _load_tma_auth(self):
        pharmacy = self.pharmacy_name_var.get().strip()
        if pharmacy:
            lik = get_tma_auth_amount(pharmacy, True)
            dd = get_tma_auth_amount(pharmacy, False)
            self.tma_auth_preview.config(text=f'ТМА: ЛІК {lik:.2f} / ДД {dd:.2f}')
        else:
            self.tma_auth_preview.config(text='ТМА: виберіть мережу')

    def _get_period_range(self):
        year = int(self.year_var.get())
        month = int(self.month_var.get()[:2])
        start = f'{year}-{month:02d}-01'
        if month == 12:
            end = f'{year + 1}-01-01'
        else:
            end = f'{year}-{month + 1:02d}-01'
        from datetime import timedelta
        end_date = datetime.strptime(end, '%Y-%m-%d') - timedelta(days=1)
        return start, end_date.strftime('%Y-%m-%d')

    def _validate_period(self):
        return self._get_period_range()

    def select_requisites_file(self, category):
        """Select requisites file for ЛИК or ДД."""
        title = 'Виберіть файл реквізитів для ЛІК' if category == 'lik' else 'Виберіть файл реквізитів для ДД'
        filename = filedialog.askopenfilename(
            title=title,
            initialdir=str(Path.cwd()),
            filetypes=[('Усі таблиці', '*.csv *.xlsx *.xls'), ('CSV', '*.csv'), ('Excel', '*.xlsx *.xls')]
        )
        if not filename:
            return
        try:
            df_req = load_requisites_file(filename)
            if df_req.empty:
                raise ValueError('Файл реквізитів порожній.')
        except (OSError, ValueError, UnicodeError) as e:
            messagebox.showerror('Помилка', str(e))
            return

        if category == 'lik':
            self.df_req_lik = df_req
            self.requisites_lik_var.set(Path(filename).name)
        else:
            self.df_req_dd = df_req
            self.requisites_dd_var.set(Path(filename).name)

        self._update_pharmacy_combo()

    def _update_pharmacy_combo(self):
        """Update pharmacy combo with chains from both requisites files."""
        all_chains = set()
        for df_req in [self.df_req_lik, self.df_req_dd]:
            if df_req is None or df_req.empty:
                continue
            target_col = None
            for col in ['chain_name', 'Клиент из базы данных', 'Мережа', 'Назва']:
                if col in df_req.columns:
                    target_col = col
                    break
            if target_col:
                chains = df_req[target_col].dropna().astype(str).str.strip()
                all_chains.update(c for c in chains.unique() if c and c.lower() != 'nan')

        if all_chains:
            unique = sorted(all_chains)
            self.pharmacy_combo['values'] = tuple(unique)
            self.pharmacy_combo.current(0)
            self._load_tma_auth()
            self.status_var.set(f'Завантажено мереж: {len(unique)}')
        else:
            self.pharmacy_combo['values'] = ()
            self.status_var.set('Реквізити не завантажені')

    def _filtered_pharmacy_data(self):
        if self.df is None or self.df.empty:
            return pd.DataFrame(columns=['CHAIN_NAME', 'PHARMACY_NAME', 'BRAND', 'MERCH', 'Price'])
        selected = self.pharmacy_name_var.get().strip()
        if not selected:
            raise ValueError('Оберіть мережу.')
        if 'CHAIN_NAME' in self.df.columns:
            filtered = self.df[
                self.df['CHAIN_NAME'].fillna('').astype(str).str.strip().str.lower() == selected.lower()
            ]
        else:
            filtered = self.df[
                self.df['PHARMACY_NAME'].fillna('').astype(str).str.lower().str.contains(selected.lower())
            ]
        return filtered if not filtered.empty else self.df.iloc[0:0]

    def _filtered_mhl_data(self):
        if self.df_mhl is None or self.df_mhl.empty:
            return self.df_mhl
        selected = self.pharmacy_name_var.get().strip()
        if not selected:
            return self.df_mhl
        if 'CHAIN_NAME' in self.df_mhl.columns:
            filtered = self.df_mhl[
                self.df_mhl['CHAIN_NAME'].fillna('').astype(str).str.strip().str.lower() == selected.lower()
            ]
        else:
            return self.df_mhl.iloc[0:0]
        return filtered if not filtered.empty else self.df_mhl.iloc[0:0]

    def load_data(self):
        try:
            period = self._get_period_range()
        except ValueError as e:
            messagebox.showwarning('Попередження', str(e))
            return

        if self.loading:
            return

        self.loading = True
        self.load_button.config(state='disabled', text='⏳ Завантаження...')
        self.status_var.set(f'Завантаження даних за {period[0]} — {period[1]}...')

        selected = self.pharmacy_name_var.get().strip()

        def load():
            data = mhl_data = offtake_data = None
            data_error = mhl_error = offtake_error = None
            try:
                data = get_pharmacy_data(*period, chain_name=selected if selected else None)
            except DataLoadError as e:
                data_error = e
            except Exception as e:
                data_error = DataLoadError(f'Несподівана помилка: {e}')
            try:
                mhl_data = get_mhl_data(*period, chain_name=selected if selected else None)
            except DataLoadError as e:
                mhl_error = e
            except Exception as e:
                mhl_error = DataLoadError(f'Несподівана помилка MHL: {e}')
            try:
                offtake_data = get_tma_offtake_data(*period, chain_name=selected if selected else None)
            except DataLoadError as e:
                offtake_error = e
            except Exception as e:
                offtake_error = DataLoadError(f'Несподівана помилка offtake: {e}')

            self.root.after(0, lambda: self._data_loaded(data, mhl_data, offtake_data, period,
                                                        data_error, mhl_error, offtake_error))

        threading.Thread(target=load, daemon=True).start()

    def _data_loaded(self, data, mhl_data, offtake_data, period, data_error=None, mhl_error=None, offtake_error=None):
        self.loading = False
        self.load_button.config(state='normal', text='🔄 Завантажити дані')
        self.df = data
        self.df_mhl = mhl_data
        self.df_tma_offtake = offtake_data

        has_pharm = data is not None and not data.empty
        has_mhl = mhl_data is not None and not mhl_data.empty

        if data_error:
            self.loaded_period = None
            self.status_var.set(f'Помилка: {data_error}')
        elif not has_pharm and not has_mhl:
            self.loaded_period = None
            self.status_var.set('Дані відсутні за цей період.')
        else:
            self.loaded_period = period
            lik_c, dd_c, unk_c = mhl_brand_counts(mhl_data)
            status = f'Звіт: {len(data) if has_pharm else 0} рядків'
            if mhl_error:
                status += f'; помилка MHL: {mhl_error}'
            else:
                unk_s = mhl_unknown_samples(mhl_data)
                status += f'; MHL: {len(mhl_data) if has_mhl else 0} рядків; ЛІК: {lik_c}; ДД: {dd_c}; невідомі: {unk_c}'
                if unk_s:
                    status += f'; приклади: {" | ".join(unk_s)}'
            self.status_var.set(status)

    def preview(self):
        try:
            report = self._get_preview_data()
        except (ValueError, KeyError) as e:
            messagebox.showwarning('Попередження', str(e))
            return

        # Show in new window
        win = tk.Toplevel(self.root)
        win.title('Попередній перегляд')
        win.geometry('900x500')
        win.configure(bg=BG_COLOR)

        tree = ttk.Treeview(win, show='headings')
        tree['columns'] = list(report.columns)
        for col in report.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor='center')
        for row in report.head(200).itertuples(index=False):
            tree.insert('', 'end', values=[str(v) for v in row])

        vsb = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(win, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)

        tk.Label(win, text=f'Показано {min(len(report), 200)} з {len(report)} рядків',
                font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).grid(row=2, column=0, sticky='w', padx=PAD_X, pady=8)

    def _get_preview_data(self):
        self._selected_requisites()
        period = self._get_period_range()
        if self.loaded_period != period:
            raise ValueError('Спочатку завантажте дані для вказаного періоду.')
        pharmacy_df = self._filtered_pharmacy_data()
        mhl_df = self.df_mhl
        if (pharmacy_df is None or pharmacy_df.empty) and (mhl_df is None or mhl_df.empty):
            raise ValueError('Немає даних для цієї мережі.')
        report = build_report_dataframe(pharmacy_df)
        if (report is None or report.empty) and mhl_df is not None and not mhl_df.empty:
            report = pd.DataFrame({'Повідомлення': [f'Дані викладки відсутні. MHL: {len(mhl_df)} рядків']})
        return report

    def _selected_requisites(self):
        """Validate that at least one requisites file is loaded."""
        if (self.df_req_lik is None or self.df_req_lik.empty) and (self.df_req_dd is None or self.df_req_dd.empty):
            raise ValueError('Оберіть файли реквізитів для ЛІК та/або ДД!')
        return self.df_req_lik, self.df_req_dd

    def download(self):
        try:
            self._get_preview_data()
        except (ValueError, KeyError) as e:
            messagebox.showwarning('Попередження', str(e))
            return

        # Validate TMA auth
        pharmacy = self.pharmacy_name_var.get().strip()
        if pharmacy:
            lik = get_tma_auth_amount(pharmacy, True)
            dd = get_tma_auth_amount(pharmacy, False)
            if lik == 0 and dd == 0:
                if not messagebox.askyesno('Увага', f'Для "{pharmacy}" ТМА авторизація = 0.\nПродовжити?'):
                    return

        # Filename
        name = pharmacy or 'Мережа'
        safe = ''.join(c for c in name if c.isalnum() or c in ' _-').strip()
        default = f'{safe}_ЗВІТ+АКТ.xlsx'

        filename = filedialog.asksaveasfilename(
            title='Зберегти звіт як',
            initialdir=str(Path.cwd()),
            initialfile=default,
            defaultextension='.xlsx',
            filetypes=[('Excel', '*.xlsx')]
        )
        if not filename:
            return

        try:
            start, end = self._validate_period()
            df_req_lik, df_req_dd = self._selected_requisites()
            generate_excel_report(
                self._filtered_pharmacy_data(),
                df_req_lik,
                df_req_dd,
                filename,
                period_start=start, period_end=end,
                selected_pharmacy=self.pharmacy_name_var.get(),
                mhl_df=self._filtered_mhl_data(),
                full_mhl_df=self.df_mhl,
                offtake_df=self.df_tma_offtake,
            )
        except DataLoadError as e:
            messagebox.showerror('Помилка даних', str(e))
            return
        except (ValueError, KeyError) as e:
            messagebox.showerror('Помилка', str(e))
            return
        except Exception as e:
            messagebox.showerror('Несподівана помилка', f'{type(e).__name__}: {e}')
            return

        self.status_var.set(f'Звіт збережено: {filename}')
        messagebox.showinfo('Успіх', f'Файл збережено:\n{filename}')


def main():
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()