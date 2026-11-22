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
    MHL_REWARD_MULTIPLIER,
    TMA_ACTION_PERCENT_FILE,
    MHL_ACTION_PERCENT_FILE,
    LA_PRICE_FILE,
)
from convertor import (
    build_report_dataframe,
    generate_excel_report,
    mhl_brand_counts,
    mhl_unknown_samples,
    get_mhl_lik_sku,
    get_mhl_dd_sku,
    get_mhl_lik_brands,
    get_mhl_dd_brands,
    get_tma_auth_amount,
    set_tma_auth_amount,
    get_all_penetration_targets,
    set_penetration_target,
    load_mhl_action_percents,
    set_mhl_action_percent_file,
    load_la_price,
    set_la_price_file,
    get_all_la_prices,
    set_la_price,
    get_mhl_unique_brands,
    get_all_sku_classifications,
    set_sku_classification,
    remove_sku_classification,
    get_sku_category_local,
    get_combined_lik_skus,
    get_combined_dd_skus,
)
from Data.Controller import get_mhl_data, get_pharmacy_data, load_requisites_file, get_tma_offtake_data, get_all_active_skus, DataLoadError


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

        # --- Tab 1: TMA Authorization ---
        tma_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(tma_frame, text='  ТМА Авторизація  ')
        self._build_tma_tab(tma_frame)

        # --- Tab 2: Contract Defaults ---
        contract_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(contract_frame, text='  Договори  ')
        self._build_contract_tab(contract_frame)

        # --- Tab 3: Service Prices ---
        prices_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(prices_frame, text='  Ціни послуг  ')
        self._build_prices_tab(prices_frame)

        # --- Tab 4: Penetration Targets ---
        targets_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(targets_frame, text='  Цілі проникнення  ')
        self._build_targets_tab(targets_frame)

        # --- Tab 5: TMA Settings ---
        tma_settings_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(tma_settings_frame, text='  ТМА Налаштування  ')
        self._build_tma_settings_tab(tma_settings_frame)

        # --- Tab 6: Brand Qualification ---
        brand_frame = tk.Frame(notebook, bg=BG_COLOR)
        notebook.add(brand_frame, text='  Кваліфікація брендів  ')
        self._build_brand_tab(brand_frame)

        # Close button
        btn_frame = tk.Frame(self, bg=BG_COLOR)
        btn_frame.pack(fill='x', padx=PAD_X, pady=(0, PAD_X))
        ModernButton(btn_frame, text='Закрити', style='secondary', command=self.destroy).pack(side='right')

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
        ).grid(row=0, column=0, columnspan=2, sticky='w', padx=PAD_X, pady=(0, PAD_Y))

        fields = [
            ('DEFAULT_CONTRACT_NO', 'Номер договору', DEFAULT_CONTRACT_NO),
            ('DEFAULT_CONTRACT_DATE', 'Дата договору (напр. 01 березня 2026 року)', DEFAULT_CONTRACT_DATE),
            ('DEFAULT_TMA_CONTRACT_DATE', 'Дата договору для ТМА (напр. 27 лютого 2026 року)', DEFAULT_TMA_CONTRACT_DATE),
            ('DEFAULT_EXECUTOR_NAME', 'Виконавець', DEFAULT_EXECUTOR_NAME),
            ('MHL_REWARD_MULTIPLIER', 'MHL відсоток винагороди (0.1 = 10%)', str(MHL_REWARD_MULTIPLIER)),
        ]

        self.contract_entries = {}
        for i, (key, label, default) in enumerate(fields):
            row = i + 1  # row 0 зайнято описовим label
            tk.Label(parent, text=label, font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(
                row=row, column=0, sticky='w', padx=PAD_X, pady=(PAD_Y, 4)
            )
            entry = tk.Entry(parent, font=FONT_BODY, relief='solid', bd=1,
                            highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER, width=50)
            entry.insert(0, default)
            entry.grid(row=row, column=0, sticky='ew', padx=PAD_X, pady=(0, 4))
            self.contract_entries[key] = entry

        parent.grid_columnconfigure(0, weight=1)

        # Save button
        btn_frame = tk.Frame(parent, bg=BG_COLOR)
        btn_frame.grid(row=len(fields) + 1, column=0, sticky='ew', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='💾 Зберегти як нові дефолти', style='success',
                    command=self._save_contract_defaults).pack(side='right')

    def _save_contract_defaults(self):
        """Note: This only affects current session. For persistent changes, edit config.py."""
        from config import DEFAULT_CONTRACT_NO, DEFAULT_CONTRACT_DATE, DEFAULT_TMA_CONTRACT_DATE, DEFAULT_EXECUTOR_NAME, MHL_REWARD_MULTIPLIER
        # Update runtime config
        import config as cfg
        cfg.DEFAULT_CONTRACT_NO = self.contract_entries['DEFAULT_CONTRACT_NO'].get()
        cfg.DEFAULT_CONTRACT_DATE = self.contract_entries['DEFAULT_CONTRACT_DATE'].get()
        cfg.DEFAULT_TMA_CONTRACT_DATE = self.contract_entries['DEFAULT_TMA_CONTRACT_DATE'].get()
        cfg.DEFAULT_EXECUTOR_NAME = self.contract_entries['DEFAULT_EXECUTOR_NAME'].get()
        cfg.MHL_REWARD_MULTIPLIER = float(self.contract_entries['MHL_REWARD_MULTIPLIER'].get() or 0.1)
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

    def _build_targets_tab(self, parent):
        """Penetration target management."""
        tk.Label(
            parent,
            text='Локальні цілі проникнення (відсоток) для перевизначення значень з бази даних.\nФормат: SKU/Назва товару → ціль (напр. 0.8 для 80%).',
            font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED, wraplength=600, justify='left'
        ).pack(anchor='w', pady=(0, PAD_Y))

        # Toolbar
        toolbar = tk.Frame(parent, bg=BG_COLOR)
        toolbar.pack(fill='x', pady=(0, 8))

        ModernButton(toolbar, text='➕ Додати ціль', style='primary', command=self._add_target).pack(side='left', padx=(0, 8))
        ModernButton(toolbar, text='🔄 Оновити', style='secondary', command=self._refresh_targets_table).pack(side='left')

        # Table
        table_frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        table_frame.pack(fill='both', expand=True)

        columns = ('SKU_UA', 'TARGET')
        self.targets_tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='extended')

        col_config = {
            'SKU_UA': ('SKU / Назва товару', 350),
            'TARGET': ('Ціль (0-1)', 100),
        }
        for col, (title, width) in col_config.items():
            self.targets_tree.heading(col, text=title)
            self.targets_tree.column(col, width=width, anchor='w' if col != 'TARGET' else 'center')

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.targets_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient='horizontal', command=self.targets_tree.xview)
        self.targets_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.targets_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.targets_tree.bind('<Double-1>', self._on_target_double_click)
        self.targets_tree.bind('<Delete>', lambda e: self._delete_selected_targets())

        self._refresh_targets_table()

    def _refresh_targets_table(self):
        for item in self.targets_tree.get_children():
            self.targets_tree.delete(item)
        df = get_all_penetration_targets()
        for _, row in df.iterrows():
            self.targets_tree.insert('', 'end', values=(row['SKU_UA'], f'{row["TARGET"]:.2f}'))

    def _on_target_double_click(self, event):
        item = self.targets_tree.selection()[0] if self.targets_tree.selection() else None
        if not item:
            return
        values = self.targets_tree.item(item, 'values')
        self._edit_target_dialog(values)

    def _add_target(self):
        self._edit_target_dialog()

    def _edit_target_dialog(self, values=None):
        """Dialog to add/edit penetration target."""
        dialog = tk.Toplevel(self)
        dialog.title('Редагувати ціль проникнення' if values else 'Додати ціль проникнення')
        dialog.geometry('420x200')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f'+{x}+{y}')

        fields = {}
        labels = [
            ('SKU_UA', 'SKU / Назва товару (обов\'язково)'),
            ('TARGET', 'Ціль проникнення (0-1, напр. 0.8 для 80%)'),
        ]
        for i, (key, label) in enumerate(labels):
            tk.Label(dialog, text=label, font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(
                row=i, column=0, sticky='w', padx=PAD_X, pady=(PAD_Y, 4)
            )
            entry = tk.Entry(dialog, font=FONT_BODY, relief='solid', bd=1,
                            highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
            entry.grid(row=i, column=0, sticky='ew', padx=PAD_X, pady=(0, 4))
            fields[key] = entry

        dialog.grid_columnconfigure(0, weight=1)

        # Pre-fill if editing
        if values:
            fields['SKU_UA'].insert(0, values[0])
            fields['SKU_UA'].config(state='readonly')
            fields['TARGET'].insert(0, values[1])

        def save():
            sku = fields['SKU_UA'].get().strip()
            target_str = fields['TARGET'].get().strip()

            if not sku:
                messagebox.showwarning('Помилка', 'SKU не може бути порожнім', parent=dialog)
                return

            try:
                target = float(target_str)
                if not 0 <= target <= 1:
                    raise ValueError
            except ValueError:
                messagebox.showwarning('Помилка', 'Ціль має бути числом від 0 до 1 (напр. 0.8)', parent=dialog)
                return

            set_penetration_target(sku, target)
            self._refresh_targets_table()
            dialog.destroy()
            self.app.status_var.set(f'Ціль проникнення оновлена: {sku} = {target:.2f}')

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.grid(row=2, column=0, sticky='ew', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='Скасувати', style='secondary', command=dialog.destroy).pack(side='right', padx=(8, 0))
        ModernButton(btn_frame, text='Зберегти', style='success', command=save).pack(side='right')

    def _delete_selected_targets(self):
        selected = self.targets_tree.selection()
        if not selected:
            return
        if messagebox.askyesno('Підтвердження', f'Видалити {len(selected)} цілі?', parent=self):
            for item in selected:
                values = self.targets_tree.item(item, 'values')
                sku = values[0]
                set_penetration_target(sku, None)  # None = remove override
            self._refresh_targets_table()

    def _build_tma_settings_tab(self, parent):
        """TMA Settings - TMA action percent file, MHL action percent file, LA_Price file path, and MHL reward multiplier."""
        tk.Label(
            parent,
            text='Налаштування для аркушів "під час акції" (ТМА).',
            font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED, wraplength=600, justify='left'
        ).pack(anchor='w', pady=(0, PAD_Y))

        # --- MHL Reward Multiplier ---
        multiplier_frame = tk.Frame(parent, bg=BG_COLOR)
        multiplier_frame.pack(fill='x', pady=(0, PAD_Y))
        multiplier_frame.columnconfigure(1, weight=1)

        tk.Label(multiplier_frame, text='MHL відсоток винагороди:', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(row=0, column=0, sticky='w', padx=(0, 10))
        tk.Label(multiplier_frame, text='(0.1 = 10%, 0.15 = 15%)', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).grid(row=0, column=2, sticky='w')
        self.mhl_reward_multiplier_var = tk.StringVar(value=str(MHL_REWARD_MULTIPLIER))
        multiplier_entry = tk.Entry(multiplier_frame, textvariable=self.mhl_reward_multiplier_var,
                                     font=FONT_BODY, relief='solid', bd=1,
                                     highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER, width=10)
        multiplier_entry.grid(row=0, column=1, sticky='w', padx=(0, 10))

        ModernButton(multiplier_frame, text='Зберегти', style='primary',
                    command=self._save_mhl_reward_multiplier).grid(row=0, column=3, sticky='e', padx=(10, 0))

        # --- TMA Action Percent File (TMA.xlsx) ---
        tk.Label(parent, text='Файл відсотків акції (TMA):', font=FONT_SUBTITLE, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', pady=(PAD_Y, 4))
        tk.Label(parent, text='Перший аркуш, колонки: SKU_UKR та %', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).pack(anchor='w', pady=(0, PAD_Y))

        tma_file_frame = tk.Frame(parent, bg=BG_COLOR)
        tma_file_frame.pack(fill='x', pady=(0, PAD_Y))
        tma_file_frame.columnconfigure(1, weight=1)

        tk.Label(tma_file_frame, text='Файл:', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(row=0, column=0, sticky='w', padx=(0, 10))
        self.tma_action_file_var = tk.StringVar(value=TMA_ACTION_PERCENT_FILE)
        tma_file_entry = tk.Entry(tma_file_frame, textvariable=self.tma_action_file_var,
                                   font=FONT_BODY, relief='solid', bd=1,
                                   highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        tma_file_entry.grid(row=0, column=1, sticky='ew', padx=(0, 10))
        ModernButton(tma_file_frame, text='Обрати...', style='outline',
                    command=self._select_tma_action_file).grid(row=0, column=2, sticky='e')

        tma_btn_frame = tk.Frame(parent, bg=BG_COLOR)
        tma_btn_frame.pack(fill='x', pady=(0, PAD_Y))
        ModernButton(tma_btn_frame, text='🔄 Перезавантажити файл відсотків', style='primary',
                    command=self._reload_tma_action_file).pack(side='left')
        ModernButton(tma_btn_frame, text='📋 Показати завантажені відсотки', style='secondary',
                    command=self._show_tma_action_percents).pack(side='left', padx=(10, 0))

        # TMA Preview area
        tma_preview_frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        tma_preview_frame.pack(fill='x', pady=(0, PAD_Y))
        tma_preview_frame.grid_rowconfigure(0, weight=1)
        tma_preview_frame.grid_columnconfigure(0, weight=1)

        self.tma_action_preview = tk.Text(tma_preview_frame, font=FONT_MONO, bg=SURFACE, fg=TEXT_MAIN,
                                           wrap='none', height=8, relief='flat', bd=0)
        self.tma_action_preview.grid(row=0, column=0, sticky='nsew', padx=8, pady=8)
        vsb = ttk.Scrollbar(tma_preview_frame, orient='vertical', command=self.tma_action_preview.yview)
        hsb = ttk.Scrollbar(tma_preview_frame, orient='horizontal', command=self.tma_action_preview.xview)
        self.tma_action_preview.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self._refresh_tma_action_preview()

        # --- MHL Action Percent File ---
        tk.Label(parent, text='Файл відсотків акції (MHL):', font=FONT_SUBTITLE, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', pady=(PAD_Y, 4))
        tk.Label(parent, text='Аркуш "для договору", колонки: SKU_UKR та Q3\'26, %', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).pack(anchor='w', pady=(0, PAD_Y))

        mhl_file_frame = tk.Frame(parent, bg=BG_COLOR)
        mhl_file_frame.pack(fill='x', pady=(0, PAD_Y))
        mhl_file_frame.columnconfigure(1, weight=1)

        tk.Label(mhl_file_frame, text='Файл:', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(row=0, column=0, sticky='w', padx=(0, 10))
        self.mhl_action_file_var = tk.StringVar(value=MHL_ACTION_PERCENT_FILE)
        mhl_file_entry = tk.Entry(mhl_file_frame, textvariable=self.mhl_action_file_var,
                                   font=FONT_BODY, relief='solid', bd=1,
                                   highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        mhl_file_entry.grid(row=0, column=1, sticky='ew', padx=(0, 10))
        ModernButton(mhl_file_frame, text='Обрати...', style='outline',
                    command=self._select_mhl_action_file).grid(row=0, column=2, sticky='e')

        mhl_btn_frame = tk.Frame(parent, bg=BG_COLOR)
        mhl_btn_frame.pack(fill='x', pady=(0, PAD_Y))
        ModernButton(mhl_btn_frame, text='🔄 Перезавантажити файл відсотків', style='primary',
                    command=self._reload_mhl_action_file).pack(side='left')
        ModernButton(mhl_btn_frame, text='📋 Показати завантажені відсотки', style='secondary',
                    command=self._show_mhl_action_percents).pack(side='left', padx=(10, 0))

        # MHL Preview area
        mhl_preview_frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        mhl_preview_frame.pack(fill='x', pady=(0, PAD_Y))
        mhl_preview_frame.grid_rowconfigure(0, weight=1)
        mhl_preview_frame.grid_columnconfigure(0, weight=1)

        self.mhl_action_preview = tk.Text(mhl_preview_frame, font=FONT_MONO, bg=SURFACE, fg=TEXT_MAIN,
                                           wrap='none', height=8, relief='flat', bd=0)
        self.mhl_action_preview.grid(row=0, column=0, sticky='nsew', padx=8, pady=8)
        vsb = ttk.Scrollbar(mhl_preview_frame, orient='vertical', command=self.mhl_action_preview.yview)
        hsb = ttk.Scrollbar(mhl_preview_frame, orient='horizontal', command=self.mhl_action_preview.xview)
        self.mhl_action_preview.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self._refresh_mhl_action_preview()

        # --- Debug: Show MHL brands ---
        debug_frame = tk.Frame(parent, bg=BG_COLOR)
        debug_frame.pack(fill='x', pady=(PAD_Y, 0))
        ModernButton(debug_frame, text='🔍 Показати бренди в MHL даних', style='secondary',
                    command=self.app._show_mhl_brands).pack(side='left')

        # --- LA_Price File (Service Prices) ---
        tk.Label(parent, text='Файл цін послуг (LA_Price):', font=FONT_SUBTITLE, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', pady=(PAD_Y, 4))
        tk.Label(parent, text='Аркуш "LA_Price" або перший аркуш, колонки: Brand та Price', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).pack(anchor='w', pady=(0, PAD_Y))

        la_file_frame = tk.Frame(parent, bg=BG_COLOR)
        la_file_frame.pack(fill='x', pady=(0, PAD_Y))
        la_file_frame.columnconfigure(1, weight=1)

        tk.Label(la_file_frame, text='Файл:', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(row=0, column=0, sticky='w', padx=(0, 10))
        self.la_price_file_var = tk.StringVar(value=LA_PRICE_FILE)
        la_file_entry = tk.Entry(la_file_frame, textvariable=self.la_price_file_var,
                                  font=FONT_BODY, relief='solid', bd=1,
                                  highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        la_file_entry.grid(row=0, column=1, sticky='ew', padx=(0, 10))
        ModernButton(la_file_frame, text='Обрати...', style='outline',
                    command=self._select_la_price_file).grid(row=0, column=2, sticky='e')

        la_btn_frame = tk.Frame(parent, bg=BG_COLOR)
        la_btn_frame.pack(fill='x', pady=(0, PAD_Y))
        ModernButton(la_btn_frame, text='🔄 Перезавантажити файл цін', style='primary',
                    command=self._reload_la_price_file).pack(side='left')
        ModernButton(la_btn_frame, text='📋 Показати завантажені ціни', style='secondary',
                    command=self._show_la_prices).pack(side='left', padx=(10, 0))

        # LA_Price Preview area
        la_preview_frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        la_preview_frame.pack(fill='x', pady=(0, PAD_Y))
        la_preview_frame.grid_rowconfigure(0, weight=1)
        la_preview_frame.grid_columnconfigure(0, weight=1)

        self.la_price_preview = tk.Text(la_preview_frame, font=FONT_MONO, bg=SURFACE, fg=TEXT_MAIN,
                                         wrap='none', height=8, relief='flat', bd=0)
        self.la_price_preview.grid(row=0, column=0, sticky='nsew', padx=8, pady=8)
        vsb2 = ttk.Scrollbar(la_preview_frame, orient='vertical', command=self.la_price_preview.yview)
        hsb2 = ttk.Scrollbar(la_preview_frame, orient='horizontal', command=self.la_price_preview.xview)
        self.la_price_preview.configure(yscrollcommand=vsb2.set, xscrollcommand=hsb2.set)
        vsb2.grid(row=0, column=1, sticky='ns')
        hsb2.grid(row=1, column=0, sticky='ew')

        self._refresh_la_price_preview()

    def _select_mhl_action_file(self):
        filename = filedialog.askopenfilename(
            title='Виберіть файл відсотків акції (MHL)',
            initialdir=str(Path.cwd()),
            filetypes=[('Excel', '*.xlsx *.xls'), ('Усі файли', '*.*')]
        )
        if filename:
            self.mhl_action_file_var.set(filename)

    def _reload_mhl_action_file(self):
        filepath = self.mhl_action_file_var.get().strip()
        if not filepath:
            messagebox.showwarning('Попередження', 'Вкажіть шлях до файлу', parent=self)
            return
        try:
            set_mhl_action_percent_file(filepath)
            self._refresh_mhl_action_preview()
            self.app.status_var.set(f'Файл відсотків акції завантажено: {filepath}')
        except Exception as e:
            messagebox.showerror('Помилка', f'Не вдалося завантажити файл: {e}', parent=self)

    def _show_mhl_action_percents(self):
        """Show loaded action percentages in a separate window."""
        from convertor import load_mhl_action_percents
        filepath = self.mhl_action_file_var.get().strip()
        percents = load_mhl_action_percents(filepath) if filepath else {}

        win = tk.Toplevel(self)
        win.title('Завантажені відсотки акції')
        win.geometry('600x400')
        win.configure(bg=BG_COLOR)

        tree = ttk.Treeview(win, columns=('SKU', 'PERCENT'), show='headings')
        tree.heading('SKU', text='SKU / Назва')
        tree.heading('PERCENT', text='Відсоток (%)')
        tree.column('SKU', width=400, anchor='w')
        tree.column('PERCENT', width=100, anchor='center')

        for sku, pct in sorted(percents.items()):
            tree.insert('', 'end', values=(sku, f'{pct:.2f}'))

        vsb = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky='nsew', padx=PAD_X, pady=PAD_X)
        vsb.grid(row=0, column=1, sticky='ns')
        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)

        tk.Label(win, text=f'Всього SKU: {len(percents)}', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).grid(row=1, column=0, sticky='w', padx=PAD_X, pady=(0, PAD_X))

    def _refresh_mhl_action_preview(self):
        from convertor import load_mhl_action_percents
        filepath = self.mhl_action_file_var.get().strip()
        percents = load_mhl_action_percents(filepath) if filepath else {}

        self.mhl_action_preview.delete('1.0', 'end')
        if not percents:
            self.mhl_action_preview.insert('end', f'Файл не знайдено або порожній: {filepath}')
            return

        self.mhl_action_preview.insert('end', f'Файл: {filepath}\n')
        self.mhl_action_preview.insert('end', f'Завантажено SKU: {len(percents)}\n\n')
        for sku, pct in sorted(percents.items()):
            self.mhl_action_preview.insert('end', f'{sku}: {pct:.2f}%\n')

    # --- TMA Action Percent Methods ---
    def _select_tma_action_file(self):
        filename = filedialog.askopenfilename(
            title='Виберіть файл відсотків акції (TMA)',
            initialdir=str(Path.cwd()),
            filetypes=[('Excel', '*.xlsx *.xls'), ('Усі файли', '*.*')]
        )
        if filename:
            self.tma_action_file_var.set(filename)

    def _reload_tma_action_file(self):
        from convertor import set_tma_action_percent_file
        filepath = self.tma_action_file_var.get().strip()
        if filepath:
            set_tma_action_percent_file(filepath)
            self.app.status_var.set(f'TMA відсотки перезавантажено: {Path(filepath).name}')
            self._refresh_tma_action_preview()
        else:
            messagebox.showwarning('Помилка', 'Спочатку оберіть файл', parent=self)

    def _show_tma_action_percents(self):
        from convertor import load_tma_action_percents
        filepath = self.tma_action_file_var.get().strip()
        percents = load_tma_action_percents(filepath) if filepath else {}

        win = tk.Toplevel(self)
        win.title('TMA Відсотки акції')
        win.geometry('400x500')
        win.configure(bg=BG_COLOR)

        text = tk.Text(win, font=FONT_MONO, bg=SURFACE, fg=TEXT_MAIN, wrap='none')
        text.pack(fill='both', expand=True, padx=PAD_X, pady=PAD_X)

        if not percents:
            text.insert('end', f'Файл не знайдено або порожній: {filepath}')
        else:
            text.insert('end', f'Файл: {filepath}\n')
            text.insert('end', f'Завантажено SKU: {len(percents)}\n\n')
            for sku, pct in sorted(percents.items()):
                text.insert('end', f'{sku}: {pct:.2f}%\n')

        tk.Label(win, text=f'Всього SKU: {len(percents)}', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).pack(anchor='w', padx=PAD_X, pady=(0, PAD_X))

    def _refresh_tma_action_preview(self):
        from convertor import load_tma_action_percents
        filepath = self.tma_action_file_var.get().strip()
        percents = load_tma_action_percents(filepath) if filepath else {}

        self.tma_action_preview.delete('1.0', 'end')
        if not percents:
            self.tma_action_preview.insert('end', f'Файл не знайдено або порожній: {filepath}')
            return

        self.tma_action_preview.insert('end', f'Файл: {filepath}\n')
        self.tma_action_preview.insert('end', f'Завантажено SKU: {len(percents)}\n\n')
        for sku, pct in sorted(percents.items()):
            self.tma_action_preview.insert('end', f'{sku}: {pct:.2f}%\n')

    def _save_mhl_reward_multiplier(self):
        """Save MHL reward multiplier to config."""
        try:
            value = float(self.mhl_reward_multiplier_var.get().strip() or 0.1)
            if not 0 <= value <= 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning('Помилка', 'Відсоток має бути числом від 0 до 1 (напр. 0.1 для 10%)', parent=self)
            return

        import config as cfg
        cfg.MHL_REWARD_MULTIPLIER = value
        self.app.status_var.set(f'MHL відсоток винагороди оновлено: {value:.2f}')
        messagebox.showinfo('Збережено', f'MHL відсоток винагороди: {value:.2f}\n(застосовано до поточної сесії)', parent=self)

    # --- LA_Price Methods ---
    def _select_la_price_file(self):
        filename = filedialog.askopenfilename(
            title='Виберіть файл цін послуг (LA_Price)',
            initialdir=str(Path.cwd()),
            filetypes=[('Excel', '*.xlsx *.xls'), ('Усі файли', '*.*')]
        )
        if filename:
            self.la_price_file_var.set(filename)

    def _reload_la_price_file(self):
        filepath = self.la_price_file_var.get().strip()
        if not filepath:
            messagebox.showwarning('Попередження', 'Вкажіть шлях до файлу', parent=self)
            return
        try:
            set_la_price_file(filepath)
            self._refresh_la_price_preview()
            self.app.status_var.set(f'Файл цін послуг завантажено: {filepath}')
        except Exception as e:
            messagebox.showerror('Помилка', f'Не вдалося завантажити файл: {e}', parent=self)

    def _show_la_prices(self):
        """Show loaded service prices in a separate window."""
        from convertor import get_all_la_prices
        filepath = self.la_price_file_var.get().strip()
        prices = get_all_la_prices() if filepath else SERVICE_PRICES

        win = tk.Toplevel(self)
        win.title('Завантажені ціни послуг')
        win.geometry('500x400')
        win.configure(bg=BG_COLOR)

        tree = ttk.Treeview(win, columns=('BRAND', 'PRICE'), show='headings')
        tree.heading('BRAND', text='Бренд')
        tree.heading('PRICE', text='Ціна (грн.)')
        tree.column('BRAND', width=300, anchor='w')
        tree.column('PRICE', width=150, anchor='e')

        for brand, price in sorted(prices.items()):
            tree.insert('', 'end', values=(brand, f'{price:.2f}'))

        vsb = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky='nsew', padx=PAD_X, pady=PAD_X)
        vsb.grid(row=0, column=1, sticky='ns')
        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)

        tk.Label(win, text=f'Всього брендів: {len(prices)}', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).grid(row=1, column=0, sticky='w', padx=PAD_X, pady=(0, PAD_X))

    def _refresh_la_price_preview(self):
        from convertor import get_all_la_prices
        filepath = self.la_price_file_var.get().strip()
        prices = get_all_la_prices() if filepath else SERVICE_PRICES

        self.la_price_preview.delete('1.0', 'end')
        if not prices:
            self.la_price_preview.insert('end', f'Файл не знайдено або порожній: {filepath}')
            return

        self.la_price_preview.insert('end', f'Файл: {filepath}\n')
        self.la_price_preview.insert('end', f'Завантажено брендів: {len(prices)}\n\n')
        for brand, price in sorted(prices.items()):
            self.la_price_preview.insert('end', f'{brand}: {price:.2f} грн.\n')

    def _build_brand_tab(self, parent):
        """SKU qualification tab - assign products (SKU) to ЛІК or ДД category, grouped by brand."""
        tk.Label(
            parent,
            text='Кваліфікація продуктів (SKU): призначайте кожен продукт категорії ЛІК або ДД.\n'
                 'Список відсортовано за брендом для зручності. Ці налаштування мають вищий пріоритет над config.py.',
            font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED, wraplength=600, justify='left'
        ).pack(anchor='w', pady=(0, PAD_Y))

        # Toolbar
        toolbar = tk.Frame(parent, bg=BG_COLOR)
        toolbar.pack(fill='x', pady=(0, 8))

        ModernButton(toolbar, text='➕ Додати продукт', style='primary', command=self._add_sku).pack(side='left', padx=(0, 8))
        ModernButton(toolbar, text='🔄 Оновити', style='secondary', command=self._refresh_sku_table).pack(side='left')
        ModernButton(toolbar, text='📥 Завантажити з MHL', style='outline', command=self._load_skus_from_mhl).pack(side='left', padx=(8, 0))
        ModernButton(toolbar, text='🗄 Завантажити всі з БД', style='outline', command=self._load_all_skus_from_db).pack(side='left', padx=(8, 0))

        # Bulk category change buttons
        tk.Frame(toolbar, bg=BG_COLOR, width=20).pack(side='left')  # spacer
        ModernButton(toolbar, text='🔄 В ЛІК', style='success', command=lambda: self._bulk_change_category('ЛІК')).pack(side='left', padx=(8, 0))
        ModernButton(toolbar, text='🔄 В ДД', style='warning', command=lambda: self._bulk_change_category('ДД')).pack(side='left', padx=(8, 0))

        # File selector for "для договорів" file
        file_frame = tk.Frame(parent, bg=BG_COLOR)
        file_frame.pack(fill='x', pady=(0, 8))
        file_frame.columnconfigure(1, weight=1)

        tk.Label(file_frame, text='Файл "для договорів":', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(row=0, column=0, sticky='w', padx=(0, 10))
        self.contract_file_var = tk.StringVar(value='')
        contract_file_entry = tk.Entry(file_frame, textvariable=self.contract_file_var,
                                        font=FONT_BODY, relief='solid', bd=1,
                                        highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        contract_file_entry.grid(row=0, column=1, sticky='ew', padx=(0, 10))
        ModernButton(file_frame, text='Обрати...', style='outline',
                    command=self._select_contract_file).grid(row=0, column=2, sticky='e', padx=(0, 8))
        ModernButton(file_frame, text='📥 Завантажити з файлу договорів', style='primary',
                    command=self._load_skus_from_contract_file).grid(row=0, column=3, sticky='e')

        # Table
        table_frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        table_frame.pack(fill='both', expand=True)

        columns = ('SKU', 'BRAND', 'CATEGORY')
        self.sku_tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='extended')

        col_config = {
            'SKU': ('Продукт (SKU)', 350),
            'BRAND': ('Бренд', 200),
            'CATEGORY': ('Категорія', 100),
        }
        for col, (title, width) in col_config.items():
            self.sku_tree.heading(col, text=title)
            self.sku_tree.column(col, width=width, anchor='w' if col != 'CATEGORY' else 'center')

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.sku_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient='horizontal', command=self.sku_tree.xview)
        self.sku_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.sku_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.sku_tree.bind('<Double-1>', self._on_sku_double_click)
        self.sku_tree.bind('<Delete>', lambda e: self._delete_selected_skus())

        # Legend
        legend_frame = tk.Frame(parent, bg=BG_COLOR)
        legend_frame.pack(fill='x', pady=(PAD_Y, 0))
        tk.Label(legend_frame, text='Легенда:', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).pack(side='left')
        tk.Label(legend_frame, text='🟦 ЛІК = Лікарські засоби', font=FONT_SMALL, bg=BG_COLOR, fg='#1a56a0').pack(side='left', padx=(16, 8))
        tk.Label(legend_frame, text='🟧 ДД = Не лікарські засоби', font=FONT_SMALL, bg=BG_COLOR, fg='#e67e22').pack(side='left', padx=(0, 8))
        tk.Label(legend_frame, text='(подвійний клік для редагування, Delete для видалення)', font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).pack(side='left', padx=(16, 0))

        self._refresh_sku_table()

    def _refresh_sku_table(self):
        for item in self.sku_tree.get_children():
            self.sku_tree.delete(item)
        df = get_all_sku_classifications()
        for _, row in df.iterrows():
            cat = row['CATEGORY']
            # Color coding by category
            item = self.sku_tree.insert('', 'end', values=(row['SKU'], row['BRAND'], cat))
            if cat == 'ЛІК':
                self.sku_tree.item(item, tags=('lik',))
            elif cat == 'ДД':
                self.sku_tree.item(item, tags=('dd',))
        # Configure tags for color
        self.sku_tree.tag_configure('lik', foreground='#1a56a0')
        self.sku_tree.tag_configure('dd', foreground='#e67e22')

    def _on_sku_double_click(self, event):
        item = self.sku_tree.selection()[0] if self.sku_tree.selection() else None
        if not item:
            return
        values = self.sku_tree.item(item, 'values')
        self._edit_sku_dialog(values)

    def _add_sku(self):
        self._edit_sku_dialog()

    def _edit_sku_dialog(self, values=None):
        """Dialog to add/edit SKU classification."""
        dialog = tk.Toplevel(self)
        dialog.title('Редагувати продукт' if values else 'Додати продукт')
        dialog.geometry('500x250')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f'+{x}+{y}')

        fields = {}
        labels = [
            ('SKU', 'Назва продукту / SKU (як в базі даних, обов\'язково)'),
            ('BRAND', 'Бренд (для групування в списку)'),
            ('CATEGORY', 'Категорія (ЛІК або ДД)'),
        ]
        for i, (key, label) in enumerate(labels):
            tk.Label(dialog, text=label, font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).grid(
                row=i, column=0, sticky='w', padx=PAD_X, pady=(PAD_Y, 4)
            )
            if key == 'CATEGORY':
                var = tk.StringVar(value='ЛІК')
                combo = ttk.Combobox(dialog, textvariable=var, values=['ЛІК', 'ДД'],
                                      state='readonly', font=FONT_BODY, width=55)
                combo.grid(row=i, column=0, sticky='ew', padx=PAD_X, pady=(0, 4))
                fields[key] = var
            else:
                entry = tk.Entry(dialog, font=FONT_BODY, relief='solid', bd=1,
                                highlightthickness=1, highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
                entry.grid(row=i, column=0, sticky='ew', padx=PAD_X, pady=(0, 4))
                fields[key] = entry

        dialog.grid_columnconfigure(0, weight=1)

        # Pre-fill if editing
        if values:
            fields['SKU'].insert(0, values[0])
            fields['SKU'].config(state='readonly')
            fields['BRAND'].insert(0, values[1])
            fields['CATEGORY'].set(values[2])

        def save():
            sku = fields['SKU'].get().strip()
            brand = fields['BRAND'].get().strip()
            category = fields['CATEGORY'].get().strip()

            if not sku:
                messagebox.showwarning('Помилка', 'Назва продукту не може бути порожньою', parent=dialog)
                return

            if category not in ('ЛІК', 'ДД'):
                messagebox.showwarning('Помилка', 'Категорія має бути ЛІК або ДД', parent=dialog)
                return

            try:
                set_sku_classification(sku, category, brand)
                self._refresh_sku_table()
                dialog.destroy()
                self.app.status_var.set(f'Продукт оновлено: {sku} = {category}')
            except Exception as e:
                messagebox.showerror('Помилка', f'Не вдалося зберегти: {e}', parent=dialog)

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.grid(row=3, column=0, sticky='ew', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='Скасувати', style='secondary', command=dialog.destroy).pack(side='right', padx=(8, 0))
        ModernButton(btn_frame, text='Зберегти', style='success', command=save).pack(side='right')

    def _delete_selected_skus(self):
        selected = self.sku_tree.selection()
        if not selected:
            return
        if messagebox.askyesno('Підтвердження', f'Видалити {len(selected)} продукт(ів) з кваліфікації?', parent=self):
            for item in selected:
                values = self.sku_tree.item(item, 'values')
                sku = values[0]
                remove_sku_classification(sku)
            self._refresh_sku_table()

    def _bulk_change_category(self, new_category):
        """Change category of selected SKUs to new_category (ЛІК or ДД)."""
        selected = self.sku_tree.selection()
        if not selected:
            messagebox.showinfo('Інформація', 'Виберіть продукти для зміни категорії.', parent=self)
            return

        # Confirm if changing many items
        if len(selected) > 10:
            if not messagebox.askyesno('Підтвердження',
                    f'Змінити категорію на "{new_category}" для {len(selected)} продуктів?', parent=self):
                return
        else:
            # Show which items will be changed
            items_info = []
            for item in selected:
                values = self.sku_tree.item(item, 'values')
                items_info.append(f'  • {values[0]} [{values[1]}] — {values[2]} → {new_category}')
            if not messagebox.askyesno('Підтвердження',
                    f'Змінити категорію на "{new_category}" для:\n' + '\n'.join(items_info), parent=self):
                return

        changed = 0
        for item in selected:
            values = self.sku_tree.item(item, 'values')
            sku = values[0]
            brand = values[1]
            old_cat = values[2]
            if old_cat != new_category:
                set_sku_classification(sku, new_category, brand)
                changed += 1

        if changed > 0:
            self._refresh_sku_table()
            self.app.status_var.set(f'Категорію змінено для {changed} продукт(ів) на {new_category}')
            messagebox.showinfo('Готово', f'Категорію змінено для {changed} продукт(ів) на {new_category}.', parent=self)
        else:
            messagebox.showinfo('Інформація', 'Всі вибрані продукти вже мають цю категорію.', parent=self)

    def _load_skus_from_mhl(self):
        """Load unique SKUs from current MHL data and add unknown ones to classification."""
        if self.app.df_mhl is None or self.app.df_mhl.empty:
            messagebox.showinfo('MHL дані', 'MHL дані не завантажені. Спочатку натисніть "Завантажити дані".', parent=self)
            return

        sku_col = 'SKU_UA' if 'SKU_UA' in self.app.df_mhl.columns else 'Назва товару'
        brand_col = 'BRAND' if 'BRAND' in self.app.df_mhl.columns else 'brand'

        if sku_col not in self.app.df_mhl.columns or brand_col not in self.app.df_mhl.columns:
            messagebox.showinfo('MHL дані', 'В MHL даних відсутні колонки SKU_UA або BRAND.', parent=self)
            return

        # Get unique SKU-Brand pairs from MHL data
        work_df = self.app.df_mhl[[sku_col, brand_col]].dropna().drop_duplicates()
        work_df[sku_col] = work_df[sku_col].astype(str).str.strip()
        work_df[brand_col] = work_df[brand_col].astype(str).str.strip()
        work_df = work_df[(work_df[sku_col] != '') & (work_df[brand_col] != '')]

        if work_df.empty:
            messagebox.showinfo('MHL дані', 'SKU не знайдені в MHL даних.', parent=self)
            return

        # Get current classifications
        current_df = get_all_sku_classifications()
        current_skus = set(current_df['SKU'].str.casefold()) if not current_df.empty else set()

        # Get config-based SKUs (from brand classification)
        lik_skus = set()
        dd_skus = set()
        if self.app.df_mhl is not None:
            lik_sku_list = get_combined_lik_skus(self.app.df_mhl)
            dd_sku_list = get_combined_dd_skus(self.app.df_mhl)
            lik_skus = {s['sku'].casefold() for s in lik_sku_list}
            dd_skus = {s['sku'].casefold() for s in dd_sku_list}
        all_known_skus = lik_skus | dd_skus

        # Find unknown SKUs (not in local classification and not in config)
        unknown = work_df[~work_df[sku_col].str.casefold().isin(current_skus | all_known_skus)]

        if unknown.empty:
            messagebox.showinfo('MHL дані', 'Всі продукти вже мають кваліфікацію.', parent=self)
            return

        # Show dialog to select which unknown SKUs to add
        dialog = tk.Toplevel(self)
        dialog.title('Невідомі продукти з MHL')
        dialog.geometry('600x500')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f'+{x}+{y}')

        tk.Label(dialog, text=f'Знайдено {len(unknown)} невідомих продуктів у MHL даних:',
                 font=FONT_SUBTITLE, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', padx=PAD_X, pady=PAD_Y)

        # List with checkboxes - show SKU + Brand
        frame = tk.Frame(dialog, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill='both', expand=True, padx=PAD_X, pady=(0, PAD_X))

        canvas = tk.Canvas(frame, bg=SURFACE, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollable_frame = tk.Frame(canvas, bg=SURFACE)
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')

        checkboxes = {}
        for _, row in unknown.iterrows():
            sku = row[sku_col]
            brand = row[brand_col]
            display_text = f'{sku}  [{brand}]'
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(scrollable_frame, text=display_text, variable=var,
                                 font=FONT_BODY, bg=SURFACE, fg=TEXT_MAIN,
                                 activebackground=SURFACE, anchor='w')
            cb.pack(fill='x', padx=8, pady=2)
            checkboxes[sku] = {'var': var, 'brand': brand}

        scrollable_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox('all'))
        canvas.pack(side='left', fill='both', expand=True, padx=8, pady=8)
        scrollbar.pack(side='right', fill='y', pady=8)

        def add_selected():
            added = 0
            for sku, data in checkboxes.items():
                if data['var'].get():
                    # Try to guess category from config
                    cat = None
                    sku_lower = sku.casefold()
                    if sku_lower in lik_skus:
                        cat = 'ЛІК'
                    elif sku_lower in dd_skus:
                        cat = 'ДД'
                    else:
                        cat = 'ЛІК'  # default
                    set_sku_classification(sku, cat, data['brand'])
                    added += 1
            self._refresh_sku_table()
            dialog.destroy()
            self.app.status_var.set(f'Додано {added} продукт(ів) до кваліфікації')
            messagebox.showinfo('Готово', f'Додано {added} продукт(ів) до кваліфікації.', parent=self)

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.pack(fill='x', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='Скасувати', style='secondary', command=dialog.destroy).pack(side='right', padx=(8, 0))
        ModernButton(btn_frame, text='Додати вибране', style='success', command=add_selected).pack(side='right')

    def _load_all_skus_from_db(self):
        """Load all unique SKUs with brands from the database for classification."""
        # Show loading indicator
        self.app.status_var.set('Завантаження продуктів з бази даних...')
        self.update_idletasks()

        try:
            df_skus = get_all_active_skus()
        except DataLoadError as e:
            messagebox.showerror('Помилка', f'Не вдалося завантажити продукти з БД: {e}', parent=self)
            self.app.status_var.set('Помилка завантаження')
            return
        except Exception as e:
            messagebox.showerror('Помилка', f'Неочікувана помилка: {e}', parent=self)
            self.app.status_var.set('Помилка завантаження')
            return

        if df_skus.empty:
            messagebox.showinfo('База даних', 'Продукти не знайдені.', parent=self)
            self.app.status_var.set('Продукти не знайдені')
            return

        # Get current classifications
        current_df = get_all_sku_classifications()
        current_skus = set(current_df['SKU'].str.casefold()) if not current_df.empty else set()

        # Get config-based SKUs
        lik_skus = set()
        dd_skus = set()
        if self.app.df_mhl is not None:
            lik_sku_list = get_combined_lik_skus(self.app.df_mhl)
            dd_sku_list = get_combined_dd_skus(self.app.df_mhl)
            lik_skus = {s['sku'].casefold() for s in lik_sku_list}
            dd_skus = {s['sku'].casefold() for s in dd_sku_list}
        all_known_skus = lik_skus | dd_skus

        # Find SKUs not yet classified
        df_skus['SKU_lower'] = df_skus['SKU_UA'].str.casefold()
        unknown = df_skus[~df_skus['SKU_lower'].isin(current_skus | all_known_skus)]

        # Show dialog with ALL products, but pre-check unknown ones
        dialog = tk.Toplevel(self)
        dialog.title('Всі продукти з бази даних')
        dialog.geometry('700x600')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f'+{x}+{y}')

        tk.Label(dialog, text=f'Знайдено {len(df_skus)} продуктів у базі даних ({len(unknown)} нових):',
                 font=FONT_SUBTITLE, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', padx=PAD_X, pady=PAD_Y)

        # Filter/search frame
        filter_frame = tk.Frame(dialog, bg=BG_COLOR)
        filter_frame.pack(fill='x', padx=PAD_X, pady=(0, PAD_Y))
        tk.Label(filter_frame, text='Фільтр:', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).pack(side='left')
        filter_var = tk.StringVar()
        filter_entry = tk.Entry(filter_frame, textvariable=filter_var, font=FONT_BODY, width=40,
                                 relief='solid', bd=1, highlightthickness=1,
                                 highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        filter_entry.pack(side='left', padx=(8, 0), fill='x', expand=True)

        # List with checkboxes - show SKU + Brand
        frame = tk.Frame(dialog, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill='both', expand=True, padx=PAD_X, pady=(0, PAD_X))

        canvas = tk.Canvas(frame, bg=SURFACE, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollable_frame = tk.Frame(canvas, bg=SURFACE)
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')

        checkboxes = {}
        row_frames = []  # Store frames for filtering

        for _, row in df_skus.iterrows():
            sku = row['SKU_UA']
            brand = row['BRAND']
            display_text = f'{sku}  [{brand}]'
            # Pre-check if unknown (not yet classified)
            is_unknown = sku.casefold() in unknown['SKU_UA'].str.casefold().values
            var = tk.BooleanVar(value=is_unknown)
            cb = tk.Checkbutton(scrollable_frame, text=display_text, variable=var,
                                 font=FONT_BODY, bg=SURFACE, fg=TEXT_MAIN,
                                 activebackground=SURFACE, anchor='w')
            cb.pack(fill='x', padx=8, pady=2)
            checkboxes[sku] = {'var': var, 'brand': brand, 'frame': cb}
            row_frames.append((sku.casefold(), cb))

        def apply_filter(*args):
            """Filter the list based on search text."""
            search = filter_var.get().strip().casefold()
            for sku_lower, cb in row_frames:
                if search in sku_lower:
                    cb.pack(fill='x', padx=8, pady=2)
                else:
                    cb.pack_forget()
            scrollable_frame.update_idletasks()
            canvas.config(scrollregion=canvas.bbox('all'))

        filter_var.trace_add('write', apply_filter)

        scrollable_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox('all'))
        canvas.pack(side='left', fill='both', expand=True, padx=8, pady=8)
        scrollbar.pack(side='right', fill='y', pady=8)

        def add_selected():
            added = 0
            for sku, data in checkboxes.items():
                if data['var'].get():
                    # Try to guess category from config
                    cat = None
                    sku_lower = sku.casefold()
                    if sku_lower in lik_skus:
                        cat = 'ЛІК'
                    elif sku_lower in dd_skus:
                        cat = 'ДД'
                    else:
                        cat = 'ЛІК'  # default
                    set_sku_classification(sku, cat, data['brand'])
                    added += 1
            self._refresh_sku_table()
            dialog.destroy()
            self.app.status_var.set(f'Додано/оновлено {added} продукт(ів) до кваліфікації')
            messagebox.showinfo('Готово', f'Додано/оновлено {added} продукт(ів) до кваліфікації.', parent=self)

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.pack(fill='x', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='Скасувати', style='secondary', command=dialog.destroy).pack(side='right', padx=(8, 0))
        ModernButton(btn_frame, text='Додати вибране', style='success', command=add_selected).pack(side='right')

    def _select_contract_file(self):
        """Select the 'для договорів' Excel file."""
        filename = filedialog.askopenfilename(
            title='Виберіть файл "для договорів"',
            initialdir=str(Path.cwd()),
            filetypes=[('Excel', '*.xlsx *.xls'), ('Усі файли', '*.*')]
        )
        if filename:
            self.contract_file_var.set(filename)

    def _load_skus_from_contract_file(self):
        """Load SKUs from the 'для договорів' file with two tables (ЛИК and ДД)."""
        filepath = self.contract_file_var.get().strip()
        if not filepath:
            messagebox.showwarning('Попередження', 'Спочатку оберіть файл.', parent=self)
            return

        self.app.status_var.set('Завантаження продуктів з файлу договорів...')
        self.update_idletasks()

        try:
            import openpyxl
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            # Find sheet "для договорів" or first sheet
            sheet_name = None
            for name in wb.sheetnames:
                if 'для договор' in name.lower() or 'договор' in name.lower():
                    sheet_name = name
                    break
            if sheet_name is None:
                sheet_name = wb.sheetnames[0]
            ws = wb[sheet_name]

            # Read all rows to find the two tables
            rows = list(ws.iter_rows(values_only=True))

            # Find header row with SKU_ENG, SKU_UKR, Q3'26, %
            header_row_idx = None
            sku_ukr_col = None
            for i, row in enumerate(rows):
                if not row:
                    continue
                row_str = [str(c).strip() if c else '' for c in row]
                if any('SKU_UKR' in c.upper() or 'SKU_УКР' in c.upper() for c in row_str):
                    header_row_idx = i
                    # Find SKU_UKR column index
                    for j, c in enumerate(row_str):
                        if 'SKU_UKR' in c.upper() or 'SKU_УКР' in c.upper():
                            sku_ukr_col = j
                            break
                    break

            if header_row_idx is None or sku_ukr_col is None:
                messagebox.showerror('Помилка', 'Не знайдено колонку SKU_UKR у файлі.', parent=self)
                self.app.status_var.set('Помилка завантаження')
                return

            # Parse both tables: first ЛІК, then after empty row - ДД
            lik_skus = []
            dd_skus = []
            current_category = 'ЛІК'  # First table is ЛІК

            for i in range(header_row_idx + 1, len(rows)):
                row = rows[i]
                if not row or all(c is None or str(c).strip() == '' for c in row):
                    # Empty row - switch to ДД
                    current_category = 'ДД'
                    continue

                sku_ukr = row[sku_ukr_col] if sku_ukr_col < len(row) else None
                if sku_ukr and str(sku_ukr).strip():
                    sku_clean = str(sku_ukr).strip()
                    if current_category == 'ЛІК':
                        lik_skus.append(sku_clean)
                    else:
                        dd_skus.append(sku_clean)

            wb.close()

        except Exception as e:
            messagebox.showerror('Помилка', f'Не вдалося прочитати файл: {e}', parent=self)
            self.app.status_var.set('Помилка завантаження')
            return

        if not lik_skus and not dd_skus:
            messagebox.showinfo('Файл договорів', 'Продукти не знайдені у файлі.', parent=self)
            self.app.status_var.set('Продукти не знайдені')
            return

        # Get current classifications
        current_df = get_all_sku_classifications()
        current_skus = set(current_df['SKU'].str.casefold()) if not current_df.empty else set()

        # Get all active SKUs with brands from database for auto-matching
        self.app.status_var.set('Завантаження брендів з бази даних...')
        self.update_idletasks()
        try:
            from Data.Controller import get_all_active_skus
            db_skus_df = get_all_active_skus()
            # Create mapping: SKU_UA (lowercase) -> BRAND
            sku_to_brand = {}
            if not db_skus_df.empty:
                for _, row in db_skus_df.iterrows():
                    sku_key = str(row['SKU_UA']).strip().casefold()
                    brand = str(row['BRAND']).strip()
                    if sku_key and brand:
                        sku_to_brand[sku_key] = brand
        except Exception as e:
            print(f'Warning: Could not load SKUs from database: {e}')
            sku_to_brand = {}

        # Show dialog with found SKUs
        dialog = tk.Toplevel(self)
        dialog.title('Продукти з файлу договорів')
        dialog.geometry('700x550')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f'+{x}+{y}')

        tk.Label(dialog, text=f'ЛІК: {len(lik_skus)} | ДД: {len(dd_skus)} (всього {len(lik_skus) + len(dd_skus)} продуктів)',
                 font=FONT_SUBTITLE, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', padx=PAD_X, pady=PAD_Y)

        # Filter/search frame
        filter_frame = tk.Frame(dialog, bg=BG_COLOR)
        filter_frame.pack(fill='x', padx=PAD_X, pady=(0, PAD_Y))
        tk.Label(filter_frame, text='Фільтр:', font=FONT_BODY, bg=BG_COLOR, fg=TEXT_MAIN).pack(side='left')
        filter_var = tk.StringVar()
        filter_entry = tk.Entry(filter_frame, textvariable=filter_var, font=FONT_BODY, width=40,
                                 relief='solid', bd=1, highlightthickness=1,
                                 highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
        filter_entry.pack(side='left', padx=(8, 0), fill='x', expand=True)

        # Legend
        legend_frame = tk.Frame(dialog, bg=BG_COLOR)
        legend_frame.pack(fill='x', padx=PAD_X, pady=(0, PAD_Y))
        tk.Label(legend_frame, text='🟢 = знайдено в БД  |  🔴 = не знайдено (вкажіть бренд вручну)',
                 font=FONT_SMALL, bg=BG_COLOR, fg=TEXT_MUTED).pack(anchor='w')

        # List with checkboxes and brand entry
        frame = tk.Frame(dialog, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill='both', expand=True, padx=PAD_X, pady=(0, PAD_X))

        canvas = tk.Canvas(frame, bg=SURFACE, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollable_frame = tk.Frame(canvas, bg=SURFACE)
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')

        # Store data for each SKU: {sku: {'var': BooleanVar, 'category': str, 'brand_var': StringVar, 'frame': widget, 'matched': bool}}
        sku_data = {}
        row_frames = []

        def create_sku_row(parent, sku, category, is_unknown, matched_brand):
            """Create a row with checkbox, SKU name, category badge, and brand entry."""
            row_frame = tk.Frame(parent, bg=SURFACE)
            row_frame.pack(fill='x', padx=8, pady=2)

            var = tk.BooleanVar(value=is_unknown)
            cb = tk.Checkbutton(row_frame, variable=var, bg=SURFACE, activebackground=SURFACE)
            cb.pack(side='left', padx=(0, 8))

            # Status indicator
            if matched_brand:
                status_lbl = tk.Label(row_frame, text='🟢', font=FONT_BODY, bg=SURFACE)
                status_lbl.pack(side='left', padx=(0, 4))
            else:
                status_lbl = tk.Label(row_frame, text='🔴', font=FONT_BODY, bg=SURFACE)
                status_lbl.pack(side='left', padx=(0, 4))

            # SKU name
            sku_lbl = tk.Label(row_frame, text=sku, font=FONT_BODY, bg=SURFACE, fg=TEXT_MAIN, anchor='w', width=35)
            sku_lbl.pack(side='left', fill='x', expand=True)

            # Category badge
            cat_color = '#1a56a0' if category == 'ЛІК' else '#e67e22'
            cat_lbl = tk.Label(row_frame, text=f'[{category}]', font=FONT_SMALL, bg=SURFACE, fg=cat_color)
            cat_lbl.pack(side='left', padx=8)

            # Brand entry (pre-filled if matched, editable)
            brand_var = tk.StringVar(value=matched_brand if matched_brand else '')
            brand_entry = tk.Entry(row_frame, textvariable=brand_var, font=FONT_BODY, width=20,
                                    relief='solid', bd=1, highlightthickness=1,
                                    highlightcolor=BORDER_FOCUS, highlightbackground=BORDER)
            brand_entry.pack(side='left', padx=8)

            # If matched, make it readonly-ish (gray background) but still editable
            if matched_brand:
                brand_entry.configure(bg='#f0f0f0')

            return {
                'var': var,
                'category': category,
                'brand_var': brand_var,
                'frame': row_frame,
                'matched': bool(matched_brand),
                'sku': sku
            }

        # Add ЛІК SKUs
        if lik_skus:
            lbl = tk.Label(scrollable_frame, text='=== ЛІК ===', font=FONT_BOLD if 'FONT_BOLD' in globals() else ('Segoe UI', 10, 'bold'),
                           bg=SURFACE, fg='#1a56a0', anchor='w')
            lbl.pack(fill='x', padx=8, pady=(8, 2))
            for sku in lik_skus:
                is_unknown = sku.casefold() not in current_skus
                matched_brand = sku_to_brand.get(sku.casefold(), '')
                data = create_sku_row(scrollable_frame, sku, 'ЛІК', is_unknown, matched_brand)
                sku_data[sku] = data
                row_frames.append((sku.casefold(), data['frame']))

        # Add ДД SKUs
        if dd_skus:
            lbl = tk.Label(scrollable_frame, text='=== ДД ===', font=FONT_BOLD if 'FONT_BOLD' in globals() else ('Segoe UI', 10, 'bold'),
                           bg=SURFACE, fg='#e67e22', anchor='w')
            lbl.pack(fill='x', padx=8, pady=(8, 2))
            for sku in dd_skus:
                is_unknown = sku.casefold() not in current_skus
                matched_brand = sku_to_brand.get(sku.casefold(), '')
                data = create_sku_row(scrollable_frame, sku, 'ДД', is_unknown, matched_brand)
                sku_data[sku] = data
                row_frames.append((sku.casefold(), data['frame']))

        def apply_filter(*args):
            search = filter_var.get().strip().casefold()
            for sku_lower, row_frame in row_frames:
                if search in sku_lower:
                    row_frame.pack(fill='x', padx=8, pady=2)
                else:
                    row_frame.pack_forget()
            # Also handle section headers
            for widget in scrollable_frame.winfo_children():
                if isinstance(widget, tk.Label) and widget.cget('text').startswith('==='):
                    pass  # Keep headers visible
            scrollable_frame.update_idletasks()
            canvas.config(scrollregion=canvas.bbox('all'))

        filter_var.trace_add('write', apply_filter)

        scrollable_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox('all'))
        canvas.pack(side='left', fill='both', expand=True, padx=8, pady=8)
        scrollbar.pack(side='right', fill='y', pady=8)

        def add_selected():
            added = 0
            for sku, data in sku_data.items():
                if data['var'].get():
                    brand = data['brand_var'].get().strip()
                    if not brand and not data['matched']:
                        # Warn if no brand entered for unmatched
                        if not messagebox.askyesno('Бренд не вказано',
                            f'Для продукту "{sku}" не знайдено бренд в БД і не вказано вручну.\nПродовжити без бренду?',
                            parent=dialog):
                            return
                    set_sku_classification(sku, data['category'], brand)
                    added += 1
            self._refresh_sku_table()
            dialog.destroy()
            self.app.status_var.set(f'Додано/оновлено {added} продукт(ів) з файлу договорів')
            messagebox.showinfo('Готово', f'Додано/оновлено {added} продукт(ів) з файлу договорів.', parent=self)

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.pack(fill='x', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='Скасувати', style='secondary', command=dialog.destroy).pack(side='right', padx=(8, 0))
        ModernButton(btn_frame, text='Додати вибране', style='success', command=add_selected).pack(side='right')


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
                    command=self.preview).pack(side='left', padx=(0, 10))

        # Debug button
        ModernButton(actions, text='🔍 Бренди MHL', style='outline',
                    command=self._show_mhl_brands).pack(side='right')

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

    def _show_mhl_brands(self):
        """Show unique brands in MHL data for debugging."""
        from tkinter import messagebox
        if self.df_mhl is None or self.df_mhl.empty:
            messagebox.showinfo('MHL Brands', 'MHL data not loaded', parent=self.root)
            return

        brands = get_mhl_unique_brands(self.df_mhl)
        if not brands:
            messagebox.showinfo('MHL Brands', 'No brands found in MHL data', parent=self.root)
            return

        # Show in a dialog
        dialog = tk.Toplevel(self.root)
        dialog.title('MHL Brands Debug')
        dialog.geometry('500x600')
        dialog.configure(bg=BG_COLOR)
        dialog.transient(self.root)
        dialog.grab_set()

        # Center on parent
        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f'+{x}+{y}')

        tk.Label(dialog, text=f'Unique brands in MHL data: {len(brands)}',
                 font=FONT_SUBTITLE, bg=BG_COLOR, fg=TEXT_MAIN).pack(anchor='w', padx=PAD_X, pady=PAD_Y)

        # Text area with scrollbar
        frame = tk.Frame(dialog, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill='both', expand=True, padx=PAD_X, pady=(0, PAD_X))

        text = tk.Text(frame, font=FONT_MONO, bg=SURFACE, fg=TEXT_MAIN, wrap='none', relief='flat', bd=0)
        text.pack(side='left', fill='both', expand=True, padx=8, pady=8)

        vsb = ttk.Scrollbar(frame, orient='vertical', command=text.yview)
        vsb.pack(side='right', fill='y', pady=8)
        text.configure(yscrollcommand=vsb.set)

        # Also show which ones match our LIK/DD lists
        lik_brands = {b.lower() for b in get_mhl_lik_brands()}
        dd_brands = {b.lower() for b in get_mhl_dd_brands()}

        for brand in brands:
            brand_lower = brand.lower()
            status = ''
            if brand_lower in lik_brands:
                status = ' ✓ ЛІК'
            elif brand_lower in dd_brands:
                status = ' ✓ ДД'
            else:
                status = ' ✗ UNKNOWN'
            text.insert('end', f'{brand}{status}\n')

        text.config(state='disabled')

        btn_frame = tk.Frame(dialog, bg=BG_COLOR)
        btn_frame.pack(fill='x', padx=PAD_X, pady=PAD_Y)
        ModernButton(btn_frame, text='Закрити', style='secondary', command=dialog.destroy).pack(side='right')

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