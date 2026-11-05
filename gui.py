import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd
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
)
from Data.Controller import get_mhl_data, get_pharmacy_data, load_requisites_file


# ============================================================
# UI CONSTANTS — modern, calm color palette
# ============================================================
BG_COLOR = '#f5f5f5'
CARD_BG = '#ffffff'
PRIMARY = '#1a56a0'
PRIMARY_HOVER = '#154380'
PRIMARY_LIGHT = '#e8eef7'
ACCENT = '#007b5e'
DANGER = '#c0392b'
TEXT_MAIN = '#212529'
TEXT_MUTED = '#6c757d'
BORDER = '#dee2e6'
SHADOW = '#0000001a'

FONT_TITLE = ('Segoe UI', 16, 'bold')
FONT_SUBTITLE = ('Segoe UI', 11)
FONT_BODY = ('Segoe UI', 10)
FONT_MONO = ('Consolas', 9)
FONT_BUTTON = ('Segoe UI', 10, 'bold')


def _find_requisites_file():
    """Auto-detect requisites file in project root."""
    root = Path(__file__).resolve().parent
    patterns = ['реквізити*', 'requisites*', 'реквизиты*']
    for pattern in patterns:
        matches = list(root.glob(f'{pattern}.*'))
        for m in matches:
            if m.suffix.lower() in ('.xlsx', '.xls', '.csv'):
                return m
    return None


class ModernButton(tk.Button):
    """Custom button with hover effects."""
    def __init__(self, master, style='primary', **kwargs):
        self.style_type = style
        colors = {
            'primary': (PRIMARY, PRIMARY_HOVER, 'white'),
            'secondary': ('#6c757d', '#5a6268', 'white'),
            'success': (ACCENT, '#00664d', 'white'),
            'danger': (DANGER, '#a02d22', 'white'),
            'outline': ('white', PRIMARY_LIGHT, PRIMARY),
        }
        bg, hover_bg, fg = colors.get(style, colors['primary'])
        super().__init__(
            master,
            bg=bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            font=FONT_BUTTON,
            relief='flat',
            bd=0,
            padx=18,
            pady=8,
            cursor='hand2',
            **kwargs
        )
        self.default_bg = bg
        self.hover_bg = hover_bg
        self.bind('<Enter>', lambda e: self.config(bg=hover_bg))
        self.bind('<Leave>', lambda e: self.config(bg=bg))


class Card(tk.Frame):
    """Card container with subtle shadow."""
    def __init__(self, master, **kwargs):
        super().__init__(master, bg=CARD_BG, **kwargs)
        self.configure(highlightbackground=BORDER, highlightthickness=1)


class ReportApp:

    def __init__(self, root):
        self.root = root
        self.root.title('Генератор маркетингових звітів та актів')
        self.root.geometry('900x600')
        self.root.minsize(800, 550)
        self.root.configure(bg=BG_COLOR)

        # Data
        self.df = None
        self.df_mhl = None
        self.df_req = None
        self.loaded_period = None  # (year, month)
        self.loading = False
        self.requisites_file_var = tk.StringVar(value='')
        self.pharmacy_name_var = tk.StringVar(value='')
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        self.month_var = tk.StringVar(value=f'{datetime.now().month:02d}')

        # Auto-detect requisites on startup
        self._auto_detect_requisites()

        self._build_ui()

    def _auto_detect_requisites(self):
        """Try to find requisites file automatically on startup."""
        found = _find_requisites_file()
        if found:
            try:
                self.df_req = load_requisites_file(found)
                self.requisites_file_var.set(found.name)
                self._update_pharmacy_combo()
            except Exception:
                pass  # Silently fail, user can select manually

    def _build_ui(self):
        # Root container
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = tk.Frame(self.root, bg=BG_COLOR)
        main.grid(row=0, column=0, sticky='nsew', padx=24, pady=24)
        main.columnconfigure(0, weight=1)

        # Header
        header = tk.Frame(main, bg=BG_COLOR)
        header.grid(row=0, column=0, sticky='ew', pady=(0, 24))
        header.columnconfigure(1, weight=1)

        tk.Label(
            header,
            text='Генератор звітів та актів',
            font=FONT_TITLE,
            bg=BG_COLOR,
            fg=TEXT_MAIN
        ).grid(row=0, column=0, sticky='w')

        # Requisites card
        req_card = Card(main, padx=20, pady=16)
        req_card.grid(row=1, column=0, sticky='ew', pady=(0, 16))
        req_card.columnconfigure(1, weight=1)

        tk.Label(
            req_card,
            text='Файл реквізитів',
            font=FONT_SUBTITLE,
            bg=CARD_BG,
            fg=TEXT_MAIN
        ).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 8))

        self.req_entry = tk.Entry(
            req_card,
            textvariable=self.requisites_file_var,
            font=FONT_BODY,
            state='readonly',
            readonlybackground=CARD_BG,
            relief='solid',
            bd=1,
            highlightthickness=1,
            highlightcolor=PRIMARY,
            highlightbackground=BORDER
        )
        self.req_entry.grid(row=1, column=0, columnspan=2, sticky='ew', padx=(0, 10))

        ModernButton(
            req_card, text='Обрати файл...', style='outline',
            command=self.select_requisites_file
        ).grid(row=1, column=2, sticky='e')

        # Pharmacy selection card
        pharm_card = Card(main, padx=20, pady=16)
        pharm_card.grid(row=2, column=0, sticky='ew', pady=(0, 16))
        pharm_card.columnconfigure(1, weight=1)

        tk.Label(
            pharm_card,
            text='Мережа / Клієнт',
            font=FONT_SUBTITLE,
            bg=CARD_BG,
            fg=TEXT_MAIN
        ).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 8))

        self.pharmacy_combo = ttk.Combobox(
            pharm_card,
            textvariable=self.pharmacy_name_var,
            state='readonly',
            font=FONT_BODY,
            width=50
        )
        self.pharmacy_combo.grid(row=1, column=0, sticky='ew')
        self.pharmacy_combo.bind('<<ComboboxSelected>>', lambda e: self._on_pharmacy_change())

        # Period selection card (Month/Year only)
        period_card = Card(main, padx=20, pady=16)
        period_card.grid(row=3, column=0, sticky='ew', pady=(0, 24))
        period_card.columnconfigure(1, weight=1)

        tk.Label(
            period_card,
            text='Звітний період',
            font=FONT_SUBTITLE,
            bg=CARD_BG,
            fg=TEXT_MAIN
        ).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 8))

        # Year dropdown
        years = [str(y) for y in range(2024, datetime.now().year + 2)]
        tk.Label(period_card, text='Рік', font=FONT_BODY, bg=CARD_BG, fg=TEXT_MUTED).grid(row=1, column=0, sticky='w')
        year_combo = ttk.Combobox(
            period_card, textvariable=self.year_var, values=years,
            state='readonly', font=FONT_BODY, width=10
        )
        year_combo.grid(row=2, column=0, sticky='w', padx=(0, 16), pady=(0, 8))

        # Month dropdown
        months = [f'{m:02d}' for m in range(1, 13)]
        month_names = ['Січень', 'Лютий', 'Березень', 'Квітень', 'Травень', 'Червень',
                       'Липень', 'Серпень', 'Вересень', 'Жовтень', 'Листопад', 'Грудень']
        tk.Label(period_card, text='Місяць', font=FONT_BODY, bg=CARD_BG, fg=TEXT_MUTED).grid(row=1, column=1, sticky='w')
        self.month_combo = ttk.Combobox(
            period_card, textvariable=self.month_var,
            values=[f'{m:02d} — {n}' for m, n in zip(range(1, 13), month_names)],
            state='readonly', font=FONT_BODY, width=22
        )
        self.month_combo.grid(row=2, column=1, sticky='w', pady=(0, 8))

        # Action buttons
        actions = tk.Frame(main, bg=BG_COLOR)
        actions.grid(row=4, column=0, sticky='ew', pady=(8, 16))
        actions.columnconfigure(0, weight=1)

        ModernButton(
            actions, text='🔄 Завантажити дані', style='primary',
            command=self.load_data
        ).pack(side='left', padx=(0, 10))

        ModernButton(
            actions, text='📥 Згенерувати Excel', style='success',
            command=self.generate_report
        ).pack(side='left', padx=(0, 10))

        ModernButton(
            actions, text='⚙️ SKU', style='secondary',
            command=self.manage_sku_active
        ).pack(side='left', padx=(0, 10))

        ModernButton(
            actions, text='🏷️ Класифікувати SKU', style='secondary',
            command=self.classify_unknown_skus
        ).pack(side='left')

        # Status bar
        self.status_var = tk.StringVar(value='Готово до роботи')
        status_bar = tk.Frame(main, bg=CARD_BG, highlightbackground=BORDER, highlightthickness=1)
        status_bar.grid(row=5, column=0, sticky='ew', pady=(16, 0))
        status_bar.columnconfigure(0, weight=1)

        self.status_label = tk.Label(
            status_bar,
            textvariable=self.status_var,
            font=FONT_MONO,
            bg=CARD_BG,
            fg=TEXT_MUTED,
            anchor='w',
            padx=12,
            pady=8
        )
        self.status_label.grid(row=0, column=0, sticky='ew')

        # Progress indicator (hidden by default)
        self.progress = ttk.Progressbar(
            status_bar, mode='indeterminate', length=200
        )
        self.progress.grid(row=0, column=1, padx=12, pady=8)

    def _on_pharmacy_change(self):
        """Called when pharmacy selection changes."""
        pass  # Can add validation here if needed

    def _get_period_range(self):
        """Get period start/end as YYYY-MM-DD for the selected month/year."""
        year = int(self.year_var.get())
        month = int(self.month_var.get()[:2])
        start = f'{year}-{month:02d}-01'
        # Last day of month
        if month == 12:
            end = f'{year + 1}-01-01'
        else:
            end = f'{year}-{month + 1:02d}-01'
        from datetime import timedelta
        end_date = datetime.strptime(end, '%Y-%m-%d') - timedelta(days=1)
        end = end_date.strftime('%Y-%m-%d')
        return start, end

    def _get_period_display(self):
        """Get human-readable period string."""
        year = self.year_var.get()
        month_idx = int(self.month_var.get()[:2]) - 1
        month_names = ['січня', 'лютого', 'березня', 'квітня', 'травня', 'червня',
                       'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня']
        return f'{month_names[month_idx]} {year}'

    def _selected_requisites(self):
        if self.df_req is None or self.df_req.empty:
            raise ValueError('Будь ласка, оберіть файл реквізитів!')
        return self.df_req

    def _query_period(self):
        return self._get_period_range()
        start_date = datetime.strptime(start, '%d.%m.%Y').strftime('%Y-%m-%d')
        end_date = datetime.strptime(end, '%d.%m.%Y').strftime('%Y-%m-%d')
        return start_date, end_date

    def load_data(self):
        try:
            period = self._query_period()
        except ValueError as error:
            messagebox.showwarning('Попередження', str(error))
            return

        if self.loading:
            return

        self._load_data_async(period)

    def _load_data_async(self, period):
        self.loading = True
        self.load_button.configure(state='disabled')
        self.status_var.set(
            f'Завантаження даних з БД за період {period[0]} — {period[1]}...'
        )

        def load():
            data = None
            mhl_data = None
            data_error = None
            mhl_error = None
            try:
                data = get_pharmacy_data(*period)
            except (OSError, TypeError, ValueError) as load_error:
                data_error = load_error
            try:
                # Load all MHL data (filtering will be done in Python for flexibility)
                mhl_data = get_mhl_data(*period)
            except (OSError, TypeError, ValueError) as load_error:
                mhl_error = load_error
            self.root.after(
                0,
                lambda: self._data_loaded(
                    data, mhl_data, period, data_error, mhl_error
                ),
            )

        threading.Thread(target=load, daemon=True).start()

    def _data_loaded(self, data, mhl_data, period, data_error=None, mhl_error=None):
        self.loading = False
        self.load_button.configure(state='normal')
        self.df = data
        self.df_mhl = mhl_data

        has_pharmacy = data is not None and not data.empty
        has_mhl = mhl_data is not None and not mhl_data.empty

        if data_error is not None:
            self.loaded_period = None
            self.status_var.set(f'Помилка завантаження звіту з БД: {data_error}')
        elif not has_pharmacy and not has_mhl:
            self.loaded_period = None
            self.status_var.set('За вказаний період дані відсутні (жодного джерела).')
        else:
            # Set loaded_period if EITHER pharmacy OR MHL data exists
            self.loaded_period = period
            lik_count, dd_count, unknown_count = mhl_brand_counts(mhl_data)
            status = f'Звіт: {len(data) if has_pharmacy else 0} рядків'
            if mhl_error is not None:
                status += f'; помилка MHL: {mhl_error}'
            else:
                unknown_samples = mhl_unknown_samples(mhl_data)
                status += (
                    f'; MHL SQL: {len(mhl_data) if has_mhl else 0} рядків; '
                    f'ЛІК: {lik_count}; ДД: {dd_count}; '
                    f'невідомі: {unknown_count}'
                )
                if unknown_samples:
                    status += f'; приклади: {" | ".join(unknown_samples)}'
            self.status_var.set(status)

    def _update_pharmacy_combo(self):
        """Заповнює комбобокс мережами з файлу реквізитів"""
        if self.df_req is not None and not self.df_req.empty:
            target_col = None
            for col_candidate in [
                'chain_name',
                'Клиент из базы данных',
                'Мережа',
                'Назва',
            ]:
                if col_candidate in self.df_req.columns:
                    target_col = col_candidate
                    break

            if target_col:
                chains = (
                    self.df_req[target_col].dropna().astype(str).str.strip()
                )
                unique_chains = sorted(
                    c for c in chains.unique() if c and c.lower() != 'nan'
                )
                # Only show specific networks, no "Усі аптеки" option
                self.pharmacy_combo['values'] = tuple(unique_chains)
                if unique_chains:
                    self.pharmacy_combo.current(0)
                self.status_var.set(
                    f'Завантажено мереж з реквізитів: {len(unique_chains)}'
                )

    def _filtered_pharmacy_data(self):
        """Filter merchandising data by selected network (always a specific network).
        Returns empty DataFrame (not error) if no pharmacy data — caller handles MHL fallback."""
        if self.df is None or self.df.empty:
            # Return empty DataFrame with expected columns instead of raising
            return pd.DataFrame(columns=['CHAIN_NAME', 'PHARMACY_NAME', 'BRAND', 'MERCH', 'Price'])

        selected_name = self.pharmacy_name_var.get().strip()
        if not selected_name:
            raise ValueError('Будь ласка, оберіть мережу.')

        # Перевіряємо наявність колонки CHAIN_NAME в датасеті з бази даних
        if 'CHAIN_NAME' in self.df.columns:
            filtered = self.df[
                self.df['CHAIN_NAME'].fillna('').astype(str).str.strip().str.lower()
                == selected_name.lower()
            ]
        else:
            # Запасний варіант, якщо колонка називається інакше
            filtered = self.df[
                self.df['PHARMACY_NAME'].fillna('').astype(str).str.lower().str.contains(selected_name.lower())
            ]

        if filtered.empty:
            return self.df.iloc[0:0]

        return filtered

    def _filtered_mhl_data(self):
        """Filter MHL data by selected network (always a specific network)."""
        if self.df_mhl is None or self.df_mhl.empty:
            return self.df_mhl

        selected_name = self.pharmacy_name_var.get().strip()
        if not selected_name:
            return self.df_mhl

        # MHL data uses CHAIN_NAME column
        if 'CHAIN_NAME' in self.df_mhl.columns:
            filtered = self.df_mhl[
                self.df_mhl['CHAIN_NAME'].fillna('').astype(str).str.strip().str.lower()
                == selected_name.lower()
            ]
        else:
            # Fallback - return empty if no chain column
            return self.df_mhl.iloc[0:0]

        if filtered.empty:
            return self.df_mhl.iloc[0:0]

        return filtered
    def select_requisites_file(self):
        filename = filedialog.askopenfilename(
            title='Виберіть файл реквізитів',
            initialdir=str(Path.cwd()),
            filetypes=[
                ('Усі таблиці', '*.csv *.xlsx *.xls'),
                ('CSV файли', '*.csv'),
                ('Excel файли', '*.xlsx *.xls'),
            ],
        )
        if not filename:
            return
        try:
            requisites = load_requisites_file(filename)
            if requisites.empty:
                raise ValueError('Файл реквізитів порожній.')
        except (OSError, ValueError, UnicodeError) as error:
            messagebox.showerror('Помилка', str(error))
            return

        self.df_req = requisites
        self.requisites_file_var.set(filename)
        self._update_pharmacy_combo()

    def _validate_period(self):
        try:
            start = datetime.strptime(self.start_var.get().strip(), '%d.%m.%Y')
            end = datetime.strptime(self.end_var.get().strip(), '%d.%m.%Y')
        except ValueError as error:
            raise ValueError('Невірний формат дати. Використовуйте ДД.ММ.РРРР.') from error
        if start > end:
            raise ValueError('Дата початку не може бути більшою за дату кінця.')
        return self.start_var.get().strip(), self.end_var.get().strip()

    def _selected_requisites(self):
        if self.df_req is None or self.df_req.empty:
            raise ValueError('Будь ласка, оберіть файл реквізитів!')
        return self.df_req

    def _get_preview_data(self):
        self._selected_requisites()
        period = self._query_period()
        if self.loaded_period != period:
            raise ValueError(
                'Спочатку завантажте дані для вказаного періоду.'
            )

        # Get filtered pharmacy data for selected chain
        pharmacy_df = self._filtered_pharmacy_data()
        # Use FULL MHL data (not filtered by chain) - chain names may differ between requisites and MHL data
        mhl_df = self.df_mhl

        # Allow if either pharmacy data OR MHL data exists
        if (pharmacy_df is None or pharmacy_df.empty) and (mhl_df is None or mhl_df.empty):
            raise ValueError('Немає даних для формування звіту за цією мережею.')

        # Build preview from pharmacy data (if any)
        report = build_report_dataframe(pharmacy_df)
        # If pharmacy data empty but MHL exists, create minimal preview
        if (report is None or report.empty) and mhl_df is not None and not mhl_df.empty:
            report = pd.DataFrame({'Повідомлення': [f'Дані викладки відсутні. MHL даних: {len(mhl_df)} рядків (усі мережі)']})

        return report

    def preview(self):
        try:
            report = self._get_preview_data()
        except (ValueError, KeyError) as error:
            messagebox.showwarning('Попередження', str(error))
            return

        self.tree.delete(*self.tree.get_children())
        columns = [str(column) for column in report.columns]
        self.tree['columns'] = columns
        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=150, minwidth=80, anchor='center')

        for row in report.head(200).itertuples(index=False, name=None):
            self.tree.insert(
                '', 'end', values=tuple(str(value) for value in row)
            )

        self.status_var.set(
            f'Відображено рядків: {min(len(report), 200)} з {len(report)}'
        )

    def download(self):
        try:
            self._get_preview_data()
        except (ValueError, KeyError) as error:
            messagebox.showwarning('Попередження', str(error))
            return

        filename = filedialog.asksaveasfilename(
            title='Зберегти звіт як',
            initialdir=str(Path.cwd()),
            initialfile='marketing_report.xlsx',
            defaultextension='.xlsx',
            filetypes=[('Excel файли', '*.xlsx')],
        )
        if not filename:
            return

        try:
            start, end = self._validate_period()
            generate_excel_report(
                self._filtered_pharmacy_data(),
                self._selected_requisites(),
                filename,
                period_start=start,
                period_end=end,
                selected_pharmacy=self.pharmacy_name_var.get(),
                mhl_df=self.df_mhl,  # Use full MHL data (not filtered by chain)
            )
        except (OSError, ValueError, KeyError) as error:
            messagebox.showerror('Помилка', str(error))
            return

        self.status_var.set(f'Звіт збережено: {filename}')
        messagebox.showinfo('Успіх', f'Файл успішно збережено:\n{filename}')

    def classify_unknown_skus(self):
        """Open dialog to classify unknown SKUs from MHL data."""
        if self.df_mhl is None or self.df_mhl.empty:
            messagebox.showwarning('Попередження', 'Спочатку завантажте дані MHL.')
            return

        # Get filtered MHL data for current network selection
        try:
            mhl_data = self._filtered_mhl_data()
        except Exception as e:
            messagebox.showerror('Помилка', f'Помилка фільтрації MHL даних: {e}')
            return

        if mhl_data is None or mhl_data.empty:
            messagebox.showinfo('Інформація', 'Немає даних MHL для вибраної мережі.')
            return

        unknown_skus = get_unknown_skus(mhl_data)
        if not unknown_skus:
            messagebox.showinfo('Інформація', 'Невідомих SKU не знайдено. Усі препарати уже класифіковані.')
            return

        # Create classification dialog
        self._show_classification_dialog(unknown_skus, mhl_data)

    def _show_classification_dialog(self, unknown_skus, mhl_data):
        """Show dialog to classify unknown SKUs."""
        dialog = tk.Toplevel(self.root)
        dialog.title('Класифікація невідомих SKU')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.minsize(600, 400)  # Prevent squishing on Windows

        # Center dialog - use after() to ensure proper sizing
        def center_dialog():
            dialog.update_idletasks()
            w, h = 700, 500
            x = (dialog.winfo_screenwidth() // 2) - (w // 2)
            y = (dialog.winfo_screenheight() // 2) - (h // 2)
            dialog.geometry(f'{w}x{h}+{x}+{y}')

        dialog.after(1, center_dialog)

        # === MAIN LAYOUT using grid ===
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)  # Tree area expands

        # Row 0: Instructions
        ttk.Label(dialog, text='Оберіть категорію для кожного невідомого препарата (дефолт: ЛІК):',
                  font=('TkDefaultFont', 10, 'bold')).grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 5))

        # Row 1: Tree + Scrollbars in container frame (expands)
        tree_container = ttk.Frame(dialog)
        tree_container.grid(row=1, column=0, sticky='nsew', padx=10, pady=5)
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        columns = ['SKU_UA', 'BRAND', 'Категорія']
        tree = ttk.Treeview(tree_container, columns=columns, show='headings', height=15)
        tree.grid(row=0, column=0, sticky='nsew')

        tree.heading('SKU_UA', text='Назва препарату (SKU)')
        tree.heading('BRAND', text='Бренд')
        tree.heading('Категорія', text='Категорія')

        tree.column('SKU_UA', width=300, minwidth=200)
        tree.column('BRAND', width=150, minwidth=100)
        tree.column('Категорія', width=100, minwidth=80, anchor='center')

        # Scrollbars INSIDE tree_container
        y_scroll = ttk.Scrollbar(tree_container, orient='vertical', command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky='ns')
        tree.configure(yscrollcommand=y_scroll.set)

        # Populate with unknown SKUs — default category = ЛІК
        sku_to_brand = {}
        for sku in unknown_skus:
            # Get brand from MHL data
            brand_rows = mhl_data[mhl_data['SKU_UA'] == sku]
            brand = brand_rows['BRAND'].iloc[0] if not brand_rows.empty else 'Невідомий'
            sku_to_brand[sku] = brand
            tree.insert('', 'end', values=(sku, brand, 'ЛІК'), tags=(sku,))

        # --- Inline combobox editor for category column ---
        def on_tree_click(event):
            """Open combobox on single click in Category column."""
            region = tree.identify('region', event.x, event.y)
            if region != 'cell':
                return
            column = tree.identify_column(event.x)
            if column != '#3':  # Category column
                return
            item = tree.identify_row(event.y)
            if not item:
                return
            _open_category_combo(tree, item)

        def _open_category_combo(tree, item):
            """Create and show combobox over the category cell."""
            bbox = tree.bbox(item, column='#3')
            if not bbox:
                return
            x, y, width, height = bbox
            current_category = tree.item(item, 'values')[2]

            combo = ttk.Combobox(tree, values=['ЛІК', 'ДД'], state='readonly', width=8)
            combo.place(x=x, y=y, width=width, height=height)
            combo.set(current_category if current_category in ('ЛІК', 'ДД') else 'ЛІК')
            combo.focus_set()
            # Auto-open dropdown
            combo.event_generate('<Button-1>')

            def on_select(event=None):
                new_category = combo.get()
                if new_category in ('ЛІК', 'ДД'):
                    values = list(tree.item(item, 'values'))
                    values[2] = new_category
                    tree.item(item, values=values)
                combo.destroy()

            def on_focus_out(event):
                # Small delay to allow combobox selection to register
                dialog.after(100, combo.destroy)

            combo.bind('<<ComboboxSelected>>', on_select)
            combo.bind('<FocusOut>', on_focus_out)
            combo.bind('<Return>', on_select)
            combo.bind('<Escape>', lambda e: combo.destroy())

        # Single click to edit category
        tree.bind('<Button-1>', on_tree_click)

        # Row 2: Buttons frame (fixed at bottom)
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, sticky='ew', padx=10, pady=10)

        def save_all():
            saved_count = 0
            for item in tree.get_children():
                values = tree.item(item, 'values')
                sku = values[0]
                brand = values[1]
                category = values[2]
                if category and category in ('ЛІК', 'ДД'):
                    if save_sku_classification(sku, brand, category):
                        saved_count += 1
            if saved_count > 0:
                messagebox.showinfo('Успіх', f'Збережено класифікацій: {saved_count}')
                dialog.destroy()
            else:
                messagebox.showwarning('Попередження', 'Не обрано категорію для жодного препарату.')

        ttk.Button(btn_frame, text='Зберегти всі', command=save_all).pack(side='right', padx=5)
        ttk.Button(btn_frame, text='Скасувати', command=dialog.destroy).pack(side='right', padx=5)

        # Focus tree for keyboard navigation
        tree.focus_set()

    def manage_sku_active(self):
        """Open dialog to manage SKU active status (include in calculations or not)."""
        dialog = tk.Toplevel(self.root)
        dialog.title('Управління активністю SKU')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.minsize(800, 500)  # Prevent squishing on Windows

        # Center dialog - use after() to ensure proper sizing
        def center_dialog():
            dialog.update_idletasks()
            w, h = 900, 600
            x = (dialog.winfo_screenwidth() // 2) - (w // 2)
            y = (dialog.winfo_screenheight() // 2) - (h // 2)
            dialog.geometry(f'{w}x{h}+{x}+{y}')

        dialog.after(1, center_dialog)

        # Load all SKUs with status
        try:
            all_skus = get_all_skus_with_status()
        except Exception as e:
            messagebox.showerror('Помилка', f'Не вдалося завантажити SKU: {e}')
            dialog.destroy()
            return

        if all_skus.empty:
            messagebox.showinfo('Інформація', 'Список SKU порожній.')
            dialog.destroy()
            return

        # === MAIN LAYOUT using grid for predictable positioning ===
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(2, weight=1)  # Tree area expands

        # Row 0: Instructions
        ttk.Label(dialog, text='Керуйте активністю SKU — неактивні SKU не беруть участь у розрахунках MHL та звітах.\n'
                  'Подвійний клік по рядку / Пробел / кнопка «Переключити» — змінити статус.',
                  font=('TkDefaultFont', 10, 'bold')).grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 5))

        # Row 1: Search/Filter frame
        search_frame = ttk.Frame(dialog)
        search_frame.grid(row=1, column=0, sticky='ew', padx=10, pady=5)
        search_frame.columnconfigure(1, weight=1)  # Search entry expands

        ttk.Label(search_frame, text='Пошук:').grid(row=0, column=0, padx=(0, 5), sticky='w')
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=40)
        search_entry.grid(row=0, column=1, padx=(0, 10), sticky='ew')

        ttk.Label(search_frame, text='Категорія:').grid(row=0, column=2, padx=(10, 5), sticky='w')
        cat_var = tk.StringVar(value='Всі')
        cat_combo = ttk.Combobox(search_frame, textvariable=cat_var, values=['Всі', 'ЛІК', 'ДД'], state='readonly', width=10)
        cat_combo.grid(row=0, column=3, padx=(0, 10), sticky='w')

        ttk.Label(search_frame, text='Статус:').grid(row=0, column=4, padx=(10, 5), sticky='w')
        status_var = tk.StringVar(value='Всі')
        status_combo = ttk.Combobox(search_frame, textvariable=status_var, values=['Всі', 'Активні', 'Неактивні'], state='readonly', width=12)
        status_combo.grid(row=0, column=5, padx=(0, 10), sticky='w')

        # Row 2: Tree + Scrollbars in a container frame (expands)
        tree_container = ttk.Frame(dialog)
        tree_container.grid(row=2, column=0, sticky='nsew', padx=10, pady=5)
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        columns = ['SKU_UA', 'BRAND', 'CATEGORY', 'ACTIVE']
        tree = ttk.Treeview(tree_container, columns=columns, show='headings', height=20, selectmode='browse')
        tree.grid(row=0, column=0, sticky='nsew')

        tree.heading('SKU_UA', text='Назва препарату (SKU)')
        tree.heading('BRAND', text='Бренд')
        tree.heading('CATEGORY', text='Категорія')
        tree.heading('ACTIVE', text='Активний')

        tree.column('SKU_UA', width=350, minwidth=250)
        tree.column('BRAND', width=180, minwidth=120)
        tree.column('CATEGORY', width=80, minwidth=60, anchor='center')
        tree.column('ACTIVE', width=80, minwidth=60, anchor='center')

        # Scrollbars INSIDE tree_container (not on dialog)
        y_scroll = ttk.Scrollbar(tree_container, orient='vertical', command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky='ns')
        x_scroll = ttk.Scrollbar(tree_container, orient='horizontal', command=tree.xview)
        x_scroll.grid(row=1, column=0, sticky='ew')
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        # Row 3: Buttons frame (fixed at bottom, never pushed out)
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=3, column=0, sticky='ew', padx=10, pady=10)

        # Store reference to data for filtering
        self._sku_all_data = all_skus

        def populate_tree(filter_text='', filter_cat='Всі', filter_status='Всі'):
            """Populate tree with filtered data."""
            tree.delete(*tree.get_children())
            filter_text = filter_text.strip().casefold()

            for _, row in all_skus.iterrows():
                sku = str(row['SKU_UA'])
                brand = str(row['BRAND'])
                cat = str(row['CATEGORY'])
                active = bool(row['active'])

                # Apply filters
                if filter_text and filter_text not in sku.casefold() and filter_text not in brand.casefold():
                    continue
                if filter_cat != 'Всі' and cat != filter_cat:
                    continue
                if filter_status == 'Активні' and not active:
                    continue
                if filter_status == 'Неактивні' and active:
                    continue

                active_str = '✓' if active else '✗'
                tree.insert('', 'end', values=(sku, brand, cat, active_str), tags=(sku,))

        # Initial populate
        populate_tree()

        # Filter functions
        def on_search_change(*_):
            populate_tree(search_var.get(), cat_var.get(), status_var.get())

        def on_cat_change(*_):
            populate_tree(search_var.get(), cat_var.get(), status_var.get())

        def on_status_change(*_):
            populate_tree(search_var.get(), cat_var.get(), status_var.get())

        search_var.trace_add('write', on_search_change)
        cat_var.trace_add('write', on_cat_change)
        status_var.trace_add('write', on_status_change)

        # --- Toggle active status ---
        def toggle_selected_sku():
            """Toggle active status for currently selected SKU."""
            selection = tree.selection()
            if not selection:
                return
            item = selection[0]
            sku = tree.item(item, 'values')[0]
            _do_toggle(sku)

        def _do_toggle(sku):
            row = all_skus[all_skus['SKU_UA'] == sku]
            if row.empty:
                return
            current_active = bool(row.iloc[0]['active'])
            new_active = not current_active

            if set_sku_active(sku, new_active):
                idx = all_skus.index[all_skus['SKU_UA'] == sku]
                if len(idx) > 0:
                    all_skus.loc[idx[0], 'active'] = new_active
                # Re-populate preserving selection position
                populate_tree(search_var.get(), cat_var.get(), status_var.get())
                self.status_var.set(f'Оновлено статус SKU: {sku} → {"Активний" if new_active else "Неактивний"}')
            else:
                messagebox.showerror('Помилка', f'Не вдалося оновити статус для {sku}')

        # Double-click on row to toggle
        def on_double_click(event):
            toggle_selected_sku()

        # Space key to toggle focused/selected row
        def on_space(event):
            toggle_selected_sku()
            return 'break'  # prevent default space behavior (scroll)

        tree.bind('<Double-1>', on_double_click)
        tree.bind('<space>', on_space)
        # Also allow Enter key
        tree.bind('<Return>', on_space)

        # Button callbacks
        def select_all_active():
            """Set all visible SKUs to active."""
            visible_skus = [tree.item(item, 'values')[0] for item in tree.get_children()]
            count = 0
            for sku in visible_skus:
                if set_sku_active(sku, True):
                    idx = all_skus.index[all_skus['SKU_UA'] == sku]
                    if len(idx) > 0:
                        all_skus.loc[idx[0], 'active'] = True
                    count += 1
            populate_tree(search_var.get(), cat_var.get(), status_var.get())
            self.status_var.set(f'Активовано SKU: {count}')

        def select_all_inactive():
            """Set all visible SKUs to inactive."""
            visible_skus = [tree.item(item, 'values')[0] for item in tree.get_children()]
            count = 0
            for sku in visible_skus:
                if set_sku_active(sku, False):
                    idx = all_skus.index[all_skus['SKU_UA'] == sku]
                    if len(idx) > 0:
                        all_skus.loc[idx[0], 'active'] = False
                    count += 1
            populate_tree(search_var.get(), cat_var.get(), status_var.get())
            self.status_var.set(f'Деактивовано SKU: {count}')

        ttk.Button(btn_frame, text='Переключити виділене', command=toggle_selected_sku).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='Активувати видимі', command=select_all_active).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='Деактивувати видимі', command=select_all_inactive).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='Закрити', command=dialog.destroy).pack(side='right', padx=5)

        # Ensure tree gets focus AFTER dialog is mapped
        def set_focus():
            tree.focus_set()
            # Select first item if any
            children = tree.get_children()
            if children:
                tree.selection_set(children[0])
                tree.see(children[0])

        dialog.after(100, set_focus)


def main():
    root = tk.Tk()
    ReportApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()