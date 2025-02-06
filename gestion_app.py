import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook

class FinanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Financial Manager Pro")
        self.root.geometry("1200x700")
        self.root.configure(bg='#f0f0f0')
        
        # Configurar estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        # Configurar almacenamiento
        self.filename = "finanzas.xlsx"
        self.transactions = []
        
        # Variables de control
        self.selected_period = tk.StringVar(value="Día")
        self.selected_transaction = None
        
        # Crear interfaz
        self.create_widgets()
        self.load_transactions()
        self.update_table()
        
    def configure_styles(self):
        self.colors = {
            'primary': '#2a73ff',
            'success': '#28a745',
            'danger': '#dc3545',
            'light': '#f8f9fa',
            'dark': '#212529'
        }
        
        # Estilos generales
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', foreground='#212529')
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'), foreground='#2a73ff')
        
        # Botones
        self.style.configure('Primary.TButton', 
                           font=('Arial', 10, 'bold'),
                           borderwidth=0,
                           relief='flat',
                           background=self.colors['primary'],
                           foreground='white',
                           padding=10)
        
        self.style.configure('Secondary.TButton', 
                           font=('Arial', 10),
                           borderwidth=0,
                           relief='flat',
                           background=self.colors['light'],
                           foreground=self.colors['dark'],
                           padding=10)
        
        self.style.map('Primary.TButton',
            background=[('active', '#1a5ee6'), ('disabled', '#cccccc')],
            foreground=[('active', 'white'), ('disabled', '#666666')]
        )
        
        self.style.map('Secondary.TButton',
            background=[('active', '#e9ecef'), ('disabled', '#cccccc')],
            foreground=[('active', self.colors['dark']), ('disabled', '#666666')]
        )

        # ComboBox
        self.style.configure('TCombobox',
            selectbackground=self.colors['primary'],
            fieldbackground='white',
            background='white',
            bordercolor='#ced4da',
            darkcolor='#ffffff',
            lightcolor='#ffffff',
            arrowsize=12,
            padding=8,
            relief='flat'
        )

        self.style.map('TCombobox',
            fieldbackground=[('readonly', 'white')],
            background=[('readonly', 'white')],
            bordercolor=[('focus', self.colors['primary']), ('!focus', '#ced4da')],
            arrowcolor=[('!disabled', self.colors['dark']), ('disabled', '#868e96')]
        )
        
        # Tabla
        self.style.configure('Treeview.Heading', 
                           font=('Arial', 10, 'bold'), 
                           background='#e9ecef',
                           relief='flat',
                           borderwidth=0)
        
        self.style.configure('Treeview', 
                           rowheight=30,
                           background='white',
                           fieldbackground='white',
                           borderwidth=0)
        
        self.style.map('Treeview',
            background=[('selected', '#e2e5e9')]
        )
        
        # Entradas
        self.style.configure('TEntry',
                           bordercolor='#ced4da',
                           lightcolor='#ffffff',
                           darkcolor='#ffffff',
                           padding=8,
                           relief='flat')
        
        self.style.configure('TCombobox',
                           selectbackground=self.colors['primary'],
                           fieldbackground='white',
                           padding=8,
                           relief='flat')
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, text="Gestión Financiera", style='Header.TLabel').pack(side=tk.LEFT)
        
        # Controles
        control_frame = ttk.Frame(header_frame)
        control_frame.pack(side=tk.RIGHT)
        
        # Selector de período
        period_selector = ttk.Combobox(
            control_frame,
            textvariable=self.selected_period,
            values=["Día", "Semana", "Mes"],
            state='readonly',
            width=10,
            style='Custom.TCombobox'
        )
        period_selector.bind('<<ComboboxSelected>>', lambda e: self.update_table())
        period_selector.pack(side=tk.LEFT, padx=10)
        
        # Botones
        btn_frame = ttk.Frame(control_frame)
        btn_frame.pack(side=tk.LEFT)
        
        ttk.Button(btn_frame, text="Nuevo", style='Primary.TButton', command=self.open_add_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Editar", style='Secondary.TButton', command=self.open_edit_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Eliminar", style='Secondary.TButton', command=self.delete_transaction).pack(side=tk.LEFT, padx=5)
        
        # Tabla
        columns = ("Fecha", "Tipo", "Categoría", "Monto", "Descripción")
        self.tree = ttk.Treeview(
            main_frame,
            columns=columns,
            show="headings",
            selectmode="browse"
        )
        
        for col in columns:
            self.tree.heading(col, text=col, anchor=tk.CENTER)
            self.tree.column(col, width=120, anchor=tk.CENTER)
            
        self.tree.column("Descripción", width=300)
        self.tree.pack(fill=tk.BOTH, expand=True)

    def update_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        start, end = self.get_period_range()
        
        for transaction in self.transactions:
            trans_date = datetime.strptime(transaction["Fecha"], "%Y-%m-%d")
            if start <= trans_date < end:
                self.tree.insert("", tk.END, values=(
                    transaction["Fecha"],
                    transaction["Tipo"],
                    transaction["Categoría"],
                    f"${float(transaction['Monto']):.2f}",
                    transaction["Descripción"]
                ))
        
    def load_transactions(self):
        try:
            workbook = load_workbook(self.filename)
            sheet = workbook.active
            self.transactions = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row:
                    self.transactions.append({
                        "Fecha": row[0],
                        "Tipo": row[1],
                        "Categoría": row[2],
                        "Monto": row[3],
                        "Descripción": row[4]
                    })
        except FileNotFoundError:
            self.transactions = []
            
    def save_transactions(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Fecha", "Tipo", "Categoría", "Monto", "Descripción"])
        
        for transaction in self.transactions:
            sheet.append([
                transaction["Fecha"],
                transaction["Tipo"],
                transaction["Categoría"],
                transaction["Monto"],
                transaction["Descripción"]
            ])
        
        workbook.save(self.filename)
        self.update_table()
    
    def get_period_range(self):
        today = datetime.today()
        if self.selected_period.get() == "Día":
            start = today.replace(hour=0, minute=0, second=0, microsecond=0)
            end = start + timedelta(days=1)
        elif self.selected_period.get() == "Semana":
            start = today - timedelta(days=today.weekday())
            start = start.replace(hour=0, minute=0, second=0, microsecond=0)
            end = start + timedelta(weeks=1)
        else:  # Mes
            start = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            end = (start + timedelta(days=32)).replace(day=1)
        return start, end
    
    def open_add_window(self):
        self.transaction_window("Agregar Transacción")
        
    def open_edit_window(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Seleccione una transacción para editar")
            return
            
        self.selected_transaction = self.tree.item(selected_item)["values"]
        self.transaction_window("Editar Transacción", edit_mode=True)
    
    def transaction_window(self, title, edit_mode=False):
        window = Toplevel(self.root)
        window.title(title)
        window.geometry("600x500")
        window.configure(bg='#f0f0f0')
        
        main_frame = ttk.Frame(window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        labels = ["Fecha (AAAA-MM-DD):", "Tipo:", "Categoría:", "Monto:", "Descripción:"]
        default_values = [
            datetime.today().strftime("%Y-%m-%d"),
            "Ingreso",
            "Otros",
            "",
            ""
        ]
        
        if edit_mode:
            default_values = [
                self.selected_transaction[0],
                self.selected_transaction[1],
                self.selected_transaction[2],
                self.selected_transaction[3][1:],
                self.selected_transaction[4]
            ]
        
        fields = {}
        for i, label in enumerate(labels[:-1]):
            frame = ttk.Frame(main_frame)
            frame.pack(fill=tk.X, pady=8)
            
            ttk.Label(frame, text=label).pack(side=tk.LEFT, padx=(0, 10))
            
            if label == "Tipo:":
                field = ttk.Combobox(frame, values=["Ingreso", "Gasto"], width=24)
            elif label == "Categoría:":
                field = ttk.Combobox(frame, values=["Efectivo", "Banco", "Otros"], width=24)
            else:
                field = ttk.Entry(frame, width=30)
            
            field.pack(side=tk.RIGHT, fill=tk.X, expand=True)
            field.insert(0, default_values[i])
            fields[label] = field
        
        # Campo descripción
        desc_frame = ttk.Frame(main_frame)
        desc_frame.pack(fill=tk.BOTH, expand=True, pady=8)
        
        ttk.Label(desc_frame, text=labels[-1]).pack(anchor=tk.NW, pady=(0, 5))
        desc_field = tk.Text(desc_frame, height=5, wrap=tk.WORD, bd=1, relief='flat',
                           font=('Arial', 10), bg='white', padx=8, pady=8)
        desc_field.pack(fill=tk.BOTH, expand=True)
        desc_field.insert("1.0", default_values[-1])
        fields[labels[-1]] = desc_field
        
        # Botones
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=15)
        
        ttk.Button(
            btn_frame,
            text="Guardar",
            style='Primary.TButton',
            command=lambda: self.save_transaction(fields, window, edit_mode)
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="Cancelar",
            style='Secondary.TButton',
            command=window.destroy
        ).pack(side=tk.RIGHT, padx=5)
        
    def save_transaction(self, fields, window, edit_mode):
        try:
            new_transaction = {
                "Fecha": fields["Fecha (AAAA-MM-DD):"].get(),
                "Tipo": fields["Tipo:"].get(),
                "Categoría": fields["Categoría:"].get(),
                "Monto": float(fields["Monto:"].get()),
                "Descripción": fields["Descripción:"].get("1.0", tk.END).strip()
            }
            
            datetime.strptime(new_transaction["Fecha"], "%Y-%m-%d")
            
            if edit_mode:
                for i, t in enumerate(self.transactions):
                    if t["Fecha"] == self.selected_transaction[0] and t["Descripción"] == self.selected_transaction[4]:
                        del self.transactions[i]
                        break
            
            self.transactions.append(new_transaction)
            self.save_transactions()
            self.update_table()
            window.destroy()
            
        except ValueError as e:
            messagebox.showerror("Error", f"Dato inválido: {str(e)}")
    
    def delete_transaction(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Seleccione una transacción para eliminar")
            return
            
        if messagebox.askyesno("Confirmar", "¿Eliminar esta transacción?"):
            selected_values = self.tree.item(selected_item)["values"]
            for i, t in enumerate(self.transactions):
                if (t["Fecha"] == selected_values[0] and 
                    t["Descripción"] == selected_values[4] and
                    float(t["Monto"]) == float(selected_values[3][1:])):
                    del self.transactions[i]
                    self.save_transactions()
                    self.update_table()
                    break

if __name__ == "__main__":
    root = tk.Tk()
    app = FinanceApp(root)
    root.mainloop()