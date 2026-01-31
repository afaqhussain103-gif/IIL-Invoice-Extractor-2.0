"""
IIL Invoice Extractor v3.0 - Stable Edition
International Industries Limited
No external UI dependencies - Pure tkinter + PyMuPDF only
"""

import os
import shutil
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import re

class IILInvoiceExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("IIL Invoice Extractor v3.0 - Stable")
        self.root.geometry("850x780")
        self.root.resizable(False, False)
        
        # OFFICIAL IIL BRAND COLORS
        self.colors = {
            'green': '#01783f',        # IIL Primary Green
            'yellow': '#c2d501',       # IIL Primary Yellow-Green
            'accent': '#76690f',       # IIL Accent Olive
            'white': '#ffffff',
            'bg': '#f8f9fa',
            'text': '#2c3e50',
            'gray': '#6c757d'
        }
        
        self.root.configure(bg=self.colors['white'])
        self.setup_ui()
    
    def setup_ui(self):
        # HEADER
        header = tk.Frame(self.root, bg=self.colors['green'], height=110)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="INTERNATIONAL INDUSTRIES LIMITED",
            font=('Arial', 20, 'bold'),
            bg=self.colors['green'],
            fg=self.colors['white']
        ).pack(pady=(22, 2))
        
        tk.Label(
            header,
            text="Building Pakistan's Future | Invoice Management System",
            font=('Arial', 10),
            bg=self.colors['green'],
            fg=self.colors['yellow']
        ).pack(pady=2)
        
        version_frame = tk.Frame(header, bg=self.colors['yellow'], padx=12, pady=3)
        version_frame.place(relx=0.95, rely=0.15, anchor='ne')
        tk.Label(
            version_frame,
            text="v3.0",
            font=('Arial', 8, 'bold'),
            bg=self.colors['yellow'],
            fg=self.colors['text']
        ).pack()
        
        # MAIN CONTENT
        main = tk.Frame(self.root, bg=self.colors['white'])
        main.pack(fill=tk.BOTH, expand=True, padx=40, pady=25)
        
        # SOURCE FOLDER
        self.create_section_header(main, "📁 SOURCE FOLDER (Invoices)", 0)
        self.source_entry = self.create_folder_input(main, 1, self.browse_source)
        
        # DESTINATION FOLDER
        self.create_section_header(main, "📂 DESTINATION FOLDER (Output)", 2)
        self.dest_entry = self.create_folder_input(main, 3, self.browse_destination)
        
        self.create_separator(main, 4)
        
        # FILTER TYPE
        self.create_section_header(main, "🔍 FILTER TYPE", 5)
        
        filter_frame = tk.Frame(main, bg=self.colors['white'])
        filter_frame.grid(row=6, column=0, columnspan=3, sticky='w', pady=5)
        
        self.filter_var = tk.StringVar(value="name")
        
        tk.Radiobutton(
            filter_frame,
            text="Customer Name",
            variable=self.filter_var,
            value="name",
            font=('Arial', 10),
            bg=self.colors['white'],
            fg=self.colors['text'],
            selectcolor=self.colors['yellow'],
            activebackground=self.colors['white']
        ).pack(side=tk.LEFT, padx=(0, 30))
        
        tk.Radiobutton(
            filter_frame,
            text="Account ID",
            variable=self.filter_var,
            value="account",
            font=('Arial', 10),
            bg=self.colors['white'],
            fg=self.colors['text'],
            selectcolor=self.colors['yellow'],
            activebackground=self.colors['white']
        ).pack(side=tk.LEFT)
        
        # SEARCH VALUE
        self.create_section_header(main, "🔎 SEARCH VALUE", 7)
        
        self.search_entry = tk.Entry(
            main,
            font=('Arial', 11),
            bd=0,
            relief=tk.FLAT,
            highlightthickness=2,
            highlightbackground=self.colors['bg'],
            highlightcolor=self.colors['green']
        )
        self.search_entry.grid(row=8, column=0, columnspan=3, sticky='ew', pady=5, ipady=10)
        
        self.create_separator(main, 9)
        
        # DATE RANGE (Simple text entry)
        self.create_section_header(main, "📅 DATE RANGE FILTER (Optional)", 10)
        
        date_frame = tk.Frame(main, bg=self.colors['white'])
        date_frame.grid(row=11, column=0, columnspan=3, sticky='w', pady=8)
        
        tk.Label(
            date_frame,
            text="From:",
            font=('Arial', 10, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['text']
        ).pack(side=tk.LEFT, padx=(0, 8))
        
        self.from_date_entry = tk.Entry(
            date_frame,
            font=('Arial', 10),
            width=15,
            bd=1,
            relief=tk.SOLID
        )
        self.from_date_entry.pack(side=tk.LEFT, padx=(0, 35))
        self.from_date_entry.insert(0, "01/01/2024")
        
        tk.Label(
            date_frame,
            text="To:",
            font=('Arial', 10, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['text']
        ).pack(side=tk.LEFT, padx=(0, 8))
        
        self.to_date_entry = tk.Entry(
            date_frame,
            font=('Arial', 10),
            width=15,
            bd=1,
            relief=tk.SOLID
        )
        self.to_date_entry.pack(side=tk.LEFT)
        self.to_date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        tk.Label(
            main,
            text="💡 Format: DD/MM/YYYY (e.g., 31/01/2025)",
            font=('Arial', 8, 'italic'),
            bg=self.colors['white'],
            fg=self.colors['gray']
        ).grid(row=12, column=0, columnspan=3, sticky='w', pady=(0, 5))
        
        # OPTIONS
        options_frame = tk.Frame(main, bg=self.colors['white'])
        options_frame.grid(row=13, column=0, columnspan=3, sticky='w', pady=10)
        
        self.use_date_filter = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options_frame,
            text="✓ Enable Date Range Filter",
            variable=self.use_date_filter,
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green'],
            selectcolor=self.colors['yellow'],
            activebackground=self.colors['white']
        ).pack(anchor='w', pady=2)
        
        self.overwrite_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options_frame,
            text="⚠️  Overwrite existing files (uncheck to skip duplicates)",
            variable=self.overwrite_var,
            font=('Arial', 9),
            bg=self.colors['white'],
            fg=self.colors['accent'],
            selectcolor=self.colors['yellow'],
            activebackground=self.colors['white']
        ).pack(anchor='w', pady=2)
        
        self.create_separator(main, 14)
        
        # PROGRESS
        progress_container = tk.Frame(main, bg=self.colors['bg'], bd=0)
        progress_container.grid(row=15, column=0, columnspan=3, sticky='ew', pady=15, ipady=15, ipadx=20)
        
        self.progress_label = tk.Label(
            progress_container,
            text="Ready to extract invoices",
            font=('Arial', 11, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['text']
        )
        self.progress_label.pack(pady=(5, 8))
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "IIL.Horizontal.TProgressbar",
            troughcolor=self.colors['white'],
            background=self.colors['green'],
            bordercolor=self.colors['green'],
            lightcolor=self.colors['yellow'],
            darkcolor=self.colors['accent'],
            borderwidth=1,
            thickness=28
        )
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_container,
            variable=self.progress_var,
            maximum=100,
            style="IIL.Horizontal.TProgressbar",
            length=700
        )
        self.progress_bar.pack(pady=5)
        
        self.percent_label = tk.Label(
            progress_container,
            text="0%",
            font=('Arial', 13, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['green']
        )
        self.percent_label.pack(pady=(5, 5))
        
        # EXTRACT BUTTON
        extract_btn = tk.Button(
            main,
            text="🚀 START EXTRACTION",
            font=('Arial', 14, 'bold'),
            bg=self.colors['green'],
            fg=self.colors['white'],
            activebackground=self.colors['yellow'],
            activeforeground=self.colors['text'],
            cursor='hand2',
            bd=0,
            padx=50,
            pady=18,
            command=self.extract_invoices
        )
        extract_btn.grid(row=16, column=0, columnspan=3, pady=15)
        
        def on_enter(e):
            extract_btn.config(bg=self.colors['yellow'], fg=self.colors['text'])
        
        def on_leave(e):
            extract_btn.config(bg=self.colors['green'], fg=self.colors['white'])
        
        extract_btn.bind('<Enter>', on_enter)
        extract_btn.bind('<Leave>', on_leave)
        
        # FOOTER
        footer = tk.Frame(self.root, bg=self.colors['green'], height=35)
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        
        tk.Label(
            footer,
            text="© 2025 International Industries Limited | Powered by Advanced AI",
            font=('Arial', 8),
            bg=self.colors['green'],
            fg=self.colors['white']
        ).pack(pady=10)
        
        main.grid_columnconfigure(0, weight=1)
    
    def create_section_header(self, parent, text, row):
        label = tk.Label(
            parent,
            text=text,
            font=('Arial', 11, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green']
        )
        label.grid(row=row, column=0, sticky='w', pady=(12, 5))
    
    def create_separator(self, parent, row):
        sep = tk.Frame(parent, bg=self.colors['bg'], height=2)
        sep.grid(row=row, column=0, columnspan=3, sticky='ew', pady=15)
    
    def create_folder_input(self, parent, row, command):
        entry = tk.Entry(
            parent,
            font=('Arial', 10),
            bd=0,
            relief=tk.FLAT,
            highlightthickness=2,
            highlightbackground=self.colors['bg'],
            highlightcolor=self.colors['green']
        )
        entry.grid(row=row, column=0, columnspan=2, sticky='ew', pady=5, ipady=9)
        
        btn = tk.Button(
            parent,
            text="📁 Browse",
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green'],
            activebackground=self.colors['green'],
            activeforeground=self.colors['white'],
            cursor='hand2',
            bd=1,
            relief=tk.SOLID,
            highlightbackground=self.colors['green'],
            padx=22,
            pady=7,
            command=command
        )
        btn.grid(row=row, column=2, padx=(12, 0), pady=5)
        
        return entry
    
    def browse_source(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, folder)
    
    def browse_destination(self):
        folder = filedialog.askdirectory(title="Select Destination Folder")
        if folder:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, folder)
    
    def parse_date(self, date_str):
        """Parse DD/MM/YYYY format"""
        try:
            return datetime.strptime(date_str.strip(), "%d/%m/%Y")
        except:
            return None
    
    def extract_date_from_pdf(self, pdf_path):
        """Extract date from PDF"""
        try:
            doc = fitz.open(pdf_path)
            text = ""
            for page in doc:
                text += page.get_text()
            doc.close()
            
            patterns = [
                r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b',
                r'\b(\d{4}[/-]\d{1,2}[/-]\d{1,2})\b'
            ]
            
            for pattern in patterns:
                matches = re.findall(pattern, text)
                if matches:
                    date_str = matches[0]
                    for fmt in ['%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d', '%Y-%m-%d']:
                        try:
                            return datetime.strptime(date_str, fmt)
                        except:
                            continue
            return None
        except:
            return None
    
    def extract_invoices(self):
        source = self.source_entry.get().strip()
        destination = self.dest_entry.get().strip()
        search_value = self.search_entry.get().strip().lower()
        
        if not source or not destination or not search_value:
            messagebox.showwarning(
                "⚠️ Missing Information",
                "Please fill in all required fields!"
            )
            return
        
        if not os.path.exists(source):
            messagebox.showerror("❌ Error", "Source folder does not exist!")
            return
        
        if not os.path.exists(destination):
            try:
                os.makedirs(destination)
            except:
                messagebox.showerror("❌ Error", "Cannot create destination folder!")
                return
        
        filter_type = self.filter_var.get()
        use_date = self.use_date_filter.get()
        overwrite = self.overwrite_var.get()
        
        from_date = None
        to_date = None
        
        if use_date:
            from_date = self.parse_date(self.from_date_entry.get())
            to_date = self.parse_date(self.to_date_entry.get())
            
            if not from_date or not to_date:
                messagebox.showerror("❌ Date Error", "Invalid date format! Use DD/MM/YYYY")
                return
        
        pdf_files = [f for f in os.listdir(source) if f.lower().endswith('.pdf')]
        total_files = len(pdf_files)
        
        if total_files == 0:
            messagebox.showinfo("ℹ️ No Files", "No PDF files found!")
            return
        
        extracted = 0
        skipped_date = 0
        skipped_duplicate = 0
        errors = 0
        
        for idx, filename in enumerate(pdf_files, 1):
            pdf_path = os.path.join(source, filename)
            
            progress = (idx / total_files) * 100
            self.progress_var.set(progress)
            self.percent_label.config(text=f"{int(progress)}%")
            
            display_name = filename[:50] + "..." if len(filename) > 50 else filename
            self.progress_label.config(text=f"Processing: {display_name}")
            self.root.update_idletasks()
            
            try:
                doc = fitz.open(pdf_path)
                text = ""
                for page in doc:
                    text += page.get_text().lower()
                doc.close()
                
                match_found = search_value in text
                
                if match_found and use_date:
                    pdf_date = self.extract_date_from_pdf(pdf_path)
                    if pdf_date:
                        if not (from_date <= pdf_date <= to_date):
                            match_found = False
                            skipped_date += 1
                    else:
                        match_found = False
                        skipped_date += 1
                
                if match_found:
                    dest_path = os.path.join(destination, filename)
                    
                    if os.path.exists(dest_path) and not overwrite:
                        skipped_duplicate += 1
                    else:
                        shutil.copy2(pdf_path, dest_path)
                        extracted += 1
            
            except Exception as e:
                errors += 1
        
        self.progress_var.set(100)
        self.percent_label.config(text="100% ✓", fg=self.colors['green'])
        self.progress_label.config(text="✅ Extraction Complete!")
        
        result = "╔═════════════════════════════════╗\n"
        result += "║   EXTRACTION COMPLETE   ║\n"
        result += "╚═════════════════════════════════╝\n\n"
        result += f"📊 Total PDFs Scanned: {total_files}\n"
        result += f"✅ Invoices Extracted: {extracted}\n"
        
        if skipped_duplicate > 0:
            result += f"⏭️  Skipped (Duplicates): {skipped_duplicate}\n"
        
        if use_date and skipped_date > 0:
            result += f"📅 Filtered by Date: {skipped_date}\n"
        
        if errors > 0:
            result += f"⚠️  Errors: {errors}\n"
        
        result += f"\n📂 Location:\n{destination}"
        
        messagebox.showinfo("✅ Success", result)
        
        self.root.after(2000, lambda: [
            self.progress_var.set(0),
            self.percent_label.config(text="0%", fg=self.colors['green']),
            self.progress_label.config(text="Ready to extract invoices")
        ])

if __name__ == "__main__":
    root = tk.Tk()
    app = IILInvoiceExtractor(root)
    root.mainloop()
