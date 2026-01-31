"""
IIL Invoice Extractor v4.1 - Compact & Resizable
International Industries Limited
"""

import os
import shutil
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class IILInvoiceExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("IIL Invoice Extractor v4.1")
        self.root.geometry("700x550")  # Smaller size
        self.root.resizable(True, True)  # Enable resizing
        self.root.minsize(600, 500)  # Minimum size
        
        # IIL BRAND COLORS
        self.colors = {
            'green': '#01783f',
            'yellow': '#c2d501',
            'white': '#ffffff',
            'bg': '#f8f9fa',
            'text': '#2c3e50'
        }
        
        self.root.configure(bg=self.colors['white'])
        self.setup_ui()
    
    def setup_ui(self):
        # HEADER - Compact
        header = tk.Frame(self.root, bg=self.colors['green'], height=70)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="IIL INVOICE EXTRACTOR",
            font=('Arial', 14, 'bold'),
            bg=self.colors['green'],
            fg=self.colors['white']
        ).pack(pady=(12, 2))
        
        tk.Label(
            header,
            text="v4.1 - International Industries Limited",
            font=('Arial', 8),
            bg=self.colors['green'],
            fg=self.colors['yellow']
        ).pack()
        
        # MAIN CONTENT - Scrollable
        main_container = tk.Frame(self.root, bg=self.colors['white'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        # SOURCE FOLDER
        tk.Label(
            main_container,
            text="📁 Source Folder:",
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green']
        ).grid(row=0, column=0, sticky='w', pady=(0, 3))
        
        source_frame = tk.Frame(main_container, bg=self.colors['white'])
        source_frame.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        
        self.source_entry = tk.Entry(source_frame, font=('Arial', 9))
        self.source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        
        tk.Button(
            source_frame,
            text="Browse",
            font=('Arial', 8),
            bg=self.colors['green'],
            fg=self.colors['white'],
            command=self.browse_source,
            padx=12,
            pady=3
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # DESTINATION FOLDER
        tk.Label(
            main_container,
            text="📂 Destination Folder:",
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green']
        ).grid(row=2, column=0, sticky='w', pady=(0, 3))
        
        dest_frame = tk.Frame(main_container, bg=self.colors['white'])
        dest_frame.grid(row=3, column=0, sticky='ew', pady=(0, 10))
        
        self.dest_entry = tk.Entry(dest_frame, font=('Arial', 9))
        self.dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        
        tk.Button(
            dest_frame,
            text="Browse",
            font=('Arial', 8),
            bg=self.colors['green'],
            fg=self.colors['white'],
            command=self.browse_dest,
            padx=12,
            pady=3
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # SEPARATOR
        tk.Frame(main_container, bg=self.colors['bg'], height=1).grid(
            row=4, column=0, sticky='ew', pady=10
        )
        
        # SEARCH VALUE
        tk.Label(
            main_container,
            text="🔎 Search (Customer Name or Account ID):",
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green']
        ).grid(row=5, column=0, sticky='w', pady=(0, 3))
        
        self.search_entry = tk.Entry(main_container, font=('Arial', 10))
        self.search_entry.grid(row=6, column=0, sticky='ew', pady=(0, 10), ipady=7)
        
        # OPTIONS
        self.overwrite_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            main_container,
            text="⚠️  Overwrite existing files",
            variable=self.overwrite_var,
            font=('Arial', 8),
            bg=self.colors['white'],
            fg=self.colors['text'],
            selectcolor=self.colors['yellow']
        ).grid(row=7, column=0, sticky='w', pady=(0, 10))
        
        # SEPARATOR
        tk.Frame(main_container, bg=self.colors['bg'], height=1).grid(
            row=8, column=0, sticky='ew', pady=10
        )
        
        # PROGRESS - Compact
        progress_frame = tk.Frame(main_container, bg=self.colors['bg'])
        progress_frame.grid(row=9, column=0, sticky='ew', pady=(0, 10), ipady=10)
        
        self.progress_label = tk.Label(
            progress_frame,
            text="Ready to extract",
            font=('Arial', 9, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['text']
        )
        self.progress_label.pack(pady=(3, 5))
        
        style = ttk.Style()
        style.configure(
            "IIL.Horizontal.TProgressbar",
            background=self.colors['green'],
            thickness=20
        )
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            style="IIL.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, padx=10, pady=3)
        
        self.percent_label = tk.Label(
            progress_frame,
            text="0%",
            font=('Arial', 10, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['green']
        )
        self.percent_label.pack(pady=3)
        
        # EXTRACT BUTTON - Compact
        tk.Button(
            main_container,
            text="🚀 START EXTRACTION",
            font=('Arial', 11, 'bold'),
            bg=self.colors['green'],
            fg=self.colors['white'],
            command=self.extract,
            padx=30,
            pady=10,
            cursor='hand2'
        ).grid(row=10, column=0, pady=10)
        
        # Configure grid weights for resizing
        main_container.grid_columnconfigure(0, weight=1)
        main_container.grid_rowconfigure(9, weight=1)
    
    def browse_source(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, folder)
    
    def browse_dest(self):
        folder = filedialog.askdirectory(title="Select Destination Folder")
        if folder:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, folder)
    
    def extract(self):
        source = self.source_entry.get().strip()
        dest = self.dest_entry.get().strip()
        search = self.search_entry.get().strip().lower()
        
        if not source or not dest or not search:
            messagebox.showwarning("Missing Info", "Please fill all fields!")
            return
        
        if not os.path.exists(source):
            messagebox.showerror("Error", "Source folder doesn't exist!")
            return
        
        if not os.path.exists(dest):
            os.makedirs(dest)
        
        pdf_files = [f for f in os.listdir(source) if f.lower().endswith('.pdf')]
        total = len(pdf_files)
        
        if total == 0:
            messagebox.showinfo("No Files", "No PDFs found in source folder!")
            return
        
        extracted = 0
        skipped = 0
        errors = 0
        
        for idx, filename in enumerate(pdf_files, 1):
            pdf_path = os.path.join(source, filename)
            
            progress = (idx / total) * 100
            self.progress_var.set(progress)
            self.percent_label.config(text=f"{int(progress)}%")
            
            # Truncate long filenames
            display_name = filename[:30] + "..." if len(filename) > 30 else filename
            self.progress_label.config(text=f"Processing: {display_name}")
            self.root.update_idletasks()
            
            try:
                doc = fitz.open(pdf_path)
                text = ""
                for page in doc:
                    text += page.get_text().lower()
                doc.close()
                
                if search in text:
                    dest_path = os.path.join(dest, filename)
                    
                    if os.path.exists(dest_path) and not self.overwrite_var.get():
                        skipped += 1
                    else:
                        shutil.copy2(pdf_path, dest_path)
                        extracted += 1
            except Exception as e:
                errors += 1
        
        self.progress_var.set(100)
        self.percent_label.config(text="100% ✓")
        self.progress_label.config(text="Complete!")
        
        result = f"✅ EXTRACTION COMPLETE\n\n"
        result += f"Scanned: {total} PDFs\n"
        result += f"Extracted: {extracted}\n"
        result += f"Skipped: {skipped}\n"
        result += f"Errors: {errors}\n\n"
        result += f"📂 {dest}"
        
        messagebox.showinfo("Success", result)
        
        # Reset after 2 seconds
        self.root.after(2000, lambda: [
            self.progress_var.set(0),
            self.percent_label.config(text="0%"),
            self.progress_label.config(text="Ready to extract")
        ])

if __name__ == "__main__":
    root = tk.Tk()
    app = IILInvoiceExtractor(root)
    root.mainloop()
