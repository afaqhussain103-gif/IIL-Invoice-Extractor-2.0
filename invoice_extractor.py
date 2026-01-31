"""
IIL Invoice Extractor v4.2 - Date Range + Auto-Overwrite
International Industries Limited
"""

import os
import shutil
import re
from datetime import datetime
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class IILInvoiceExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("IIL Invoice Extractor v4.2")
        self.root.geometry("700x620")  # Slightly taller for date fields
        self.root.resizable(True, True)
        self.root.minsize(600, 550)
        
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
            text="v4.2 - International Industries Limited",
            font=('Arial', 8),
            bg=self.colors['green'],
            fg=self.colors['yellow']
        ).pack()
        
        # MAIN CONTENT
        main_container = tk.Frame(self.root, bg=self.colors['white'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        # SOURCE FOLDER
        tk.Label(
            main_container,
            text="📁 Source Folder:",
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green']
        ).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 3))
        
        source_frame = tk.Frame(main_container, bg=self.colors['white'])
        source_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        
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
        ).grid(row=2, column=0, columnspan=2, sticky='w', pady=(0, 3))
        
        dest_frame = tk.Frame(main_container, bg=self.colors['white'])
        dest_frame.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        
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
            row=4, column=0, columnspan=2, sticky='ew', pady=10
        )
        
        # DATE RANGE
        tk.Label(
            main_container,
            text="📅 Date Range (Optional - Format: YYYY-MM-DD):",
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green']
        ).grid(row=5, column=0, columnspan=2, sticky='w', pady=(0, 3))
        
        date_frame = tk.Frame(main_container, bg=self.colors['white'])
        date_frame.grid(row=6, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        
        tk.Label(date_frame, text="From:", font=('Arial', 8), bg=self.colors['white']).pack(side=tk.LEFT, padx=(0, 5))
        self.date_from_entry = tk.Entry(date_frame, font=('Arial', 9), width=12)
        self.date_from_entry.pack(side=tk.LEFT, ipady=4)
        self.date_from_entry.insert(0, "2024-01-01")
        
        tk.Label(date_frame, text="To:", font=('Arial', 8), bg=self.colors['white']).pack(side=tk.LEFT, padx=(15, 5))
        self.date_to_entry = tk.Entry(date_frame, font=('Arial', 9), width=12)
        self.date_to_entry.pack(side=tk.LEFT, ipady=4)
        self.date_to_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # SEARCH VALUE
        tk.Label(
            main_container,
            text="🔎 Search (Customer Name or Account ID):",
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green']
        ).grid(row=7, column=0, columnspan=2, sticky='w', pady=(0, 3))
        
        self.search_entry = tk.Entry(main_container, font=('Arial', 10))
        self.search_entry.grid(row=8, column=0, columnspan=2, sticky='ew', pady=(0, 10), ipady=7)
        
        # INFO NOTE
        tk.Label(
            main_container,
            text="ℹ️  Files with the same name will be automatically overwritten",
            font=('Arial', 7, 'italic'),
            bg=self.colors['white'],
            fg=self.colors['text']
        ).grid(row=9, column=0, columnspan=2, sticky='w', pady=(0, 10))
        
        # SEPARATOR
        tk.Frame(main_container, bg=self.colors['bg'], height=1).grid(
            row=10, column=0, columnspan=2, sticky='ew', pady=10
        )
        
        # PROGRESS - Compact
        progress_frame = tk.Frame(main_container, bg=self.colors['bg'])
        progress_frame.grid(row=11, column=0, columnspan=2, sticky='ew', pady=(0, 10), ipady=10)
        
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
        
        # EXTRACT BUTTON
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
        ).grid(row=12, column=0, columnspan=2, pady=10)
        
        # Configure grid weights
        main_container.grid_columnconfigure(0, weight=1)
    
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
    
    def parse_date_from_text(self, text):
        """Extract date from PDF text in various formats"""
        date_patterns = [
            r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b',  # DD/MM/YYYY or DD-MM-YYYY
            r'\b(\d{4})[/-](\d{1,2})[/-](\d{1,2})\b',  # YYYY-MM-DD
            r'\b(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{4})\b',  # DD Month YYYY
        ]
        
        for pattern in date_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                try:
                    match = matches[0]
                    if len(match) == 3:
                        if match[2].isdigit() and len(match[2]) == 4:  # DD/MM/YYYY
                            return datetime(int(match[2]), int(match[1]), int(match[0]))
                        elif match[0].isdigit() and len(match[0]) == 4:  # YYYY-MM-DD
                            return datetime(int(match[0]), int(match[1]), int(match[2]))
                        else:  # DD Month YYYY
                            months = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
                                    'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
                            month = months.get(match[1][:3].lower())
                            if month:
                                return datetime(int(match[2]), month, int(match[0]))
                except:
                    continue
        return None
    
    def extract(self):
        source = self.source_entry.get().strip()
        dest = self.dest_entry.get().strip()
        search = self.search_entry.get().strip().lower()
        date_from_str = self.date_from_entry.get().strip()
        date_to_str = self.date_to_entry.get().strip()
        
        if not source or not dest or not search:
            messagebox.showwarning("Missing Info", "Please fill source, destination, and search fields!")
            return
        
        if not os.path.exists(source):
            messagebox.showerror("Error", "Source folder doesn't exist!")
            return
        
        # Parse date range
        date_from = None
        date_to = None
        
        if date_from_str:
            try:
                date_from = datetime.strptime(date_from_str, "%Y-%m-%d")
            except:
                messagebox.showerror("Error", "Invalid 'From' date format! Use YYYY-MM-DD")
                return
        
        if date_to_str:
            try:
                date_to = datetime.strptime(date_to_str, "%Y-%m-%d")
            except:
                messagebox.showerror("Error", "Invalid 'To' date format! Use YYYY-MM-DD")
                return
        
        if not os.path.exists(dest):
            os.makedirs(dest)
        
        pdf_files = [f for f in os.listdir(source) if f.lower().endswith('.pdf')]
        total = len(pdf_files)
        
        if total == 0:
            messagebox.showinfo("No Files", "No PDFs found in source folder!")
            return
        
        extracted = 0
        skipped_search = 0
        skipped_date = 0
        errors = 0
        
        for idx, filename in enumerate(pdf_files, 1):
            pdf_path = os.path.join(source, filename)
            
            progress = (idx / total) * 100
            self.progress_var.set(progress)
            self.percent_label.config(text=f"{int(progress)}%")
            
            display_name = filename[:30] + "..." if len(filename) > 30 else filename
            self.progress_label.config(text=f"Processing: {display_name}")
            self.root.update_idletasks()
            
            try:
                doc = fitz.open(pdf_path)
                text = ""
                for page in doc:
                    text += page.get_text()
                doc.close()
                
                text_lower = text.lower()
                
                # Check search term
                if search not in text_lower:
                    skipped_search += 1
                    continue
                
                # Check date range
                if date_from or date_to:
                    pdf_date = self.parse_date_from_text(text)
                    if pdf_date:
                        if date_from and pdf_date < date_from:
                            skipped_date += 1
                            continue
                        if date_to and pdf_date > date_to:
                            skipped_date += 1
                            continue
                    else:
                        # If no date found, skip if date filtering is active
                        if date_from or date_to:
                            skipped_date += 1
                            continue
                
                # Copy file (always overwrite)
                dest_path = os.path.join(dest, filename)
                shutil.copy2(pdf_path, dest_path)
                extracted += 1
                
            except Exception as e:
                errors += 1
        
        self.progress_var.set(100)
        self.percent_label.config(text="100% ✓")
        self.progress_label.config(text="Complete!")
        
        result = f"✅ EXTRACTION COMPLETE\n\n"
        result += f"📊 Total Scanned: {total} PDFs\n"
        result += f"✅ Extracted: {extracted}\n"
        result += f"🔍 Skipped (no match): {skipped_search}\n"
        result += f"📅 Skipped (date): {skipped_date}\n"
        result += f"❌ Errors: {errors}\n\n"
        result += f"📂 Location:\n{dest}"
        
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


if __name__ == "__main__":
    root = tk.Tk()
    app = IILInvoiceExtractor(root)
    root.mainloop()
