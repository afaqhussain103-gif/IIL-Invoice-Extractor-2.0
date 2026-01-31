"""
IIL Invoice Extractor v4.3
International Industries Limited
Back to basics - proven working version with date range
"""

import os
import shutil
import re
from datetime import datetime
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class InvoiceExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("IIL Invoice Extractor v4.3")
        self.root.geometry("650x550")
        
        # IIL Colors
        self.green = '#01783f'
        self.yellow = '#c2d501'
        
        self.setup_ui()
    
    def setup_ui(self):
        # Header
        header = tk.Frame(self.root, bg=self.green, height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="IIL INVOICE EXTRACTOR",
            font=('Arial', 16, 'bold'),
            bg=self.green,
            fg='white'
        ).pack(pady=15)
        
        # Main frame
        main = tk.Frame(self.root, padx=30, pady=20)
        main.pack(fill=tk.BOTH, expand=True)
        
        # Source folder
        tk.Label(main, text="Source Folder:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
        self.source_entry = tk.Entry(main, width=40, font=('Arial', 10))
        self.source_entry.grid(row=1, column=0, pady=5)
        tk.Button(main, text="Browse", command=self.browse_source, bg=self.green, fg='white').grid(row=1, column=1, padx=5)
        
        # Destination folder
        tk.Label(main, text="Destination Folder:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', pady=5)
        self.dest_entry = tk.Entry(main, width=40, font=('Arial', 10))
        self.dest_entry.grid(row=3, column=0, pady=5)
        tk.Button(main, text="Browse", command=self.browse_dest, bg=self.green, fg='white').grid(row=3, column=1, padx=5)
        
        # Date From
        tk.Label(main, text="Date From (YYYY-MM-DD) - Optional:", font=('Arial', 10, 'bold')).grid(row=4, column=0, sticky='w', pady=5)
        self.date_from = tk.Entry(main, width=40, font=('Arial', 10))
        self.date_from.grid(row=5, column=0, pady=5)
        self.date_from.insert(0, "2024-01-01")
        
        # Date To
        tk.Label(main, text="Date To (YYYY-MM-DD) - Optional:", font=('Arial', 10, 'bold')).grid(row=6, column=0, sticky='w', pady=5)
        self.date_to = tk.Entry(main, width=40, font=('Arial', 10))
        self.date_to.grid(row=7, column=0, pady=5)
        self.date_to.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # Search text
        tk.Label(main, text="Search Text (Customer Name/ID):", font=('Arial', 10, 'bold')).grid(row=8, column=0, sticky='w', pady=5)
        self.search_entry = tk.Entry(main, width=40, font=('Arial', 10))
        self.search_entry.grid(row=9, column=0, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main, length=400, mode='determinate')
        self.progress.grid(row=10, column=0, columnspan=2, pady=15)
        
        self.status_label = tk.Label(main, text="Ready", font=('Arial', 9))
        self.status_label.grid(row=11, column=0, columnspan=2)
        
        # Extract button
        tk.Button(
            main,
            text="START EXTRACTION",
            command=self.extract_invoices,
            bg=self.yellow,
            font=('Arial', 12, 'bold'),
            padx=30,
            pady=10
        ).grid(row=12, column=0, columnspan=2, pady=15)
    
    def browse_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, folder)
    
    def browse_dest(self):
        folder = filedialog.askdirectory()
        if folder:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, folder)
    
    def extract_date_from_pdf(self, text):
        """Extract date from PDF text"""
        # Try multiple date patterns
        patterns = [
            r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})',  # DD/MM/YYYY
            r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',  # YYYY-MM-DD
            r'(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})',  # DD Month YYYY
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                try:
                    groups = match.groups()
                    if len(groups[0]) == 4:  # YYYY-MM-DD
                        return datetime(int(groups[0]), int(groups[1]), int(groups[2]))
                    elif groups[1].isalpha():  # DD Month YYYY
                        months = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
                                'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
                        month = months.get(groups[1][:3].lower())
                        if month:
                            return datetime(int(groups[2]), month, int(groups[0]))
                    else:  # DD/MM/YYYY
                        return datetime(int(groups[2]), int(groups[1]), int(groups[0]))
                except:
                    continue
        return None
    
    def extract_invoices(self):
        source = self.source_entry.get()
        dest = self.dest_entry.get()
        search_text = self.search_entry.get().lower()
        date_from_str = self.date_from.get().strip()
        date_to_str = self.date_to.get().strip()
        
        if not source or not dest or not search_text:
            messagebox.showerror("Error", "Please fill all required fields!")
            return
        
        if not os.path.exists(source):
            messagebox.showerror("Error", f"Source folder not found:\n{source}")
            return
        
        # Parse dates
        date_from = None
        date_to = None
        
        if date_from_str:
            try:
                date_from = datetime.strptime(date_from_str, "%Y-%m-%d")
            except:
                messagebox.showerror("Error", "Invalid 'From' date! Use YYYY-MM-DD")
                return
        
        if date_to_str:
            try:
                date_to = datetime.strptime(date_to_str, "%Y-%m-%d")
            except:
                messagebox.showerror("Error", "Invalid 'To' date! Use YYYY-MM-DD")
                return
        
        # Create destination folder
        os.makedirs(dest, exist_ok=True)
        
        # Get all PDF files
        all_files = os.listdir(source)
        pdf_files = [f for f in all_files if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            messagebox.showerror("Error", f"No PDF files found in:\n{source}\n\nFiles found: {len(all_files)}")
            return
        
        total = len(pdf_files)
        extracted = 0
        skipped = 0
        
        self.progress['maximum'] = total
        
        for i, filename in enumerate(pdf_files):
            self.progress['value'] = i + 1
            self.status_label.config(text=f"Processing {i+1}/{total}: {filename}")
            self.root.update()
            
            filepath = os.path.join(source, filename)
            
            try:
                # Open PDF and extract text
                doc = fitz.open(filepath)
                text = ""
                for page in doc:
                    text += page.get_text()
                doc.close()
                
                # Check if search text is in PDF
                if search_text not in text.lower():
                    skipped += 1
                    continue
                
                # Check date range if specified
                if date_from or date_to:
                    pdf_date = self.extract_date_from_pdf(text)
                    if pdf_date:
                        if date_from and pdf_date < date_from:
                            skipped += 1
                            continue
                        if date_to and pdf_date > date_to:
                            skipped += 1
                            continue
                
                # Copy file (always overwrite)
                dest_path = os.path.join(dest, filename)
                shutil.copy2(filepath, dest_path)
                extracted += 1
                
            except Exception as e:
                print(f"Error processing {filename}: {e}")
        
        self.status_label.config(text="Complete!")
        messagebox.showinfo(
            "Success",
            f"Extraction Complete!\n\n"
            f"Total PDFs: {total}\n"
            f"Extracted: {extracted}\n"
            f"Skipped: {skipped}\n\n"
            f"Saved to: {dest}"
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceExtractor(root)
    root.mainloop()
