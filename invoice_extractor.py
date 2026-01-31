"""
IIL Invoice Extractor v5.0 - PAGE EXTRACTOR
Extracts individual pages containing search term
"""

import os
import re
from datetime import datetime
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class IILPageExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("IIL Invoice Extractor v5.0")
        self.root.geometry("650x550")
        
        # IIL Colors
        self.green = '#01783f'
        self.yellow = '#c2d501'
        
        self.setup_ui()
    
    def setup_ui(self):
        # Header
        header = tk.Frame(self.root, bg=self.green, height=70)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="IIL INVOICE EXTRACTOR",
            font=('Arial', 14, 'bold'),
            bg=self.green,
            fg='white'
        ).pack(pady=20)
        
        # Main frame
        main = tk.Frame(self.root, padx=25, pady=15)
        main.pack(fill=tk.BOTH, expand=True)
        
        # Source
        tk.Label(main, text="Source Folder (with PDFs):", font=('Arial', 10, 'bold')).pack(anchor='w', pady=3)
        sf = tk.Frame(main)
        sf.pack(fill=tk.X, pady=5)
        self.source_entry = tk.Entry(sf, font=('Arial', 9))
        self.source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(sf, text="Browse", command=self.browse_source, bg=self.green, fg='white', padx=10).pack(side=tk.LEFT, padx=5)
        
        # Destination
        tk.Label(main, text="Destination Folder:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=3)
        df = tk.Frame(main)
        df.pack(fill=tk.X, pady=5)
        self.dest_entry = tk.Entry(df, font=('Arial', 9))
        self.dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(df, text="Browse", command=self.browse_dest, bg=self.green, fg='white', padx=10).pack(side=tk.LEFT, padx=5)
        
        # Date From
        tk.Label(main, text="Date From (YYYY-MM-DD, optional):", font=('Arial', 9)).pack(anchor='w', pady=3)
        self.date_from = tk.Entry(main, font=('Arial', 9))
        self.date_from.pack(fill=tk.X, pady=3)
        self.date_from.insert(0, "2024-01-01")
        
        # Date To
        tk.Label(main, text="Date To (YYYY-MM-DD, optional):", font=('Arial', 9)).pack(anchor='w', pady=3)
        self.date_to = tk.Entry(main, font=('Arial', 9))
        self.date_to.pack(fill=tk.X, pady=3)
        self.date_to.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # Search
        tk.Label(main, text="Customer Name or Account ID:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=3)
        self.search_entry = tk.Entry(main, font=('Arial', 10))
        self.search_entry.pack(fill=tk.X, pady=5)
        
        # Progress
        self.progress = ttk.Progressbar(main, length=500, mode='determinate')
        self.progress.pack(pady=10)
        
        self.status = tk.Label(main, text="Ready to extract pages", font=('Arial', 9))
        self.status.pack(pady=5)
        
        # Extract button
        tk.Button(
            main,
            text="EXTRACT PAGES",
            command=self.extract_pages,
            bg=self.yellow,
            font=('Arial', 11, 'bold'),
            padx=30,
            pady=10
        ).pack(pady=10)
    
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
    
    def extract_date_from_text(self, text):
        """Extract date from PDF text"""
        patterns = [
            r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})',
            r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',
            r'(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})',
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
    
    def extract_pages(self):
        source = self.source_entry.get()
        dest = self.dest_entry.get()
        search = self.search_entry.get().lower().strip()
        date_from_str = self.date_from.get().strip()
        date_to_str = self.date_to.get().strip()
        
        if not source or not dest or not search:
            messagebox.showerror("Error", "Please fill all required fields!")
            return
        
        if not os.path.exists(source):
            messagebox.showerror("Error", "Source folder doesn't exist!")
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
        
        os.makedirs(dest, exist_ok=True)
        
        # Get all PDFs
        pdf_files = [f for f in os.listdir(source) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            messagebox.showerror("Error", f"No PDF files found in:\n{source}")
            return
        
        total_pdfs = len(pdf_files)
        total_pages_extracted = 0
        pdfs_with_matches = 0
        
        self.progress['maximum'] = total_pdfs
        
        # Create a single output PDF for all extracted pages
        output_pdf = fitz.open()
        
        for idx, filename in enumerate(pdf_files):
            self.progress['value'] = idx + 1
            self.status.config(text=f"Scanning {idx+1}/{total_pdfs}: {filename[:40]}")
            self.root.update()
            
            filepath = os.path.join(source, filename)
            pages_found_in_this_pdf = 0
            
            try:
                doc = fitz.open(filepath)
                
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    text = page.get_text()
                    
                    # Check if search term is in page
                    if search in text.lower():
                        # Check date if specified
                        if date_from or date_to:
                            page_date = self.extract_date_from_text(text)
                            if page_date:
                                if date_from and page_date < date_from:
                                    continue
                                if date_to and page_date > date_to:
                                    continue
                        
                        # Extract this page
                        output_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        total_pages_extracted += 1
                        pages_found_in_this_pdf += 1
                
                if pages_found_in_this_pdf > 0:
                    pdfs_with_matches += 1
                
                doc.close()
                
            except Exception as e:
                print(f"Error processing {filename}: {e}")
        
        # Save the output PDF (overwrite if exists)
        if total_pages_extracted > 0:
            output_filename = f"{search.replace(' ', '_')}_extracted.pdf"
            output_path = os.path.join(dest, output_filename)
            output_pdf.save(output_path)
            output_pdf.close()
            
            self.status.config(text="Extraction Complete!")
            messagebox.showinfo(
                "Success",
                f"✅ EXTRACTION COMPLETE\n\n"
                f"📊 Scanned: {total_pdfs} PDFs\n"
                f"📑 Pages extracted: {total_pages_extracted}\n"
                f"📁 From {pdfs_with_matches} different PDFs\n\n"
                f"💾 Saved as:\n{output_filename}\n\n"
                f"📂 Location:\n{dest}"
            )
        else:
            output_pdf.close()
            messagebox.showinfo(
                "No Matches",
                f"No pages found containing:\n'{search}'\n\n"
                f"Scanned {total_pdfs} PDFs"
            )
        
        self.progress['value'] = 0
        self.status.config(text="Ready to extract pages")

if __name__ == "__main__":
    root = tk.Tk()
    app = IILPageExtractor(root)
    root.mainloop()
":
    root = tk.Tk()
    app = InvoiceExtractor(root)
    root.mainloop()
