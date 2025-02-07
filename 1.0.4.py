import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

# ReportLab imports
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

class DataProcessorGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Data Processor and Report Generator")
        self.master.geometry("1000x700")
        
        # Will store multiple file paths
        self.csv_files = []
        
        # Create GUI components
        self.create_widgets()
    
    def create_widgets(self):
        # Frame for file selection
        frame_file = tk.Frame(self.master)
        frame_file.pack(pady=10)
        
        btn_browse = tk.Button(frame_file, text="Browse CSV(s)", command=self.browse_files)
        btn_browse.pack(side=tk.LEFT, padx=5)
        
        self.lbl_file = tk.Label(frame_file, text="No files selected", width=80, anchor="w")
        self.lbl_file.pack(side=tk.LEFT, padx=5)
        
        # Process Button
        btn_process = tk.Button(self.master, text="Process Data", command=self.process_data, bg="green", fg="white")
        btn_process.pack(pady=10)
        
        # Frame for data display
        frame_data = tk.Frame(self.master)
        frame_data.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview for displaying grouped data
        self.tree = ttk.Treeview(
            frame_data, 
            columns=("File Source", "Measurement", "Min Loss", "Max Loss"), 
            show='headings'
        )
        self.tree.heading("File Source", text="File Source")
        self.tree.heading("Measurement", text="Measurement")
        self.tree.heading("Min Loss", text="Min Loss")
        self.tree.heading("Max Loss", text="Max Loss")
        
        self.tree.column("File Source", width=200)
        self.tree.column("Measurement", width=200)
        self.tree.column("Min Loss", width=100, anchor='center')
        self.tree.column("Max Loss", width=100, anchor='center')
        
        scrollbar = ttk.Scrollbar(frame_data, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Frame for plots
        frame_plot = tk.Frame(self.master)
        frame_plot.pack(pady=10)
        
        self.fig, self.ax = plt.subplots(figsize=(8,4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame_plot)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack()
    
    def browse_files(self):
        """
        Allows user to select multiple CSV files.
        """
        file_paths = filedialog.askopenfilenames(
            filetypes=[("CSV files", "*.csv")],
            title="Select One or More CSV Files"
        )
        if file_paths:
            self.csv_files = list(file_paths)
            # Update label to show how many files are selected
            self.lbl_file.config(text=f"{len(self.csv_files)} file(s) selected")
    
    def process_data(self):
        """
        Process each selected CSV file one by one.
        For each file:
          - Detect header row ("Time,Metric,Value,Measurement")
          - Clean & group data
          - Update UI Treeview
          - Generate a bar chart
          - Automatically save CSV and PDF
        """
        if not self.csv_files:
            messagebox.showerror("Error", "Please select at least one CSV file.")
            return
        
        try:
            # Clear the Treeview from any old data
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            for file_path in self.csv_files:
                # 1. Locate custom header row
                header_row = None
                with open(file_path, 'r', encoding='utf-8') as f:
                    for idx, line in enumerate(f):
                        if line.strip().startswith("Time,Metric,Value,Measurement"):
                            header_row = idx
                            break
                
                if header_row is None:
                    messagebox.showwarning(
                        "Warning", 
                        f"Header row not found in {os.path.basename(file_path)}. Skipped."
                    )
                    continue

                # 2. Read CSV from the detected header row
                df = pd.read_csv(
                    file_path,
                    skiprows=range(0, header_row),
                    header=0,
                    encoding="utf-8"
                )
                
                # 3. Data cleaning
                df['Percentage Loss'] = pd.to_numeric(df['Percentage Loss'], errors='coerce')
                df.dropna(subset=['Percentage Loss'], inplace=True)
                
                # 4. Group by Measurement
                grouped_data = df.groupby("Measurement")["Percentage Loss"].agg(['min', 'max']).reset_index()
                grouped_data.rename(columns={'min': 'Min Loss', 'max': 'Max Loss'}, inplace=True)
                
                # 5. Update the Treeview for each file
                for _, row in grouped_data.iterrows():
                    self.tree.insert("", tk.END, values=(
                        os.path.basename(file_path),
                        row['Measurement'],
                        row['Min Loss'],
                        row['Max Loss']
                    ))
                
                # 6. Generate a plot for THIS file
                self.generate_plot_for_file(grouped_data, file_path)
                
                # 7. Automatically save a CSV with processed results
                self.save_csv_for_file(grouped_data, file_path)
                
                # 8. Automatically generate a PDF with table + plot
                self.generate_pdf_for_file(grouped_data, file_path)
            
            messagebox.showinfo("Success", "All selected files have been processed and saved.")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing the data:\n{e}")
    
    def generate_plot_for_file(self, grouped_data, file_path):
        """
        Generate and display a bar chart for the current file's grouped data,
        then save the figure as <original_filename>_plot.png
        """
        self.ax.clear()
        
        import numpy as np
        x = np.arange(len(grouped_data))
        width = 0.35
        
        self.ax.bar(x - width/2, grouped_data['Min Loss'], width, label='Min Loss')
        self.ax.bar(x + width/2, grouped_data['Max Loss'], width, label='Max Loss')
        
        self.ax.set_xlabel('Measurement')
        self.ax.set_ylabel('Percentage Loss')
        self.ax.set_title(f'Min/Max Percentage Loss - {os.path.basename(file_path)}')
        self.ax.set_xticks(x)
        self.ax.set_xticklabels(grouped_data['Measurement'], rotation=45, ha='right')
        self.ax.legend()
        
        self.fig.tight_layout()
        self.canvas.draw()
        
        # Construct a plot filename (e.g., "data_plot.png")
        base_name = os.path.splitext(file_path)[0]  # full path without .csv
        plot_file = base_name + "_plot.png"
        
        # Save plot as image
        self.fig.savefig(plot_file)
    
    def save_csv_for_file(self, grouped_data, file_path):
        """
        Save grouped data to <original_filename>_processed.csv
        """
        base_name = os.path.splitext(file_path)[0]
        save_path = base_name + "_processed.csv"
        grouped_data.to_csv(save_path, index=False)
    
    def generate_pdf_for_file(self, grouped_data, file_path):
        """
        Generate a PDF report for each file, containing a table and the plot image.
        Saves to <original_filename>_report.pdf
        """
        base_name = os.path.splitext(file_path)[0]
        pdf_path = base_name + "_report.pdf"
        
        # Create a canvas
        c = canvas.Canvas(pdf_path, pagesize=letter)
        width, height = letter
        styles = getSampleStyleSheet()
        
        # Title
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(width / 2, height - 50, "Data Processing Report")
        
        # Sub-title / date
        c.setFont("Helvetica", 12)
        c.drawString(50, height - 80, f"Report for: {os.path.basename(file_path)}")
        c.drawString(50, height - 100, f"Generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Prepare table data
        pdf_data = [["Measurement", "Min Loss", "Max Loss"]]
        for _, row in grouped_data.iterrows():
            pdf_data.append([row['Measurement'], row['Min Loss'], row['Max Loss']])
        
        # Create table
        table = Table(pdf_data, colWidths=[200, 100, 100])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
            ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 12),
            ('BACKGROUND',(0,1),(-1,-1),colors.beige),
            ('GRID', (0,0), (-1,-1), 1, colors.black),
        ]))
        
        # Draw the table
        table_width, table_height = table.wrap(0, 0)
        table.drawOn(c, 50, height - 150 - table_height)
        
        # Attempt to embed the plot
        plot_file = base_name + "_plot.png"
        y_position = height - 160 - table_height
        if os.path.exists(plot_file):
            # Insert the plot image below the table
            c.drawImage(plot_file, 50, y_position - 300, width=400, height=300, preserveAspectRatio=True)
        
        # Save PDF
        c.save()

def main():
    root = tk.Tk()
    app = DataProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
