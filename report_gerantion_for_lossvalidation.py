import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Image, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import os

class DataProcessorGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Data Processor and Report Generator")
        self.master.geometry("1000x700")
        
        # Initialize variables
        self.csv_file = ""
        self.grouped_data = pd.DataFrame()
        self.plot_path = "plot.png"
        
        # Create GUI components
        self.create_widgets()
    
    def create_widgets(self):
        # Frame for file selection
        frame_file = tk.Frame(self.master)
        frame_file.pack(pady=10)
        
        btn_browse = tk.Button(frame_file, text="Browse CSV", command=self.browse_file)
        btn_browse.pack(side=tk.LEFT, padx=5)
        
        self.lbl_file = tk.Label(frame_file, text="No file selected", width=80, anchor="w")
        self.lbl_file.pack(side=tk.LEFT, padx=5)
        
        # Process Button
        btn_process = tk.Button(self.master, text="Process Data", command=self.process_data, bg="green", fg="white")
        btn_process.pack(pady=10)
        
        # Frame for data display
        frame_data = tk.Frame(self.master)
        frame_data.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbar for the table
        self.tree = ttk.Treeview(frame_data, columns=("Measurement", "Min Loss", "Max Loss"), show='headings')
        self.tree.heading("Measurement", text="Measurement")
        self.tree.heading("Min Loss", text="Min Loss")
        self.tree.heading("Max Loss", text="Max Loss")
        
        self.tree.column("Measurement", width=300)
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
        
        # Generate Report Button
        btn_report = tk.Button(self.master, text="Generate PDF Report", command=self.generate_report, bg="blue", fg="white")
        btn_report.pack(pady=10)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")],
            title="Select CSV File"
        )
        if file_path:
            self.csv_file = file_path
            self.lbl_file.config(text=os.path.basename(file_path))
    
    def process_data(self):
        if not self.csv_file:
            messagebox.showerror("Error", "Please select a CSV file first.")
            return
        
        try:
            # Detect header row by searching for the processed data header
            header_identifier = "Time,Metric,Value,Measurement,InfluxDB Field Name,Available in Valid File?,CAN Dictionary MAP,Time ,Expected Count,Loss,Percentage Loss"
            with open(self.csv_file, 'r', encoding='utf-8') as f:
                for idx, line in enumerate(f):
                    if line.strip().startswith("Time,Metric,Value,Measurement"):
                        header_row = idx
                        break
                else:
                    messagebox.showerror("Error", "Header row not found in the CSV file.")
                    return
            
            # Read the CSV, skipping metadata rows
            df = pd.read_csv(
                self.csv_file,
                skiprows=range(0, header_row),
                header=0,
                encoding="utf-8"
            )
            
            # Data Cleaning
            df['Percentage Loss'] = pd.to_numeric(df['Percentage Loss'], errors='coerce')
            initial_row_count = df.shape[0]
            df = df.dropna(subset=['Percentage Loss'])
            cleaned_row_count = df.shape[0]
            dropped_rows = initial_row_count - cleaned_row_count
            
            if dropped_rows > 0:
                messagebox.showwarning("Warning", f"Dropped {dropped_rows} rows due to invalid 'Percentage Loss' values.")
            
            # Grouping
            grouped = df.groupby("Measurement")["Percentage Loss"].agg(['min', 'max']).reset_index()
            grouped.rename(columns={'min': 'Min Loss', 'max': 'Max Loss'}, inplace=True)
            self.grouped_data = grouped
            
            # Update Treeview
            self.update_table(grouped)
            
            # Generate Plot
            self.generate_plot(grouped)
            
            messagebox.showinfo("Success", "Data processed successfully.")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing the data:\n{e}")
    
    def update_table(self, data):
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Insert new data
        for _, row in data.iterrows():
            self.tree.insert("", tk.END, values=(row['Measurement'], row['Min Loss'], row['Max Loss']))
    
    def generate_plot(self, data):
        self.ax.clear()
        
        # Set positions and width for the bars
        import numpy as np
        x = np.arange(len(data))
        width = 0.35
        
        # Plot min and max losses
        self.ax.bar(x - width/2, data['Min Loss'], width, label='Min Loss')
        self.ax.bar(x + width/2, data['Max Loss'], width, label='Max Loss')
        
        # Labels and titles
        self.ax.set_xlabel('Measurement')
        self.ax.set_ylabel('Percentage Loss')
        self.ax.set_title('Minimum and Maximum Percentage Loss per Measurement')
        self.ax.set_xticks(x)
        self.ax.set_xticklabels(data['Measurement'], rotation=45, ha='right')
        self.ax.legend()
        
        self.fig.tight_layout()
        self.canvas.draw()
        
        # Save plot as image for PDF
        self.fig.savefig(self.plot_path)
    
    def generate_report(self):
        if self.grouped_data.empty:
            messagebox.showerror("Error", "No data to generate report. Please process data first.")
            return
        
        try:
            # Ask user where to save the PDF
            report_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Report As"
            )
            if not report_path:
                return  # User cancelled
            
            c = canvas.Canvas(report_path, pagesize=letter)
            width, height = letter
            styles = getSampleStyleSheet()
            styleN = styles['Normal']
            styleH = styles['Heading1']
            
            # Title
            c.setFont("Helvetica-Bold", 20)
            c.drawCentredString(width / 2, height - 50, "Data Processing Report")
            
            # Spacer
            c.setFont("Helvetica", 12)
            c.drawString(50, height - 80, f"Report generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            # Table Data
            data = [self.grouped_data.columns.tolist()] + self.grouped_data.values.tolist()
            table = Table(data, colWidths=[250, 100, 100])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                ('BACKGROUND',(0,1),(-1,-1),colors.beige),
                ('GRID', (0,0), (-1,-1), 1, colors.black),
            ]))
            
            # Position the table
            table_width, table_height = table.wrap(0, 0)
            table.drawOn(c, 50, height - 150 - table_height)
            
            # Spacer
            y_position = height - 160 - table_height
            
            # Add the plot image
            if os.path.exists(self.plot_path):
                c.drawImage(self.plot_path, 50, y_position - 300, width=500, height=300)
                y_position -= 320  # Adjust y position after the image
            else:
                c.drawString(50, y_position - 20, "Plot image not found.")
                y_position -= 30
            
            # Save the PDF
            c.save()
            messagebox.showinfo("Success", f"Report generated successfully at:\n{report_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while generating the report:\n{e}")

def main():
    root = tk.Tk()
    app = DataProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
