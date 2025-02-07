import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Image, Paragraph, Spacer, SimpleDocTemplate
from reportlab.lib.styles import getSampleStyleSheet
import os
import numpy as np
import tempfile
import atexit

class DataProcessorGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Data Processor and Report Generator")
        self.master.geometry("1000x700")
        
        # Initialize variables
        self.data_file = ""
        self.grouped_data = pd.DataFrame()
        self.plot_path = "plot.png"
        
        # Create GUI components
        self.create_widgets()
    
    def create_widgets(self):
        # Frame for file selection
        frame_file = tk.Frame(self.master)
        frame_file.pack(pady=10)
        
        btn_browse = tk.Button(frame_file, text="Browse Data File", command=self.browse_file)
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
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")],
            title="Select Data File"
        )
        if file_path:
            self.data_file = file_path
            self.lbl_file.config(text=os.path.basename(file_path))
    
    def process_data(self):
        if not self.data_file:
            messagebox.showerror("Error", "Please select a data file first.")
            return

        try:
            file_extension = os.path.splitext(self.data_file)[1].lower()
            
            # Determine file type and read accordingly
            if file_extension == '.csv':
                # Detect header row for CSV
                df = self.read_csv_with_header_detection(self.data_file)
            elif file_extension in ['.xlsx', '.xls']:
                # Detect header row for Excel
                df = self.read_excel_with_header_detection(self.data_file)
            else:
                messagebox.showerror("Error", "Unsupported file type selected.")
                return

            if df is None:
                return  # Error message already shown in the read functions

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
    
    def read_csv_with_header_detection(self, file_path):
        try:
            # Detect header row by searching for the processed data header
            with open(file_path, 'r', encoding='utf-8') as f:
                for idx, line in enumerate(f):
                    if line.strip().startswith("Time,Metric,Value,Measurement"):
                        header_row = idx
                        break
                else:
                    messagebox.showerror("Error", "Header row not found in the CSV file.")
                    return None

            # Read the CSV, skipping metadata rows
            df = pd.read_csv(
                file_path,
                skiprows=range(0, header_row),
                header=0,
                encoding="utf-8"
            )

            # Validate required columns
            required_columns = ["Measurement", "Percentage Loss"]
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", f"CSV file must contain the following columns: {', '.join(required_columns)}")
                return None

            return df

        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", "The selected CSV file is empty.")
            return None
        except pd.errors.ParserError:
            messagebox.showerror("Error", "Error parsing the CSV file. Please ensure it is properly formatted.")
            return None
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while reading the CSV file:\n{e}")
            return None

    def read_excel_with_header_detection(self, file_path):
        try:
            # Read the entire Excel file to find the header row
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names  # List of sheet names
            df = None
            selected_sheet = None

            # Search for header in all sheets
            for sheet in sheet_names:
                temp_df = pd.read_excel(file_path, sheet_name=sheet, header=None)
                for idx, row in temp_df.iterrows():
                    if isinstance(row[0], str) and row[0].startswith("Time"):
                        # Assume this is the header row
                        df = pd.read_excel(file_path, sheet_name=sheet, skiprows=range(0, idx), header=0)
                        selected_sheet = sheet
                        break
                if df is not None:
                    break

            if df is None:
                messagebox.showerror("Error", "Header row not found in the Excel file.")
                return None

            # If multiple sheets have the header, ask the user to choose
            if sheet_names.count(selected_sheet) > 1:
                selected_sheet = simpledialog.askstring("Select Sheet", f"Multiple sheets found. Enter the sheet name to use (Available sheets: {', '.join(sheet_names)}):")
                if selected_sheet not in sheet_names:
                    messagebox.showerror("Error", "Invalid sheet name entered.")
                    return None
                df = pd.read_excel(file_path, sheet_name=selected_sheet, skiprows=range(0, idx), header=0)

            # Validate required columns
            required_columns = ["Measurement", "Percentage Loss"]
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", f"Excel file must contain the following columns: {', '.join(required_columns)}")
                return None

            return df

        except FileNotFoundError:
            messagebox.showerror("Error", "The selected Excel file was not found.")
            return None
        except ValueError as ve:
            messagebox.showerror("Error", f"Value error: {ve}")
            return None
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while reading the Excel file:\n{e}")
            return None

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
        # Use a temporary file to avoid conflicts
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmpfile:
            self.plot_path = tmpfile.name
            self.fig.savefig(self.plot_path)
        
        # Register cleanup for temporary plot image
        atexit.register(lambda: os.remove(self.plot_path) if os.path.exists(self.plot_path) else None)
    
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
            
            # Use ReportLab's Platypus for better layout management
            doc = SimpleDocTemplate(report_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            
            # Title
            title = Paragraph("Data Processing Report", styles['Heading1'])
            elements.append(title)
            elements.append(Spacer(1, 12))
            
            # Timestamp
            timestamp = Paragraph(f"Report generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal'])
            elements.append(timestamp)
            elements.append(Spacer(1, 12))
            
            # Table Data
            table_data = [self.grouped_data.columns.tolist()] + self.grouped_data.values.tolist()
            table = Table(table_data, colWidths=[250, 100, 100])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                ('BACKGROUND',(0,1),(-1,-1),colors.beige),
                ('GRID', (0,0), (-1,-1), 1, colors.black),
            ]))
            elements.append(table)
            elements.append(Spacer(1, 12))
            
            # Add the plot image
            if os.path.exists(self.plot_path):
                img = Image(self.plot_path, width=500, height=300)
                elements.append(img)
            else:
                elements.append(Paragraph("Plot image not found.", styles['Normal']))
            
            # Build the PDF
            doc.build(elements)
            messagebox.showinfo("Success", f"Report generated successfully at:\n{report_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while generating the report:\n{e}")

def main():
    root = tk.Tk()
    app = DataProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
