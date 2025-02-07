import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
import os

class DataProcessorGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Data Processor and Report Generator")
        self.master.geometry("1000x700")
        
        # Initialize variables
        self.csv_files = []  # Will store multiple file paths
        self.combined_data = pd.DataFrame()
        self.plot_path = "plot.png"
        
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
        
        # Save CSV Button
        btn_csv = tk.Button(self.master, text="Save CSV", command=self.save_csv, bg="orange", fg="black")
        btn_csv.pack(pady=5)
        
        # Frame for data display
        frame_data = tk.Frame(self.master)
        frame_data.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview for displaying grouped data
        self.tree = ttk.Treeview(frame_data, columns=("File Source", "Measurement", "Min Loss", "Max Loss"), show='headings')
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
        
        # Generate Report Button
        btn_report = tk.Button(self.master, text="Generate PDF Report", command=self.generate_report, bg="blue", fg="white")
        btn_report.pack(pady=10)
    
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
            # Update label to show the count of files selected
            self.lbl_file.config(text=f"{len(self.csv_files)} file(s) selected")
    
    def process_data(self):
        if not self.csv_files:
            messagebox.showerror("Error", "Please select at least one CSV file.")
            return
        
        # Clear any old data
        self.combined_data = pd.DataFrame()
        
        try:
            all_grouped_frames = []

            # Process each CSV file one by one
            for file_path in self.csv_files:
                # Attempt to locate the custom header
                with open(file_path, 'r', encoding='utf-8') as f:
                    header_row = None
                    for idx, line in enumerate(f):
                        if line.strip().startswith("Time,Metric,Value,Measurement"):
                            header_row = idx
                            break
                    else:
                        messagebox.showwarning("Warning", f"Header row not found in {os.path.basename(file_path)}. Skipped.")
                        continue

                # Read the CSV from the correct header row
                df = pd.read_csv(
                    file_path,
                    skiprows=range(0, header_row),
                    header=0,
                    encoding="utf-8"
                )
                
                # Data Cleaning
                df['Percentage Loss'] = pd.to_numeric(df['Percentage Loss'], errors='coerce')
                df.dropna(subset=['Percentage Loss'], inplace=True)
                
                # Grouping
                grouped = df.groupby("Measurement")["Percentage Loss"].agg(['min', 'max']).reset_index()
                grouped.rename(columns={'min': 'Min Loss', 'max': 'Max Loss'}, inplace=True)
                
                # Add a column to track which file the data came from
                grouped['File Source'] = os.path.basename(file_path)
                
                all_grouped_frames.append(grouped)
            
            # Combine all grouped data into one DataFrame
            if all_grouped_frames:
                self.combined_data = pd.concat(all_grouped_frames, ignore_index=True)
            else:
                messagebox.showerror("Error", "No valid data processed from the selected files.")
                return
            
            # Update Treeview
            self.update_table(self.combined_data)
            
            # Generate Plot (optional: you could show combined or first file, etc.)
            # Here, we demonstrate a combined plot by grouping again on Measurement.
            # If you have measurements from multiple files with the same name, they’ll be aggregated together.
            # If you want them separate, you’ll need different logic.
            aggregated_for_plot = self.combined_data.groupby("Measurement").agg({
                'Min Loss': 'min',
                'Max Loss': 'max'
            }).reset_index()
            
            self.generate_plot(aggregated_for_plot)
            
            messagebox.showinfo("Success", "All files processed successfully.")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing the data:\n{e}")
    
    def update_table(self, data):
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Insert new data
        for _, row in data.iterrows():
            self.tree.insert("", tk.END, values=(
                row['File Source'],
                row['Measurement'],
                row['Min Loss'],
                row['Max Loss']
            ))
    
    def generate_plot(self, data):
        self.ax.clear()
        
        import numpy as np
        x = np.arange(len(data))
        width = 0.35
        
        self.ax.bar(x - width/2, data['Min Loss'], width, label='Min Loss')
        self.ax.bar(x + width/2, data['Max Loss'], width, label='Max Loss')
        
        self.ax.set_xlabel('Measurement')
        self.ax.set_ylabel('Percentage Loss')
        self.ax.set_title('Minimum and Maximum Percentage Loss (Combined)')
        self.ax.set_xticks(x)
        self.ax.set_xticklabels(data['Measurement'], rotation=45, ha='right')
        self.ax.legend()
        
        self.fig.tight_layout()
        self.canvas.draw()
        
        # Save plot as image for PDF
        self.fig.savefig(self.plot_path)
    
    def generate_report(self):
        """
        Generates a PDF report with a consolidated table and plot of all processed data.
        """
        if self.combined_data.empty:
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
            
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            c = canvas.Canvas(report_path, pagesize=letter)
            width, height = letter
            styles = getSampleStyleSheet()
            
            # Title
            c.setFont("Helvetica-Bold", 20)
            c.drawCentredString(width / 2, height - 50, "Data Processing Report")
            
            # Sub-title / Date
            c.setFont("Helvetica", 12)
            c.drawString(50, height - 80, f"Report generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")

            # We need to create a table with all columns
            # Reorganize columns for the PDF table
            display_columns = ["File Source", "Measurement", "Min Loss", "Max Loss"]
            pdf_data = [display_columns] + self.combined_data[display_columns].values.tolist()
            
            table = Table(pdf_data, colWidths=[120, 150, 80, 80])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                ('BACKGROUND',(0,1),(-1,-1),colors.beige),
                ('GRID', (0,0), (-1,-1), 1, colors.black),
            ]))
            
            table_width, table_height = table.wrap(0, 0)
            table.drawOn(c, 50, height - 150 - table_height)
            
            # Position just below table
            y_position = height - 160 - table_height
            
            # Add plot image if exists
            if os.path.exists(self.plot_path):
                c.drawImage(self.plot_path, 50, y_position - 300, width=500, height=300)
            
            c.save()
            messagebox.showinfo("Success", f"Report generated successfully at:\n{report_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while generating the report:\n{e}")

    def save_csv(self):
        """
        Allows the user to save the combined_data DataFrame to a CSV file.
        """
        if self.combined_data.empty:
            messagebox.showerror("Error", "No data to save. Please process data first.")
            return
        
        # Ask user where to save the CSV
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save CSV As"
        )
        if not save_path:
            return  # User cancelled
        
        try:
            self.combined_data.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"Data saved successfully at:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the CSV:\n{e}")

def main():
    root = tk.Tk()
    app = DataProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
