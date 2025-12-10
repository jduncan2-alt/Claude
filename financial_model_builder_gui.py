#!/usr/bin/env python3
"""
Financial Model Builder - GUI Version
Graphical interface for creating 3-statement financial models from Capital IQ exports
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import sys
import os
from pathlib import Path

# Import the core functionality from the main script
from build_financial_model import CapIQDataExtractor, FinancialModelBuilder, ExcelFormatter


class FinancialModelBuilderGUI:
    """GUI application for financial model builder"""

    def __init__(self, root):
        self.root = root
        self.root.title("Financial Model Builder")
        self.root.geometry("800x700")
        self.root.resizable(True, True)

        # Variables
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.company_name = tk.StringVar()
        self.num_projection_years = tk.IntVar(value=3)
        self.revenue_growth_override = tk.StringVar(value="")
        self.use_auto_growth = tk.BooleanVar(value=True)

        self.setup_ui()

    def setup_ui(self):
        """Setup the user interface"""
        # Title
        title_frame = ttk.Frame(self.root, padding="10")
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

        title_label = ttk.Label(
            title_frame,
            text="Financial Model Builder",
            font=("Helvetica", 18, "bold")
        )
        title_label.grid(row=0, column=0, sticky=tk.W)

        subtitle_label = ttk.Label(
            title_frame,
            text="Create 3-statement financial models from Capital IQ exports",
            font=("Helvetica", 10)
        )
        subtitle_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 10))

        # Main content frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights for responsive layout
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        main_frame.columnconfigure(1, weight=1)

        current_row = 0

        # Input File Section
        ttk.Label(main_frame, text="Input Capital IQ File:", font=("Helvetica", 10, "bold")).grid(
            row=current_row, column=0, sticky=tk.W, pady=(10, 5)
        )
        current_row += 1

        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)

        ttk.Entry(input_frame, textvariable=self.input_file_path, width=50).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(input_frame, text="Browse...", command=self.browse_input_file).grid(
            row=0, column=1
        )
        current_row += 1

        # Company Name
        ttk.Label(main_frame, text="Company Name:", font=("Helvetica", 10, "bold")).grid(
            row=current_row, column=0, sticky=tk.W, pady=(10, 5)
        )
        current_row += 1

        ttk.Entry(main_frame, textvariable=self.company_name, width=50).grid(
            row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10)
        )
        current_row += 1

        # Output File Section
        ttk.Label(main_frame, text="Output Excel File:", font=("Helvetica", 10, "bold")).grid(
            row=current_row, column=0, sticky=tk.W, pady=(10, 5)
        )
        current_row += 1

        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)

        ttk.Entry(output_frame, textvariable=self.output_file_path, width=50).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(output_frame, text="Browse...", command=self.browse_output_file).grid(
            row=0, column=1
        )
        current_row += 1

        # Settings Section
        settings_frame = ttk.LabelFrame(main_frame, text="Projection Settings", padding="10")
        settings_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 10))
        settings_frame.columnconfigure(1, weight=1)
        current_row += 1

        # Number of projection years
        ttk.Label(settings_frame, text="Number of Projection Years:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        ttk.Spinbox(
            settings_frame,
            from_=1,
            to=10,
            textvariable=self.num_projection_years,
            width=10
        ).grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # Growth rate options
        ttk.Label(settings_frame, text="Revenue Growth Rate:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )

        growth_frame = ttk.Frame(settings_frame)
        growth_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)

        ttk.Radiobutton(
            growth_frame,
            text="Auto (calculate from historical)",
            variable=self.use_auto_growth,
            value=True,
            command=self.toggle_growth_rate
        ).grid(row=0, column=0, sticky=tk.W)

        manual_frame = ttk.Frame(growth_frame)
        manual_frame.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))

        ttk.Radiobutton(
            manual_frame,
            text="Manual:",
            variable=self.use_auto_growth,
            value=False,
            command=self.toggle_growth_rate
        ).grid(row=0, column=0, sticky=tk.W)

        self.growth_entry = ttk.Entry(manual_frame, textvariable=self.revenue_growth_override, width=10)
        self.growth_entry.grid(row=0, column=1, padx=(5, 5))
        self.growth_entry.config(state='disabled')

        ttk.Label(manual_frame, text="% (e.g., 5 for 5%)").grid(row=0, column=2, sticky=tk.W)

        # Action Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=current_row, column=0, columnspan=3, pady=(20, 10))
        current_row += 1

        self.build_button = ttk.Button(
            button_frame,
            text="Build Financial Model",
            command=self.build_model,
            style="Accent.TButton"
        )
        self.build_button.grid(row=0, column=0, padx=5)

        ttk.Button(
            button_frame,
            text="Clear",
            command=self.clear_form
        ).grid(row=0, column=1, padx=5)

        # Progress/Log Section
        ttk.Label(main_frame, text="Progress Log:", font=("Helvetica", 10, "bold")).grid(
            row=current_row, column=0, sticky=tk.W, pady=(10, 5)
        )
        current_row += 1

        # Progress bar
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))
        current_row += 1

        # Log text area
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(current_row, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=12,
            wrap=tk.WORD,
            font=("Courier", 9)
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.config(state='disabled')

        # Style configuration
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Helvetica", 10, "bold"))

    def toggle_growth_rate(self):
        """Enable/disable manual growth rate entry"""
        if self.use_auto_growth.get():
            self.growth_entry.config(state='disabled')
        else:
            self.growth_entry.config(state='normal')

    def browse_input_file(self):
        """Browse for input file"""
        filename = filedialog.askopenfilename(
            title="Select Capital IQ Export File",
            filetypes=[
                ("Excel files", "*.xls *.xlsx"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.input_file_path.set(filename)

            # Auto-set company name from filename if not already set
            if not self.company_name.get():
                base_name = Path(filename).stem
                # Try to extract company name (remove "Financials" and other common suffixes)
                clean_name = base_name.replace(" Financials", "").replace("_Financials", "")
                self.company_name.set(clean_name)

            # Auto-set output file name
            if not self.output_file_path.get():
                output_dir = Path(filename).parent
                output_name = f"{self.company_name.get()}_Financial_Model.xlsx"
                self.output_file_path.set(str(output_dir / output_name))

    def browse_output_file(self):
        """Browse for output file location"""
        filename = filedialog.asksaveasfilename(
            title="Save Financial Model As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.output_file_path.set(filename)

    def log(self, message):
        """Add message to log"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

    def clear_log(self):
        """Clear the log"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def clear_form(self):
        """Clear all form fields"""
        self.input_file_path.set("")
        self.output_file_path.set("")
        self.company_name.set("")
        self.num_projection_years.set(3)
        self.revenue_growth_override.set("")
        self.use_auto_growth.set(True)
        self.toggle_growth_rate()
        self.clear_log()

    def validate_inputs(self):
        """Validate user inputs"""
        if not self.input_file_path.get():
            messagebox.showerror("Error", "Please select an input file")
            return False

        if not os.path.exists(self.input_file_path.get()):
            messagebox.showerror("Error", "Input file does not exist")
            return False

        if not self.company_name.get():
            messagebox.showerror("Error", "Please enter a company name")
            return False

        if not self.output_file_path.get():
            messagebox.showerror("Error", "Please specify an output file location")
            return False

        if not self.use_auto_growth.get():
            try:
                growth_rate = float(self.revenue_growth_override.get())
                if growth_rate < -100 or growth_rate > 1000:
                    messagebox.showerror("Error", "Growth rate must be between -100% and 1000%")
                    return False
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid number for growth rate")
                return False

        return True

    def build_model(self):
        """Build the financial model"""
        if not self.validate_inputs():
            return

        # Disable build button during processing
        self.build_button.config(state='disabled')
        self.clear_log()

        # Run in separate thread to keep UI responsive
        thread = threading.Thread(target=self._build_model_thread)
        thread.daemon = True
        thread.start()

    def _build_model_thread(self):
        """Build model in separate thread"""
        try:
            self.progress_bar.start()

            self.log("=" * 70)
            self.log("Financial Model Builder")
            self.log("=" * 70)
            self.log("")

            # Extract data
            self.log(f"1. Reading Capital IQ export: {Path(self.input_file_path.get()).name}")
            extractor = CapIQDataExtractor(self.input_file_path.get())

            self.log("2. Extracting Income Statement...")
            is_df = extractor.extract_income_statement()
            if is_df is not None:
                self.log(f"   ✓ Found {len(is_df)} line items across {len(is_df.columns)} periods")
            else:
                raise ValueError("Failed to extract Income Statement data")

            self.log("3. Extracting Balance Sheet...")
            bs_df = extractor.extract_balance_sheet()
            if bs_df is not None:
                self.log(f"   ✓ Found {len(bs_df)} line items across {len(bs_df.columns)} periods")
            else:
                raise ValueError("Failed to extract Balance Sheet data")

            self.log("4. Extracting Cash Flow Statement...")
            cf_df = extractor.extract_cash_flow()
            if cf_df is not None:
                self.log(f"   ✓ Found {len(cf_df)} line items across {len(cf_df.columns)} periods")
            else:
                raise ValueError("Failed to extract Cash Flow data")

            self.log("")
            self.log("5. Building financial model with projections...")

            # Build model
            model_builder = FinancialModelBuilder(
                self.company_name.get(),
                is_df,
                bs_df,
                cf_df
            )

            # Override growth rate if manual
            if not self.use_auto_growth.get():
                manual_growth = float(self.revenue_growth_override.get()) / 100
                model_builder.calculate_growth_rates = lambda: {
                    'revenue': manual_growth,
                    'operating_margin': model_builder.calculate_growth_rates.__func__(model_builder).get('operating_margin', 0.15)
                }

            model_data = model_builder.build_model()

            # Update projection years based on user input
            num_years = self.num_projection_years.get()
            historical_years = model_data['historical_years']
            last_year = int(historical_years[-1])
            projection_years = [str(last_year + i + 1) for i in range(num_years)]
            model_data['projection_years'] = projection_years

            # Re-project with correct number of years
            is_projected, _ = model_builder.project_financials(num_years=num_years)
            model_data['income_statement'] = is_projected

            self.log(f"   ✓ Historical years: {', '.join(model_data['historical_years'])}")
            self.log(f"   ✓ Projection years: {', '.join(model_data['projection_years'])}")
            self.log(f"   ✓ Revenue growth rate: {model_data['growth_rates'].get('revenue', 0)*100:.1f}%")
            self.log("")

            # Format and save
            self.log(f"6. Formatting Excel output...")
            formatter = ExcelFormatter(self.company_name.get(), model_data)
            formatter.save(self.output_file_path.get())

            self.log("")
            self.log("=" * 70)
            self.log("✓ SUCCESS! Financial model created successfully.")
            self.log("=" * 70)
            self.log(f"Output saved to: {self.output_file_path.get()}")

            # Show success message
            self.root.after(0, lambda: messagebox.showinfo(
                "Success",
                f"Financial model created successfully!\n\nSaved to:\n{self.output_file_path.get()}"
            ))

        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self.log("")
            self.log("=" * 70)
            self.log(f"✗ ERROR: {error_msg}")
            self.log("=" * 70)

            self.root.after(0, lambda: messagebox.showerror(
                "Error",
                f"Failed to build financial model:\n\n{error_msg}"
            ))

        finally:
            self.progress_bar.stop()
            self.root.after(0, lambda: self.build_button.config(state='normal'))


def main():
    """Main entry point"""
    root = tk.Tk()
    app = FinancialModelBuilderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
