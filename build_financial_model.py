#!/usr/bin/env python3
"""
Financial Model Builder
Reads Capital IQ Excel exports and creates a 3-statement financial model
"""
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import re
import warnings
warnings.filterwarnings('ignore')


class CapIQDataExtractor:
    """Extract financial data from Capital IQ exports"""

    def __init__(self, filepath):
        self.filepath = filepath
        self.excel_file = pd.ExcelFile(filepath)

    def _find_data_start_row(self, df, sheet_name):
        """Find the row where actual data starts"""
        for idx, row in df.iterrows():
            # Look for period headers (contains fiscal period info)
            if pd.notna(row[0]) and 'For the Fiscal Period' in str(row[0]):
                return idx
            if pd.notna(row[0]) and 'Balance Sheet as of' in str(row[0]):
                return idx
        return None

    def _extract_periods(self, header_row):
        """Extract period information from header row"""
        periods = []
        for col_val in header_row:
            if pd.notna(col_val):
                str_val = str(col_val)
                # Look for year patterns like "Dec-31-2021" or "2021"
                if 'Dec-' in str_val or 'Sep-' in str_val:
                    # Extract 4-digit year from date strings
                    year_match = re.search(r'(20\d{2})', str_val)
                    if year_match:
                        year = year_match.group(1)
                        # Skip LTM periods if desired, or include them
                        if 'LTM' not in str_val:
                            periods.append(year)
        return periods

    def extract_income_statement(self):
        """Extract income statement data"""
        df = pd.read_excel(self.excel_file, sheet_name='Income Statement', header=None)

        # Find where data starts
        data_start_row = self._find_data_start_row(df, 'Income Statement')
        if data_start_row is None:
            return None

        # Extract periods from header row
        periods = self._extract_periods(df.iloc[data_start_row])

        # Extract data rows
        data_dict = {}
        for idx in range(data_start_row + 2, len(df)):  # Skip header and currency rows
            row = df.iloc[idx]
            line_item = str(row[0]).strip() if pd.notna(row[0]) else ''

            if line_item and line_item != 'nan':
                values = []
                for col_idx in range(1, len(periods) + 1):
                    val = row[col_idx]
                    # Handle various non-numeric values
                    if pd.notna(val) and val not in ['-', 'NM', 'NA', 'N/A']:
                        try:
                            values.append(float(val))
                        except (ValueError, TypeError):
                            values.append(0)
                    else:
                        values.append(0)

                if len(values) == len(periods):
                    data_dict[line_item] = values

        # Create DataFrame
        is_df = pd.DataFrame(data_dict, index=periods).T
        return is_df

    def extract_balance_sheet(self):
        """Extract balance sheet data"""
        df = pd.read_excel(self.excel_file, sheet_name='Balance Sheet', header=None)

        # Find where data starts
        data_start_row = None
        for idx, row in df.iterrows():
            if pd.notna(row[0]) and 'Balance Sheet as of' in str(row[0]):
                data_start_row = idx
                break

        if data_start_row is None:
            return None

        # Extract periods (years) from header
        periods = []
        header_row = df.iloc[data_start_row]
        for col_val in header_row[1:]:
            if pd.notna(col_val):
                date_str = str(col_val)
                if '-' in date_str:
                    year = date_str.split('-')[0]
                    periods.append(year)

        # Extract data rows
        data_dict = {}
        for idx in range(data_start_row + 2, len(df)):
            row = df.iloc[idx]
            line_item = str(row[0]).strip() if pd.notna(row[0]) else ''

            if line_item and line_item != 'nan':
                values = []
                for col_idx in range(1, len(periods) + 1):
                    val = row[col_idx]
                    # Handle various non-numeric values
                    if pd.notna(val) and val not in ['-', 'NM', 'NA', 'N/A']:
                        try:
                            values.append(float(val))
                        except (ValueError, TypeError):
                            values.append(0)
                    else:
                        values.append(0)

                if len(values) == len(periods):
                    data_dict[line_item] = values

        # Create DataFrame
        bs_df = pd.DataFrame(data_dict, index=periods).T
        return bs_df

    def extract_cash_flow(self):
        """Extract cash flow data"""
        df = pd.read_excel(self.excel_file, sheet_name='Cash Flow', header=None)

        # Find where data starts
        data_start_row = self._find_data_start_row(df, 'Cash Flow')
        if data_start_row is None:
            return None

        # Extract periods from header row
        periods = self._extract_periods(df.iloc[data_start_row])

        # Extract data rows
        data_dict = {}
        for idx in range(data_start_row + 2, len(df)):
            row = df.iloc[idx]
            line_item = str(row[0]).strip() if pd.notna(row[0]) else ''

            if line_item and line_item != 'nan':
                values = []
                for col_idx in range(1, len(periods) + 1):
                    val = row[col_idx]
                    # Handle various non-numeric values
                    if pd.notna(val) and val not in ['-', 'NM', 'NA', 'N/A']:
                        try:
                            values.append(float(val))
                        except (ValueError, TypeError):
                            values.append(0)
                    else:
                        values.append(0)

                if len(values) == len(periods):
                    data_dict[line_item] = values

        # Create DataFrame
        cf_df = pd.DataFrame(data_dict, index=periods).T
        return cf_df


class FinancialModelBuilder:
    """Build 3-statement financial model with projections"""

    def __init__(self, company_name, is_df, bs_df, cf_df):
        self.company_name = company_name
        self.is_df = is_df
        self.bs_df = bs_df
        self.cf_df = cf_df

        # Extract historical years (keep only valid years, exclude LTM)
        self.historical_years = [y for y in is_df.columns if len(y) == 4 and y.isdigit()]

    def calculate_growth_rates(self):
        """Calculate historical growth rates for projections"""
        growth_rates = {}

        # Revenue growth rate (average of last 3 years)
        if 'Total Revenue' in self.is_df.index:
            revenue_values = [self.is_df.loc['Total Revenue', year]
                            for year in self.historical_years if year in self.is_df.columns]

            if len(revenue_values) >= 2:
                growth_rates_list = []
                for i in range(1, len(revenue_values)):
                    if revenue_values[i-1] != 0:
                        growth = (revenue_values[i] - revenue_values[i-1]) / revenue_values[i-1]
                        growth_rates_list.append(growth)

                growth_rates['revenue'] = np.mean(growth_rates_list[-3:]) if growth_rates_list else 0.05
            else:
                growth_rates['revenue'] = 0.05

        # Operating margin (average of last 3 years)
        if 'Total Revenue' in self.is_df.index and 'Operating Income' in self.is_df.index:
            margins = []
            for year in self.historical_years[-3:]:
                if year in self.is_df.columns:
                    revenue = self.is_df.loc['Total Revenue', year]
                    op_income = self.is_df.loc['Operating Income', year]
                    if revenue != 0:
                        margins.append(op_income / revenue)

            growth_rates['operating_margin'] = np.mean(margins) if margins else 0.15

        return growth_rates

    def project_financials(self, num_years=3):
        """Project financial statements"""
        growth_rates = self.calculate_growth_rates()

        last_year = int(self.historical_years[-1])
        projection_years = [str(last_year + i + 1) for i in range(num_years)]

        # Project Income Statement
        is_projected = self.is_df.copy()

        for year in projection_years:
            prev_year = str(int(year) - 1)

            if prev_year in is_projected.columns:
                # Revenue projection
                if 'Total Revenue' in is_projected.index:
                    prev_revenue = is_projected.loc['Total Revenue', prev_year]
                    is_projected.loc['Total Revenue', year] = prev_revenue * (1 + growth_rates.get('revenue', 0.05))

                # Project other line items as % of revenue
                if 'Total Revenue' in is_projected.index:
                    revenue = is_projected.loc['Total Revenue', year]

                    # COGS (maintain historical margin)
                    if 'COGS' in is_projected.index and 'Total Revenue' in is_projected.index:
                        hist_cogs_pct = self.is_df.loc['COGS', self.historical_years[-1]] / self.is_df.loc['Total Revenue', self.historical_years[-1]]
                        is_projected.loc['COGS', year] = revenue * hist_cogs_pct

                    # Gross Profit
                    if 'Gross Profit' in is_projected.index:
                        total_rev = is_projected.loc['Total Revenue', year] if 'Total Revenue' in is_projected.index else 0
                        cogs = is_projected.loc['COGS', year] if 'COGS' in is_projected.index else 0
                        is_projected.loc['Gross Profit', year] = total_rev - cogs

                    # Operating expenses (as % of revenue)
                    for item in ['SG&A', 'R&D', 'Operating Expenses']:
                        if item in is_projected.index:
                            hist_pct = self.is_df.loc[item, self.historical_years[-1]] / self.is_df.loc['Total Revenue', self.historical_years[-1]]
                            is_projected.loc[item, year] = revenue * hist_pct

                    # Operating Income
                    if 'Operating Income' in is_projected.index:
                        gross_profit = is_projected.loc['Gross Profit', year] if 'Gross Profit' in is_projected.index else 0
                        op_exp = is_projected.loc['Operating Expenses', year] if 'Operating Expenses' in is_projected.index else 0
                        is_projected.loc['Operating Income', year] = gross_profit - op_exp

                    # Interest and taxes (as % of historical)
                    for item in ['Interest Expense', 'Interest and Other, net']:
                        if item in is_projected.index:
                            is_projected.loc[item, year] = is_projected.loc[item, prev_year]

                    # Pre-tax income
                    if 'Pretax Income' in is_projected.index:
                        op_income = is_projected.loc['Operating Income', year] if 'Operating Income' in is_projected.index else 0
                        interest = is_projected.loc['Interest and Other, net', year] if 'Interest and Other, net' in is_projected.index else 0
                        is_projected.loc['Pretax Income', year] = op_income + interest

                    # Income tax (maintain historical rate)
                    if 'Income Tax Expense' in is_projected.index and 'Pretax Income' in is_projected.index:
                        hist_tax_rate = abs(self.is_df.loc['Income Tax Expense', self.historical_years[-1]] /
                                          self.is_df.loc['Pretax Income', self.historical_years[-1]]) if self.is_df.loc['Pretax Income', self.historical_years[-1]] != 0 else 0.21
                        is_projected.loc['Income Tax Expense', year] = is_projected.loc['Pretax Income', year] * hist_tax_rate

                    # Net Income
                    if 'Net Income' in is_projected.index:
                        pretax = is_projected.loc['Pretax Income', year] if 'Pretax Income' in is_projected.index else 0
                        tax = is_projected.loc['Income Tax Expense', year] if 'Income Tax Expense' in is_projected.index else 0
                        is_projected.loc['Net Income', year] = pretax - tax

        return is_projected, projection_years

    def build_model(self):
        """Build complete financial model"""
        is_projected, projection_years = self.project_financials()

        return {
            'income_statement': is_projected,
            'balance_sheet': self.bs_df,
            'cash_flow': self.cf_df,
            'historical_years': self.historical_years,
            'projection_years': projection_years,
            'growth_rates': self.calculate_growth_rates()
        }


class ExcelFormatter:
    """Format Excel output to match template style"""

    def __init__(self, company_name, model_data):
        self.company_name = company_name
        self.model_data = model_data
        self.wb = Workbook()

    def _apply_header_style(self, cell):
        """Apply header cell style"""
        cell.font = Font(name='Calibri', size=11, bold=True)
        cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        cell.font = Font(color='FFFFFF', bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    def _apply_section_header_style(self, cell):
        """Apply section header style"""
        cell.font = Font(name='Calibri', size=11, bold=True)
        cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    def _apply_number_format(self, cell, value):
        """Apply number formatting"""
        if isinstance(value, (int, float)):
            cell.number_format = '#,##0.0'
            cell.alignment = Alignment(horizontal='right')

    def create_fsm_sheet(self):
        """Create main FSM (Financial Statement Model) sheet"""
        ws = self.wb.active
        ws.title = 'FSM'

        # Header
        ws['A1'] = f'{self.company_name} - Financial Statement Model'
        ws['A1'].font = Font(name='Calibri', size=14, bold=True)

        ws['A2'] = f'Source: Capital IQ Export | All figures in $ millions | Generated {datetime.now().strftime("%Y-%m-%d")}'
        ws['A2'].font = Font(name='Calibri', size=9, italic=True)

        current_row = 4

        # Key Assumptions Section
        ws[f'A{current_row}'] = 'KEY ASSUMPTIONS'
        self._apply_section_header_style(ws[f'A{current_row}'])
        current_row += 1

        growth_rates = self.model_data['growth_rates']

        ws[f'A{current_row}'] = 'Assumption'
        ws[f'B{current_row}'] = 'Value'
        ws[f'C{current_row}'] = 'Source / Notes'
        for col in ['A', 'B', 'C']:
            self._apply_header_style(ws[f'{col}{current_row}'])
        current_row += 1

        ws[f'A{current_row}'] = 'Revenue Growth Rate'
        ws[f'B{current_row}'] = growth_rates.get('revenue', 0.05)
        ws[f'B{current_row}'].number_format = '0.0%'
        ws[f'C{current_row}'] = 'Average of last 3 years historical growth'
        current_row += 1

        ws[f'A{current_row}'] = 'Operating Margin'
        ws[f'B{current_row}'] = growth_rates.get('operating_margin', 0.15)
        ws[f'B{current_row}'].number_format = '0.0%'
        ws[f'C{current_row}'] = 'Average of last 3 years'
        current_row += 2

        # Column headers for periods
        current_row += 1
        ws[f'A{current_row}'] = 'Fiscal Year'
        col_idx = 2

        # Historical years
        historical_years = self.model_data['historical_years']
        projection_years = self.model_data['projection_years']

        all_years = historical_years + projection_years

        for year in all_years:
            col_letter = get_column_letter(col_idx)
            suffix = 'A' if year in historical_years else 'E'
            ws[f'{col_letter}{current_row}'] = f'FY{year}{suffix}'
            self._apply_header_style(ws[f'{col_letter}{current_row}'])
            col_idx += 1

        current_row += 2

        # Income Statement Section
        ws[f'A{current_row}'] = 'INCOME STATEMENT'
        self._apply_section_header_style(ws[f'A{current_row}'])
        current_row += 1

        is_df = self.model_data['income_statement']

        # Key income statement items to display
        is_items = [
            'Total Revenue',
            'COGS',
            'Gross Profit',
            'Operating Expenses',
            'Operating Income',
            'Interest and Other, net',
            'Pretax Income',
            'Income Tax Expense',
            'Net Income'
        ]

        for item in is_items:
            if item in is_df.index:
                ws[f'A{current_row}'] = item
                col_idx = 2

                for year in all_years:
                    if year in is_df.columns:
                        col_letter = get_column_letter(col_idx)
                        value = is_df.loc[item, year]
                        ws[f'{col_letter}{current_row}'] = value
                        self._apply_number_format(ws[f'{col_letter}{current_row}'], value)
                    col_idx += 1

                current_row += 1

        current_row += 2

        # Balance Sheet Section
        ws[f'A{current_row}'] = 'BALANCE SHEET'
        self._apply_section_header_style(ws[f'A{current_row}'])
        current_row += 1

        bs_df = self.model_data['balance_sheet']

        # Key balance sheet items
        bs_items = [
            'Cash And Equivalents',
            'Total Current Assets',
            'PP&E, Net',
            'Total Assets',
            'Total Current Liabilities',
            'Long-Term Debt',
            'Total Liabilities',
            'Total Equity',
            'Total Liabilities And Equity'
        ]

        for item in bs_items:
            if item in bs_df.index:
                ws[f'A{current_row}'] = item
                col_idx = 2

                # Only show historical years for balance sheet
                for year in historical_years:
                    if year in bs_df.columns:
                        col_letter = get_column_letter(col_idx)
                        value = bs_df.loc[item, year]
                        ws[f'{col_letter}{current_row}'] = value
                        self._apply_number_format(ws[f'{col_letter}{current_row}'], value)
                    col_idx += 1

                current_row += 1

        current_row += 2

        # Cash Flow Section
        ws[f'A{current_row}'] = 'CASH FLOW STATEMENT'
        self._apply_section_header_style(ws[f'A{current_row}'])
        current_row += 1

        cf_df = self.model_data['cash_flow']

        # Key cash flow items
        cf_items = [
            'Net Income',
            'Depreciation & Amort.',
            'Change In Working Capital',
            'Cash From Operations',
            'Capital Expenditures',
            'Cash From Investing',
            'Cash From Financing',
            'Net Change In Cash'
        ]

        for item in cf_items:
            if item in cf_df.index:
                ws[f'A{current_row}'] = item
                col_idx = 2

                # Only show historical years for cash flow
                for year in historical_years:
                    if year in cf_df.columns:
                        col_letter = get_column_letter(col_idx)
                        value = cf_df.loc[item, year]
                        ws[f'{col_letter}{current_row}'] = value
                        self._apply_number_format(ws[f'{col_letter}{current_row}'], value)
                    col_idx += 1

                current_row += 1

        # Adjust column widths
        ws.column_dimensions['A'].width = 35
        for col_idx in range(2, len(all_years) + 2):
            ws.column_dimensions[get_column_letter(col_idx)].width = 12

    def save(self, output_path):
        """Save workbook"""
        self.create_fsm_sheet()
        self.wb.save(output_path)
        print(f"Financial model saved to: {output_path}")


def main():
    """Main execution function"""
    print("=" * 80)
    print("Financial Model Builder")
    print("=" * 80)

    # Input file
    input_file = 'Boston Scientific Corporation NYSE BSX Financials.xls'
    company_name = 'Boston Scientific Corporation (NYSE:BSX)'

    # Output file
    output_file = 'Boston_Scientific_Financial_Model.xlsx'

    print(f"\n1. Reading Capital IQ export: {input_file}")
    extractor = CapIQDataExtractor(input_file)

    print("2. Extracting Income Statement...")
    is_df = extractor.extract_income_statement()
    if is_df is not None:
        print(f"   - Found {len(is_df)} line items across {len(is_df.columns)} periods")

    print("3. Extracting Balance Sheet...")
    bs_df = extractor.extract_balance_sheet()
    if bs_df is not None:
        print(f"   - Found {len(bs_df)} line items across {len(bs_df.columns)} periods")

    print("4. Extracting Cash Flow Statement...")
    cf_df = extractor.extract_cash_flow()
    if cf_df is not None:
        print(f"   - Found {len(cf_df)} line items across {len(cf_df.columns)} periods")

    print("\n5. Building financial model with projections...")
    model_builder = FinancialModelBuilder(company_name, is_df, bs_df, cf_df)
    model_data = model_builder.build_model()

    print(f"   - Historical years: {', '.join(model_data['historical_years'])}")
    print(f"   - Projection years: {', '.join(model_data['projection_years'])}")
    print(f"   - Revenue growth rate: {model_data['growth_rates'].get('revenue', 0)*100:.1f}%")

    print(f"\n6. Formatting Excel output...")
    formatter = ExcelFormatter(company_name, model_data)
    formatter.save(output_file)

    print("\n" + "=" * 80)
    print("SUCCESS! Financial model created successfully.")
    print("=" * 80)


if __name__ == "__main__":
    main()
