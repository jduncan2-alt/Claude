# Financial Model Builder - GUI Version

## Quick Start

There are two GUI options available:

### Option 1: Web-Based GUI (Recommended) ⭐

The web-based GUI works in your browser and is the easiest to use.

**To run:**
```bash
python3 financial_model_builder_web.py
```

Then open your browser and go to: **http://localhost:5000**

You'll see a beautiful web interface where you can:
- Upload your Capital IQ Excel file
- Enter company name
- Set projection parameters
- Build and download your financial model

**Features:**
- Modern, user-friendly interface
- Drag-and-drop file upload
- Real-time progress logs
- Automatic company name detection
- Custom growth rate settings
- Direct download of results

### Option 2: Desktop GUI (Tkinter)

If you prefer a desktop application:

**Prerequisites:**
```bash
# Install tkinter (if not already installed)
sudo apt-get install python3-tk  # On Linux
```

**To run:**
```bash
python3 financial_model_builder_gui.py
```

## Command Line Version

If you prefer the command line:

```bash
python3 build_financial_model.py
```

Edit the `main()` function in `build_financial_model.py` to specify your input file and company name.

## How It Works

1. **Upload**: Select your Capital IQ Excel export file (must include Income Statement, Balance Sheet, and Cash Flow sheets)

2. **Configure**:
   - Enter the company name
   - Choose number of projection years (default: 3)
   - Select automatic or manual growth rate

3. **Build**: Click "Build Financial Model" and watch the progress

4. **Download**: Download your formatted Excel file with 3-statement model

## Example Files

- **Input**: `Boston Scientific Corporation NYSE BSX Financials.xls`
- **Output**: `Boston_Scientific_Financial_Model.xlsx`

## Projection Features

- **Historical Data**: Automatically extracts all available historical periods
- **Revenue Growth**: Calculate automatically from historical trends or set manually
- **Operating Margins**: Maintained at historical averages
- **Income Statement**: Full P&L with projections
- **Balance Sheet**: Historical balance sheet data
- **Cash Flow**: Historical cash flow statement

## Output Format

The generated Excel file includes:
- Key assumptions section
- Historical and projected periods
- Formatted headers and sections
- Professional color scheme matching template style
- Single revenue line (simplified from multi-stream templates)

## Troubleshooting

**"No module named 'tkinter'" error?**
- Use the web-based GUI instead (Option 1)

**"File not found" error?**
- Make sure the Capital IQ export file is in the correct location
- Check that the file has the required sheets: Income Statement, Balance Sheet, Cash Flow

**Import errors?**
- Install required packages: `pip install pandas openpyxl xlrd flask flask-cors`

## Support

For issues or questions, check the log output in the GUI for detailed error messages.
