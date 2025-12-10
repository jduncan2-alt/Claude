#!/usr/bin/env python3
"""
Financial Model Builder - Web GUI
Modern web-based interface for creating 3-statement financial models
"""
from flask import Flask, render_template_string, request, jsonify, send_file
import os
import sys
from pathlib import Path
import traceback
from werkzeug.utils import secure_filename

# Import the core functionality
from build_financial_model import CapIQDataExtractor, FinancialModelBuilder, ExcelFormatter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = '/tmp/financial_models'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Financial Model Builder</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }

        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 700;
        }

        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }

        .content {
            padding: 40px;
        }

        .form-group {
            margin-bottom: 25px;
        }

        label {
            display: block;
            font-weight: 600;
            margin-bottom: 8px;
            color: #333;
            font-size: 0.95em;
        }

        input[type="text"],
        input[type="number"],
        input[type="file"] {
            width: 100%;
            padding: 12px 15px;
            border: 2px solid #e1e8ed;
            border-radius: 8px;
            font-size: 1em;
            transition: all 0.3s ease;
        }

        input[type="text"]:focus,
        input[type="number"]:focus,
        input[type="file"]:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }

        .file-input-wrapper {
            position: relative;
            overflow: hidden;
            display: inline-block;
            width: 100%;
        }

        .file-input-wrapper input[type=file] {
            position: absolute;
            left: 0;
            top: 0;
            opacity: 0;
            cursor: pointer;
            width: 100%;
            height: 100%;
        }

        .file-input-button {
            display: inline-block;
            padding: 12px 20px;
            background: #667eea;
            color: white;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.3s ease;
        }

        .file-input-button:hover {
            background: #5568d3;
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.3);
        }

        .file-name {
            display: inline-block;
            margin-left: 15px;
            color: #666;
            font-style: italic;
        }

        .settings-group {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 25px;
        }

        .settings-group h3 {
            margin-bottom: 15px;
            color: #333;
            font-size: 1.1em;
        }

        .radio-group {
            margin: 10px 0;
        }

        .radio-group label {
            display: inline-block;
            font-weight: normal;
            margin-right: 20px;
            cursor: pointer;
        }

        .radio-group input[type="radio"] {
            margin-right: 5px;
        }

        .manual-growth {
            display: none;
            margin-top: 10px;
            padding-left: 25px;
        }

        .manual-growth.active {
            display: block;
        }

        .btn {
            padding: 15px 40px;
            font-size: 1.1em;
            font-weight: 600;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s ease;
            display: inline-block;
            text-align: center;
        }

        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }

        .btn-primary:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 10px 25px rgba(102, 126, 234, 0.4);
        }

        .btn-primary:disabled {
            opacity: 0.6;
            cursor: not-allowed;
        }

        .btn-secondary {
            background: #e1e8ed;
            color: #333;
            margin-left: 10px;
        }

        .btn-secondary:hover {
            background: #cbd5dd;
        }

        .button-group {
            margin-top: 30px;
            text-align: center;
        }

        .progress-section {
            margin-top: 30px;
            display: none;
        }

        .progress-section.active {
            display: block;
        }

        .progress-bar {
            width: 100%;
            height: 6px;
            background: #e1e8ed;
            border-radius: 10px;
            overflow: hidden;
            margin-bottom: 15px;
        }

        .progress-bar-fill {
            height: 100%;
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
            width: 0%;
            transition: width 0.3s ease;
            animation: progress-animation 1.5s ease-in-out infinite;
        }

        @keyframes progress-animation {
            0% { width: 0%; }
            50% { width: 70%; }
            100% { width: 100%; }
        }

        .log-container {
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 20px;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 0.9em;
            max-height: 300px;
            overflow-y: auto;
            line-height: 1.6;
        }

        .log-line {
            margin-bottom: 5px;
        }

        .log-line.success {
            color: #4ec9b0;
        }

        .log-line.error {
            color: #f48771;
        }

        .log-line.separator {
            color: #608b4e;
        }

        .result-section {
            margin-top: 20px;
            padding: 20px;
            background: #d4edda;
            border: 1px solid #c3e6cb;
            border-radius: 8px;
            display: none;
        }

        .result-section.active {
            display: block;
        }

        .result-section h3 {
            color: #155724;
            margin-bottom: 10px;
        }

        .result-section p {
            color: #155724;
            margin-bottom: 10px;
        }

        .download-btn {
            background: #28a745;
            color: white;
            padding: 12px 30px;
            text-decoration: none;
            border-radius: 8px;
            display: inline-block;
            font-weight: 600;
            margin-top: 10px;
            transition: all 0.3s ease;
        }

        .download-btn:hover {
            background: #218838;
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(40, 167, 69, 0.3);
        }

        .error-section {
            margin-top: 20px;
            padding: 20px;
            background: #f8d7da;
            border: 1px solid #f5c6cb;
            border-radius: 8px;
            display: none;
        }

        .error-section.active {
            display: block;
        }

        .error-section h3 {
            color: #721c24;
            margin-bottom: 10px;
        }

        .error-section p {
            color: #721c24;
        }

        .info-box {
            background: #e7f3ff;
            border-left: 4px solid #2196F3;
            padding: 15px;
            margin-bottom: 25px;
            border-radius: 4px;
        }

        .info-box p {
            color: #0c5460;
            line-height: 1.6;
        }

        @media (max-width: 768px) {
            .container {
                margin: 10px;
            }

            .header h1 {
                font-size: 1.8em;
            }

            .content {
                padding: 20px;
            }

            .btn {
                width: 100%;
                margin-bottom: 10px;
            }

            .btn-secondary {
                margin-left: 0;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Financial Model Builder</h1>
            <p>Create 3-statement financial models from Capital IQ exports</p>
        </div>

        <div class="content">
            <div class="info-box">
                <p><strong>How it works:</strong> Upload your Capital IQ Excel export file, and we'll automatically extract historical financials and create projections for the Income Statement, Balance Sheet, and Cash Flow Statement.</p>
            </div>

            <form id="modelForm" enctype="multipart/form-data">
                <div class="form-group">
                    <label for="inputFile">Capital IQ Export File *</label>
                    <div class="file-input-wrapper">
                        <label class="file-input-button">
                            Choose File
                            <input type="file" id="inputFile" name="inputFile" accept=".xls,.xlsx" required>
                        </label>
                        <span class="file-name" id="fileName">No file chosen</span>
                    </div>
                </div>

                <div class="form-group">
                    <label for="companyName">Company Name *</label>
                    <input type="text" id="companyName" name="companyName" placeholder="e.g., Boston Scientific Corporation (NYSE:BSX)" required>
                </div>

                <div class="settings-group">
                    <h3>Projection Settings</h3>

                    <div class="form-group">
                        <label for="numYears">Number of Projection Years</label>
                        <input type="number" id="numYears" name="numYears" value="3" min="1" max="10" required>
                    </div>

                    <div class="form-group">
                        <label>Revenue Growth Rate</label>
                        <div class="radio-group">
                            <label>
                                <input type="radio" name="growthType" value="auto" checked>
                                Auto (calculate from historical data)
                            </label>
                        </div>
                        <div class="radio-group">
                            <label>
                                <input type="radio" name="growthType" value="manual">
                                Manual
                            </label>
                        </div>
                        <div class="manual-growth" id="manualGrowth">
                            <label for="manualGrowthRate">Growth Rate (%)</label>
                            <input type="number" id="manualGrowthRate" name="manualGrowthRate" step="0.1" placeholder="e.g., 5 for 5%">
                        </div>
                    </div>
                </div>

                <div class="button-group">
                    <button type="submit" class="btn btn-primary" id="buildBtn">
                        Build Financial Model
                    </button>
                    <button type="button" class="btn btn-secondary" onclick="resetForm()">
                        Clear
                    </button>
                </div>
            </form>

            <div class="progress-section" id="progressSection">
                <h3 style="margin-bottom: 15px;">Building Model...</h3>
                <div class="progress-bar">
                    <div class="progress-bar-fill"></div>
                </div>
                <div class="log-container" id="logContainer"></div>
            </div>

            <div class="result-section" id="resultSection">
                <h3>✅ Success!</h3>
                <p>Your financial model has been created successfully.</p>
                <a href="#" class="download-btn" id="downloadBtn">Download Financial Model</a>
            </div>

            <div class="error-section" id="errorSection">
                <h3>❌ Error</h3>
                <p id="errorMessage"></p>
            </div>
        </div>
    </div>

    <script>
        // File input handler
        document.getElementById('inputFile').addEventListener('change', function(e) {
            const fileName = e.target.files[0]?.name || 'No file chosen';
            document.getElementById('fileName').textContent = fileName;

            // Auto-fill company name if empty
            if (!document.getElementById('companyName').value && fileName) {
                const cleanName = fileName.replace(/\\.(xls|xlsx)$/i, '')
                                         .replace(/ Financials$/i, '')
                                         .replace(/_Financials$/i, '');
                document.getElementById('companyName').value = cleanName;
            }
        });

        // Growth rate type toggle
        document.querySelectorAll('input[name="growthType"]').forEach(radio => {
            radio.addEventListener('change', function() {
                const manualDiv = document.getElementById('manualGrowth');
                if (this.value === 'manual') {
                    manualDiv.classList.add('active');
                    document.getElementById('manualGrowthRate').required = true;
                } else {
                    manualDiv.classList.remove('active');
                    document.getElementById('manualGrowthRate').required = false;
                }
            });
        });

        // Form submission
        document.getElementById('modelForm').addEventListener('submit', async function(e) {
            e.preventDefault();

            // Reset UI
            document.getElementById('progressSection').classList.add('active');
            document.getElementById('resultSection').classList.remove('active');
            document.getElementById('errorSection').classList.remove('active');
            document.getElementById('logContainer').innerHTML = '';
            document.getElementById('buildBtn').disabled = true;

            // Prepare form data
            const formData = new FormData(this);

            try {
                const response = await fetch('/build', {
                    method: 'POST',
                    body: formData
                });

                const result = await response.json();

                if (result.success) {
                    // Show logs
                    if (result.logs) {
                        result.logs.forEach(log => {
                            addLog(log);
                        });
                    }

                    // Show success
                    document.getElementById('resultSection').classList.add('active');
                    document.getElementById('downloadBtn').href = '/download/' + result.filename;

                } else {
                    // Show error
                    document.getElementById('errorSection').classList.add('active');
                    document.getElementById('errorMessage').textContent = result.error;

                    // Show logs if available
                    if (result.logs) {
                        result.logs.forEach(log => {
                            addLog(log);
                        });
                    }
                }

            } catch (error) {
                document.getElementById('errorSection').classList.add('active');
                document.getElementById('errorMessage').textContent = 'Network error: ' + error.message;
            } finally {
                document.getElementById('buildBtn').disabled = false;
                document.getElementById('progressSection').classList.remove('active');
            }
        });

        function addLog(message) {
            const logContainer = document.getElementById('logContainer');
            const logLine = document.createElement('div');
            logLine.className = 'log-line';

            if (message.includes('✓') || message.includes('SUCCESS')) {
                logLine.classList.add('success');
            } else if (message.includes('Error') || message.includes('ERROR') || message.includes('✗')) {
                logLine.classList.add('error');
            } else if (message.includes('===')) {
                logLine.classList.add('separator');
            }

            logLine.textContent = message;
            logContainer.appendChild(logLine);
            logContainer.scrollTop = logContainer.scrollHeight;
        }

        function resetForm() {
            document.getElementById('modelForm').reset();
            document.getElementById('fileName').textContent = 'No file chosen';
            document.getElementById('progressSection').classList.remove('active');
            document.getElementById('resultSection').classList.remove('active');
            document.getElementById('errorSection').classList.remove('active');
            document.getElementById('manualGrowth').classList.remove('active');
        }
    </script>
</body>
</html>
"""


@app.route('/')
def index():
    """Render the main page"""
    return render_template_string(HTML_TEMPLATE)


@app.route('/build', methods=['POST'])
def build_model():
    """Build the financial model"""
    logs = []

    def log(message):
        logs.append(message)

    try:
        # Get form data
        if 'inputFile' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'})

        file = request.files['inputFile']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})

        company_name = request.form.get('companyName', '')
        if not company_name:
            return jsonify({'success': False, 'error': 'Company name is required'})

        num_years = int(request.form.get('numYears', 3))
        growth_type = request.form.get('growthType', 'auto')
        manual_growth_rate = request.form.get('manualGrowthRate', '')

        # Save uploaded file
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)

        log("=" * 70)
        log("Financial Model Builder")
        log("=" * 70)
        log("")

        # Extract data
        log(f"1. Reading Capital IQ export: {filename}")
        extractor = CapIQDataExtractor(input_path)

        log("2. Extracting Income Statement...")
        is_df = extractor.extract_income_statement()
        if is_df is not None:
            log(f"   ✓ Found {len(is_df)} line items across {len(is_df.columns)} periods")
        else:
            raise ValueError("Failed to extract Income Statement data")

        log("3. Extracting Balance Sheet...")
        bs_df = extractor.extract_balance_sheet()
        if bs_df is not None:
            log(f"   ✓ Found {len(bs_df)} line items across {len(bs_df.columns)} periods")
        else:
            raise ValueError("Failed to extract Balance Sheet data")

        log("4. Extracting Cash Flow Statement...")
        cf_df = extractor.extract_cash_flow()
        if cf_df is not None:
            log(f"   ✓ Found {len(cf_df)} line items across {len(cf_df.columns)} periods")
        else:
            raise ValueError("Failed to extract Cash Flow data")

        log("")
        log("5. Building financial model with projections...")

        # Build model
        model_builder = FinancialModelBuilder(company_name, is_df, bs_df, cf_df)

        # Override growth rate if manual
        if growth_type == 'manual' and manual_growth_rate:
            manual_growth = float(manual_growth_rate) / 100
            original_calc = model_builder.calculate_growth_rates

            def custom_growth():
                rates = original_calc()
                rates['revenue'] = manual_growth
                return rates

            model_builder.calculate_growth_rates = custom_growth

        # Project financials
        is_projected, projection_years = model_builder.project_financials(num_years=num_years)

        model_data = {
            'income_statement': is_projected,
            'balance_sheet': bs_df,
            'cash_flow': cf_df,
            'historical_years': model_builder.historical_years,
            'projection_years': projection_years,
            'growth_rates': model_builder.calculate_growth_rates()
        }

        log(f"   ✓ Historical years: {', '.join(model_data['historical_years'])}")
        log(f"   ✓ Projection years: {', '.join(model_data['projection_years'])}")
        log(f"   ✓ Revenue growth rate: {model_data['growth_rates'].get('revenue', 0)*100:.1f}%")
        log("")

        # Format and save
        log(f"6. Formatting Excel output...")
        output_filename = f"{company_name.replace(' ', '_')}_Financial_Model.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        formatter = ExcelFormatter(company_name, model_data)
        formatter.save(output_path)

        log("")
        log("=" * 70)
        log("✓ SUCCESS! Financial model created successfully.")
        log("=" * 70)

        return jsonify({
            'success': True,
            'logs': logs,
            'filename': output_filename
        })

    except Exception as e:
        log("")
        log("=" * 70)
        log(f"✗ ERROR: {str(e)}")
        log("=" * 70)

        return jsonify({
            'success': False,
            'error': str(e),
            'logs': logs
        })


@app.route('/download/<filename>')
def download_file(filename):
    """Download the generated file"""
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(filename))
    return send_file(filepath, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    print("=" * 80)
    print("Financial Model Builder - Web Interface")
    print("=" * 80)
    print("\nStarting server...")
    print("Open your browser and navigate to: http://localhost:5000")
    print("\nPress Ctrl+C to stop the server")
    print("=" * 80)
    app.run(debug=True, host='0.0.0.0', port=5000)
