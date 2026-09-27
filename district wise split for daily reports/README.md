# IFSC District Splitter

A Python Flask web application that processes Excel files containing bank branch data and splits them by Gujarat districts using IFSC codes with pincode fallback.

## Features

- 📁 **Excel File Upload** - Support for .xlsx, .xls, and .csv files
- 🔍 **IFSC Code Lookup** - Retrieves district information from Razorpay IFSC API
- 📍 **Pincode Fallback** - Falls back to Indian Postal Code API if IFSC district not found or not from Gujarat
- 📦 **ZIP File Generation** - Creates district-wise ZIP files containing CSV data
- 🎯 **Gujarat Districts Only** - Validates and filters for valid Gujarat districts
- 📊 **Processing Summary** - Displays statistics about processed records
- 💾 **Download Management** - View and download all processed ZIP files

## Prerequisites

- Python 3.8+
- pip (Python package manager)

## Installation

1. Clone or extract this project:
```bash
cd ifsc-district-splite
```

2. Create a virtual environment (optional but recommended):
```bash
python -m venv venv
# On Windows
venv\Scripts\activate
# On macOS/Linux
source venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Project Structure

```
ifsc-district-splite/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Web interface
├── uploads/              # Temporary storage for uploaded files (auto-created)
├── downloads/            # Storage for generated ZIP files (auto-created)
└── README.md             # This file
```

## Usage

1. **Start the application:**
```bash
python app.py
```

The app will run on `http://localhost:5000`

2. **Open in browser:**
Navigate to `http://localhost:5000` in your web browser

3. **Upload Excel File:**
   - Click "Choose File" or drag & drop your Excel file
   - File should contain bank branch data with:
     - **IFSC** (or ifsc) - IFSC code for lookup
     - **Pincode** (or pincode/PIN) - Pincode as fallback
     - Any other relevant columns (will be preserved in output)

4. **Process File:**
   - Click "Process File" button
   - App will:
     - Look up district from IFSC code
     - Validate it's a Gujarat district
     - Fall back to pincode if needed
     - Group records by district
     - Create ZIP files

5. **Download Results:**
   - Master ZIP file contains individual district ZIP files
   - Each district ZIP contains a CSV with that district's data
   - Download from the "Available Downloads" section

## Excel File Format Example

| IFSC | Pincode | Branch Name | Bank Name | Contact |
|------|---------|-------------|-----------|---------|
| AUBL0002567 | 380001 | Ahmedabad Main | Axis Bank | +91-79-xxxx |
| SBIN0012345 | 380050 | Ahmedabad West | SBI | +91-79-yyyy |
| HDFC0009876 | 395001 | Surat Main | HDFC Bank | +91-261-zzzz |

## How It Works

### District Lookup Process

1. **Primary: IFSC Code Lookup**
   - Queries Razorpay IFSC API with IFSC code
   - Extracts DISTRICT, STATE, and ISO3166 code
   - Validates if district is in Gujarat (ISO3166 = IN-GJ or STATE = GUJARAT)
   - Validates against predefined list of valid Gujarat districts

2. **Fallback: Pincode Lookup**
   - If IFSC lookup fails or district is not from Gujarat
   - Queries Indian Postal Code API with pincode
   - Extracts district information
   - Validates against Gujarat districts list

3. **Grouping & ZIP Creation**
   - Groups validated records by district
   - Creates individual CSV files for each district
   - Packages all district ZIPs into a master ZIP

## API Sources

- **IFSC API**: https://github.com/razorpay/ifsc-api
  - Base URL: `https://ifsc.razorpay.com`
  - Endpoint: `/search?code={IFSC_CODE}`

- **Postal Code API**: https://github.com/nstack-in/indian-postal-code-api
  - Base URL: `https://indian-postal-code-api.herokuapp.com`
  - Endpoint: `/api/pin/find/{PINCODE}`

## Configuration

Edit `app.py` to modify:

- **UPLOAD_FOLDER** - Directory for temporary uploaded files (default: `uploads/`)
- **DOWNLOAD_FOLDER** - Directory for generated ZIP files (default: `downloads/`)
- **MAX_CONTENT_LENGTH** - Maximum file size in bytes (default: 50MB)
- **GUJARAT_DISTRICTS** - List of valid Gujarat districts (predefined)

## Error Handling

The application handles various error scenarios:

- Invalid file types - Only .xlsx, .xls, .csv allowed
- File too large - Maximum 50MB
- API timeouts - Gracefully falls back to pincode
- Invalid IFSC codes - Falls back to pincode
- Non-Gujarat districts - Attempts pincode lookup
- No district found - Records are reported as errors

Errors are logged and displayed in the UI without stopping processing.

## Output Format

### Master ZIP Structure
```
District_Split_20240101_120000.zip
├── AHMEDABAD_data.zip
│   └── AHMEDABAD_data.csv
├── SURAT_data.zip
│   └── SURAT_data.csv
├── RAJKOT_data.zip
│   └── RAJKOT_data.csv
└── ...
```

### CSV Format
Each district's CSV contains:
- All original columns from input file
- Added `District` column with verified district name
- UTF-8 encoding

## Performance Notes

- Processing speed depends on API response times
- For large files (1000+ records), expect 1-5 minutes processing time
- APIs have request rate limits (typically 100-1000 requests/minute)
- Uploaded files are automatically cleaned up after processing

## Troubleshooting

**"File too large" error:**
- Maximum file size is 50MB
- Reduce file size or split into multiple files

**"Could not determine district" errors:**
- Verify IFSC codes are valid
- Verify pincodes are correct
- Check if branch is actually in Gujarat

**API connection errors:**
- Check internet connection
- APIs might be temporarily unavailable
- Try uploading file later

**Port already in use:**
- Change port: `app.run(port=5001)`
- Or kill process using port 5000

## Development

To run in development mode with debugging:

```bash
# Already enabled by default in app.py
python app.py
```

For production deployment, set `debug=False`:
```python
app.run(debug=False, host='0.0.0.0', port=5000)
```

## License

This project uses public APIs from:
- Razorpay IFSC API (MIT License)
- Indian Postal Code API

## Support

For issues or feature requests, please check:
- IFSC API: https://github.com/razorpay/ifsc-api
- Postal Code API: https://github.com/nstack-in/indian-postal-code-api
