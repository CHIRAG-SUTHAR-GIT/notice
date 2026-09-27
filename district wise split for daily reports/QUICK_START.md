# Quick Start Guide

Get IFSC District Splitter running in 5 minutes!

## System Requirements

- Windows, macOS, or Linux
- Python 3.8 or higher
- Internet connection (for API calls)
- 50MB free disk space

## Step 1: Verify Python Installation

Open terminal/command prompt and run:

**Windows (Command Prompt):**
```cmd
python --version
```

**macOS/Linux (Terminal):**
```bash
python3 --version
```

Should show Python 3.8+. If not, install from https://www.python.org

## Step 2: Navigate to Project Directory

```bash
cd path/to/ifsc-district-splite
```

## Step 3: One-Click Setup

### Option A: Windows Users
Simply double-click: **`run.bat`**

It will automatically:
- Create required folders
- Start the application
- Open http://localhost:5000 in your browser

### Option B: macOS/Linux Users
Run in terminal:
```bash
chmod +x run.sh
./run.sh
```

### Option C: Manual Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Create folders (if not exists)
mkdir uploads
mkdir downloads
mkdir templates

# Start the app
python app.py
```

## Step 4: Open in Browser

Once you see:
```
Running on http://127.0.0.1:5000
```

Open your browser and go to: **http://localhost:5000**

You should see the upload interface.

## Step 5: Prepare Your Excel File

Create an Excel file with columns:
- **IFSC** - Bank branch IFSC code (e.g., AUBL0002567)
- **Pincode** - Postal code as fallback (e.g., 380001)
- Any other columns you want to keep

### Example Format:

| IFSC | Pincode | Branch_Name | Bank_Name |
|------|---------|------------|-----------|
| AUBL0002567 | 380001 | Ahmedabad Main | Axis Bank |
| HDFC0001234 | 395001 | Surat Branch | HDFC Bank |

## Step 6: Upload and Process

1. Click **"Choose File"** or drag & drop your Excel file
2. Click **"Process File"**
3. Wait for processing to complete

## Step 7: Download Results

Once processing is complete, you'll see:
- Summary of processed records by district
- Download link for ZIP file containing all district-wise data

Click the download link to get your `District_Split_*.zip` file.

## Troubleshooting

### "Python not found"
- Install Python 3.8+ from https://www.python.org
- On Windows: Make sure "Add Python to PATH" is checked during installation

### "Port 5000 already in use"
- Edit `app.py` line: `app.run(port=5001)` (change 5000 to 5001)
- Then access http://localhost:5001

### "No module named 'flask'"
```bash
pip install -r requirements.txt
```

### "No data found / Could not determine district"
- Verify IFSC codes are valid
- Verify pincodes are correct
- Check if branches are actually in Gujarat
- See error list for specific row issues

### File upload fails
- Maximum file size: 50MB
- Supported formats: .xlsx, .xls, .csv
- Make sure file is not corrupted

### Very slow processing
- Large files (1000+ records) take 1-5 minutes
- This is normal (API calls take ~400-1000ms per record)
- Don't close browser during processing

## File Structure After First Run

```
ifsc-district-splite/
├── app.py
├── requirements.txt
├── run.bat / run.sh
├── README.md
├── templates/
│   └── index.html
├── uploads/              (created on first upload)
│   └── [temporary files]
└── downloads/            (created on first upload)
    └── District_Split_*.zip
```

## Understanding the Output

### Downloaded ZIP Structure

```
District_Split_20240115_143025.zip
├── AHMEDABAD_data.zip
│   └── AHMEDABAD_data.csv
├── SURAT_data.zip
│   └── SURAT_data.csv
├── RAJKOT_data.zip
│   └── RAJKOT_data.csv
└── ...
```

Each CSV contains:
- All original columns from your Excel file
- New `District` column with verified district name
- Only records from that district

## Common Use Cases

### Use Case 1: Bank Branch Master Data
**Input:** Excel with all bank branches and IFSC codes
**Output:** Separate CSV files for each district
**Benefit:** Organize branches geographically

### Use Case 2: Customer Database with Pincodes
**Input:** Excel with customer data and pincodes
**Output:** Customer distribution by district
**Benefit:** Regional analysis, marketing, operations

### Use Case 3: Branch Network Analysis
**Input:** Excel with branches, IFSC, and metadata
**Output:** District-wise branch count and details
**Benefit:** Network planning, optimization

## Next Steps

1. **For detailed documentation**, see: [README.md](README.md)
2. **For API details**, see: [API_DOCUMENTATION.md](API_DOCUMENTATION.md)
3. **For configuration**, see: [config.py](config.py)

## Tips for Success

✅ **Do:**
- Use valid, complete IFSC codes (11 characters)
- Include valid Indian pincodes as backup
- Test with small file first (5-10 records)
- Wait for processing to complete (watch for ✓ message)

❌ **Don't:**
- Close browser during processing
- Modify files in upload/download folders manually
- Upload files larger than 50MB
- Use spaces in IFSC codes

## Support Resources

- **IFSC Code Lookup**: https://ifsc.razorpay.com
- **Postal Code Lookup**: https://pincode.in
- **Gujarat Districts**: https://en.wikipedia.org/wiki/Districts_of_Gujarat

## Still Need Help?

1. Check [README.md](README.md) FAQ section
2. Review [API_DOCUMENTATION.md](API_DOCUMENTATION.md)
3. Check browser console (F12 > Console tab) for JavaScript errors
4. Look at server console for Python errors

## Keyboard Shortcuts

- **Ctrl+C** in terminal - Stop the server
- **F5** in browser - Refresh page
- **Ctrl+Shift+Delete** - Clear browser cache if issues persist

---

**Enjoy using IFSC District Splitter!** 🎉
