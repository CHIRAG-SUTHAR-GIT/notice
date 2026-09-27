#!/usr/bin/env python
"""
Quick setup and test script for IFSC District Splitter
"""
import os
import sys
import subprocess

def check_python_version():
    """Check if Python version is 3.8+"""
    if sys.version_info < (3, 8):
        print("❌ Python 3.8+ required!")
        print(f"   Current version: {sys.version}")
        sys.exit(1)
    print(f"✓ Python {sys.version_info.major}.{sys.version_info.minor} detected")

def install_requirements():
    """Install required packages"""
    print("\n📦 Installing requirements...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✓ Requirements installed successfully")
    except subprocess.CalledProcessError:
        print("❌ Failed to install requirements")
        sys.exit(1)

def create_directories():
    """Create necessary directories"""
    print("\n📁 Creating directories...")
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('downloads', exist_ok=True)
    os.makedirs('templates', exist_ok=True)
    print("✓ Directories created")

def create_sample_excel():
    """Create a sample Excel file for testing"""
    print("\n📄 Creating sample Excel file...")
    try:
        import pandas as pd
        
        sample_data = {
            'IFSC': [
                'AUBL0000123',
                'HDFC0001234',
                'SBIN0005678',
                'YESB0000890',
                'ICIC0000345'
            ],
            'Pincode': [
                '380001',
                '395001',
                '361001',
                '302018',
                '400001'
            ],
            'Branch_Name': [
                'Ahmedabad Main',
                'Surat Central',
                'Rajkot Branch',
                'Jaipur Main',
                'Mumbai Main'
            ],
            'Bank_Name': [
                'Axis Bank',
                'HDFC Bank',
                'State Bank of India',
                'Yes Bank',
                'ICICI Bank'
            ]
        }
        
        df = pd.DataFrame(sample_data)
        df.to_excel('sample_input.xlsx', index=False)
        print("✓ Sample file created: sample_input.xlsx")
        print("  (Note: Most of these will show as errors as they're not all Gujarat branches)")
    except Exception as e:
        print(f"⚠ Could not create sample file: {e}")

def main():
    print("=" * 50)
    print("   IFSC District Splitter - Setup")
    print("=" * 50)
    
    check_python_version()
    create_directories()
    
    # Only install if not already in venv or venv exists
    if not os.path.exists('requirements.txt'):
        print("❌ requirements.txt not found!")
        sys.exit(1)
    
    install_requirements()
    create_sample_excel()
    
    print("\n" + "=" * 50)
    print("✓ Setup complete!")
    print("=" * 50)
    print("\n📝 To start the application, run:")
    print("   python app.py")
    print("\n🌐 Then open: http://localhost:5000")
    print("\n📚 For help, read: README.md")
    print("=" * 50)

if __name__ == '__main__':
    main()
