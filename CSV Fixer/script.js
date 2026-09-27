document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    
    const uploadCard = document.querySelector('.upload-card');
    const statusContainer = document.getElementById('status-container');
    const statusIcon = document.getElementById('status-icon');
    const statusTitle = document.getElementById('status-title');
    const statusMessage = document.getElementById('status-message');
    const progressBar = document.getElementById('progress-bar');
    const actionButtons = document.getElementById('action-buttons');
    
    const downloadBtn = document.getElementById('download-btn');
    const resetBtn = document.getElementById('reset-btn');

    let fixedCsvData = null;
    let originalFileName = '';

    // Drag and drop event listeners
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.add('drag-over');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove('drag-over');
        }, false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });

    fileInput.addEventListener('change', function(e) {
        if (this.files.length > 0) {
            handleFile(this.files[0]);
        }
    });

    function handleFile(file) {
        if (!file.name.endsWith('.csv')) {
            showError("Please upload a valid CSV file.");
            return;
        }

        originalFileName = file.name;
        showProcessing();

        const reader = new FileReader();
        reader.onload = function(e) {
            const text = e.target.result;
            // Simulate processing time for better UX
            setTimeout(() => {
                try {
                    fixedCsvData = fixCSV(text);
                    showSuccess();
                } catch (error) {
                    showError("Failed to process the CSV file. It might be too corrupted or in an unknown format.");
                    console.error(error);
                }
            }, 800);
        };
        
        reader.onerror = function() {
            showError("Failed to read the file.");
        };

        reader.readAsText(file);
    }

    function fixCSV(csvText) {
        const lines = csvText.split(/\r?\n/);
        const fixedLines = [];
        
        // The file structure has 16 expected columns.
        // We know that the last 14 columns are numeric data and the 1st is S No.
        // The problem is the 2nd column (Bank Name) contains commas without quotes.
        
        const EXPECTED_COLS = 16;
        const NUM_TRAILING_COLS = 14;

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            if (!line.trim()) {
                fixedLines.push(line);
                continue;
            }

            // Split by comma. Note: this naive split works here because 
            // the problem explicitly states the commas are raw unquoted commas.
            // If there were already quotes, a more complex parser would be needed.
            const parts = line.split(',');

            if (parts.length > EXPECTED_COLS) {
                // There are extra commas in the name field
                const sNo = parts[0];
                
                // Extract trailing numerical columns
                const trailingStartIdx = parts.length - NUM_TRAILING_COLS;
                const trailingCols = parts.slice(trailingStartIdx);
                
                // The name field is everything in between
                const nameParts = parts.slice(1, trailingStartIdx);
                
                // Join name parts and wrap in double quotes to escape commas
                const fixedName = `"${nameParts.join(',')}"`;
                
                // Reconstruct the row
                const fixedRow = [sNo, fixedName, ...trailingCols].join(',');
                fixedLines.push(fixedRow);
            } else {
                // Line is fine (or has fewer columns, which we leave alone)
                fixedLines.push(line);
            }
        }

        return fixedLines.join('\n');
    }

    // UI State Management
    function showProcessing() {
        uploadCard.classList.add('hidden');
        statusContainer.classList.remove('hidden');
        
        statusIcon.className = 'status-icon processing';
        statusIcon.innerHTML = `
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>
            </svg>
        `;
        
        statusTitle.textContent = 'Processing...';
        statusMessage.textContent = 'Analyzing and fixing your CSV structure.';
        
        progressBar.classList.remove('hidden');
        actionButtons.classList.add('hidden');
    }

    function showSuccess() {
        statusIcon.className = 'status-icon success';
        statusIcon.innerHTML = `
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
                <polyline points="22 4 12 14.01 9 11.01"></polyline>
            </svg>
        `;
        
        statusTitle.textContent = 'Success!';
        statusMessage.textContent = 'Your CSV file has been successfully fixed and aligned.';
        
        progressBar.classList.add('hidden');
        actionButtons.classList.remove('hidden');
    }

    function showError(msg) {
        uploadCard.classList.add('hidden');
        statusContainer.classList.remove('hidden');
        
        statusIcon.className = 'status-icon error';
        statusIcon.innerHTML = `
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="8" x2="12" y2="12"></line>
                <line x1="12" y1="16" x2="12.01" y2="16"></line>
            </svg>
        `;
        
        statusTitle.textContent = 'Error';
        statusMessage.textContent = msg;
        
        progressBar.classList.add('hidden');
        actionButtons.classList.remove('hidden');
        downloadBtn.classList.add('hidden'); // Hide download on error
    }

    // Button actions
    downloadBtn.addEventListener('click', () => {
        if (!fixedCsvData) return;
        
        const blob = new Blob([fixedCsvData], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement('a');
        link.setAttribute('href', url);
        
        // Add "fixed_" prefix to original filename
        const downloadName = "fixed_" + originalFileName;
        link.setAttribute('download', downloadName);
        
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });

    resetBtn.addEventListener('click', () => {
        statusContainer.classList.add('hidden');
        uploadCard.classList.remove('hidden');
        fileInput.value = '';
        fixedCsvData = null;
        originalFileName = '';
        downloadBtn.classList.remove('hidden');
    });
});
