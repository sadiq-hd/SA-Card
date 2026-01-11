// ============================
// Global Variables
// ============================
let excelData = [];
let frontTemplate = null;
let backTemplate = null;
let generatedCards = [];

// ============================
// Load Templates on Page Load
// ============================
window.addEventListener('DOMContentLoaded', function() {
    // Load Front Template
    const frontImg = new Image();
    frontImg.crossOrigin = 'anonymous';
    frontImg.onload = function() {
        frontTemplate = frontImg;
        document.getElementById('frontBox').classList.add('uploaded');
        document.getElementById('frontBox').querySelector('label').innerHTML = 
            '✅<br>Front Template Loaded<br><small>(Click to change)</small>';
        console.log('Front template loaded:', frontImg.width, 'x', frontImg.height);
    };
    frontImg.onerror = function() {
        console.error('Failed to load Front.png');
        alert('⚠️ Front.png not found in the same folder as HTML file');
    };
    frontImg.src = 'Front.png'; // Must be in same folder as HTML
    
    // Load Back Template
    const backImg = new Image();
    backImg.crossOrigin = 'anonymous';
    backImg.onload = function() {
        backTemplate = backImg;
        document.getElementById('backBox').classList.add('uploaded');
        document.getElementById('backBox').querySelector('label').innerHTML = 
            '✅<br>Back Template Loaded<br><small>(Click to change)</small>';
        console.log('Back template loaded:', backImg.width, 'x', backImg.height);
    };
    backImg.onerror = function() {
        console.error('Failed to load Back.png');
        alert('⚠️ Back.png not found in the same folder as HTML file');
    };
    backImg.src = 'Back.png'; // Must be in same folder as HTML
});

// ============================
// File Upload Handlers
// ============================

// Handle Excel Upload
document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            excelData = jsonData.map(row => ({
                badge: String(row['Badge #'] || row['Badge'] || '').trim(),
                firstName: String(row['First name'] || row['FirstName'] || '').trim(),
                lastName: String(row['Last name'] || row['LastName'] || '').trim()
            }));
            
            document.getElementById('excelBox').classList.add('uploaded');
            alert(`✅ Loaded ${excelData.length} employees from Excel`);
        } catch (error) {
            alert('❌ Error reading Excel file: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
});

// Handle Front Template Upload
document.getElementById('frontImage').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(event) {
        const img = new Image();
        img.crossOrigin = 'anonymous'; // Important for canvas
        img.onload = function() {
            frontTemplate = img;
            document.getElementById('frontBox').classList.add('uploaded');
            console.log('Front template loaded:', img.width, 'x', img.height);
            alert('✅ Front template loaded successfully');
        };
        img.onerror = function() {
            alert('❌ Error loading front template image');
        };
        img.src = event.target.result;
    };
    reader.readAsDataURL(file);
});

// Handle Back Template Upload
document.getElementById('backImage').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(event) {
        const img = new Image();
        img.crossOrigin = 'anonymous'; // Important for canvas
        img.onload = function() {
            backTemplate = img;
            document.getElementById('backBox').classList.add('uploaded');
            console.log('Back template loaded:', img.width, 'x', img.height);
            alert('✅ Back template loaded successfully');
        };
        img.onerror = function() {
            alert('❌ Error loading back template image');
        };
        img.src = event.target.result;
    };
    reader.readAsDataURL(file);
});

// ============================
// Card Generation Functions
// ============================

function getCardDimensions() {
    const width = parseFloat(document.getElementById('cardWidth').value);
    const height = parseFloat(document.getElementById('cardHeight').value);
    const dpi = parseInt(document.getElementById('cardDPI').value);
    
    const pixelWidth = Math.round((width / 2.54) * dpi);
    const pixelHeight = Math.round((height / 2.54) * dpi);
    
    // Update display
    const display = document.getElementById('dimensionDisplay');
    if (display) {
        display.textContent = `${pixelWidth} × ${pixelHeight} pixels`;
    }
    
    return {
        width: pixelWidth,
        height: pixelHeight
    };
}

// Update dimension display on input change
document.addEventListener('DOMContentLoaded', function() {
    ['cardWidth', 'cardHeight', 'cardDPI'].forEach(id => {
        const input = document.getElementById(id);
        if (input) {
            input.addEventListener('input', getCardDimensions);
        }
    });
    
    // Initial update
    setTimeout(getCardDimensions, 100);
});

function getSettings() {
    return {
        nameX: parseInt(document.getElementById('nameX').value),
        nameY: parseInt(document.getElementById('nameY').value),
        nameFontSize: parseInt(document.getElementById('nameFontSize').value),
        nameFontWeight: document.getElementById('nameFontWeight').value,
        nameColor: document.getElementById('nameColor').value,
        
        badgeLabelX: parseInt(document.getElementById('badgeLabelX').value),
        badgeLabelY: parseInt(document.getElementById('badgeLabelY').value),
        badgeLabelSize: parseInt(document.getElementById('badgeLabelSize').value),
        
        badgeNumX: parseInt(document.getElementById('badgeNumX').value),
        badgeNumY: parseInt(document.getElementById('badgeNumY').value),
        badgeNumSize: parseInt(document.getElementById('badgeNumSize').value),
        badgeFontWeight: document.getElementById('badgeFontWeight').value,
        
        barcodeX: parseInt(document.getElementById('barcodeX').value),
        barcodeY: parseInt(document.getElementById('barcodeY').value),
        barcodeWidth: parseInt(document.getElementById('barcodeWidth').value),
        barcodeHeight: parseInt(document.getElementById('barcodeHeight').value)
    };
}

function generateFrontCard(employee, dimensions, settings) {
    const canvas = document.createElement('canvas');
    canvas.width = dimensions.width;
    canvas.height = dimensions.height;
    
    drawFrontCard(canvas, employee, dimensions, settings);
    return canvas;
}

function drawFrontCard(canvas, employee, dimensions, settings) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    // Clear canvas
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    
    // White background first
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    // Draw template image
    if (frontTemplate && frontTemplate.complete && frontTemplate.naturalWidth > 0) {
        try {
            ctx.drawImage(frontTemplate, 0, 0, dimensions.width, dimensions.height);
            console.log('Front image drawn successfully');
        } catch (e) {
            console.error('Error drawing front image:', e);
        }
    } else {
        console.error('Front template not ready');
    }
    
    // Setup text rendering
    ctx.textBaseline = 'top';
    ctx.textAlign = 'left';
    
    // Full name
    const fullName = `${employee.firstName} ${employee.lastName}`.toUpperCase();
    ctx.font = `${settings.nameFontWeight} ${settings.nameFontSize}px Amiri, Arial, sans-serif`;
    ctx.fillStyle = settings.nameColor;
    ctx.fillText(fullName, settings.nameX, settings.nameY);
    
    // "Badge NO:"
    ctx.font = `${settings.badgeFontWeight} ${settings.badgeLabelSize}px Amiri, Arial, sans-serif`;
    ctx.fillStyle = '#000000';
    ctx.fillText('Badge NO:', settings.badgeLabelX, settings.badgeLabelY);
    
    // Badge number
    ctx.font = `${settings.badgeFontWeight} ${settings.badgeNumSize}px Amiri, Arial, sans-serif`;
    ctx.fillText(employee.badge, settings.badgeNumX, settings.badgeNumY);
}

function generateBackCard(employee, dimensions, settings) {
    const canvas = document.createElement('canvas');
    canvas.width = dimensions.width;
    canvas.height = dimensions.height;
    
    drawBackCard(canvas, employee, dimensions, settings);
    return canvas;
}

function drawBackCard(canvas, employee, dimensions, settings) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    // Clear canvas
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    
    // White background first
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    // Draw template image
    if (backTemplate && backTemplate.complete && backTemplate.naturalWidth > 0) {
        try {
            ctx.drawImage(backTemplate, 0, 0, dimensions.width, dimensions.height);
            console.log('Back image drawn successfully');
        } catch (e) {
            console.error('Error drawing back image:', e);
        }
    } else {
        console.error('Back template not ready');
    }
    
    // Generate and draw barcode
    const barcodeCanvas = document.createElement('canvas');
    try {
        JsBarcode(barcodeCanvas, employee.badge, {
            format: 'CODE128',
            width: 2,
            height: 100,
            displayValue: false,
            margin: 0
        });
        
        // Draw barcode on main canvas
        if (barcodeCanvas.width > 0) {
            ctx.drawImage(
                barcodeCanvas,
                settings.barcodeX,
                settings.barcodeY,
                settings.barcodeWidth,
                settings.barcodeHeight
            );
            console.log('Barcode drawn successfully');
        }
    } catch (error) {
        console.error('Barcode generation error:', error);
    }
}

// ============================
// Preview Function
// ============================
function generatePreview() {
    if (!excelData.length) {
        alert('❌ Please upload Excel file first');
        return;
    }
    if (!frontTemplate || !backTemplate) {
        alert('❌ Templates not loaded yet. Please wait...');
        return;
    }
    
    const dimensions = getCardDimensions();
    const settings = getSettings();
    const employee = excelData[0];
    
    console.log('Generating preview with dimensions:', dimensions);
    console.log('Employee:', employee);
    
    const previewDiv = document.getElementById('preview');
    previewDiv.innerHTML = `
        <h2 class="preview-title">Preview: ${employee.firstName} ${employee.lastName} - Badge #${employee.badge}</h2>
        <div style="display: flex; gap: 20px; justify-content: center; flex-wrap: wrap;">
            <div class="card-preview">
                <h4>Front</h4>
                <canvas id="previewFront"></canvas>
            </div>
            <div class="card-preview">
                <h4>Back</h4>
                <canvas id="previewBack"></canvas>
            </div>
        </div>
    `;
    
    // Wait for DOM to update, then draw
    setTimeout(() => {
        const frontCanvas = document.getElementById('previewFront');
        const backCanvas = document.getElementById('previewBack');
        
        frontCanvas.width = dimensions.width;
        frontCanvas.height = dimensions.height;
        backCanvas.width = dimensions.width;
        backCanvas.height = dimensions.height;
        
        drawFrontCard(frontCanvas, employee, dimensions, settings);
        drawBackCard(backCanvas, employee, dimensions, settings);
    }, 100);
}

// ============================
// Generate All Cards
// ============================
function generateAllCards() {
    if (!excelData.length) {
        alert('❌ Please upload Excel file first');
        return;
    }
    if (!frontTemplate || !backTemplate) {
        alert('❌ Please upload both front and back templates');
        return;
    }
    
    const dimensions = getCardDimensions();
    const settings = getSettings();
    
    generatedCards = [];
    const progressBar = document.getElementById('progressBar');
    const progressBarFill = document.getElementById('progressBarFill');
    progressBar.style.display = 'block';
    
    let processed = 0;
    
    excelData.forEach((employee, index) => {
        setTimeout(() => {
            const frontCanvas = generateFrontCard(employee, dimensions, settings);
            const backCanvas = generateBackCard(employee, dimensions, settings);
            
            generatedCards.push({
                employee: employee,
                front: frontCanvas,
                back: backCanvas
            });
            
            processed++;
            const progress = Math.round((processed / excelData.length) * 100);
            progressBarFill.style.width = progress + '%';
            progressBarFill.textContent = progress + '%';
            
            if (processed === excelData.length) {
                setTimeout(() => {
                    progressBar.style.display = 'none';
                    alert(`✅ Generated ${processed} cards successfully!`);
                    document.getElementById('printBtn').style.display = 'inline-block';
                }, 500);
            }
        }, index * 50);
    });
}

// ============================
// Print All Cards
// ============================
function printAllCards() {
    if (!generatedCards.length) {
        alert('❌ Please generate cards first');
        return;
    }
    
    const allCardsDiv = document.getElementById('allCards');
    allCardsDiv.innerHTML = '';
    
    generatedCards.forEach(card => {
        // Front page
        const frontPage = document.createElement('div');
        frontPage.className = 'print-page';
        frontPage.appendChild(card.front.cloneNode(true));
        allCardsDiv.appendChild(frontPage);
        
        // Back page
        const backPage = document.createElement('div');
        backPage.className = 'print-page';
        backPage.appendChild(card.back.cloneNode(true));
        allCardsDiv.appendChild(backPage);
    });
    
    // Trigger print
    window.print();
}