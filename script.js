// ============================
// Global Variables
// ============================
let excelData = [];
let frontTemplate = null;
let backTemplate = null;
let generatedCards = [];

// Fixed positions - no dragging needed
let positions = {
    front: {
        nameY: 391,          // Name vertical position
        nameFontSize: 62,    // Name font size
        badgeY: 457,         // Badge number vertical position
        badgeFontSize: 62    // Badge number font size
    },
    back: {
        barcodeY: 717,       // Barcode vertical position
        barcodeWidth: 450,   // Barcode width
        barcodeHeight: 100,  // Barcode height
        textY: 820,          // Badge text vertical position
        textSize: 28         // Badge text font size
    }
};

// ============================
// Load Templates on Page Load
// ============================
window.addEventListener('DOMContentLoaded', function() {
    const previewBtn = document.querySelector('.btn-preview');
    const generateBtn = document.querySelector('.btn-generate');
    
    previewBtn.disabled = true;
    generateBtn.disabled = true;
    previewBtn.textContent = '⏳ Loading Templates...';
    generateBtn.textContent = '⏳ Loading Templates...';
    
    // Load Front Template
    const frontImg = new Image();
    frontImg.onload = function() {
        frontTemplate = frontImg;
        console.log('Front template loaded:', frontImg.width, 'x', frontImg.height);
        checkTemplatesReady();
    };
    frontImg.onerror = function() {
        console.error('Failed to load front template. Check file path: assets/Front.png');
        alert('❌ Failed to load Front template!\nCheck that assets/Front.png exists.');
    };
    frontImg.src = 'assets/Front.png';
    
    // Load Back Template
    const backImg = new Image();
    backImg.onload = function() {
        backTemplate = backImg;
        console.log('Back template loaded:', backImg.width, 'x', backImg.height);
        checkTemplatesReady();
    };
    backImg.onerror = function() {
        console.error('Failed to load back template. Check file path: assets/Back.png');
        alert('❌ Failed to load Back template!\nCheck that assets/Back.png exists.');
    };
    backImg.src = 'assets/Back.png';
    
    // Enable buttons when both templates are loaded
    function checkTemplatesReady() {
        if (frontTemplate && backTemplate) {
            previewBtn.disabled = false;
            generateBtn.disabled = false;
            previewBtn.textContent = '🔍 Preview First Card';
            generateBtn.textContent = '📦 Generate All Cards';
            console.log('✅ All templates loaded successfully!');
        }
    }
});

// ============================
// File Upload Handler
// ============================
document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
            
            if (jsonData.length > 0) {
                const columns = Object.keys(jsonData[0]);
                console.log('=== EXCEL COLUMNS ===');
                columns.forEach((col, idx) => {
                    console.log(`Column ${idx}: "${col}" (length: ${col.length})`);
                });
                console.log('=== FIRST ROW DATA ===');
                console.log(jsonData[0]);
            }
            
            excelData = jsonData.map((row, idx) => {
                const allKeys = Object.keys(row);
                
                let badge = '';
                let firstName = '';
                let lastName = '';
                
                // Find badge column
                for (let key of allKeys) {
                    const lowerKey = key.toLowerCase().trim();
                    if (lowerKey.includes('badge')) {
                        badge = String(row[key] || '').trim();
                        if (badge) break;
                    }
                }
                
                // Find first name column
                for (let key of allKeys) {
                    const lowerKey = key.toLowerCase().trim();
                    if (lowerKey.includes('first')) {
                        firstName = String(row[key] || '').trim();
                        if (firstName) break;
                    }
                }
                
                // Find last name column
                for (let key of allKeys) {
                    const lowerKey = key.toLowerCase().trim();
                    if (lowerKey.includes('last')) {
                        lastName = String(row[key] || '').trim();
                        if (lastName) break;
                    }
                }
                
                if (idx === 0) {
                    console.log('=== FIRST EMPLOYEE MAPPED ===');
                    console.log('Badge:', badge);
                    console.log('First Name:', firstName);
                    console.log('Last Name:', lastName);
                }
                
                return { badge, firstName, lastName };
            });
            
            document.getElementById('excelBox').classList.add('uploaded');
            document.getElementById('excelBox').querySelector('label').innerHTML = 
                `✅<br>Excel Loaded<br><small>${excelData.length} employees</small>`;
            
            const msg = `✅ Loaded ${excelData.length} employees\n\n` +
                        `First employee:\n` +
                        `First Name: "${excelData[0].firstName}"\n` +
                        `Last Name: "${excelData[0].lastName}"\n` +
                        `Badge: "${excelData[0].badge}"\n\n` +
                        `Check browser console (F12) for details`;
            
            alert(msg);
        } catch (error) {
            alert('❌ Error reading Excel file: ' + error.message);
            console.error('Excel error:', error);
        }
    };
    reader.readAsArrayBuffer(file);
});

// ============================
// Canvas Drawing Functions
// ============================

function drawFrontCard(canvas, employee) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    if (frontTemplate) {
        ctx.drawImage(frontTemplate, 0, 0, frontTemplate.width, frontTemplate.height);
    }
    
    ctx.fillStyle = '#323232';
    ctx.textBaseline = 'top';
    
    // Draw Name (centered with auto-sizing)
    const fullName = `${employee.firstName} ${employee.lastName}`;
    let nameFontSize = positions.front.nameFontSize;
    const maxNameWidth = canvas.width * 0.85;
    const minFontSize = 30;
    
    ctx.font = `900 ${nameFontSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
    let nameWidth = ctx.measureText(fullName).width;
    
    while (nameWidth > maxNameWidth && nameFontSize > minFontSize) {
        nameFontSize -= 2;
        ctx.font = `900 ${nameFontSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
        nameWidth = ctx.measureText(fullName).width;
    }
    
    ctx.textAlign = 'center';
    const nameX = canvas.width / 2;
    ctx.fillText(fullName, nameX, positions.front.nameY);
    
    // Draw Badge Number (centered with offset and auto-sizing)
    let badgeFontSize = positions.front.badgeFontSize;
    const maxBadgeWidth = canvas.width * 0.85;

    ctx.font = `900 ${badgeFontSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
    let badgeWidth = ctx.measureText(employee.badge).width;

    while (badgeWidth > maxBadgeWidth && badgeFontSize > minFontSize) {
        badgeFontSize -= 2;
        ctx.font = `900 ${badgeFontSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
        badgeWidth = ctx.measureText(employee.badge).width;
    }

    // Center with slight offset (adjusts based on badge length)
    ctx.textAlign = 'center';
    const badgeLength = employee.badge.length;
    let badgeOffset;

    if (badgeLength <= 6) {
        badgeOffset = 30;
    } else if (badgeLength === 7) {
        badgeOffset = 40;
    } else {
        badgeOffset = 60;
    }

    const badgeX = (canvas.width / 2) + badgeOffset;
    ctx.fillText(employee.badge, badgeX, positions.front.badgeY);
}

function drawBackCard(canvas, employee) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    if (backTemplate) {
        // Rotate 180 degrees for proper printing orientation
        ctx.save();
        ctx.translate(canvas.width, canvas.height);
        ctx.rotate(Math.PI);
        ctx.drawImage(backTemplate, 0, 0, backTemplate.width, backTemplate.height);
        ctx.restore();
    }
    
    // Rotate text 180 degrees to match template
    ctx.save();
    ctx.translate(canvas.width, canvas.height);
    ctx.rotate(Math.PI);
    
    // Draw Barcode (centered)
    const barcodeCanvas = document.createElement('canvas');
    try {
        JsBarcode(barcodeCanvas, employee.badge, {
            format: 'CODE39',
            width: 3,
            height: 150,
            displayValue: false,
            margin: 0
        });
        
        if (barcodeCanvas.width > 0) {
            const barcodeX = (canvas.width - positions.back.barcodeWidth) / 2;
            
            ctx.drawImage(
                barcodeCanvas,
                barcodeX,
                positions.back.barcodeY,
                positions.back.barcodeWidth,
                positions.back.barcodeHeight
            );
        }
    } catch (error) {
        console.error('Barcode generation error:', error);
    }
    
    // Draw Badge Text (centered)
    ctx.fillStyle = '#323232';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.font = `900 ${positions.back.textSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
    
    const badgeTextX = canvas.width / 2;
    ctx.fillText(employee.badge, badgeTextX, positions.back.textY);
    
    ctx.restore();
}

// ============================
// Preview Function
// ============================
function generatePreview() {
    if (!excelData.length) {
        alert('❌ Please upload Excel file first');
        return;
    }
    
    const employee = excelData[0];
    const previewDiv = document.getElementById('preview');
    
    previewDiv.innerHTML = `
        <h2 class="preview-title">Preview: ${employee.firstName} ${employee.lastName} - Badge #${employee.badge}</h2>
        <div class="preview-cards">
            <div class="card-preview">
                <h4>Front Card</h4>
                <canvas id="previewFront"></canvas>
            </div>
            <div class="card-preview">
                <h4>Back Card (Rotated for Printing)</h4>
                <canvas id="previewBack"></canvas>
            </div>
        </div>
    `;
    
    setTimeout(() => {
        const frontCanvas = document.getElementById('previewFront');
        const backCanvas = document.getElementById('previewBack');
        
        frontCanvas.width = frontTemplate.width;
        frontCanvas.height = frontTemplate.height;
        backCanvas.width = backTemplate.width;
        backCanvas.height = backTemplate.height;
        
        drawFrontCard(frontCanvas, employee);
        drawBackCard(backCanvas, employee);
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
    
    if (frontTemplate.width !== backTemplate.width || frontTemplate.height !== backTemplate.height) {
        const proceed = confirm(
            `⚠️ Warning: Template dimensions don't match!\n\n` +
            `Front: ${frontTemplate.width} × ${frontTemplate.height}px\n` +
            `Back: ${backTemplate.width} × ${backTemplate.height}px\n\n` +
            `This may cause printing issues.\n\n` +
            `Do you want to continue anyway?`
        );
        if (!proceed) return;
    }
    
    generatedCards = [];
    const progressBar = document.getElementById('progressBar');
    const progressBarFill = document.getElementById('progressBarFill');
    progressBar.style.display = 'block';
    
    let processed = 0;
    
    const maxWidth = Math.max(frontTemplate.width, backTemplate.width);
    const maxHeight = Math.max(frontTemplate.height, backTemplate.height);
    
    excelData.forEach((employee, index) => {
        setTimeout(() => {
            const frontCanvas = document.createElement('canvas');
            frontCanvas.width = maxWidth;
            frontCanvas.height = maxHeight;
            drawFrontCard(frontCanvas, employee);
            
            const backCanvas = document.createElement('canvas');
            backCanvas.width = maxWidth;
            backCanvas.height = maxHeight;
            drawBackCard(backCanvas, employee);
            
            generatedCards.push({
                employee: employee,
                front: frontCanvas.toDataURL('image/png', 1.0),
                back: backCanvas.toDataURL('image/png', 1.0)
            });
            
            processed++;
            const progress = Math.round((processed / excelData.length) * 100);
            progressBarFill.style.width = progress + '%';
            progressBarFill.textContent = progress + '%';
            
            if (processed === excelData.length) {
                setTimeout(() => {
                    progressBar.style.display = 'none';
                    alert(`✅ Generated ${processed} cards successfully!\n\nCard size: ${maxWidth} × ${maxHeight}px`);
                    document.getElementById('previewBtn').style.display = 'inline-block';
                    document.getElementById('downloadZipBtn').style.display = 'inline-block';
                }, 500);
            }
        }, index * 100);
    });
}

// ============================
// Cards Preview Function
// ============================
function showCardsPreview() {
    if (!generatedCards.length) {
        alert('❌ Please generate cards first');
        return;
    }
    
    const cardsContent = document.getElementById('cardsContent');
    const cardsPreview = document.getElementById('cardsPreview');
    
    cardsContent.innerHTML = '';
    
    generatedCards.forEach((card, index) => {
        const cardPair = document.createElement('div');
        cardPair.className = 'card-pair';
        
        cardPair.innerHTML = `
            <h3>Card ${index + 1}: ${card.employee.firstName} ${card.employee.lastName} - Badge #${card.employee.badge}</h3>
            <div class="card-sides">
                <div class="card-side">
                    <h4>📄 Front Side</h4>
                    <img src="${card.front}" alt="Front Card">
                </div>
                <div class="card-side">
                    <h4>📄 Back Side (Rotated for Printing)</h4>
                    <img src="${card.back}" alt="Back Card">
                </div>
            </div>
        `;
        
        cardsContent.appendChild(cardPair);
    });
    
    cardsPreview.style.display = 'block';
    cardsPreview.scrollIntoView({ behavior: 'smooth' });
    
    alert(`✅ Preview ready!\n\n${generatedCards.length} cards loaded.\n\nScroll down to review all cards.`);
}

function closeCardsPreview() {
    document.getElementById('cardsPreview').style.display = 'none';
    document.getElementById('preview').scrollIntoView({ behavior: 'smooth' });
}

// ============================
// Download All Cards as ZIP
// ============================
async function downloadAllCards() {
    if (!generatedCards.length) {
        alert('❌ Please generate cards first');
        return;
    }
    
    try {
        alert('⏳ Preparing ZIP file...\n\nThis may take a moment for large batches.');
        
        const zip = new JSZip();
        const cardsFolder = zip.folder("Badge_Cards");
        
        for (let i = 0; i < generatedCards.length; i++) {
            const card = generatedCards[i];
            const paddedIndex = String(i + 1).padStart(4, '0');
            const baseName = `${paddedIndex}_${card.employee.firstName}_${card.employee.lastName}_Badge_${card.employee.badge}`;
            
            const frontBlob = await fetch(card.front).then(r => r.blob());
            const backBlob = await fetch(card.back).then(r => r.blob());
            
            cardsFolder.file(`${baseName}_FRONT.png`, frontBlob);
            cardsFolder.file(`${baseName}_BACK.png`, backBlob);
        }
        
        const content = await zip.generateAsync({
            type: "blob",
            compression: "DEFLATE",
            compressionOptions: { level: 6 }
        });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(content);
        link.download = `Badge_Cards_${generatedCards.length}_Employees.zip`;
        link.click();
        
        setTimeout(() => URL.revokeObjectURL(link.href), 100);
        
        alert(`✅ ZIP file downloaded successfully!\n\n${generatedCards.length} employees\n${generatedCards.length * 2} images (Front + Back)`);
        
    } catch (error) {
        console.error('ZIP creation error:', error);
        alert('❌ Error creating ZIP file: ' + error.message);
    }
}
