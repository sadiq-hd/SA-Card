// ============================
// Global Variables
// ============================
let excelData = [];
let frontTemplate = null;
let backTemplate = null;
let generatedCards = [];

// Position settings (stored in memory)
let positions = {
    front: {
        nameX: 150,
        nameY: 300,
        nameFontSize: 62,
        badgeX: 150,
        badgeY: 380,
        badgeFontSize: 62
    },
    back: {
        barcodeX: 50,
        barcodeY: 150,
        barcodeWidth: 450,
        barcodeHeight: 100,
        textX: 275,
        textY: 320,
        textSize: 28
    }
};

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
                
                for (let key of allKeys) {
                    const lowerKey = key.toLowerCase().trim();
                    if (lowerKey.includes('badge')) {
                        badge = String(row[key] || '').trim();
                        if (badge) break;
                    }
                }
                
                for (let key of allKeys) {
                    const lowerKey = key.toLowerCase().trim();
                    if (lowerKey.includes('first')) {
                        firstName = String(row[key] || '').trim();
                        if (firstName) break;
                    }
                }
                
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

// Handle Front Template Upload
document.getElementById('frontImage').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(event) {
        const img = new Image();
        img.onload = function() {
            frontTemplate = img;
            document.getElementById('frontBox').classList.add('uploaded');
            document.getElementById('frontBox').querySelector('label').innerHTML = 
                '✅<br>Front Template<br><small>Ready to use</small>';
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
        img.onload = function() {
            backTemplate = img;
            document.getElementById('backBox').classList.add('uploaded');
            document.getElementById('backBox').querySelector('label').innerHTML = 
                '✅<br>Back Template<br><small>Ready to use</small>';
            console.log('Back template loaded:', img.width, 'x', img.height);
        };
        img.src = event.target.result;
    };
    reader.readAsDataURL(file);
});

// ============================
// Canvas Drawing Functions
// ============================

function drawFrontCard(canvas, employee, hideText = false) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    if (frontTemplate) {
        ctx.drawImage(frontTemplate, 0, 0, frontTemplate.width, frontTemplate.height);
    }
    
    if (!hideText) {
        ctx.fillStyle = '#323232';
        ctx.textBaseline = 'top';
        
        // ========== الاسم (مع توسيط وتصغير تلقائي) ==========
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
        
        // توسيط الاسم
        ctx.textAlign = 'center';
        const nameX = canvas.width / 2;
        ctx.fillText(fullName, nameX, positions.front.nameY);
        
        // ========== رقم البادج (بعد "Badge NO:") ==========
        let badgeFontSize = positions.front.badgeFontSize;
const maxBadgeWidth = canvas.width * 0.85;

ctx.font = `900 ${badgeFontSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
let badgeWidth = ctx.measureText(employee.badge).width;

while (badgeWidth > maxBadgeWidth && badgeFontSize > minFontSize) {
    badgeFontSize -= 2;
    ctx.font = `900 ${badgeFontSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
    badgeWidth = ctx.measureText(employee.badge).width;
}

// توسيط مع إزاحة بسيطة لليمين
ctx.textAlign = 'center';
const badgeLength = employee.badge.length;
let badgeOffset;

if (badgeLength <= 6) {
    badgeOffset = 30;  // 6 أرقام أو أقل
} else if (badgeLength === 7) {
    badgeOffset = 40;  // 7 أرقام
} else {
    badgeOffset = 60;  // 8 أرقام أو أكثر
}

const badgeX = (canvas.width / 2) + badgeOffset;
ctx.fillText(employee.badge, badgeX, positions.front.badgeY);
    }
}

function drawBackCard(canvas, employee, hideText = false) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    if (backTemplate) {
        ctx.drawImage(backTemplate, 0, 0, backTemplate.width, backTemplate.height);
    }
    
    if (!hideText) {
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
                ctx.drawImage(
                    barcodeCanvas,
                    positions.back.barcodeX,
                    positions.back.barcodeY,
                    positions.back.barcodeWidth,
                    positions.back.barcodeHeight
                );
            }
        } catch (error) {
            console.error('Barcode generation error:', error);
        }
        
        ctx.fillStyle = '#323232';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        ctx.font = `900 ${positions.back.textSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
        ctx.fillText(employee.badge, positions.back.textX, positions.back.textY);
    }
}

// ============================
// Draggable Elements
// ============================

function makeDraggable(element, side, itemType) {
    let isDragging = false;
    let startX, startY, initialLeft, initialTop;
    
    element.addEventListener('mousedown', function(e) {
        isDragging = true;
        element.classList.add('active');
        
        startX = e.clientX;
        startY = e.clientY;
        initialLeft = parseInt(element.style.left) || 0;
        initialTop = parseInt(element.style.top) || 0;
        
        e.preventDefault();
        e.stopPropagation();
    });
    
    const handleMouseMove = function(e) {
        if (!isDragging) return;
        
        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;
        
        const newLeft = initialLeft + deltaX;
        const newTop = initialTop + deltaY;
        
        element.style.left = newLeft + 'px';
        element.style.top = newTop + 'px';
        
        e.preventDefault();
    };
    
    const handleMouseUp = function(e) {
        if (!isDragging) return;
        
        isDragging = false;
        element.classList.remove('active');
        
        const container = element.parentElement;
        const canvas = container.querySelector('canvas');
        const canvasRect = canvas.getBoundingClientRect();
        const elementRect = element.getBoundingClientRect();
        
        const scaleX = canvas.width / canvasRect.width;
        const scaleY = canvas.height / canvasRect.height;
        
        const finalX = Math.round((elementRect.left - canvasRect.left) * scaleX);
        const finalY = Math.round((elementRect.top - canvasRect.top) * scaleY);
        
        if (side === 'front') {
            if (itemType === 'name') {
                positions.front.nameX = finalX;
                positions.front.nameY = finalY;
            } else if (itemType === 'badge') {
                positions.front.badgeX = finalX;
                positions.front.badgeY = finalY;
            }
        } else if (side === 'back') {
            if (itemType === 'barcode') {
                positions.back.barcodeX = finalX;
                positions.back.barcodeY = finalY;
            } else if (itemType === 'text') {
                positions.back.textX = finalX;
                positions.back.textY = finalY;
            }
        }
        
        const employee = excelData[0];
        if (side === 'front') {
            drawFrontCard(canvas, employee, true);
        } else {
            drawBackCard(canvas, employee, true);
        }
        
        console.log('Updated positions:', positions);
    };
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
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
        alert('❌ Please upload both templates');
        return;
    }
    
    const employee = excelData[0];
    const previewDiv = document.getElementById('preview');
    
    previewDiv.innerHTML = `
        <h2 class="preview-title">Preview: ${employee.firstName} ${employee.lastName} - Badge #${employee.badge}</h2>
        <div class="preview-cards">
            <div class="card-preview">
                <h4>Front Card</h4>
                <div class="positioning-container" id="frontContainer">
                    <canvas id="previewFront"></canvas>
                </div>
            </div>
            <div class="card-preview">
                <h4>Back Card</h4>
                <div class="positioning-container" id="backContainer">
                    <canvas id="previewBack"></canvas>
                </div>
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
        
        drawFrontCard(frontCanvas, employee, true);
        drawBackCard(backCanvas, employee, true);
        
        const frontContainer = document.getElementById('frontContainer');
        const canvasRect = frontCanvas.getBoundingClientRect();
        const scaleX = canvasRect.width / frontCanvas.width;
        const scaleY = canvasRect.height / frontCanvas.height;
        
        const nameDiv = document.createElement('div');
        nameDiv.className = 'draggable-item';
        nameDiv.textContent = `${employee.firstName} ${employee.lastName}`;
        nameDiv.style.left = (positions.front.nameX * scaleX) + 'px';
        nameDiv.style.top = (positions.front.nameY * scaleY) + 'px';
        nameDiv.style.fontSize = (positions.front.nameFontSize * scaleX) + 'px';
        frontContainer.appendChild(nameDiv);
        makeDraggable(nameDiv, 'front', 'name');
        
        const badgeDiv = document.createElement('div');
        badgeDiv.className = 'draggable-item';
        badgeDiv.textContent = employee.badge;
        badgeDiv.style.left = (positions.front.badgeX * scaleX) + 'px';
        badgeDiv.style.top = (positions.front.badgeY * scaleY) + 'px';
        badgeDiv.style.fontSize = (positions.front.badgeFontSize * scaleX) + 'px';
        frontContainer.appendChild(badgeDiv);
        makeDraggable(badgeDiv, 'front', 'badge');
        
        const backContainer = document.getElementById('backContainer');
        const backCanvasRect = backCanvas.getBoundingClientRect();
        const backScaleX = backCanvasRect.width / backCanvas.width;
        const backScaleY = backCanvasRect.height / backCanvas.height;
        
        const barcodeDiv = document.createElement('div');
        barcodeDiv.className = 'draggable-item';
        barcodeDiv.textContent = `Barcode`;
        barcodeDiv.style.left = (positions.back.barcodeX * backScaleX) + 'px';
        barcodeDiv.style.top = (positions.back.barcodeY * backScaleY) + 'px';
        barcodeDiv.style.width = (positions.back.barcodeWidth * backScaleX) + 'px';
        barcodeDiv.style.height = (positions.back.barcodeHeight * backScaleY) + 'px';
        barcodeDiv.style.display = 'flex';
        barcodeDiv.style.alignItems = 'center';
        barcodeDiv.style.justifyContent = 'center';
        backContainer.appendChild(barcodeDiv);
        makeDraggable(barcodeDiv, 'back', 'barcode');
        
        const barcodeTextDiv = document.createElement('div');
        barcodeTextDiv.className = 'draggable-item';
        barcodeTextDiv.textContent = employee.badge;
        barcodeTextDiv.style.left = (positions.back.textX * backScaleX) + 'px';
        barcodeTextDiv.style.top = (positions.back.textY * backScaleY) + 'px';
        barcodeTextDiv.style.fontSize = (positions.back.textSize * backScaleX) + 'px';
        barcodeTextDiv.style.width = '200px';
        barcodeTextDiv.style.textAlign = 'center';
        backContainer.appendChild(barcodeTextDiv);
        makeDraggable(barcodeTextDiv, 'back', 'text');
        
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
        alert('❌ Please upload both templates');
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
        document.getElementById('downloadBtn').style.display = 'inline-block';
        document.getElementById('downloadZipBtn').style.display = 'inline-block';  // ← أضف هذا السطر
    }, 500);
}
        }, index * 100);
    });
}

// ============================
// Print Preview and Print Functions
// ============================
function showPrintPreview() {
    if (!generatedCards.length) {
        alert('❌ Please generate cards first');
        return;
    }
    
    const printContent = document.getElementById('printContent');
    const printPreview = document.getElementById('printPreview');
    
    printContent.innerHTML = '';
    
    // كل بطاقة Front ثم Back في صفحة منفصلة
    generatedCards.forEach((card, index) => {
        // صفحة Front
        const frontPage = document.createElement('div');
        frontPage.className = 'print-page';
        frontPage.innerHTML = `
            <div class="preview-info">
                <p>Card ${index + 1}/${generatedCards.length} - FRONT</p>
                <p>${card.employee.firstName} ${card.employee.lastName} - Badge #${card.employee.badge}</p>
            </div>
            <img src="${card.front}" alt="Front Card" class="card-image">
        `;
        printContent.appendChild(frontPage);
        
        // صفحة Back
        const backPage = document.createElement('div');
        backPage.className = 'print-page';
        backPage.innerHTML = `
            <div class="preview-info">
                <p>Card ${index + 1}/${generatedCards.length} - BACK</p>
                <p>${card.employee.firstName} ${card.employee.lastName} - Badge #${card.employee.badge}</p>
            </div>
            <img src="${card.back}" alt="Back Card" class="card-image">
        `;
        printContent.appendChild(backPage);
    });
    
    printPreview.style.display = 'block';
    printPreview.scrollIntoView({ behavior: 'smooth' });
    
    const totalPages = generatedCards.length * 2;
    alert(`✅ Print preview ready!\n\n${generatedCards.length} employees\n${totalPages} pages (Front + Back)\n\nScroll down to review, then click "Print All Cards"`);
}
function printCards() {
    window.print();
}

function closePrintPreview() {
    document.getElementById('printPreview').style.display = 'none';
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
    const paddedIndex = String(i + 1).padStart(4, '0'); // 0001, 0002, 0003...
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