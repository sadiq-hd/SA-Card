let excelData = [];
let frontTemplate = null;
let backTemplate = null;
let generatedCards = [];

// Position settings (stored in memory)
let positions = {
    front: {
        nameX: 150,
        nameY: 300,
        nameFontSize: 72,
        badgeX: 150,
        badgeY: 380,
        badgeFontSize: 72
    },
    back: {
        barcodeX: 50,
        barcodeY: 150,
        barcodeWidth: 700,
        barcodeHeight: 150,
        textY: 320,
        textSize: 48
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
            
            // Debug: Show all column names
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
                // Get all possible values for each field
                const allKeys = Object.keys(row);
                
                let badge = '';
                let firstName = '';
                let lastName = '';
                
                // Find badge
                for (let key of allKeys) {
                    const lowerKey = key.toLowerCase().trim();
                    if (lowerKey.includes('badge')) {
                        badge = String(row[key] || '').trim();
                        if (badge) break;
                    }
                }
                
                // Find first name
                for (let key of allKeys) {
                    const lowerKey = key.toLowerCase().trim();
                    if (lowerKey.includes('first')) {
                        firstName = String(row[key] || '').trim();
                        if (firstName) break;
                    }
                }
                
                // Find last name
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
            console.log('Front template loaded:', img.width, 'x', img.height);
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

function drawFrontCard(canvas, employee) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    // Clear and draw template
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    if (frontTemplate) {
        // Center the template if canvas is larger
        const x = (canvas.width - frontTemplate.width) / 2;
        const y = (canvas.height - frontTemplate.height) / 2;
        ctx.drawImage(frontTemplate, x, y, frontTemplate.width, frontTemplate.height);
    }
    
    // Setup text rendering with bold
    ctx.fillStyle = '#323232';
    ctx.textBaseline = 'top';
    ctx.textAlign = 'left';
    
    // Draw full name (First + Last)
    const fullName = `${employee.firstName} ${employee.lastName}`;
    ctx.font = `900 ${positions.front.nameFontSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
    ctx.fillText(fullName, positions.front.nameX, positions.front.nameY);
    
    // Draw badge number only (no label)
    ctx.font = `900 ${positions.front.badgeFontSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
    ctx.fillText(employee.badge, positions.front.badgeX, positions.front.badgeY);
}

function drawBackCard(canvas, employee) {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    // Clear and draw template
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    if (backTemplate) {
        // Center the template if canvas is larger
        const x = (canvas.width - backTemplate.width) / 2;
        const y = (canvas.height - backTemplate.height) / 2;
        ctx.drawImage(backTemplate, x, y, backTemplate.width, backTemplate.height);
    }
    
    // Generate and draw barcode
    const barcodeCanvas = document.createElement('canvas');
    try {
        JsBarcode(barcodeCanvas, employee.badge, {
            format: 'CODE128',
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
    
    // Draw badge number text below barcode
    ctx.fillStyle = '#323232';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.font = `900 ${positions.back.textSize}px 'Noto Kufi Arabic', Arial, sans-serif`;
    const centerX = positions.back.barcodeX + (positions.back.barcodeWidth / 2);
    ctx.fillText(employee.badge, centerX, positions.back.textY);
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
        
        // Calculate final position relative to canvas
        const container = element.parentElement;
        const canvas = container.querySelector('canvas');
        const canvasRect = canvas.getBoundingClientRect();
        const elementRect = element.getBoundingClientRect();
        
        const scaleX = canvas.width / canvasRect.width;
        const scaleY = canvas.height / canvasRect.height;
        
        const finalX = Math.round((elementRect.left - canvasRect.left) * scaleX);
        const finalY = Math.round((elementRect.top - canvasRect.top) * scaleY);
        
        // Update positions
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
                positions.back.textY = finalY;
            }
        }
        
        // Redraw canvas with new positions
        const employee = excelData[0];
        if (side === 'front') {
            drawFrontCard(canvas, employee);
        } else {
            drawBackCard(canvas, employee);
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
        
        // Set canvas dimensions to match templates
        frontCanvas.width = frontTemplate.width;
        frontCanvas.height = frontTemplate.height;
        backCanvas.width = backTemplate.width;
        backCanvas.height = backTemplate.height;
        
        // Draw cards
        drawFrontCard(frontCanvas, employee);
        drawBackCard(backCanvas, employee);
        
        // Add draggable elements for front card
        const frontContainer = document.getElementById('frontContainer');
        const canvasRect = frontCanvas.getBoundingClientRect();
        const scaleX = canvasRect.width / frontCanvas.width;
        const scaleY = canvasRect.height / frontCanvas.height;
        
        // Name draggable
        const nameDiv = document.createElement('div');
        nameDiv.className = 'draggable-item';
        nameDiv.textContent = `${employee.firstName} ${employee.lastName}`;
        nameDiv.style.left = (positions.front.nameX * scaleX) + 'px';
        nameDiv.style.top = (positions.front.nameY * scaleY) + 'px';
        nameDiv.style.fontSize = (positions.front.nameFontSize * scaleX) + 'px';
        frontContainer.appendChild(nameDiv);
        makeDraggable(nameDiv, 'front', 'name');
        
        // Badge draggable
        const badgeDiv = document.createElement('div');
        badgeDiv.className = 'draggable-item';
        badgeDiv.textContent = employee.badge;
        badgeDiv.style.left = (positions.front.badgeX * scaleX) + 'px';
        badgeDiv.style.top = (positions.front.badgeY * scaleY) + 'px';
        badgeDiv.style.fontSize = (positions.front.badgeFontSize * scaleX) + 'px';
        frontContainer.appendChild(badgeDiv);
        makeDraggable(badgeDiv, 'front', 'badge');
        
        // Add draggable element for back card barcode
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
        
        // Badge number text below barcode (draggable)
        const barcodeTextDiv = document.createElement('div');
        barcodeTextDiv.className = 'draggable-item';
        barcodeTextDiv.textContent = employee.badge;
        barcodeTextDiv.style.left = ((positions.back.barcodeX + positions.back.barcodeWidth/2 - 100) * backScaleX) + 'px';
        barcodeTextDiv.style.top = (positions.back.textY * backScaleY) + 'px';
        barcodeTextDiv.style.fontSize = (positions.back.textSize * backScaleX) + 'px';
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
    
    // Check if templates have same dimensions
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
    
    // Use the larger dimensions to ensure both cards fit
    const maxWidth = Math.max(frontTemplate.width, backTemplate.width);
    const maxHeight = Math.max(frontTemplate.height, backTemplate.height);
    
    excelData.forEach((employee, index) => {
        setTimeout(() => {
            // Create front card with consistent dimensions
            const frontCanvas = document.createElement('canvas');
            frontCanvas.width = maxWidth;
            frontCanvas.height = maxHeight;
            drawFrontCard(frontCanvas, employee);
            
            // Create back card with consistent dimensions
            const backCanvas = document.createElement('canvas');
            backCanvas.width = maxWidth;
            backCanvas.height = maxHeight;
            drawBackCard(backCanvas, employee);
            
            generatedCards.push({
                employee: employee,
                front: frontCanvas.toDataURL('image/png'),
                back: backCanvas.toDataURL('image/png')
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
    
    generatedCards.forEach((card, index) => {
        const cardPair = document.createElement('div');
        cardPair.className = 'print-card-pair';
        
        cardPair.innerHTML = `
            <h3>${card.employee.firstName} ${card.employee.lastName} - Badge #${card.employee.badge}</h3>
            <div class="card-sides">
                <div class="card-side">
                    <h4>📄 Front Side</h4>
                    <img src="${card.front}" alt="Front Card">
                </div>
                <div class="card-side">
                    <h4>📄 Back Side</h4>
                    <img src="${card.back}" alt="Back Card">
                </div>
            </div>
        `;
        
        printContent.appendChild(cardPair);
    });
    
    // Show print preview and scroll to it
    printPreview.style.display = 'block';
    printPreview.scrollIntoView({ behavior: 'smooth' });
    
    alert(`✅ Print preview ready!\n\n${generatedCards.length} cards loaded.\n\nScroll down to review, then click "Print All Cards"`);
}

function printCards() {
    window.print();
}

function closePrintPreview() {
    document.getElementById('printPreview').style.display = 'none';
    document.getElementById('preview').scrollIntoView({ behavior: 'smooth' });
}

// ============================
// Download All Cards (Keep as backup option)
// ============================
function downloadAllCards() {
    if (!generatedCards.length) {
        alert('❌ Please generate cards first');
        return;
    }
    
    // Create a combined PDF-style download (front and back together)
    const downloadMethod = confirm('Choose download method:\n\nOK = Combined PDF (Front + Back together)\nCancel = Separate images');
    
    if (downloadMethod) {
        // Download as combined images (front + back side by side)
        generatedCards.forEach((card, index) => {
            setTimeout(() => {
                const combinedCanvas = document.createElement('canvas');
                const img1 = new Image();
                const img2 = new Image();
                
                img1.onload = function() {
                    img2.onload = function() {
                        // Set canvas to hold both images side by side
                        combinedCanvas.width = img1.width + img2.width + 40; // 40px gap
                        combinedCanvas.height = Math.max(img1.height, img2.height) + 40;
                        
                        const ctx = combinedCanvas.getContext('2d');
                        ctx.fillStyle = '#ffffff';
                        ctx.fillRect(0, 0, combinedCanvas.width, combinedCanvas.height);
                        
                        // Draw front
                        ctx.fillStyle = '#333';
                        ctx.font = 'bold 20px Arial';
                        ctx.fillText('FRONT', 20, 30);
                        ctx.drawImage(img1, 20, 50);
                        
                        // Draw back
                        ctx.fillText('BACK', img1.width + 60, 30);
                        ctx.drawImage(img2, img1.width + 60, 50);
                        
                        // Download
                        const link = document.createElement('a');
                        link.href = combinedCanvas.toDataURL('image/png');
                        link.download = `${card.employee.firstName}_${card.employee.lastName}_Badge_${card.employee.badge}.png`;
                        link.click();
                    };
                    img2.src = card.back;
                };
                img1.src = card.front;
            }, index * 500);
        });
        
        alert(`✅ Downloading ${generatedCards.length} combined cards...`);
    } else {
        // Download separate images
        generatedCards.forEach((card, index) => {
            setTimeout(() => {
                // Download front
                const frontLink = document.createElement('a');
                frontLink.href = card.front;
                frontLink.download = `${card.employee.firstName}_${card.employee.lastName}_${card.employee.badge}_FRONT.png`;
                frontLink.click();
                
                // Download back
                setTimeout(() => {
                    const backLink = document.createElement('a');
                    backLink.href = card.back;
                    backLink.download = `${card.employee.firstName}_${card.employee.lastName}_${card.employee.badge}_BACK.png`;
                    backLink.click();
                }, 100);
            }, index * 400);
        });
        
        alert(`✅ Downloading ${generatedCards.length * 2} images...`);
    }
}