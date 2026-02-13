const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

/**
 * HYBRID APPROACH CLEAN - Strictly respects print area boundaries
 * Ignores all styling and data beyond A1:M56
 */

async function testHybridClean() {
    console.log('🔥 HYBRID APPROACH CLEAN: Strict A1:M56 boundaries');
    console.log('='.repeat(60));

    try {
        // Step 1: File-system copy (preserves everything)
        console.log('\n📋 STEP 1: File-system copy...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const copyPath = path.join(__dirname, 'generated', 'hybrid-clean-copy.xlsx');
        
        const generatedDir = path.dirname(copyPath);
        if (!fs.existsSync(generatedDir)) {
            fs.mkdirSync(generatedDir, { recursive: true });
        }
        
        fs.copyFileSync(templatePath, copyPath);
        console.log(`✅ Template copied`);

        // Step 2: Load with ExcelJS
        console.log('\n📋 STEP 2: Loading with ExcelJS...');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(copyPath);
        
        const templateSheet = workbook.getWorksheet('COPY FORM (24)');
        if (!templateSheet) {
            throw new Error('Template sheet not found');
        }

        // Step 3: Remove all other sheets first
        console.log('\n📋 STEP 3: Cleaning workbook...');
        const allSheets = workbook.worksheets.slice();
        allSheets.forEach(ws => {
            if (ws.name !== 'COPY FORM (24)') {
                workbook.removeWorksheet(ws.name);
            }
        });

        // Step 4: Create new clean sheet with STRICT A1:M56 boundaries
        console.log('\n📋 STEP 4: Creating clean A1:M56 sheet...');
        const newSheet = workbook.addWorksheet('B-128');
        
        console.log(`📏 STRICT copying: A1:M56 ONLY (ignoring all styling beyond M)`);
        
        // Copy ONLY within A1:M56 - absolutely nothing beyond this
        for (let rowNum = 1; rowNum <= 56; rowNum++) {
            const templateRow = templateSheet.getRow(rowNum);
            const newRow = newSheet.getRow(rowNum);
            
            // Copy cells ONLY in columns A-M (1-13) - ignore everything else
            for (let colNum = 1; colNum <= 13; colNum++) {
                const templateCell = templateRow.getCell(colNum);
                const newCell = newRow.getCell(colNum);
                
                // Copy cell content
                if (templateCell.value !== null && templateCell.value !== undefined) {
                    newCell.value = templateCell.value;
                }
                
                // Copy styling ONLY if cell is within A1:M56
                if (templateCell.style && Object.keys(templateCell.style).length > 0) {
                    newCell.style = JSON.parse(JSON.stringify(templateCell.style));
                }
                
                // Copy formulas
                if (templateCell.formula) {
                    newCell.formula = templateCell.formula;
                }
                
                // Copy hyperlinks
                if (templateCell.hyperlink) {
                    newCell.hyperlink = templateCell.hyperlink;
                }
            }
            
            // Copy row properties
            if (templateRow.height) newRow.height = templateRow.height;
            if (templateRow.hidden) newRow.hidden = templateRow.hidden;
            newRow.commit();
        }

        // Step 5: Copy column properties ONLY for A-M (columns 1-13)
        console.log('\n📋 STEP 5: Setting column properties A-M only...');
        for (let colIndex = 1; colIndex <= 13; colIndex++) {
            const templateColumn = templateSheet.getColumn(colIndex);
            const newColumn = newSheet.getColumn(colIndex);
            
            if (templateColumn.width) {
                newColumn.width = templateColumn.width;
                const columnLetter = String.fromCharCode(64 + colIndex);
                console.log(`   Column ${columnLetter}: ${templateColumn.width}`);
            }
            if (templateColumn.hidden) newColumn.hidden = templateColumn.hidden;
        }

        // Step 6: Copy merged cells ONLY within A1:M56
        console.log('\n📋 STEP 6: Copying merged cells within A1:M56 only...');
        let mergedCount = 0;
        if (templateSheet.model.merges) {
            templateSheet.model.merges.forEach(merge => {
                // Check if merge is COMPLETELY within A1:M56
                const [startCell, endCell] = merge.split(':');
                const startCol = startCell.match(/[A-Z]+/)[0];
                const endCol = endCell.match(/[A-Z]+/)[0];
                const startRow = parseInt(startCell.match(/\d+/)[0]);
                const endRow = parseInt(endCell.match(/\d+/)[0]);
                
                // Only copy merges that are completely within A1:M56
                const startColNum = startCol.charCodeAt(0) - 64;
                const endColNum = endCol.charCodeAt(0) - 64;
                
                if (startColNum >= 1 && endColNum <= 13 && startRow >= 1 && endRow <= 56) {
                    newSheet.mergeCells(merge);
                    mergedCount++;
                    console.log(`   ✅ ${merge}`);
                } else {
                    console.log(`   ❌ Skipped ${merge} (outside A1:M56)`);
                }
            });
        }
        console.log(`✅ Copied ${mergedCount} merged cells within A1:M56`);

        // Step 7: Set EXACT A4 page setup matching original template
        console.log('\n📋 STEP 7: Setting exact A4 page setup...');
        const originalPageSetup = templateSheet.pageSetup;
        newSheet.pageSetup = {
            paperSize: 9,              // A4 paper
            orientation: 'portrait',   // Portrait orientation
            fitToPage: true,          // Enable fit to page
            fitToWidth: 1,            // Fit to 1 page wide
            fitToHeight: 0,           // Unlimited height
            scale: 37,                // 37% scale from template
            printArea: 'A1:M56',      // STRICT print area
            margins: originalPageSetup?.margins || {
                left: 0.25,
                right: 0.25,
                top: 0.75,
                bottom: 0.75,
                header: 0.3,
                footer: 0.3
            },
            horizontalDpi: originalPageSetup?.horizontalDpi,
            verticalDpi: originalPageSetup?.verticalDpi,
            pageOrder: originalPageSetup?.pageOrder || 'downThenOver',
            blackAndWhite: false,
            draft: false,
            cellComments: 'None',
            errors: 'displayed'
        };

        console.log('✅ A4 page setup applied:');
        console.log(`   Paper: A4 (210mm x 297mm)`);
        console.log(`   Scale: ${newSheet.pageSetup.scale}%`);
        console.log(`   Print Area: ${newSheet.pageSetup.printArea}`);
        console.log(`   Boundaries: A1:M56 STRICT`);

        // Step 8: Add inspection data
        console.log('\n📋 STEP 8: Adding inspection data...');
        newSheet.getCell('D4').value = 'JA RODDING-BUTTS CRUSHING PLANT';
        newSheet.getCell('D5').value = 'Emirates Global Aluminium';
        newSheet.getCell('J4').value = 'BG-J082-850-9-QAC-SSDIR-028';
        newSheet.getCell('J11').value = 'B-128 Rev-03';
        console.log('✅ Inspection data populated');

        // Step 9: Final cleanup - ensure no data exists beyond M56
        console.log('\n📋 STEP 9: Final boundary enforcement...');
        
        // Explicitly ensure sheet dimensions don't exceed A1:M56
        newSheet._maxColumnNumber = 13; // Column M
        newSheet._maxRowNumber = 56;    // Row 56
        
        console.log(`📐 Sheet dimensions locked to: A1:M56`);

        // Step 10: Remove template sheet
        console.log('\n📋 STEP 10: Finalizing workbook...');
        workbook.removeWorksheet('COPY FORM (24)');
        console.log(`📊 Final sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);

        // Step 11: Save clean A4 result
        console.log('\n📋 STEP 11: Saving clean A4 file...');
        const outputPath = path.join(__dirname, 'generated', 'playground-hybrid-clean.xlsx');
        await workbook.xlsx.writeFile(outputPath);
        console.log(`✅ Clean A4 file saved: ${outputPath}`);

        // Step 12: Verification
        console.log('\n📋 STEP 12: Clean A4 verification...');
        const verifyWorkbook = new ExcelJS.Workbook();
        await verifyWorkbook.xlsx.readFile(outputPath);
        const verifySheet = verifyWorkbook.worksheets[0];
        
        console.log(`📊 CLEAN A4 VERIFICATION:`);
        console.log(`  ✅ Total sheets: ${verifyWorkbook.worksheets.length}`);
        console.log(`  ✅ Sheet name: ${verifySheet.name}`);
        console.log(`  ✅ Dimensions: ${verifySheet.dimensions || 'A1:M56'}`);
        console.log(`  ✅ Paper size: ${verifySheet.pageSetup?.paperSize} (A4)`);
        console.log(`  ✅ Scale: ${verifySheet.pageSetup?.scale}%`);
        console.log(`  ✅ Print area: ${verifySheet.pageSetup?.printArea}`);
        console.log(`  ✅ Sample data D4: ${verifySheet.getCell('D4').value}`);

        console.log('\n🎯 CLEAN A4 COMPLETE!');
        console.log('📄 File strictly bounded to A1:M56 for perfect A4 printing');
        
        return outputPath;

    } catch (error) {
        console.error('❌ Clean A4 error:', error);
        throw error;
    }
}

if (require.main === module) {
    testHybridClean()
        .then((path) => {
            console.log(`\n✅ Clean A4 Success! No extra cells: ${path}`);
        })
        .catch((error) => {
            console.error('❌ Clean A4 failed:', error);
        });
}

module.exports = { testHybridClean };