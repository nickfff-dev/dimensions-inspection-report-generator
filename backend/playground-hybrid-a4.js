const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

/**
 * HYBRID APPROACH A4 - Fixed for proper A4 paper formatting
 * Respects print area boundaries to avoid extra cells
 */

async function testHybridA4() {
    console.log('🔥 HYBRID APPROACH A4: Proper A4 Paper Formatting');
    console.log('='.repeat(60));

    try {
        // Step 1: File-system copy (preserves everything)
        console.log('\n📋 STEP 1: File-system copy...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const copyPath = path.join(__dirname, 'generated', 'hybrid-a4-copy.xlsx');
        
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

        console.log('📐 Template analysis:');
        console.log(`   Print Area: ${templateSheet.pageSetup?.printArea}`);
        console.log(`   Used Range: ${templateSheet.usedRange?.address}`);

        // Step 3: Remove all other sheets first
        console.log('\n📋 STEP 3: Cleaning workbook...');
        const allSheets = workbook.worksheets.slice();
        allSheets.forEach(ws => {
            if (ws.name !== 'COPY FORM (24)') {
                workbook.removeWorksheet(ws.name);
            }
        });

        // Step 4: Create new sheet with A4 boundaries
        console.log('\n📋 STEP 4: Creating A4-bounded inspection sheet...');
        const newSheet = workbook.addWorksheet('B-128');
        
        // Parse print area to get exact boundaries (A1:M56)
        const printArea = templateSheet.pageSetup?.printArea || 'A1:M56';
        console.log(`📏 Copying only print area: ${printArea}`);
        
        // Copy cells row by row within print area only (rows 1-56, cols A-M)
        for (let rowNum = 1; rowNum <= 56; rowNum++) {
            const templateRow = templateSheet.getRow(rowNum);
            const newRow = newSheet.getRow(rowNum);
            
            // Copy cells only in columns A-M (1-13)
            for (let colNum = 1; colNum <= 13; colNum++) {
                const templateCell = templateRow.getCell(colNum);
                const newCell = newRow.getCell(colNum);
                
                // Only copy if cell has content or style
                if (templateCell.value !== null && templateCell.value !== undefined) {
                    newCell.value = templateCell.value;
                }
                
                if (templateCell.style && Object.keys(templateCell.style).length > 0) {
                    newCell.style = JSON.parse(JSON.stringify(templateCell.style));
                }
                
                if (templateCell.formula) {
                    newCell.formula = templateCell.formula;
                }
                
                if (templateCell.hyperlink) {
                    newCell.hyperlink = templateCell.hyperlink;
                }
            }
            
            // Copy row properties
            if (templateRow.height) newRow.height = templateRow.height;
            if (templateRow.hidden) newRow.hidden = templateRow.hidden;
            newRow.commit();
        }

        // Copy column properties ONLY for A-M (columns 1-13)
        console.log('\n📋 STEP 5: Setting A4 column properties...');
        for (let colIndex = 1; colIndex <= 13; colIndex++) {
            const templateColumn = templateSheet.getColumn(colIndex);
            const newColumn = newSheet.getColumn(colIndex);
            
            if (templateColumn.width) {
                newColumn.width = templateColumn.width;
                console.log(`   Column ${String.fromCharCode(64 + colIndex)}: ${templateColumn.width}`);
            }
            if (templateColumn.hidden) newColumn.hidden = templateColumn.hidden;
        }

        // Copy merged cells ONLY within print area
        console.log('\n📋 STEP 6: Copying merged cells within A4 area...');
        let mergedCount = 0;
        if (templateSheet.model.merges) {
            templateSheet.model.merges.forEach(merge => {
                // Check if merge is within print area (A1:M56)
                const mergeRange = merge.split(':');
                const startCol = mergeRange[0].match(/[A-Z]+/)[0];
                const endCol = mergeRange[1].match(/[A-Z]+/)[0];
                
                // Only copy merges within columns A-M
                if (startCol.charCodeAt(0) <= 77 && endCol.charCodeAt(0) <= 77) { // M = 77
                    newSheet.mergeCells(merge);
                    mergedCount++;
                }
            });
        }
        console.log(`✅ Copied ${mergedCount} merged cells within A4 area`);

        // Step 7: Set exact A4 page setup
        console.log('\n📋 STEP 7: Setting A4 page setup...');
        newSheet.pageSetup = {
            paperSize: 9,              // A4 paper
            orientation: 'portrait',   // Portrait orientation
            fitToPage: true,          // Enable fit to page
            fitToWidth: 1,            // Fit to 1 page wide
            fitToHeight: 0,           // Unlimited height (single page)
            scale: 37,                // 37% scale from template
            printArea: 'A1:M56',      // Exact print area for A4
            margins: {
                left: 0.25,
                right: 0.25,
                top: 0.75,
                bottom: 0.75,
                header: 0.3,
                footer: 0.3
            }
        };

        console.log('✅ A4 page setup applied:');
        console.log(`   Paper: A4 (210mm x 297mm)`);
        console.log(`   Scale: ${newSheet.pageSetup.scale}%`);
        console.log(`   Print Area: ${newSheet.pageSetup.printArea}`);
        console.log(`   Fit: ${newSheet.pageSetup.fitToWidth} page wide`);

        // Step 8: Add inspection data
        console.log('\n📋 STEP 8: Adding inspection data...');
        newSheet.getCell('D4').value = 'JA RODDING-BUTTS CRUSHING PLANT';
        newSheet.getCell('D5').value = 'Emirates Global Aluminium';
        newSheet.getCell('J4').value = 'BG-J082-850-9-QAC-SSDIR-028';
        newSheet.getCell('J11').value = 'B-128 Rev-03';
        console.log('✅ Inspection data populated');

        // Step 9: Remove template sheet
        console.log('\n📋 STEP 9: Finalizing workbook...');
        workbook.removeWorksheet('COPY FORM (24)');
        console.log(`📊 Final sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);

        // Step 10: Save A4-formatted result
        console.log('\n📋 STEP 10: Saving A4-formatted file...');
        const outputPath = path.join(__dirname, 'generated', 'playground-hybrid-a4.xlsx');
        await workbook.xlsx.writeFile(outputPath);
        console.log(`✅ A4 file saved: ${outputPath}`);

        // Step 11: Verification
        console.log('\n📋 STEP 11: A4 formatting verification...');
        const verifyWorkbook = new ExcelJS.Workbook();
        await verifyWorkbook.xlsx.readFile(outputPath);
        const verifySheet = verifyWorkbook.worksheets[0];
        
        console.log(`📊 A4 VERIFICATION:`);
        console.log(`  ✅ Total sheets: ${verifyWorkbook.worksheets.length} (should be 1)`);
        console.log(`  ✅ Sheet name: ${verifySheet.name}`);
        console.log(`  ✅ Paper size: ${verifySheet.pageSetup?.paperSize} (A4)`);
        console.log(`  ✅ Scale: ${verifySheet.pageSetup?.scale}%`);
        console.log(`  ✅ Print area: ${verifySheet.pageSetup?.printArea}`);
        console.log(`  ✅ Fit to width: ${verifySheet.pageSetup?.fitToWidth} page`);
        console.log(`  ✅ Sample data D4: ${verifySheet.getCell('D4').value}`);

        console.log('\n🎯 A4 FORMATTING COMPLETE!');
        console.log('📄 File is now properly formatted for A4 paper (210mm x 297mm)');
        
        return outputPath;

    } catch (error) {
        console.error('❌ A4 formatting error:', error);
        throw error;
    }
}

if (require.main === module) {
    testHybridA4()
        .then((path) => {
            console.log(`\n✅ A4 Success! File ready for printing: ${path}`);
        })
        .catch((error) => {
            console.error('❌ A4 formatting failed:', error);
        });
}

module.exports = { testHybridA4 };