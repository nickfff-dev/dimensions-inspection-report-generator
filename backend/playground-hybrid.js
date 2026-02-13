const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

/**
 * HYBRID APPROACH - File-system copying + ExcelJS data population
 * Based on the working TypeScript service approach
 */

async function testHybridApproach() {
    console.log('🔥 HYBRID APPROACH: File-system copying + ExcelJS');
    console.log('='.repeat(60));

    try {
        // Step 1: File-system copy of template (preserves everything)
        console.log('\n📋 STEP 1: File-system copy of template...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const copyPath = path.join(__dirname, 'generated', 'hybrid-copy.xlsx');
        
        // Ensure generated directory exists
        const generatedDir = path.dirname(copyPath);
        if (!fs.existsSync(generatedDir)) {
            fs.mkdirSync(generatedDir, { recursive: true });
        }
        
        // Copy file
        fs.copyFileSync(templatePath, copyPath);
        console.log(`✅ Template copied: ${templatePath} → ${copyPath}`);

        // Step 2: Load the copied file with ExcelJS for data population
        console.log('\n📋 STEP 2: Loading copy with ExcelJS...');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(copyPath);
        
        console.log(`📄 Total sheets: ${workbook.worksheets.length}`);
        console.log(`📝 Sheet names: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);

        // Step 3: Get the template sheet and duplicate it
        console.log('\n📋 STEP 3: Duplicating template sheet...');
        const templateSheet = workbook.getWorksheet('COPY FORM (24)');
        if (!templateSheet) {
            throw new Error('Template sheet "COPY FORM (24)" not found');
        }
        
        // Create new sheet by copying the template
        const newSheet = workbook.addWorksheet('B-128');
        
        // Copy all data from template to new sheet
        templateSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const newRow = newSheet.getRow(rowNumber);
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const newCell = newRow.getCell(colNumber);
                
                // Copy value
                newCell.value = cell.value;
                
                // Copy style if it exists
                if (cell.style) {
                    newCell.style = { ...cell.style };
                }
            });
            newRow.commit();
        });

        // Copy column widths
        templateSheet.columns.forEach((column, index) => {
            if (column.width) {
                newSheet.getColumn(index + 1).width = column.width;
            }
        });

        // Copy row heights
        for (let i = 1; i <= templateSheet.rowCount; i++) {
            const templateRow = templateSheet.getRow(i);
            if (templateRow.height) {
                newSheet.getRow(i).height = templateRow.height;
            }
        }

        // Copy page setup from template - preserve original template settings
        if (templateSheet.pageSetup) {
            console.log(`📐 Original template page setup:`, templateSheet.pageSetup);
            newSheet.pageSetup = {
                ...templateSheet.pageSetup
            };
        } else {
            // Fallback page setup based on template analysis
            newSheet.pageSetup = {
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 0,
                orientation: 'portrait',
                paperSize: 9,
                scale: 37,
                margins: {
                    left: 0.7,
                    right: 0.7,
                    top: 0.75,
                    bottom: 0.75,
                    header: 0.3,
                    footer: 0.3
                }
            };
        }

        console.log(`✅ Sheet duplicated: ${newSheet.name}`);
        console.log(`📐 Page setup applied: fitToWidth=1, fitToHeight=0`);

        // Step 4: Populate with inspection data
        console.log('\n📋 STEP 4: Populating with inspection data...');
        newSheet.getCell('D4').value = 'JA RODDING-BUTTS CRUSHING PLANT';
        newSheet.getCell('D5').value = 'Emirates Global Aluminium';
        newSheet.getCell('J4').value = 'BG-J082-850-9-QAC-SSDIR-028';
        newSheet.getCell('J11').value = 'B-128 Rev-03';
        
        console.log('✅ Inspection data populated');

        // Step 5: Remove all original template sheets
        console.log('\n📋 STEP 5: Removing template sheets...');
        const sheetsToRemove = [];
        workbook.eachSheet((worksheet) => {
            if (worksheet.name !== 'B-128') {
                sheetsToRemove.push(worksheet.name);
            }
        });
        
        console.log(`🗑️ Will remove ${sheetsToRemove.length} template sheets...`);
        sheetsToRemove.forEach(sheetName => {
            try {
                workbook.removeWorksheet(sheetName);
                console.log(`🗑️ Removed: ${sheetName}`);
            } catch (err) {
                console.log(`⚠️ Could not remove: ${sheetName} - ${err.message}`);
            }
        });

        console.log(`📊 Final sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);

        // Step 6: Save the final result
        console.log('\n📋 STEP 6: Saving final result...');
        const outputPath = path.join(__dirname, 'generated', 'playground-hybrid.xlsx');
        await workbook.xlsx.writeFile(outputPath);

        console.log(`✅ File saved: ${outputPath}`);

        // Step 7: Verification
        console.log('\n📋 STEP 7: Verifying result...');
        const verifyWorkbook = new ExcelJS.Workbook();
        await verifyWorkbook.xlsx.readFile(outputPath);
        
        const verifySheet = verifyWorkbook.worksheets[0];
        
        console.log(`📊 VERIFICATION:`);
        console.log(`  Total sheets: ${verifyWorkbook.worksheets.length} (should be 1)`);
        console.log(`  Sheet name: ${verifySheet.name}`);
        console.log(`  Sample data D4: ${verifySheet.getCell('D4').value}`);
        console.log(`  Sample data D5: ${verifySheet.getCell('D5').value}`);
        console.log(`  Sample data J4: ${verifySheet.getCell('J4').value}`);
        console.log(`  Page setup fitToWidth: ${verifySheet.pageSetup?.fitToWidth}`);
        console.log(`  Page setup fitToHeight: ${verifySheet.pageSetup?.fitToHeight}`);

        console.log('\n🎯 HYBRID APPROACH COMPLETE!');
        return outputPath;

    } catch (error) {
        console.error('❌ Hybrid approach error:', error);
        throw error;
    }
}

// Run the hybrid test
if (require.main === module) {
    testHybridApproach()
        .then((path) => {
            console.log(`\n✅ Hybrid Success! Check the file: ${path}`);
        })
        .catch((error) => {
            console.error('❌ Hybrid test failed:', error);
        });
}

module.exports = { testHybridApproach };