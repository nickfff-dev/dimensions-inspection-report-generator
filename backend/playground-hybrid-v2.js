const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

/**
 * HYBRID APPROACH V2 - Better template preservation
 */

async function testHybridV2() {
    console.log('🔥 HYBRID APPROACH V2: Enhanced Template Preservation');
    console.log('='.repeat(65));

    try {
        // Step 1: File-system copy (preserves everything)
        console.log('\n📋 STEP 1: File-system copy...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const copyPath = path.join(__dirname, 'generated', 'hybrid-v2-copy.xlsx');
        
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

        // Step 3: Better sheet duplication approach
        console.log('\n📋 STEP 3: Creating inspection sheet...');
        
        // Remove all sheets first except template
        const allSheets = workbook.worksheets.slice();
        allSheets.forEach(ws => {
            if (ws.name !== 'COPY FORM (24)') {
                workbook.removeWorksheet(ws.name);
            }
        });
        
        // Now duplicate the template sheet properly
        const newSheet = workbook.addWorksheet('B-128');
        
        // Copy all data, styles, and formatting
        templateSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const newRow = newSheet.getRow(rowNumber);
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const newCell = newRow.getCell(colNumber);
                
                // Deep copy cell properties
                newCell.value = cell.value;
                if (cell.style) {
                    newCell.style = JSON.parse(JSON.stringify(cell.style));
                }
                if (cell.formula) {
                    newCell.formula = cell.formula;
                }
                if (cell.hyperlink) {
                    newCell.hyperlink = cell.hyperlink;
                }
            });
            
            // Copy row properties
            if (row.height) newRow.height = row.height;
            if (row.hidden) newRow.hidden = row.hidden;
            if (row.outlineLevel) newRow.outlineLevel = row.outlineLevel;
            
            newRow.commit();
        });

        // Copy column properties
        templateSheet.columns.forEach((column, index) => {
            const newColumn = newSheet.getColumn(index + 1);
            if (column.width) newColumn.width = column.width;
            if (column.hidden) newColumn.hidden = column.hidden;
            if (column.outlineLevel) newColumn.outlineLevel = column.outlineLevel;
        });

        // Copy page setup exactly as is
        if (templateSheet.pageSetup) {
            console.log('📐 Copying exact page setup from template...');
            newSheet.pageSetup = JSON.parse(JSON.stringify(templateSheet.pageSetup));
            
            console.log(`   Scale: ${newSheet.pageSetup.scale}`);
            console.log(`   FitToWidth: ${newSheet.pageSetup.fitToWidth}`);
            console.log(`   FitToHeight: ${newSheet.pageSetup.fitToHeight}`);
            console.log(`   PrintArea: ${newSheet.pageSetup.printArea}`);
        }

        // Copy any merged cells
        if (templateSheet.model.merges) {
            templateSheet.model.merges.forEach(merge => {
                newSheet.mergeCells(merge);
            });
            console.log(`✅ Copied ${templateSheet.model.merges.length} merged cell ranges`);
        }

        console.log(`✅ Sheet created: ${newSheet.name}`);

        // Step 4: Populate with data
        console.log('\n📋 STEP 4: Adding inspection data...');
        newSheet.getCell('D4').value = 'JA RODDING-BUTTS CRUSHING PLANT';
        newSheet.getCell('D5').value = 'Emirates Global Aluminium';
        newSheet.getCell('J4').value = 'BG-J082-850-9-QAC-SSDIR-028';
        newSheet.getCell('J11').value = 'B-128 Rev-03';
        console.log('✅ Data populated');

        // Step 5: Remove template sheet
        console.log('\n📋 STEP 5: Removing template sheet...');
        workbook.removeWorksheet('COPY FORM (24)');
        console.log(`📊 Final sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);

        // Step 6: Save
        console.log('\n📋 STEP 6: Saving result...');
        const outputPath = path.join(__dirname, 'generated', 'playground-hybrid-v2.xlsx');
        await workbook.xlsx.writeFile(outputPath);
        console.log(`✅ File saved: ${outputPath}`);

        // Step 7: Verification
        console.log('\n📋 STEP 7: Verification...');
        const verifyWorkbook = new ExcelJS.Workbook();
        await verifyWorkbook.xlsx.readFile(outputPath);
        const verifySheet = verifyWorkbook.worksheets[0];
        
        console.log(`📊 VERIFICATION:`);
        console.log(`  Total sheets: ${verifyWorkbook.worksheets.length} (should be 1)`);
        console.log(`  Sheet name: ${verifySheet.name}`);
        console.log(`  Sample data D4: ${verifySheet.getCell('D4').value}`);
        console.log(`  Sample data J4: ${verifySheet.getCell('J4').value}`);
        console.log(`  Page setup scale: ${verifySheet.pageSetup?.scale}`);
        console.log(`  Page setup fitToWidth: ${verifySheet.pageSetup?.fitToWidth}`);
        console.log(`  Page setup printArea: ${verifySheet.pageSetup?.printArea}`);

        console.log('\n🎯 HYBRID V2 COMPLETE!');
        return outputPath;

    } catch (error) {
        console.error('❌ Hybrid V2 error:', error);
        throw error;
    }
}

if (require.main === module) {
    testHybridV2()
        .then((path) => {
            console.log(`\n✅ Success! Check: ${path}`);
        })
        .catch((error) => {
            console.error('❌ Failed:', error);
        });
}

module.exports = { testHybridV2 };