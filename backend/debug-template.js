const ExcelJS = require('exceljs');
const path = require('path');

/**
 * DEBUG TEMPLATE - Analyze what's causing extra cells
 */

async function debugTemplate() {
    console.log('🔍 DEBUGGING TEMPLATE: Finding source of extra cells');
    console.log('='.repeat(60));

    try {
        // Load original template
        console.log('\n📋 STEP 1: Loading original template...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        
        const templateSheet = workbook.getWorksheet('COPY FORM (24)');
        console.log(`✅ Template loaded: ${templateSheet.name}`);

        // Check dimensions
        console.log('\n📋 STEP 2: Analyzing template dimensions...');
        console.log(`📏 Dimensions: ${templateSheet.dimensions}`);
        console.log(`📐 Row count: ${templateSheet.rowCount}`);
        console.log(`📐 Column count: ${templateSheet.columnCount}`);
        console.log(`📐 Actual row count: ${templateSheet.actualRowCount}`);
        console.log(`📐 Actual column count: ${templateSheet.actualColumnCount}`);

        // Check used range
        const usedRange = templateSheet.usedRange;
        if (usedRange) {
            console.log(`📊 Used range: ${usedRange.address}`);
        } else {
            console.log(`📊 Used range: Not defined`);
        }

        // Check page setup
        console.log('\n📋 STEP 3: Page setup analysis...');
        const pageSetup = templateSheet.pageSetup;
        console.log(`📄 Print area: ${pageSetup?.printArea}`);
        console.log(`📄 Paper size: ${pageSetup?.paperSize}`);
        console.log(`📄 Scale: ${pageSetup?.scale}`);

        // Check columns beyond M (13)
        console.log('\n📋 STEP 4: Checking columns beyond M...');
        let foundDataBeyondM = false;
        for (let col = 14; col <= 30; col++) { // Check columns N-AD
            const column = templateSheet.getColumn(col);
            const columnLetter = String.fromCharCode(64 + col); // Convert to letter
            
            // Check if column has any properties set
            if (column.width !== undefined || column.hidden !== undefined) {
                console.log(`⚠️ Column ${columnLetter} (${col}): width=${column.width}, hidden=${column.hidden}`);
                foundDataBeyondM = true;
            }
            
            // Check if any cells in this column have data
            for (let row = 1; row <= 56; row++) {
                const cell = templateSheet.getCell(row, col);
                if (cell.value !== null && cell.value !== undefined) {
                    console.log(`⚠️ Cell ${columnLetter}${row}: ${cell.value}`);
                    foundDataBeyondM = true;
                }
                if (cell.style && Object.keys(cell.style).length > 0) {
                    console.log(`⚠️ Cell ${columnLetter}${row} has styling`);
                    foundDataBeyondM = true;
                }
            }
        }

        if (!foundDataBeyondM) {
            console.log(`✅ No data found beyond column M`);
        }

        // Check merged cells beyond M
        console.log('\n📋 STEP 5: Checking merged cells...');
        if (templateSheet.model.merges) {
            const mergesBeyondM = templateSheet.model.merges.filter(merge => {
                const endCol = merge.split(':')[1].match(/[A-Z]+/)[0];
                return endCol.charCodeAt(0) > 77; // M = 77
            });
            
            if (mergesBeyondM.length > 0) {
                console.log(`⚠️ Found ${mergesBeyondM.length} merged cells beyond column M:`);
                mergesBeyondM.forEach(merge => console.log(`   ${merge}`));
            } else {
                console.log(`✅ No merged cells beyond column M`);
            }
        }

        // Check if there are any drawings/images that might extend the range
        console.log('\n📋 STEP 6: Checking drawings/images...');
        if (templateSheet.model.drawing) {
            console.log(`⚠️ Sheet has drawings that might extend range`);
            console.log(`   Drawing anchors might be causing extended range`);
        } else {
            console.log(`✅ No drawings found`);
        }

        console.log('\n🎯 TEMPLATE DEBUG COMPLETE!');

    } catch (error) {
        console.error('❌ Debug error:', error);
        throw error;
    }
}

if (require.main === module) {
    debugTemplate()
        .then(() => {
            console.log(`\n✅ Debug complete!`);
        })
        .catch((error) => {
            console.error('❌ Debug failed:', error);
        });
}

module.exports = { debugTemplate };