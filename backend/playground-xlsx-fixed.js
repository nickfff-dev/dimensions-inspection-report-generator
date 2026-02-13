const XlsxPopulate = require("xlsx-populate");
const path = require('path');

/**
 * XLSX-POPULATE Fixed approach
 */

async function testXlsxPopulateFixed() {
    console.log('🔥 XLSX-POPULATE FIXED: Proper Template Copy');
    console.log('='.repeat(60));

    try {
        // Step 1: Load template
        console.log('\n📋 STEP 1: Loading template...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const workbook = await XlsxPopulate.fromFileAsync(templatePath);
        
        const sourceSheet = workbook.sheet("COPY FORM (24)");
        console.log(`✅ Source sheet: ${sourceSheet.name()}`);

        // Step 2: Create new workbook and add sheet first
        console.log('\n📋 STEP 2: Creating clean workbook...');
        const newWorkbook = await XlsxPopulate.fromBlankAsync();
        
        // Add our sheet first, then remove default
        const newSheet = newWorkbook.addSheet("B-128");
        const defaultSheet = newWorkbook.sheet(0); // Get first sheet (default)
        newWorkbook.deleteSheet(defaultSheet.name());
        
        console.log(`✅ Clean workbook created with sheet: ${newSheet.name()}`);

        // Step 3: Copy ALL content from source sheet
        console.log('\n📋 STEP 3: Copying complete sheet content...');
        const usedRange = sourceSheet.usedRange();
        
        if (usedRange) {
            console.log(`📐 Copying range: ${usedRange.address()}`);
            console.log(`📏 Size: ${usedRange.endCell().rowNumber()} rows x ${usedRange.endCell().columnNumber()} cols`);
            
            // Copy values
            const values = usedRange.value();
            newSheet.range(usedRange.address()).value(values);
            
            console.log('✅ Content copied');
        }

        // Step 4: Try to copy formatting/styles if possible
        console.log('\n📋 STEP 4: Attempting to copy styles...');
        try {
            // Copy column widths if available
            for (let col = 1; col <= 20; col++) {
                try {
                    const width = sourceSheet.column(col).width();
                    if (width) {
                        newSheet.column(col).width(width);
                    }
                } catch (e) {
                    // Continue if column width fails
                }
            }
            console.log('✅ Column widths copied');
        } catch (e) {
            console.log('⚠️ Style copying partially failed');
        }

        // Step 5: Add our specific data
        console.log('\n📋 STEP 5: Adding inspection data...');
        newSheet.cell("D4").value("JA RODDING-BUTTS CRUSHING PLANT");
        newSheet.cell("D5").value("Emirates Global Aluminium");
        newSheet.cell("J4").value("BG-J082-850-9-QAC-SSDIR-028");
        newSheet.cell("J11").value("B-128 Rev-03");

        // Step 6: Save result
        console.log('\n📋 STEP 6: Saving final result...');
        const outputPath = path.join(__dirname, 'generated', 'playground-xlsx-fixed.xlsx');
        await newWorkbook.toFileAsync(outputPath);

        // Step 7: Verification
        console.log('\n📋 STEP 7: Verifying result...');
        const verifyWorkbook = await XlsxPopulate.fromFileAsync(outputPath);
        const verifySheet = verifyWorkbook.sheets()[0];

        console.log(`📊 VERIFICATION:`);
        console.log(`  Total sheets: ${verifyWorkbook.sheets().length} (should be 1)`);
        console.log(`  Sheet name: ${verifySheet.name()}`);
        console.log(`  Sample data D4: ${verifySheet.cell("D4").value()}`);
        console.log(`  Sample data D5: ${verifySheet.cell("D5").value()}`);
        console.log(`  Sample data J4: ${verifySheet.cell("J4").value()}`);

        console.log(`✅ File saved: ${outputPath}`);
        console.log('\n🎯 XLSX-POPULATE FIXED COMPLETE!');
        
        return outputPath;

    } catch (error) {
        console.error('❌ Fixed xlsx-populate error:', error);
        throw error;
    }
}

// Run the test
if (require.main === module) {
    testXlsxPopulateFixed()
        .then((path) => {
            console.log(`\n✅ Fixed Success! Check the file: ${path}`);
        })
        .catch((error) => {
            console.error('❌ Fixed test failed:', error);
        });
}

module.exports = { testXlsxPopulateFixed };