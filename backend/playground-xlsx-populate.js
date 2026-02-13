const XlsxPopulate = require("xlsx-populate");
const path = require('path');

/**
 * XLSX-POPULATE approach - Using full deep clone
 */

async function playgroundXlsxPopulate() {
    console.log('🔥 XLSX-POPULATE PLAYGROUND: Full Deep Clone');
    console.log('='.repeat(60));

    try {
        // Step 1: Load template with xlsx-populate
        console.log('\n📋 STEP 1: Loading template with xlsx-populate...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const workbook = await XlsxPopulate.fromFileAsync(templatePath);
        
        console.log(`✅ Template loaded: ${templatePath}`);
        console.log(`📄 Total sheets: ${workbook.sheets().length}`);
        console.log(`📝 Sheet names: ${workbook.sheets().map(s => s.name()).join(', ')}`);

        // Step 2: Get the source template sheet
        console.log('\n📋 STEP 2: Getting source template sheet...');
        const sourceSheet = workbook.sheet("COPY FORM (24)");
        console.log(`✅ Using source sheet: ${sourceSheet.name()}`);

        // Check some source sheet properties
        console.log(`📐 Source sheet info:`);
        const usedRange = sourceSheet.usedRange();
        if (usedRange) {
            console.log(`  Rows: ${usedRange.endCell().rowNumber()}`);
            console.log(`  Cols: ${usedRange.endCell().columnNumber()}`);
        }

        // Step 3: DEEP CLONE the sheet
        console.log('\n📋 STEP 3: Deep cloning sheet...');
        const clonedSheet = sourceSheet.clone("B-128");
        console.log(`✅ Sheet cloned: ${clonedSheet.name()}`);

        // Step 4: Remove all original sheets except our clone
        console.log('\n📋 STEP 4: Removing original template sheets...');
        const sheetsToRemove = workbook.sheets().filter(sheet => sheet.name() !== "B-128");
        console.log(`🗑️ Will remove ${sheetsToRemove.length} sheets...`);
        
        sheetsToRemove.forEach(sheet => {
            try {
                workbook.deleteSheet(sheet.name());
                console.log(`🗑️ Removed: ${sheet.name()}`);
            } catch (err) {
                console.log(`⚠️ Could not remove: ${sheet.name()}`);
            }
        });

        console.log(`📊 Final sheets: ${workbook.sheets().map(s => s.name()).join(', ')}`);

        // Step 5: Add test data to the cloned sheet
        console.log('\n📋 STEP 5: Adding test data...');
        const targetSheet = workbook.sheet("B-128");
        
        targetSheet.cell("D4").value("JA RODDING-BUTTS CRUSHING PLANT");
        targetSheet.cell("D5").value("Emirates Global Aluminium");
        targetSheet.cell("J4").value("BG-J082-850-9-QAC-SSDIR-028");
        targetSheet.cell("J11").value("B-128 Rev-03");
        
        console.log('✅ Test data added');

        // Step 6: Save the result
        console.log('\n📋 STEP 6: Saving xlsx-populate result...');
        const outputPath = path.join(__dirname, 'generated', 'playground-xlsx-populate.xlsx');
        await workbook.toFileAsync(outputPath);

        console.log(`✅ File saved: ${outputPath}`);

        // Step 7: Verification (load with xlsx-populate to check)
        console.log('\n📋 STEP 7: Verifying result...');
        const verifyWorkbook = await XlsxPopulate.fromFileAsync(outputPath);
        const verifySheet = verifyWorkbook.sheets()[0];

        console.log(`📊 VERIFICATION:`);
        console.log(`  Total sheets: ${verifyWorkbook.sheets().length} (should be 1)`);
        console.log(`  Sheet name: ${verifySheet.name()}`);
        console.log(`  Sample data D4: ${verifySheet.cell("D4").value()}`);
        console.log(`  Sample data D5: ${verifySheet.cell("D5").value()}`);
        console.log(`  Sample data J4: ${verifySheet.cell("J4").value()}`);

        console.log('\n🎯 XLSX-POPULATE PLAYGROUND COMPLETE!');
        return outputPath;

    } catch (error) {
        console.error('❌ xlsx-populate playground error:', error);
        throw error;
    }
}

// Run the xlsx-populate playground
if (require.main === module) {
    playgroundXlsxPopulate()
        .then((path) => {
            console.log(`\n✅ xlsx-populate Success! Check the file: ${path}`);
        })
        .catch((error) => {
            console.error('❌ xlsx-populate playground failed:', error);
        });
}

module.exports = { playgroundXlsxPopulate };