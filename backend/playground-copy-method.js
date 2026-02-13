const XlsxPopulate = require("xlsx-populate");
const path = require('path');

/**
 * XLSX-POPULATE approach - Using correct copy method
 */

async function testXlsxPopulateCopy() {
    console.log('🔥 XLSX-POPULATE: Testing Copy Methods');
    console.log('='.repeat(60));

    try {
        // Step 1: Load template
        console.log('\n📋 STEP 1: Loading template...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const workbook = await XlsxPopulate.fromFileAsync(templatePath);
        
        const sourceSheet = workbook.sheet("COPY FORM (24)");
        console.log(`✅ Source sheet: ${sourceSheet.name()}`);

        // Step 2: Check available methods on sheet
        console.log('\n📋 STEP 2: Checking available sheet methods...');
        const sheetMethods = [];
        let obj = sourceSheet;
        do {
            sheetMethods.push(...Object.getOwnPropertyNames(obj));
        } while (obj = Object.getPrototypeOf(obj));
        
        const copyMethods = sheetMethods.filter(method => 
            method.toLowerCase().includes('copy') || 
            method.toLowerCase().includes('clone') ||
            method.toLowerCase().includes('duplicate')
        );
        console.log('📚 Copy-related methods:', copyMethods);

        // Step 3: Try workbook-level copy
        console.log('\n📋 STEP 3: Trying workbook copy sheet...');
        // Create new workbook
        const newWorkbook = await XlsxPopulate.fromBlankAsync();
        
        // Remove default sheet
        const defaultSheet = newWorkbook.sheets()[0];
        newWorkbook.deleteSheet(defaultSheet.name());
        
        // Method 1: Try copySheet if it exists
        try {
            if (typeof workbook.copySheet === 'function') {
                const copiedSheet = workbook.copySheet(sourceSheet, "B-128");
                console.log('✅ Used workbook.copySheet()');
            } else {
                console.log('⚠️ workbook.copySheet() not available');
            }
        } catch (err) {
            console.log('⚠️ copySheet failed:', err.message);
        }

        // Method 2: Manual copy approach
        console.log('\n📋 STEP 4: Using manual copy approach...');
        const newSheet = newWorkbook.addSheet("B-128");
        
        // Copy used range
        const usedRange = sourceSheet.usedRange();
        if (usedRange) {
            console.log(`📐 Copying range: ${usedRange.address()}`);
            
            // Get all values and copy them
            const values = usedRange.value();
            newSheet.range(usedRange.address()).value(values);
            
            console.log('✅ Values copied');
        }

        // Step 5: Add our test data
        console.log('\n📋 STEP 5: Adding test data...');
        newSheet.cell("D4").value("JA RODDING-BUTTS CRUSHING PLANT");
        newSheet.cell("D5").value("Emirates Global Aluminium");
        newSheet.cell("J4").value("BG-J082-850-9-QAC-SSDIR-028");
        newSheet.cell("J11").value("B-128 Rev-03");

        // Step 6: Save result
        console.log('\n📋 STEP 6: Saving result...');
        const outputPath = path.join(__dirname, 'generated', 'playground-copy-method.xlsx');
        await newWorkbook.toFileAsync(outputPath);

        console.log(`✅ File saved: ${outputPath}`);
        console.log('\n🎯 COPY METHOD TEST COMPLETE!');
        
        return outputPath;

    } catch (error) {
        console.error('❌ Copy method error:', error);
        throw error;
    }
}

// Run the test
if (require.main === module) {
    testXlsxPopulateCopy()
        .then((path) => {
            console.log(`\n✅ Success! Check the file: ${path}`);
        })
        .catch((error) => {
            console.error('❌ Test failed:', error);
        });
}

module.exports = { testXlsxPopulateCopy };