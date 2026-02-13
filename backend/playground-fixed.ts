import ExcelJS from 'exceljs';
import path from 'path';

/**
 * FIXED playground - Copy the COMPLETE page setup from template
 */

async function playgroundFixedClone() {
    console.log('🔧 FIXED PLAYGROUND: Copying Complete Page Setup');
    console.log('='.repeat(60));

    try {
        // Step 1: Load template and get the perfect sheet
        console.log('\n📋 STEP 1: Loading template...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        
        const sourceSheet = workbook.getWorksheet('COPY FORM (24)');
        console.log(`✅ Using source sheet: ${sourceSheet!.name}`);
        console.log(`📐 Source page setup:`, sourceSheet!.pageSetup);

        // Step 2: Create new workbook and clone sheet WITH page setup
        console.log('\n📋 STEP 2: Cloning sheet with COMPLETE page setup...');
        const newWorkbook = new ExcelJS.Workbook();
        newWorkbook.removeWorksheet(1); // Remove default sheet
        
        const clonedSheet = newWorkbook.addWorksheet('B-128');

        // Copy all data
        sourceSheet!.eachRow((row, rowNum) => {
            const newRow = clonedSheet.getRow(rowNum);
            row.eachCell((cell, colNum) => {
                const newCell = newRow.getCell(colNum);
                newCell.value = cell.value;
                if (cell.style) {
                    newCell.style = { ...cell.style };
                }
            });
            newRow.commit();
        });

        // 🎯 CRITICAL FIX: Copy the COMPLETE page setup object
        clonedSheet.pageSetup = { ...sourceSheet!.pageSetup };
        
        console.log(`📐 Cloned page setup:`, clonedSheet.pageSetup);

        // Step 3: Add test data
        console.log('\n📋 STEP 3: Adding test data...');
        clonedSheet.getCell('D4').value = 'JA RODDING-BUTTS CRUSHING PLANT';
        clonedSheet.getCell('D5').value = 'Emirates Global Aluminium';
        clonedSheet.getCell('J4').value = 'BG-J082-850-9-QAC-SSDIR-028';
        clonedSheet.getCell('J11').value = 'B-128 Rev-03';

        // Step 4: Save and verify
        console.log('\n📋 STEP 4: Saving fixed file...');
        const outputPath = path.join(__dirname, 'generated', 'playground-FIXED.xlsx');
        await newWorkbook.xlsx.writeFile(outputPath);

        // Verify
        const verifyWorkbook = new ExcelJS.Workbook();
        await verifyWorkbook.xlsx.readFile(outputPath);
        const verifySheet = verifyWorkbook.worksheets[0];

        console.log(`\n📊 FIXED FILE VERIFICATION:`);
        console.log(`  Sheets: ${verifyWorkbook.worksheets.length}`);
        console.log(`  Sheet name: ${verifySheet.name}`);
        console.log(`  Scale: ${verifySheet.pageSetup.scale} (should be 37)`);
        console.log(`  FitToWidth: ${verifySheet.pageSetup.fitToWidth} (should be 1)`);
        console.log(`  FitToHeight: ${verifySheet.pageSetup.fitToHeight} (should be 0)`);
        console.log(`  FitToPage: ${verifySheet.pageSetup.fitToPage} (should be true)`);

        console.log('\n🎯 FIXED PLAYGROUND COMPLETE!');
        return outputPath;

    } catch (error) {
        console.error('❌ Fixed playground error:', error);
        throw error;
    }
}

// Run the fixed playground
if (require.main === module) {
    playgroundFixedClone()
        .then((path) => {
            console.log(`\n✅ Fixed Success! Check the file: ${path}`);
        })
        .catch((error) => {
            console.error('❌ Fixed playground failed:', error);
        });
}

export { playgroundFixedClone };