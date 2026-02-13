import ExcelJS from 'exceljs';
import path from 'path';

/**
 * Playground script to test template cloning step by step
 * This helps us debug the page setup issues in isolation
 */

async function playgroundTemplateClone() {
    console.log('🧪 PLAYGROUND: Testing Template Cloning Step by Step');
    console.log('='.repeat(60));

    try {
        // Step 1: Load the template
        console.log('\n📋 STEP 1: Loading template...');
        const templatePath = path.join(__dirname, 'templates', 'template.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        
        console.log(`✅ Template loaded: ${templatePath}`);
        console.log(`📄 Total sheets in template: ${workbook.worksheets.length}`);
        console.log(`📝 Sheet names: ${workbook.worksheets.slice(0, 5).map(ws => ws.name).join(', ')}...`);

        // Step 2: Check different template sheets
        console.log('\n📋 STEP 2: Analyzing template sheets...');
        workbook.worksheets.slice(0, 5).forEach((sheet, i) => {
            console.log(`  Sheet ${i}: ${sheet.name}`);
            console.log(`    Page Setup: ${JSON.stringify(sheet.pageSetup)}`);
        });

        // Step 3: Choose the best template sheet
        console.log('\n📋 STEP 3: Selecting template sheet...');
        let sourceSheet = workbook.getWorksheet('COPY FORM (24)');
        if (!sourceSheet) {
            sourceSheet = workbook.worksheets[0]; // Fallback to first sheet
            console.log(`⚠️ Fallback to first sheet: ${sourceSheet.name}`);
        } else {
            console.log(`✅ Using source sheet: ${sourceSheet.name}`);
        }

        console.log(`📐 Source page setup:`, sourceSheet.pageSetup);

        // Step 4: Create a new workbook with just one cloned sheet
        console.log('\n📋 STEP 4: Creating clean workbook with one cloned sheet...');
        const newWorkbook = new ExcelJS.Workbook();
        
        // Remove default sheet and add cloned sheet
        newWorkbook.removeWorksheet(1); // Remove default sheet
        
        // Clone the template sheet
        const clonedSheet = newWorkbook.addWorksheet('B-128');
        
        // Copy all data from source sheet to cloned sheet
        sourceSheet.eachRow((row, rowNum) => {
            const newRow = clonedSheet.getRow(rowNum);
            row.eachCell((cell, colNum) => {
                const newCell = newRow.getCell(colNum);
                newCell.value = cell.value;
                // Copy styling
                if (cell.style) {
                    newCell.style = { ...cell.style };
                }
            });
            newRow.commit();
        });

        console.log(`✅ Cloned sheet created: ${clonedSheet.name}`);

        // Step 5: Apply critical page setup
        console.log('\n📋 STEP 5: Applying page setup...');
        console.log(`📐 Before:`, clonedSheet.pageSetup);

        // Apply the critical settings for single page
        clonedSheet.pageSetup = {
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
            scale: undefined, // Let Excel auto-calculate
            orientation: 'portrait',
            paperSize: 9 // A4
        };

        console.log(`📐 After:`, clonedSheet.pageSetup);

        // Step 6: Add some test data
        console.log('\n📋 STEP 6: Adding test data...');
        clonedSheet.getCell('D4').value = 'JA RODDING-BUTTS CRUSHING PLANT';
        clonedSheet.getCell('D5').value = 'Emirates Global Aluminium';
        clonedSheet.getCell('J4').value = 'BG-J082-850-9-QAC-SSDIR-028';
        clonedSheet.getCell('J11').value = 'B-128 Rev-03';

        console.log('✅ Test data added');

        // Step 7: Save the playground file
        console.log('\n📋 STEP 7: Saving playground file...');
        const outputPath = path.join(__dirname, 'generated', 'playground-single-sheet.xlsx');
        await newWorkbook.xlsx.writeFile(outputPath);

        console.log(`✅ Playground file saved: ${outputPath}`);

        // Step 8: Verify the saved file
        console.log('\n📋 STEP 8: Verifying saved file...');
        const verifyWorkbook = new ExcelJS.Workbook();
        await verifyWorkbook.xlsx.readFile(outputPath);
        const verifySheet = verifyWorkbook.worksheets[0];

        console.log(`📊 Final verification:`);
        console.log(`  Sheets: ${verifyWorkbook.worksheets.length} (should be 1)`);
        console.log(`  Sheet name: ${verifySheet.name}`);
        console.log(`  Page setup:`, verifySheet.pageSetup);
        console.log(`  Data sample: ${verifySheet.getCell('D4').value}`);

        console.log('\n🎯 PLAYGROUND TEST COMPLETE!');
        return outputPath;

    } catch (error) {
        console.error('❌ Playground error:', error);
        throw error;
    }
}

// Run the playground
if (require.main === module) {
    playgroundTemplateClone()
        .then((path) => {
            console.log(`\n✅ Success! Check the file: ${path}`);
        })
        .catch((error) => {
            console.error('❌ Playground failed:', error);
        });
}

export { playgroundTemplateClone };