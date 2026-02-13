import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import { v4 as uuidv4 } from "uuid";
import {
  InspectionFormData,
  InspectionItem,
  ExpandedSheetItem,
  BatchInfo,
  FileInfo,
} from "../types";

export class ExcelGeneratorService {
  private templatePath: string;
  private outputDir: string;
  private batchfileTemplatePath:string;
  private batches: Map<string, BatchInfo> = new Map();

  constructor() {
    this.templatePath = path.join(__dirname, "../../templates/template.xlsx");
    this.batchfileTemplatePath = path.join(__dirname, "../../templates/batchfile_template.xlsx");
    this.outputDir = path.join(__dirname, "../../generated");
    this.ensureOutputDirectory();
  }

  private ensureOutputDirectory(): void {
    if (!fs.existsSync(this.outputDir)) {
      fs.mkdirSync(this.outputDir, { recursive: true });
    }
  }

  async generateInspectionReports(
    formData: InspectionFormData
  ): Promise<{ files: string[]; batchId: string }> {
    const batchId = uuidv4();

    try {
      // Generate single filename for the batch
      const filename = this.generateBatchFilename(formData);
      const filepath = path.join(this.outputDir, filename);

      // Create single workbook with multiple sheets
      await this.createMultiSheetReport(formData, filepath);

      const fileInfo: FileInfo = {
        filename,
        originalName: filename,
        size: fs.statSync(filepath).size,
        path: filepath,
        createdAt: new Date(),
      };

      // Store batch information
      const batchInfo: BatchInfo = {
        batchId,
        files: [fileInfo],
        createdAt: new Date(),
        formData,
      };

      this.batches.set(batchId, batchInfo);

      return {
        files: [filename],
        batchId,
      };
    } catch (error) {
      console.error("Error generating Excel reports:", error);
      throw new Error(
        `Failed to generate Excel reports: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  private async createMultiSheetReport(
    formData: InspectionFormData,
    outputPath: string
  ): Promise<void> {
    console.log(
      "🚀 HYBRID APPROACH: Creating multi-sheet report using file-system copying + ExcelJS data population"
    );
    console.log(
      `📋 Original items: ${formData.items.map((item) => `${item.itemNo}(×${item.qty})`).join(", ")}`
    );

    // STEP 0: Expand items by quantity first
    const expandedItems = this.expandItemsByQuantity(formData.items);

    const tempDir = path.join(this.outputDir, "temp-" + uuidv4());

    try {
      // STEP 1: Create file-system copies of template (preserves ALL content including images)
      const templateCopies = await this.createTemplateCopies(
        expandedItems,
        tempDir
      );

      // STEP 2: Populate each temp file individually (no cloning, just data insertion)
      for (let i = 0; i < templateCopies.length; i++) {
        const { filePath, item } = templateCopies[i];
        await this.populateTemplateCopy(filePath, formData, item);
      }

      // If only one item, keep the populated temp file as the final output.
      // This preserves the template exactly (images, drawings, themes) by using a file-system copy.
      if (templateCopies.length === 1) {
        // For single item, still need to add backfile sheet
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templateCopies[0].filePath);
        
      
        
        // Save with backfile
        await workbook.xlsx.writeFile(outputPath);
        console.log(`📁 Single-item report with backfile saved: ${outputPath}`);
      } else {
        // STEP 3: Assemble final multi-sheet workbook
        await this.assembleMultiSheetWorkbook(templateCopies, outputPath);
      }
    } finally {
      // STEP 4: Cleanup temporary files
      await this.cleanupTempFiles(tempDir);
    }

    console.log(
      `🎉 HYBRID APPROACH COMPLETE: Multi-sheet report saved: ${outputPath}`
    );
    const batchfilePath = outputPath.replace('.xlsx', '-BATCHFILE.xlsx');
    await this.generateBatchfile(formData, batchfilePath);
    console.log(`📋 Batchfile generated: ${batchfilePath}`);
  }

  private generateBatchFilename(formData: InspectionFormData): string {
    const date = new Date().toISOString().split("T")[0];
    const projectName = formData.project
      .replace(/[^a-zA-Z0-9]/g, "-")
      .substring(0, 30);
    return `Inspection-Report-${projectName}-${date}.xlsx`;
  }

  // QUANTITY EXPANSION UTILITY

  /**
   * Expand items array based on quantity to create individual sheet items
   */
  private expandItemsByQuantity(items: InspectionItem[]): ExpandedSheetItem[] {
    console.log('📊 QUANTITY EXPANSION: Processing items with quantities...');
    
    const expandedItems: ExpandedSheetItem[] = [];
    let globalSequence = 1;

    items.forEach((item, _) => {
      const quantity = Math.max(1, item.qty || 1); // Ensure at least 1
      
      console.log(`  📄 Item ${item.itemNo}: quantity ${quantity} → ${quantity} sheet(s)`);
      
      for (let seqInItem = 1; seqInItem <= quantity; seqInItem++) {
        // Generate sheet name based on sequence
        const sheetName = seqInItem === 1 
          ? item.itemNo  // First sheet uses original name: "B-129"
          : `${item.itemNo} (${seqInItem - 1})`; // Subsequent: "B-129 (1)", "B-129 (2)"

        const expandedItem: ExpandedSheetItem = {
          // Copy original item properties
          slNo: item.slNo,
          drawingNo: item.drawingNo,
          rev: item.rev,
          itemNo: item.itemNo,
          qty: item.qty,
          remarks: item.remarks,
          
          // Add quantity-specific properties
          originalItemNo: item.itemNo,
          sheetName: sheetName,
          sequenceInItem: seqInItem,
          totalInItem: quantity,
          globalSequence: globalSequence
        };

        expandedItems.push(expandedItem);
        globalSequence++;

        console.log(`    ✅ Sheet ${globalSequence - 1}: "${sheetName}" (${seqInItem}/${quantity})`);
      }
    });

    console.log(`📊 QUANTITY EXPANSION COMPLETE: ${items.length} items → ${expandedItems.length} sheets`);
    return expandedItems;
  }

  // HYBRID APPROACH - Supporting Methods

  private async createTemplateCopies(
    expandedItems: ExpandedSheetItem[],
    tempDir: string
  ): Promise<Array<{ filePath: string; item: ExpandedSheetItem }>> {
    console.log("📁 STEP 1: Creating file-system copies of template...");

    // Ensure temp directory exists
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    const templateCopies: Array<{ filePath: string; item: ExpandedSheetItem }> =
      [];

    for (let i = 0; i < expandedItems.length; i++) {
      const item = expandedItems[i];
      // Use sheetName for filename (safe version)
      const tempFileName = `temp_${i}_${item.sheetName.replace(
        /[^a-zA-Z0-9]/g,
        "_"
      )}.xlsx`;
      const tempFilePath = path.join(tempDir, tempFileName);

      console.log(`  📋 Copying template for ${item.sheetName} (${item.sequenceInItem}/${item.totalInItem}) → ${tempFileName}`);

      // Use file-system copy to preserve ALL content (including embedded images, drawings, etc.)
      fs.copyFileSync(this.templatePath, tempFilePath);

      templateCopies.push({ filePath: tempFilePath, item });
      console.log(`  ✅ Template copied: ${tempFileName}`);
    }

    console.log(
      `📁 STEP 1 COMPLETE: Created ${templateCopies.length} template copies`
    );
    return templateCopies;
  }

  private async populateTemplateCopy(
    filePath: string,
    formData: InspectionFormData,
    item: ExpandedSheetItem
  ): Promise<void> {
    console.log(
      `📋 STEP 2: Populating ${item.sheetName} (${item.sequenceInItem}/${item.totalInItem}) at ${path.basename(filePath)}...`
    );

    try {
      // Open the file copy with ExcelJS for data population only
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);

      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        throw new Error(`No worksheet found in ${filePath}`);
      }

      // Use existing data population logic (reuse populateSheetData method)
      await this.populateSheetData(worksheet, formData, item);

      // Save the populated file
      await workbook.xlsx.writeFile(filePath);

      console.log(`  ✅ Data populated for ${item.sheetName}`);
    } catch (error) {
      console.error(`  ❌ Error populating ${item.sheetName}:`, error);
      throw error;
    }
  }

  private async assembleMultiSheetWorkbook(
    templateCopies: Array<{ filePath: string; item: ExpandedSheetItem }>,
    outputPath: string
  ): Promise<void> {
    console.log(
      "📚 STEP 3: Assembling multi-sheet workbook with image preservation..."
    );

    try {
      // Create final output workbook
      const finalWorkbook = new ExcelJS.Workbook();

      // STEP 3A: Extract and transfer workbook-level media from template
      const imageIdMap = await this.transferWorkbookMedia(
        templateCopies[0].filePath,
        finalWorkbook
      );

      // STEP 3B: Process each populated template copy
      for (const { filePath, item } of templateCopies) {
        console.log(`  📄 Adding sheet: ${item.sheetName} (${item.sequenceInItem}/${item.totalInItem})...`);

        // Read the populated template
        const tempWorkbook = new ExcelJS.Workbook();
        await tempWorkbook.xlsx.readFile(filePath);

        const sourceSheet = tempWorkbook.worksheets[0];
        if (!sourceSheet) {
          console.error(`  ❌ No worksheet found in ${filePath}`);
          continue;
        }

        // Add worksheet to final workbook with correct name (using sheetName for quantity support)
        const targetSheet = finalWorkbook.addWorksheet(item.sheetName);

        // (Do NOT set pageSetup yet - the following copy may overwrite it)

        // Copy the entire worksheet content with enhanced image support
        await this.copyWorksheetContentWithImages(
          sourceSheet,
          targetSheet,
          tempWorkbook,
          finalWorkbook,
          imageIdMap
        );

        // Now override pageSetup and print area with STRICT A4 boundaries
        try {
          // STRICT A4: Always use A1:M56 regardless of source dimensions
          console.log(`    📄 Setting STRICT A4 page setup: A1:M56 print area`);
          
          // Force A4 page setup with correct scale and boundaries
          targetSheet.pageSetup = {
            paperSize: 9,              // A4 paper
            orientation: 'portrait',   // Portrait orientation
            fitToPage: true,          // Enable fit to page
            fitToWidth: 1,            // Fit to 1 page wide
            fitToHeight: 0,           // Unlimited height (single page)
            scale: 37,                // 37% scale from template analysis
            printArea: 'A1:M56',      // STRICT A4 print area
            margins: {
              left: 0.25,
              right: 0.25,
              top: 0.75,
              bottom: 0.75,
              header: 0.3,
              footer: 0.3
            },
            pageOrder: 'downThenOver',
            blackAndWhite: false,
            draft: false,
            cellComments: 'None',
            errors: 'displayed'
          };
        } catch (psError) {
          console.log(
            "  ⚠️ Failed to set pageSetup/printArea on target sheet:",
            psError
          );
        }

        console.log(`  ✅ Sheet added: ${item.sheetName}`);
      }

      try {
        const templateWorkbook = new ExcelJS.Workbook();
        await templateWorkbook.xlsx.readFile(templateCopies[0].filePath);
        
        const refSheet = templateWorkbook.getWorksheet('REF');
        if (refSheet) {
          const targetRefSheet = finalWorkbook.addWorksheet('REF');
          
          // Copy all cells with formulas intact
          refSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
              const targetCell = targetRefSheet.getCell(rowNumber, colNumber);
              targetCell.model = cell.model; // Preserves formulas, values, styles
            });
          });
          
          // Copy column widths
          refSheet.columns.forEach((col, idx) => {
            if (col.width) targetRefSheet.getColumn(idx + 1).width = col.width;
          });
          
          console.log('  ✅ REF sheet copied successfully');
        } else {
          console.log('  ⚠️ REF sheet not found in template');
        }
      } catch (error) {
        console.error('  ❌ Failed to copy REF sheet:', error);
      }
      console.log(
        `📚 STEP 3 PROGRESS: ${finalWorkbook.worksheets.length} sheets assembled`
      );

      // Write final workbook
      await finalWorkbook.xlsx.writeFile(outputPath);
      console.log(
        `📚 STEP 3 COMPLETE: Final workbook saved with ${finalWorkbook.worksheets.length} sheets`
      );
    } catch (error) {
      console.error("❌ Error assembling workbook:", error);
      throw error;
    }
  }

  private async copyWorksheetContent(
    sourceSheet: ExcelJS.Worksheet,
    targetSheet: ExcelJS.Worksheet
  ): Promise<void> {
    // Copy all data and formatting from source to target
    // This is similar to cloneWorksheetWithExcelJS but for already-populated sheets

    console.log("    🔍 Analyzing source sheet for content copying...");

    // Try multiple ways to get dimensions
    const dimension = (sourceSheet as any).dimension || null;
    const actualDimension = (sourceSheet as any).actualDimension || null;

    console.log(
      `    🔍 Source sheet dimension: ${
        dimension ? JSON.stringify(dimension) : "null"
      }`
    );
    console.log(
      `    🔍 Source sheet actualDimension: ${
        actualDimension ? JSON.stringify(actualDimension) : "null"
      }`
    );
    console.log(`    🔍 Source sheet rowCount: ${sourceSheet.rowCount}`);
    console.log(
      `    🔍 Source sheet actualRowCount: ${
        (sourceSheet as any).actualRowCount || "undefined"
      }`
    );
    console.log(
      `    🔍 Source sheet columnCount: ${
        (sourceSheet as any).columnCount || "undefined"
      }`
    );
    console.log(
      `    🔍 Source sheet actualColumnCount: ${
        (sourceSheet as any).actualColumnCount || "undefined"
      }`
    );

    // STRICT A4 BOUNDARIES: Force copy range to A1:M56 only (ignore template's extended styling)
    const maxRow = 56; // A4 print area row limit
    const maxCol = 13; // Column M (A4 print area column limit)

    console.log(`    📏 STRICT A4 BOUNDARIES: A1:M56 only (ignoring template extensions beyond M)`);
    console.log(`    📏 Template dimension: ${dimension ? JSON.stringify(dimension) : 'null'}`);
    console.log(`    📏 Template rowCount: ${sourceSheet.rowCount}, columnCount: ${(sourceSheet as any).columnCount || 'undefined'}`);
    console.log(`    📏 Enforced copy range: A1 to M56 (A4 print area)`);
    
    if (dimension && (dimension.right > 13 || dimension.bottom > 56)) {
      console.log(`    ⚠️  Template extends beyond A1:M56 (to row ${dimension.bottom}, col ${dimension.right}) - ignoring extra content`);
    }

    // Copy all cell content and formatting
    let copiedCells = 0;
    let totalCellsChecked = 0;

    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber++) {
      const sourceRow = sourceSheet.getRow(rowNumber);
      const targetRow = targetSheet.getRow(rowNumber);

      if (sourceRow.height) {
        targetRow.height = sourceRow.height;
      }

      for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
        totalCellsChecked++;
        const sourceCell = sourceRow.getCell(colNumber);
        const targetCell = targetRow.getCell(colNumber);

        if (sourceCell.value !== null && sourceCell.value !== undefined) {
          // Preserve formulas by using model instead of value
          if (sourceCell.formula || sourceCell.type === ExcelJS.ValueType.Formula) {
            targetCell.model = sourceCell.model;
          } else {
            targetCell.value = sourceCell.value;
          }
          copiedCells++;

          // Log first few cells for debugging
          if (copiedCells <= 5) {
            console.log(
              `    🔄 A4 copy cell [${rowNumber},${colNumber}]: "${sourceCell.value}"`
            );
          }
        }

        if (sourceCell.style && Object.keys(sourceCell.style).length > 0) {
          try {
            targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
          } catch (styleError) {
            // Continue without style if copying fails
          }
        }

        // Skip hyperlink copying as it's read-only in some versions
        // if (sourceCell.hyperlink) {
        //   try {
        //     targetCell.hyperlink = sourceCell.hyperlink;
        //   } catch (hyperlinkError) {
        //     // Continue without hyperlink if copying fails
        //   }
        // }
      }
    }
    // Fix external workbook references in formulas (remove [1] prefix)
    console.log(`    🔧 Fixing formula references...`);
    let fixedFormulas = 0;
    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber++) {
      for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
        const cell = targetSheet.getCell(rowNumber, colNumber);
        if (cell.formula && typeof cell.formula === 'string') {
          const oldFormula = cell.formula;
          const newFormula = oldFormula.replace(/\[1\]/g, '');
          if (oldFormula !== newFormula) {
            // Use cell.value with formula object to update formula
            cell.value = {
              formula: newFormula,
              result: cell.result
            };
            fixedFormulas++;
          }
        }
      }
    }
    if (fixedFormulas > 0) {
      console.log(`    ✅ Fixed ${fixedFormulas} formulas (removed [1] external references)`);
    }
    // Copy merged cells ONLY within A1:M56 boundaries
    try {
      const merges = (sourceSheet as any).model?.merges;
      if (merges && Array.isArray(merges)) {
        let copiedMerges = 0;
        let skippedMerges = 0;
        
        merges.forEach((merge: any) => {
          try {
            if (typeof merge === "string") {
              // Check if merge is within A1:M56
              const [startCell, endCell] = merge.split(':');
              const startCol = (startCell as any).match(/[A-Z]+/)[0];
              const endCol = (endCell as any).match(/[A-Z]+/)[0];
              const startRow = parseInt((startCell as any).match(/\d+/)[0]);
              const endRow = parseInt((endCell as any).match(/\d+/)[0]);
              
              const startColNum = startCol.charCodeAt(0) - 64;
              const endColNum = endCol.charCodeAt(0) - 64;
              
              // Only copy merges completely within A1:M56
              if (startColNum >= 1 && endColNum <= 13 && startRow >= 1 && endRow <= 56) {
                targetSheet.mergeCells(merge);
                copiedMerges++;
              } else {
                skippedMerges++;
              }
            } else if (merge.top && merge.left && merge.bottom && merge.right) {
              // Only copy merges completely within A1:M56
              if (merge.left >= 1 && merge.right <= 13 && merge.top >= 1 && merge.bottom <= 56) {
                targetSheet.mergeCells(
                  merge.top,
                  merge.left,
                  merge.bottom,
                  merge.right
                );
                copiedMerges++;
              } else {
                skippedMerges++;
              }
            }
          } catch (error) {
            // Continue if merge fails
            skippedMerges++;
          }
        });
        
        console.log(`    📊 A4 merges: ${copiedMerges} copied, ${skippedMerges} skipped (outside A1:M56)`);
      }
    } catch (mergeError) {
      // Continue if merge copying fails
    }

    // Copy column properties
    for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
      const sourceColumn = sourceSheet.getColumn(colNumber);
      const targetColumn = targetSheet.getColumn(colNumber);

      if (sourceColumn.width) {
        targetColumn.width = sourceColumn.width;
      }
      if (sourceColumn.hidden) {
        targetColumn.hidden = sourceColumn.hidden;
      }
    }

    // Copy worksheet-level properties and media
    try {
      // const sourceModel = (sourceSheet as any).model;
      // const targetModel = (targetSheet as any).model;

      // Copy drawings, media, images, charts, shapes
      // ["drawing", "media", "images", "charts", "shapes"].forEach((prop) => {
      //   if (sourceModel?.[prop]) {
      //     targetModel[prop] = JSON.parse(JSON.stringify(sourceModel[prop]));
      //   }
      // });

      // Copy page setup and views
      if ((sourceSheet as any).pageSetup) {
        (targetSheet as any).pageSetup = JSON.parse(
          JSON.stringify((sourceSheet as any).pageSetup)
        );
      }
      if ((sourceSheet as any).views) {
        (targetSheet as any).views = JSON.parse(
          JSON.stringify((sourceSheet as any).views)
        );
      }
    } catch (propError) {
      // Continue if property copying fails
    }

    console.log(
      `    ✅ Copy summary: ${copiedCells} cells copied out of ${totalCellsChecked} checked`
    );
  }

  private async cleanupTempFiles(tempDir: string): Promise<void> {
    console.log("🧹 STEP 4: Cleaning up temporary files...");

    try {
      if (fs.existsSync(tempDir)) {
        // Remove all files in temp directory
        const files = fs.readdirSync(tempDir);
        for (const file of files) {
          const filePath = path.join(tempDir, file);
          fs.unlinkSync(filePath);
        }

        // Remove temp directory
        fs.rmdirSync(tempDir);

        console.log(
          `🧹 STEP 4 COMPLETE: Cleaned up ${files.length} temporary files`
        );
      }
    } catch (error) {
      console.log(`🧹 STEP 4 WARNING: Failed to cleanup temp files: ${error}`);
      // Don't throw - cleanup failure shouldn't break the main process
    }
  }

  // ENHANCED IMAGE PRESERVATION METHODS

  private async transferWorkbookMedia(
    templateFilePath: string,
    targetWorkbook: ExcelJS.Workbook
  ): Promise<Map<string, string>> {
    console.log(
      "  🖼️ STEP 3A: Transferring workbook-level media and images..."
    );

    const imageIdMap = new Map<string, string>(); // Map old imageId to new imageId

    try {
      // Read template workbook to extract media
      const templateWorkbook = new ExcelJS.Workbook();
      await templateWorkbook.xlsx.readFile(templateFilePath);

      const templateModel = (templateWorkbook as any).model;

      // Transfer all images using ExcelJS addImage method
      if (templateModel?.media && Array.isArray(templateModel.media)) {
        console.log(
          `    🖼️ Found ${templateModel.media.length} media items to transfer...`
        );

        for (let i = 0; i < templateModel.media.length; i++) {
          const mediaItem = templateModel.media[i];

          try {
            // Add image to target workbook using ExcelJS addImage method
            const newImageId = targetWorkbook.addImage({
              buffer: mediaItem.buffer || mediaItem.image,
              extension: mediaItem.extension || "png",
            });

            // Map old ID to new ID for worksheet references
            imageIdMap.set(String(mediaItem.index || i), String(newImageId));

            console.log(
              `    🖼️ Transferred image ${i + 1}/${
                templateModel.media.length
              } (ID: ${mediaItem.index || i} → ${newImageId})`
            );
          } catch (imageError) {
            console.log(
              `    ❌ Failed to transfer image ${i + 1}: ${imageError}`
            );
          }
        }

        console.log(
          `    ✅ Successfully transferred ${imageIdMap.size} images`
        );
      } else {
        console.log("    ℹ️ No media items found in template");
      }

      // Transfer workbook properties and themes (important for image rendering)
      if (templateWorkbook.properties) {
        targetWorkbook.properties = templateWorkbook.properties;
        console.log("    ✅ Workbook properties transferred");
      }

      // Preserve themes (critical for proper image and formatting display)
      const targetModel = (targetWorkbook as any).model;
      if (templateModel?.themes) {
        targetModel.themes = JSON.parse(JSON.stringify(templateModel.themes));
        console.log("    ✅ Workbook themes preserved");
      }

      console.log(
        "  🖼️ STEP 3A COMPLETE: Workbook-level media transfer completed"
      );
    } catch (error) {
      console.log(`  ❌ Error transferring workbook media: ${error}`);
    }

    return imageIdMap;
  }

  private async copyWorksheetContentWithImages(
    sourceSheet: ExcelJS.Worksheet,
    targetSheet: ExcelJS.Worksheet,
    _sourceWorkbook: ExcelJS.Workbook,
    _targetWorkbook: ExcelJS.Workbook,
    imageIdMap?: Map<string, string>
  ): Promise<void> {
    console.log(
      "    🖼️ Copying worksheet content with proper ExcelJS image support..."
    );

    // First, copy all basic content using the existing method
    await this.copyWorksheetContent(sourceSheet, targetSheet);

    // Then, copy images using ExcelJS getImages() and addImage() methods
    try {
      // Get all images from source worksheet
      const images = sourceSheet.getImages();

      if (images.length > 0) {
        console.log(`    🖼️ Found ${images.length} images in source worksheet`);

        for (let i = 0; i < images.length; i++) {
          const image = images[i];

          try {
            let newImageId = image.imageId;

            // If we have an image ID mapping, use the new ID
            if (imageIdMap && imageIdMap.has(image.imageId)) {
              newImageId = imageIdMap.get(image.imageId)!;
              console.log(
                `    🖼️ Remapping image ID ${image.imageId} → ${newImageId}`
              );
            }

            // Add image to target worksheet using ExcelJS addImage method
            // Convert imageId to number for addImage method
            const imageIdNumber = parseInt(newImageId);

            if (image.range && typeof image.range === 'object') {
              // If image has a range, use it
              targetSheet.addImage(imageIdNumber, image.range);
              console.log(
                `    🖼️ Added image ${i + 1} with range: ${JSON.stringify(
                  image.range
                )}`
              );
            } else {
              // Use a default range if no specific positioning is available
              console.log(
                `    ⚠️ Image ${
                  i + 1
                } has no range info, using default placement`
              );
              targetSheet.addImage(imageIdNumber, "A1:C3");
            }
          } catch (imageError) {
            console.log(`    ❌ Failed to copy image ${i + 1}: ${imageError}`);
          }
        }

        console.log(`    ✅ Successfully processed ${images.length} images`);
      } else {
        console.log("    ℹ️ No images found in source worksheet");
      }
    } catch (error) {
      console.log(`    ❌ Error in ExcelJS image support: ${error}`);

      // Fallback: Try the model-level approach as backup
      // try {
      //   console.log("    🔄 Attempting fallback model-level image transfer...");
      //   const sourceModel = (sourceSheet as any).model;
      //   const targetModel = (targetSheet as any).model;

      //   if (sourceModel?.drawing) {
      //     targetModel.drawing = JSON.parse(JSON.stringify(sourceModel.drawing));
      //     console.log("    ✅ Fallback: Drawing model transferred");
      //   }
      // } catch (fallbackError) {
      //   console.log(`    ❌ Fallback also failed: ${fallbackError}`);
      // }
    }
  }

  // Legacy cloning method - now replaced by hybrid approach
  /* private async cloneWorksheetWithExcelJS(sourceWorksheet: ExcelJS.Worksheet, targetWorksheet: ExcelJS.Worksheet): Promise<void> {
    console.log('🔄 Performing ExcelJS worksheet cloning with full formatting preservation...');
    
    // Get the actual used range from source
    const dimension = (sourceWorksheet as any).dimension || null;
    
    console.log(`🔍 Source worksheet dimension: ${dimension ? `${dimension.tl}:${dimension.br}` : 'No dimension'}`);
    console.log(`🔍 Source worksheet rowCount: ${sourceWorksheet.rowCount}`);
    console.log(`🔍 Source worksheet actualRowCount: ${sourceWorksheet.actualRowCount}`);
    
    // Check for drawings and images
    const sourceModel = (sourceWorksheet as any).model;
    if (sourceModel?.drawing || sourceModel?.media || sourceModel?.images) {
      console.log('🔍 Found drawings/images in source worksheet');
      if (sourceModel.drawing) console.log(`  - Drawing elements: ${Object.keys(sourceModel.drawing).length}`);
      if (sourceModel.media) console.log(`  - Media elements: ${sourceModel.media.length || 0}`);
      if (sourceModel.images) console.log(`  - Image elements: ${sourceModel.images.length || 0}`);
    }
    
    if (!dimension && sourceWorksheet.rowCount === 0) {
      console.log('⚠️ Source worksheet appears to be empty - no content to clone');
      return;
    }
    
    // Use a more aggressive approach - copy a larger range to ensure we get everything including drawings
    const maxRow = Math.max(dimension?.bottom || 100, sourceWorksheet.actualRowCount, 100); // Increased range
    const maxCol = Math.max(dimension?.right || 50, sourceWorksheet.actualColumnCount, 50); // Increased range
    
    console.log(`  - Cloning expanded range: A1:${this.getColumnLetter(maxCol)}${maxRow}`);
    
    let copiedCells = 0;
    
    // Copy all rows and cells with full formatting
    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber++) {
      const sourceRow = sourceWorksheet.getRow(rowNumber);
      const targetRow = targetWorksheet.getRow(rowNumber);
      
      // Copy row height
      if (sourceRow.height) {
        targetRow.height = sourceRow.height;
      }
      
      // Copy each cell in the row
      for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
        const sourceCell = sourceRow.getCell(colNumber);
        const targetCell = targetRow.getCell(colNumber);
        
        // Check if source cell has any content
        if (sourceCell.value !== null && sourceCell.value !== undefined) {
          // Copy cell value (formulas, text, numbers, etc.)
          targetCell.value = sourceCell.value;
          copiedCells++;
          
          if (copiedCells <= 10) { // Log first 10 cells for debugging
            console.log(`🔄 Copied cell ${this.getColumnLetter(colNumber)}${rowNumber}: "${sourceCell.value}"`);
          }
        }
        
        // Copy cell style completely - this is critical for formatting
        try {
          // Use deep copy to preserve all style properties
          if (sourceCell.style && Object.keys(sourceCell.style).length > 0) {
            targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
          }
        } catch (styleError) {
          // If full style copy fails, try copying individual properties
          try {
            if (sourceCell.style) {
              const newStyle: any = {};
              if (sourceCell.style.font) newStyle.font = sourceCell.style.font;
              if (sourceCell.style.fill) newStyle.fill = sourceCell.style.fill;
              if (sourceCell.style.border) newStyle.border = sourceCell.style.border;
              if (sourceCell.style.alignment) newStyle.alignment = sourceCell.style.alignment;
              if (sourceCell.style.numFmt) newStyle.numFmt = sourceCell.style.numFmt;
              if (sourceCell.style.protection) newStyle.protection = sourceCell.style.protection;
              targetCell.style = newStyle;
            }
          } catch (fallbackError) {
            // Style copying completely failed, continue without style
          }
        }
        
        // Skip hyperlink copying as it's read-only in some versions
        // if (sourceCell.hyperlink) {
        //   try {
        //     targetCell.hyperlink = sourceCell.hyperlink;
        //   } catch (hyperlinkError) {
        //     // Hyperlink copying failed, continue
        //   }
        // }
      }
    }
    
    // Copy merged cells FIRST - this is critical for proper formatting
    console.log('🔄 Copying merged cells...');
    try {
      const merges = (sourceWorksheet as any).model?.merges;
      if (merges && Array.isArray(merges)) {
        console.log(`  - Found ${merges.length} merged cell ranges`);
        merges.forEach((merge: any, index: number) => {
          try {
            if (typeof merge === 'string') {
              targetWorksheet.mergeCells(merge);
              console.log(`  ✅ Merged cells: ${merge}`);
            } else if (merge.top && merge.left && merge.bottom && merge.right) {
              targetWorksheet.mergeCells(merge.top, merge.left, merge.bottom, merge.right);
              console.log(`  ✅ Merged range: ${merge.top},${merge.left} to ${merge.bottom},${merge.right}`);
            }
          } catch (error) {
            console.log(`  ❌ Could not copy merge ${index}: ${JSON.stringify(merge)} - ${error}`);
          }
        });
      } else {
        console.log('  - No merged cells found in source');
      }
    } catch (mergeError) {
      console.log(`  ❌ Failed to access merged cells: ${mergeError}`);
    }
    
    // Copy column widths and properties
    console.log('🔄 Copying column properties...');
    let columnsCopied = 0;
    for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
      const sourceColumn = sourceWorksheet.getColumn(colNumber);
      const targetColumn = targetWorksheet.getColumn(colNumber);
      
      if (sourceColumn.width) {
        targetColumn.width = sourceColumn.width;
        columnsCopied++;
      }
      
      if (sourceColumn.hidden) {
        targetColumn.hidden = sourceColumn.hidden;
      }
    }
    console.log(`  - Copied properties for ${columnsCopied} columns`);
    
    // Copy drawings, images, and charts
    console.log('🔄 Copying drawings and media...');
    try {
      const sourceModel = (sourceWorksheet as any).model;
      const targetModel = (targetWorksheet as any).model;
      
      // Copy drawing elements if they exist
      if (sourceModel?.drawing) {
        console.log(`  - Copying drawing elements...`);
        targetModel.drawing = JSON.parse(JSON.stringify(sourceModel.drawing));
        console.log('  ✅ Drawing elements copied');
      }
      
      // Copy media elements if they exist
      if (sourceModel?.media) {
        console.log(`  - Copying media elements...`);
        targetModel.media = JSON.parse(JSON.stringify(sourceModel.media));
        console.log('  ✅ Media elements copied');
      }
      
      // Copy images if they exist
      if (sourceModel?.images) {
        console.log(`  - Copying image elements...`);
        targetModel.images = JSON.parse(JSON.stringify(sourceModel.images));
        console.log('  ✅ Image elements copied');
      }
      
      // Copy charts if they exist
      if (sourceModel?.charts) {
        console.log(`  - Copying chart elements...`);
        targetModel.charts = JSON.parse(JSON.stringify(sourceModel.charts));
        console.log('  ✅ Chart elements copied');
      }
      
      // Copy shapes if they exist
      if (sourceModel?.shapes) {
        console.log(`  - Copying shape elements...`);
        targetModel.shapes = JSON.parse(JSON.stringify(sourceModel.shapes));
        console.log('  ✅ Shape elements copied');
      }
      
      if (!sourceModel?.drawing && !sourceModel?.media && !sourceModel?.images && !sourceModel?.charts && !sourceModel?.shapes) {
        console.log('  - No drawing/media elements found in source');
      }
    } catch (drawingError) {
      console.log(`  ❌ Failed to copy drawings/media: ${drawingError}`);
    }
    
    // Copy worksheet properties for better formatting
    console.log('🔄 Copying worksheet properties...');
    try {
      // Copy page setup if it exists
      if ((sourceWorksheet as any).pageSetup) {
        (targetWorksheet as any).pageSetup = JSON.parse(JSON.stringify((sourceWorksheet as any).pageSetup));
        console.log('  ✅ Page setup copied');
      }
      
      // Copy views if they exist
      if ((sourceWorksheet as any).views) {
        (targetWorksheet as any).views = JSON.parse(JSON.stringify((sourceWorksheet as any).views));
        console.log('  ✅ Worksheet views copied');
      }
    } catch (propError) {
      console.log(`  ❌ Failed to copy worksheet properties: ${propError}`);
    }
    
    console.log(`✅ ExcelJS cloning completed - ${copiedCells} cells copied from ${maxRow} rows, ${maxCol} columns`);
  } */
  
    private async generateBatchfile(
      formData: InspectionFormData,
      outputPath: string
    ): Promise<void> {
      console.log('📋 Generating batchfile from template...');
    
      
      
      // Load batchfile template
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(this.batchfileTemplatePath);
      
      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        throw new Error('No worksheet found in batchfile template');
      }
    
      // Populate header information
      const globalMappings = [
        { cell: "D4", value: formData.project, field: "Project" },
        { cell: "D5", value: formData.company, field: "Company" },
        { cell: "D6", value: formData.contractor, field: "Contractor" },
        { cell: "D7", value: formData.subContractor, field: "Sub-Contractor" },
        { cell: "D8", value: formData.itpNo, field: "ITP No" },
        { cell: "L5", value: formData.date, field: "Date" },
        { cell: "L7", value: formData.rfiNo, field: "Work Order No" },
        { cell: "L8", value: formData.itpReferenceClause, field: "Acceptance Standard" },
        { cell: "L9", value: "Final", field: "Stage" },
      ];
    
      globalMappings.forEach(({ cell, value, field }) => {
        try {
          worksheet.getCell(cell).value = value;
          console.log(`  ✅ ${field}: "${value}" → ${cell}`);
        } catch (error) {
          console.log(`  ❌ Error setting ${field} in ${cell}: ${error}`);
        }
      });
    
      // Batchfile-specific fields
      worksheet.getCell("L4").value = `${formData.certNo}\nBATCHFILE`;
      worksheet.getCell("D9").value = "As Mentioned Below";
      worksheet.getCell("L6").value = "As Mentioned Below";
    
      // Add signature dates
      worksheet.getCell("D56").value = formData.date;
      worksheet.getCell("H56").value = formData.date;
    
      // Expand items by quantity (to match the individual sheets)
      const expandedItems = this.expandItemsByQuantity(formData.items);
    
      // Populate the table - starts at row 16
      const tableStartRow = 12;
    
      console.log(`  📋 Populating ${expandedItems.length} expanded items in table...`);
    
      expandedItems.forEach((item, index) => {
        const rowNumber = tableStartRow + index;
        
        // Based on the template structure:
        // Column A (1): Sl.No
        // Column B (2): Drawing (column 1)
        // Column C (3): Drawing (column 2)
        // Column D (4): Drawing (column 3) 
        // Column E (5): Description (column 1)
        // Column F (6): Description (column 2)
        // Column G (7): As Per Drawing (mm) - LEAVE BLANK
        // Column H (8): Actual Observation(mm) - LEAVE BLANK
        
        const drawingWithRev = `${item.drawingNo} Rev.${item.rev}`;
        
        // Remove "324-1-" prefix and create description like "B-129  1/1"
        const cleanItemNo = item.originalItemNo.replace(/^324-1-/, '');
        const description = `${cleanItemNo}  ${item.sequenceInItem}/${item.totalInItem}`;
        
        worksheet.getCell(rowNumber, 1).value = index + 1; // Sl.No
        worksheet.getCell(rowNumber, 2).value = drawingWithRev; // Drawing col 1
        worksheet.getCell(rowNumber, 3).value = drawingWithRev; // Drawing col 2
        worksheet.getCell(rowNumber, 4).value = drawingWithRev; // Drawing col 3
        worksheet.getCell(rowNumber, 5).value = description; // Description col 1
        worksheet.getCell(rowNumber, 6).value = description; // Description col 2
        // Columns G and H left blank as per customer request
        
        console.log(`  ✅ Row ${rowNumber}: ${description} (${drawingWithRev})`);
      });
    
      // Save the populated batchfile
      await workbook.xlsx.writeFile(outputPath);
      console.log(`✅ Batchfile saved with ${expandedItems.length} items: ${outputPath}`);
    }
  private async populateSheetData(
    worksheet: ExcelJS.Worksheet,
    formData: InspectionFormData,
    item: ExpandedSheetItem
  ): Promise<void> {
    console.log(
      `📋 Populating data for sheet: ${item.sheetName} (${item.sequenceInItem}/${item.totalInItem}, global: ${item.globalSequence})`
    );

    try {
      // 🟢 GLOBAL FIELDS (same in all sheets) - based on your green annotations
      const globalMappings = [
        { cell: "D4", value: formData.project, field: "Project" },
        { cell: "D5", value: formData.company, field: "Company" },
        { cell: "D6", value: formData.contractor, field: "Contractor" },
        { cell: "D7", value: formData.subContractor, field: "Sub-Contractor" },
        { cell: "D8", value: formData.itpNo, field: "ITP No" },
        { cell: "L5", value: formData.date, field: "Date" },
        { cell: "L7", value: formData.rfiNo, field: "Work Order No" },
        {
          cell: "L8",
          value: formData.itpReferenceClause,
          field: "Acceptance Standard",
        },
        { cell: "L9", value: "Final", field: "Stage" },
      ];

      // Apply global mappings
      globalMappings.forEach(({ cell, value, field }) => {
        try {
          const cellObj = worksheet.getCell(cell);
          cellObj.value = value;
          console.log(`  ✅ ${field}: "${value}" → ${cell}`);
        } catch (error) {
          console.log(`  ❌ Error setting ${field} in ${cell}: ${error}`);
        }
      });

      // 🟣 SERIAL/INDEX FIELDS (quantity-aware incrementing) - based on your purple annotations
      const annexSuffix = String(item.globalSequence).padStart(3, "0");
      const reportNo = `${formData.certNo}\nANNEX -${annexSuffix}`;

      try {
        const reportCell = worksheet.getCell("L4");
        reportCell.value = reportNo;
        console.log(`  ✅ Report No: "${reportNo}" → L4 (global sequence: ${item.globalSequence})`);
      } catch (error) {
        console.log(`  ❌ Error setting Report No: ${error}`);
      }

      // 🟡 COMPOSITE FIELDS (quantity-aware combinations) - based on your yellow annotations
      const itemDescription = formData.itemDescription; // Global item description for all sheets
      const refDrawing = `${item.drawingNo} Rev-${item.rev || "03"}`;
      const sketchText = `SKETCH : ${item.originalItemNo}`; // Use original item number

      const compositeMappings = [
        { cell: "L6", value: itemDescription, field: "Item Description" },
        { cell: "D9", value: refDrawing, field: "Ref.Drawing" },
        { cell: "G11", value: sketchText, field: "Sketch" },
      ];

      compositeMappings.forEach(({ cell, value, field }) => {
        try {
          const cellObj = worksheet.getCell(cell);
          cellObj.value = value;
          console.log(`  ✅ ${field}: "${value}" → ${cell}`);
        } catch (error) {
          console.log(`  ❌ Error setting ${field} in ${cell}: ${error}`);
        }
      });

      // Update dimensional details section with quantity sequence
      const dimensionalDetailsText = `DIMENSIONAL DETAILS : ${item.originalItemNo}  ${item.sequenceInItem}/${item.totalInItem}`;
      try {
        const detailsCell = worksheet.getCell("G26");
        detailsCell.value = dimensionalDetailsText;
        console.log(
          `  ✅ Dimensional Details: "${dimensionalDetailsText}" → G26`
        );
      } catch (error) {
        console.log(`  ❌ Error setting Dimensional Details: ${error}`);
      }

      // Add signature dates (same format as main date)
      const signatureDates = [
        { cell: "D56", field: "Signature Date 1 (RAJAN PALAIYAN)" },
        { cell: "H56", field: "Signature Date 2 (SAAD GHOBRIAL)" },
      ];

      signatureDates.forEach(({ cell, field }) => {
        try {
          const cellObj = worksheet.getCell(cell);
          cellObj.value = formData.date; // Same date format as main date
          console.log(`  ✅ ${field}: "${formData.date}" → ${cell}`);
        } catch (error) {
          console.log(`  ❌ Error setting ${field} in ${cell}: ${error}`);
        }
      });

      console.log(`🎉 Data population completed for ${item.sheetName}`);
    } catch (error) {
      console.error(
        `❌ Error during data population for ${item.sheetName}:`,
        error
      );
    }
  }

  getBatchInfo(batchId: string): BatchInfo | undefined {
    return this.batches.get(batchId);
  }

  getFilePath(filename: string): string {
    return path.join(this.outputDir, filename);
  }

  fileExists(filename: string): boolean {
    return fs.existsSync(this.getFilePath(filename));
  }

  // Cleanup old files (older than 24 hours)
  cleanupOldFiles(): void {
    const cutoffTime = Date.now() - 24 * 60 * 60 * 1000; // 24 hours ago

    try {
      const files = fs.readdirSync(this.outputDir);

      files.forEach((file) => {
        const filepath = path.join(this.outputDir, file);
        const stats = fs.statSync(filepath);

        if (stats.mtime.getTime() < cutoffTime) {
          fs.unlinkSync(filepath);
          console.log(`Cleaned up old file: ${file}`);
        }
      });

      // Also cleanup old batch info
      this.batches.forEach((batch, batchId) => {
        if (batch.createdAt.getTime() < cutoffTime) {
          this.batches.delete(batchId);
        }
      });
    } catch (error) {
      console.error("Error during cleanup:", error);
    }
  }
}

// Export singleton instance
export const excelService = new ExcelGeneratorService();

// Setup periodic cleanup (run every hour)
setInterval(() => {
  excelService.cleanupOldFiles();
}, 60 * 60 * 1000);
