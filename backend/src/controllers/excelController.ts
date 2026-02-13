import { Request, Response } from 'express';
import archiver from 'archiver';
import fs from 'fs';
import { excelService } from '../services/excelService';
import { InspectionFormData, GenerationResponse } from '../types';

export const generateExcelFiles = async (
  req: Request<{}, GenerationResponse, InspectionFormData>,
  res: Response<GenerationResponse>
): Promise<void> => {
  try {
    const formData = req.body;
    
    console.log('Generating Excel files for:', {
      project: formData.project,
      itemCount: formData.items.length
    });

    const result = await excelService.generateInspectionReports(formData);
    
    res.json({
      success: true,
      files: result.files,
      batchId: result.batchId,
      message: `Generated ${result.files.length} Excel files successfully`
    });
  } catch (error: any) {
    console.error('Error generating Excel files:', error);
    
    res.status(500).json({
      success: false,
      files: [],
      batchId: '',
      message: `Failed to generate Excel files: ${error.message}`
    });
  }
};

export const downloadFile = (
  req: Request<{ filename: string }>,
  res: Response
): void => {
  try {
    const { filename } = req.params;
    
    if (!filename || filename.includes('..') || filename.includes('/')) {
      res.status(400).json({
        success: false,
        error: 'Invalid filename'
      });
      return;
    }

    if (!excelService.fileExists(filename)) {
      res.status(404).json({
        success: false,
        error: 'File not found'
      });
      return;
    }

    const filepath = excelService.getFilePath(filename);
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    
    const fileStream = fs.createReadStream(filepath);
    fileStream.pipe(res);
    
    fileStream.on('error', (error) => {
      console.error('Error streaming file:', error);
      if (!res.headersSent) {
        res.status(500).json({
          success: false,
          error: 'Error downloading file'
        });
      }
    });
  } catch (error: any) {
    console.error('Error in downloadFile:', error);
    
    if (!res.headersSent) {
      res.status(500).json({
        success: false,
        error: 'Internal server error'
      });
    }
  }
};

export const downloadBatch = (
  req: Request<{ batchId: string }>,
  res: Response
): void => {
  try {
    const { batchId } = req.params;
    
    if (!batchId) {
      res.status(400).json({
        success: false,
        error: 'Batch ID is required'
      });
      return;
    }

    const batchInfo = excelService.getBatchInfo(batchId);
    
    if (!batchInfo) {
      res.status(404).json({
        success: false,
        error: 'Batch not found'
      });
      return;
    }

    const zipFilename = `inspection-reports-${batchId}.zip`;
    
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${zipFilename}"`);

    const archive = archiver('zip', {
      zlib: { level: 9 } // Maximum compression
    });

    archive.on('error', (error) => {
      console.error('Archive error:', error);
      if (!res.headersSent) {
        res.status(500).json({
          success: false,
          error: 'Error creating archive'
        });
      }
    });

    archive.pipe(res);

    // Add each file to the archive
    batchInfo.files.forEach((fileInfo) => {
      if (fs.existsSync(fileInfo.path)) {
        archive.file(fileInfo.path, { name: fileInfo.filename });
      }
    });

    archive.finalize();
  } catch (error: any) {
    console.error('Error in downloadBatch:', error);
    
    if (!res.headersSent) {
      res.status(500).json({
        success: false,
        error: 'Internal server error'
      });
    }
  }
};