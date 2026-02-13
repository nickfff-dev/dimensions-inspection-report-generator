import express, { Router } from 'express';
import { generateExcelFiles, downloadFile, downloadBatch } from '../controllers/excelController';
import { validateInspectionForm } from '../middleware/validation';

const router: Router = express.Router();

// Generate Excel files from form data
router.post('/generate', validateInspectionForm, generateExcelFiles);

// Download individual file
router.get('/download/:filename', downloadFile);

// Download batch as ZIP
router.get('/download/batch/:batchId', downloadBatch);

export default router;