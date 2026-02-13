import { z } from 'zod';

export const InspectionItemSchema = z.object({
  slNo: z.number().positive('Serial number must be positive'),
  drawingNo: z.string().min(1, 'Drawing number is required'),
  rev: z.string().min(1, 'Revision is required'),
  itemNo: z.string().min(1, 'Item number is required'),
  qty: z.number().positive('Quantity must be positive'),
  remarks: z.string().optional()
});

export const InspectionFormSchema = z.object({
  // Header Information
  project: z.string().min(1, 'Project name is required'),
  contractNo: z.string().min(1, 'Contract number is required'),
  company: z.string().min(1, 'Company name is required'),
  date: z.string().min(1, 'Date is required'),
  contractor: z.string().min(1, 'Contractor is required'),
  subContractor: z.string().min(1, 'Sub-contractor is required'),
  itpNo: z.string().min(1, 'ITP number is required'),
  itpRevNo: z.string().min(1, 'ITP revision number is required'),
  itpReferenceClause: z.string().min(1, 'ITP reference clause is required'),
  certNo: z.string().min(1, 'Certificate number is required'),
  rfiNo: z.string().min(1, 'RFI number is required'),
  itemDescription: z.string().min(1, 'Item description is required'),
  
  // Items array with at least one item
  items: z.array(InspectionItemSchema).min(1, 'At least one inspection item is required')
});

export type InspectionFormInput = z.infer<typeof InspectionFormSchema>;
export type InspectionItemInput = z.infer<typeof InspectionItemSchema>;