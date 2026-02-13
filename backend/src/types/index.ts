export interface InspectionItem {
  slNo: number;
  drawingNo: string;
  rev: string;
  itemNo: string;
  qty: number;
  remarks?: string;
}

export interface ExpandedSheetItem {
  slNo: number;
  drawingNo: string;
  rev: string;
  itemNo: string;
  qty: number;
  remarks?: string;
  // Quantity-specific fields
  originalItemNo: string;  // Original item number (e.g., "B-129")
  sheetName: string;       // Sheet name with quantity suffix (e.g., "B-129 (1)")
  sequenceInItem: number;  // 1, 2, 3 for this specific item
  totalInItem: number;     // Total quantity for this item (e.g., 3)
  globalSequence: number;  // Global sequence across all sheets
}

export interface InspectionFormData {
  // Header Information from Inspection Release Note
  project: string;
  contractNo: string;
  company: string;
  date: string;
  contractor: string;
  subContractor: string;
  itpNo: string;
  itpRevNo: string;
  itpReferenceClause: string;
  certNo: string;
  rfiNo: string;
  itemDescription: string;
  
  // Dynamic Table Items
  items: InspectionItem[];
}

export interface GenerationResponse {
  success: boolean;
  files: string[];
  batchId: string;
  message?: string;
}

export interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
}

export interface FileInfo {
  filename: string;
  originalName: string;
  size: number;
  path: string;
  createdAt: Date;
}

export interface BatchInfo {
  batchId: string;
  files: FileInfo[];
  createdAt: Date;
  formData: InspectionFormData;
}