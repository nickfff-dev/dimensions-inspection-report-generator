export interface InspectionItem {
  slNo: number;
  drawingNo: string;
  rev: string;
  itemNo: string;
  qty: number;
  remarks?: string;
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