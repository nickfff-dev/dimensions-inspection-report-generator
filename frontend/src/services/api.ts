import axios from 'axios';
import { InspectionFormData, GenerationResponse } from '../types';

const API_BASE_URL = '/api';

const api = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    'Content-Type': 'application/json',
  },
});

export const excelApi = {
  generateFiles: async (formData: InspectionFormData): Promise<GenerationResponse> => {
    try {
      const response = await api.post<GenerationResponse>('/generate', formData);
      return response.data;
    } catch (error: any) {
      console.error('Error generating files:', error);
      throw new Error(
        error.response?.data?.message || 'Failed to generate Excel files'
      );
    }
  },

  downloadFile: (filename: string): string => {
    return `${API_BASE_URL}/download/${filename}`;
  },

  downloadBatch: (batchId: string): string => {
    return `${API_BASE_URL}/download/batch/${batchId}`;
  },
};

export default api;