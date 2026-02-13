import React from 'react';
import { GenerationResponse } from '../types';
import { excelApi } from '../services/api';

interface DownloadResultsProps {
  results: GenerationResponse | null;
  isLoading: boolean;
}

const DownloadResults: React.FC<DownloadResultsProps> = ({ results, isLoading }) => {
  if (isLoading) {
    return (
      <div className="results-container">
        <div className="loading">
          Generating Excel files, please wait...
        </div>
      </div>
    );
  }

  if (!results) {
    return null;
  }

  if (!results.success) {
    return (
      <div className="results-container">
        <h2 style={{ color: '#e74c3c', marginBottom: '1rem' }}>❌ Generation Failed</h2>
        <p style={{ color: '#666', marginBottom: '1rem' }}>{results.message}</p>
        <div style={{ 
          backgroundColor: '#fdf2f2', 
          border: '1px solid #fecaca', 
          padding: '1rem', 
          borderRadius: '4px',
          color: '#b91c1c'
        }}>
          <strong>Error:</strong> {results.message}
        </div>
      </div>
    );
  }

  const handleDownloadFile = (filename: string) => {
    const downloadUrl = excelApi.downloadFile(filename);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleDownloadBatchfile = (mainFilename: string) => {
    // Convert "Inspection-Report-PROJECT-2026-02-12.xlsx" 
    // to "Inspection-Report-PROJECT-2026-02-12-BATCHFILE.xlsx"
    const batchfileName = mainFilename.replace('.xlsx', '-BATCHFILE.xlsx');
    
    const downloadUrl = excelApi.downloadFile(batchfileName);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = batchfileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="results-container">
      <h2 style={{ color: '#27ae60', marginBottom: '1rem' }}>✅ Files Generated Successfully!</h2>
      
      <div style={{ 
        backgroundColor: '#f0fdf4', 
        border: '1px solid #bbf7d0', 
        padding: '1rem', 
        borderRadius: '4px',
        marginBottom: '2rem',
        color: '#166534'
      }}>
        <strong>Success:</strong> {results.message} and 1 batch file
        <br />
        <strong>Generated:</strong> Multi-sheet Excel file with {results.files.length > 0 ? 'multiple inspection sheets' : 'inspection data'}
        <br />
        <strong>Batch ID:</strong> {results.batchId}
      </div>

      {/* Single File Download */}
      {/* <div style={{ marginBottom: '2rem', textAlign: 'center' }}>
        <h3 style={{ marginBottom: '1rem' }}>📊 Download Inspection Report</h3>
        {results.files.length > 0 && (
          <div>
            <button
              onClick={() => handleDownloadFile(results.files[0])}
              className="btn btn-success"
              style={{ fontSize: '1.2rem', padding: '1rem 2rem', marginBottom: '1rem' }}
            >
              📥 Download Excel Report
            </button>
            <div style={{ marginTop: '1rem' }}>
              <strong>File:</strong> <span style={{ color: '#666', fontSize: '0.9rem' }}>{results.files[0]}</span>
            </div>
          </div>
        )}
      </div> */}
      <div style={{ marginBottom: '2rem', textAlign: 'center' }}>
        <h3 style={{ marginBottom: '1rem' }}>📊 Download Reports</h3>
        {results.files.length > 0 && (
          <div>
            {/* Main Inspection Report */}
            <button
              onClick={() => {
                console.log(results)
                handleDownloadFile(results.files[0])
              }}
              className="btn btn-success"
              style={{ fontSize: '1.2rem', padding: '1rem 2rem', marginBottom: '1rem', marginRight: '1rem' }}
            >
              📥 Download Inspection Report
            </button>
            
            {/* Batchfile Button */}
            <button
              onClick={() => handleDownloadBatchfile(results.files[0])}
              className="btn btn-primary"
              style={{ fontSize: '1.2rem', padding: '1rem 2rem', marginBottom: '1rem' }}
            >
              📋 Download Batchfile
            </button>
            
            <div style={{ marginTop: '1rem' }}>
              <strong>Main File:</strong> <span style={{ color: '#666', fontSize: '0.9rem' }}>{results.files[0]}</span>
              <br />
              <strong>Batchfile:</strong> <span style={{ color: '#666', fontSize: '0.9rem' }}>{results.files[0].replace('.xlsx', '-BATCHFILE.xlsx')}</span>
            </div>
          </div>
        )}
      </div>

      {/* Instructions */}
      <div style={{ 
        marginTop: '2rem', 
        padding: '1rem', 
        backgroundColor: '#f8f9fa', 
        borderRadius: '4px',
        border: '1px solid #e9ecef'
      }}>
        <h4 style={{ marginBottom: '0.5rem', color: '#495057' }}>📋 What's Next?</h4>
        <ul style={{ paddingLeft: '1.5rem', color: '#666' }}>
          <li>Download the Excel file using the button above</li>
          <li>The file contains multiple sheets - one for each inspection item</li>
          <li>Each sheet is named with the item number (B-128, B-129, etc.)</li>
          <li>All sheets contain the complete dimensional inspection report data</li>
          <li>Files are automatically named with project name and date</li>
          <li>Files will be automatically cleaned up after 24 hours</li>
        </ul>
      </div>
    </div>
  );
};

export default DownloadResults;