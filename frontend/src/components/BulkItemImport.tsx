import React, { useState, useRef } from 'react';
import { useCSVReader } from 'react-papaparse';
import { InspectionItem } from '../types';

interface BulkItemImportProps {
  onItemsImported: (items: InspectionItem[], mode: 'replace' | 'add') => void;
  existingItemsCount: number;
}

interface ImportError {
  row: number;
  field: string;
  message: string;
}

interface ImportPreview {
  items: InspectionItem[];
  errors: ImportError[];
  isValid: boolean;
}

const BulkItemImport: React.FC<BulkItemImportProps> = ({ onItemsImported, existingItemsCount }) => {
  const [importPreview, setImportPreview] = useState<ImportPreview | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  const [importMode, setImportMode] = useState<'replace' | 'add'>('add');
  const [pasteData, setPasteData] = useState('');
  const [showPasteArea, setShowPasteArea] = useState(false);
  
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const { CSVReader } = useCSVReader();

  // Validate and process imported items
  const validateAndProcessItems = (data: any[]): ImportPreview => {
    const items: InspectionItem[] = [];
    const errors: ImportError[] = [];
    let isValid = true;

    data.forEach((row, index) => {
      // Skip empty rows
      if (!row || Object.values(row).every(val => !val || val.toString().trim() === '')) {
        return;
      }

      const item: Partial<InspectionItem> = {
        slNo: 0, // Will be set when adding to form
        drawingNo: '',
        rev: '',
        itemNo: '',
        qty: 1,
        remarks: ''
      };

      // Map CSV fields to item properties
      const fieldMappings = [
        { csv: ['Drawing No', 'drawingNo', 'drawing_no', 'drawing'], prop: 'drawingNo' },
        { csv: ['Rev', 'rev', 'revision', 'Rev No'], prop: 'rev' },
        { csv: ['Item No', 'itemNo', 'item_no', 'item'], prop: 'itemNo' },
        { csv: ['Qty', 'qty', 'quantity', 'Qty (Nos.)', 'Qty (Nos.)'], prop: 'qty' },
        { csv: ['Remarks', 'remarks', 'notes', 'comment'], prop: 'remarks' }
      ];

      fieldMappings.forEach(({ csv, prop }) => {
        const csvField = csv.find(field => row && row[field] !== undefined);
        if (csvField && row && row[csvField]) {
          if (prop === 'qty') {
            const qtyValue = parseInt(row[csvField]);
            if (isNaN(qtyValue) || qtyValue <= 0) {
              errors.push({
                row: index + 1,
                field: prop,
                message: `Quantity must be a positive number (got: "${row[csvField]}")`
              });
              isValid = false;
            } else {
              (item as any)[prop] = qtyValue;
            }
          } else {
            (item as any)[prop] = row[csvField].toString().trim();
          }
        }
      });

      // Validate required fields
      const requiredFields = ['drawingNo', 'rev', 'itemNo'] as const;
      requiredFields.forEach(field => {
        const fieldValue = item[field];
        if (!fieldValue || fieldValue.toString().trim() === '') {
          errors.push({
            row: index + 1,
            field,
            message: `${field.charAt(0).toUpperCase() + field.slice(1).replace(/([A-Z])/g, ' $1')} is required`
          });
          isValid = false;
        }
      });

      // If no major errors, add the item
      if (item.drawingNo && item.rev && item.itemNo) {
        items.push(item as InspectionItem);
      }
    });

    return { items, errors, isValid: isValid && items.length > 0 };
  };

  // Handle CSV file upload
  const handleCSVUpload = (results: any) => {
    try {
      const preview = validateAndProcessItems(results.data);
      setImportPreview(preview);
      setShowPreview(true);
    } catch (error) {
      console.error('Error processing CSV:', error);
      alert('Error processing CSV file. Please check the format and try again.');
    }
  };

  // Handle paste from clipboard (Excel format)
  const handlePaste = () => {
    if (!pasteData.trim()) {
      alert('Please paste some data first.');
      return;
    }

    try {
      // Parse tab-delimited data (Excel format)
      const lines = pasteData.trim().split('\n');
      const headers = lines[0].split('\t');
      
      const data = lines.slice(1).map(line => {
        const values = line.split('\t');
        const row: any = {};
        headers.forEach((header, index) => {
          row[header.trim()] = values[index]?.trim() || '';
        });
        return row;
      });

      const preview = validateAndProcessItems(data);
      setImportPreview(preview);
      setShowPreview(true);
      setShowPasteArea(false);
      setPasteData('');
    } catch (error) {
      console.error('Error processing paste data:', error);
      alert('Error processing pasted data. Please check the format and try again.');
    }
  };

  // Handle import confirmation
  const handleConfirmImport = () => {
    if (importPreview && importPreview.isValid) {
      // Assign serial numbers based on import mode
      const startingSlNo = importMode === 'replace' ? 1 : existingItemsCount + 1;
      const itemsWithSlNo = importPreview.items.map((item, index) => ({
        ...item,
        slNo: startingSlNo + index
      }));

      onItemsImported(itemsWithSlNo, importMode);
      setImportPreview(null);
      setShowPreview(false);
    }
  };

  return (
    <div className="bulk-import-section">
      <div style={{ 
        backgroundColor: '#f8f9fa', 
        padding: '1rem', 
        borderRadius: '8px', 
        marginBottom: '1rem',
        border: '1px solid #e9ecef'
      }}>
        <h3 style={{ marginBottom: '1rem', color: '#495057' }}>📥 Bulk Import Items</h3>
        
        <div style={{ display: 'flex', gap: '1rem', flexWrap: 'wrap', alignItems: 'center' }}>
          {/* CSV Upload */}
          <CSVReader
            onUploadAccepted={handleCSVUpload}
            config={{
              header: true,
              skipEmptyLines: true,
            }}
          >
            {({
              getRootProps,
              acceptedFile,
              getRemoveFileProps,
            }: any) => (
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <button
                  type="button"
                  {...getRootProps()}
                  className="btn btn-secondary"
                  style={{ 
                    backgroundColor: '#28a745',
                    borderColor: '#28a745',
                    color: 'white'
                  }}
                >
                  📁 Upload CSV File
                </button>
                {acceptedFile && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                    <span style={{ fontSize: '0.9rem', color: '#666' }}>
                      {acceptedFile.name}
                    </span>
                    <button
                      type="button"
                      {...getRemoveFileProps()}
                      className="btn btn-danger"
                      style={{ padding: '0.25rem 0.5rem', fontSize: '0.8rem' }}
                    >
                      ✕
                    </button>
                  </div>
                )}
              </div>
            )}
          </CSVReader>

          {/* Paste from Excel */}
          <button
            type="button"
            onClick={() => setShowPasteArea(!showPasteArea)}
            className="btn btn-secondary"
            style={{ 
              backgroundColor: '#007bff',
              borderColor: '#007bff',
              color: 'white'
            }}
          >
            📋 Paste from Excel
          </button>

          {/* Help Text */}
          <div style={{ fontSize: '0.9rem', color: '#666', fontStyle: 'italic' }}>
            Import multiple items at once via CSV or copy-paste from Excel
          </div>
        </div>

        {/* Paste Area */}
        {showPasteArea && (
          <div style={{ marginTop: '1rem', padding: '1rem', backgroundColor: 'white', border: '1px solid #ddd', borderRadius: '4px' }}>
            <h4>Paste Excel Data</h4>
            <p style={{ fontSize: '0.9rem', color: '#666', margin: '0.5rem 0' }}>
              Copy data from Excel (Ctrl+C) and paste it here (Ctrl+V). Include headers: Drawing No, Rev, Item No, Qty, Remarks
            </p>
            <textarea
              ref={textareaRef}
              value={pasteData}
              onChange={(e) => setPasteData(e.target.value)}
              placeholder="Paste your Excel data here... (tab-separated values)"
              style={{
                width: '100%',
                height: '120px',
                padding: '0.5rem',
                border: '1px solid #ccc',
                borderRadius: '4px',
                fontSize: '0.9rem',
                fontFamily: 'monospace'
              }}
            />
            <div style={{ marginTop: '0.5rem', display: 'flex', gap: '0.5rem' }}>
              <button
                type="button"
                onClick={handlePaste}
                className="btn btn-primary"
                disabled={!pasteData.trim()}
              >
                Process Pasted Data
              </button>
              <button
                type="button"
                onClick={() => {
                  setShowPasteArea(false);
                  setPasteData('');
                }}
                className="btn btn-secondary"
              >
                Cancel
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Import Preview Modal */}
      {showPreview && importPreview && (
        <div style={{
          position: 'fixed',
          top: '0',
          left: '0',
          right: '0',
          bottom: '0',
          backgroundColor: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          zIndex: 1000
        }}>
          <div style={{
            backgroundColor: 'white',
            padding: '2rem',
            borderRadius: '8px',
            maxWidth: '90vw',
            maxHeight: '80vh',
            overflow: 'auto',
            minWidth: '600px'
          }}>
            <h3>Import Preview</h3>
            
            <div style={{ marginBottom: '1rem' }}>
              <p>
                <strong>Items to import:</strong> {importPreview.items.length}<br/>
                <strong>Errors found:</strong> {importPreview.errors.length}
              </p>
              
              <div style={{ marginBottom: '1rem' }}>
                <label style={{ marginRight: '1rem' }}>
                  <input
                    type="radio"
                    name="importMode"
                    value="add"
                    checked={importMode === 'add'}
                    onChange={(e) => setImportMode(e.target.value as 'add')}
                  />
                  Add to existing {existingItemsCount} items
                </label>
                <label>
                  <input
                    type="radio"
                    name="importMode"
                    value="replace"
                    checked={importMode === 'replace'}
                    onChange={(e) => setImportMode(e.target.value as 'replace')}
                  />
                  Replace all existing items
                </label>
              </div>
            </div>

            {/* Errors */}
            {importPreview.errors.length > 0 && (
              <div style={{ marginBottom: '1rem' }}>
                <h4 style={{ color: '#dc3545' }}>Errors to Fix:</h4>
                <div style={{ maxHeight: '150px', overflow: 'auto', border: '1px solid #dc3545', padding: '0.5rem', backgroundColor: '#f8d7da' }}>
                  {importPreview.errors.map((error, index) => (
                    <div key={index} style={{ fontSize: '0.9rem', marginBottom: '0.25rem' }}>
                      <strong>Row {error.row}:</strong> {error.message}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Preview Data */}
            {importPreview.items.length > 0 && (
              <div style={{ marginBottom: '1rem' }}>
                <h4>Preview Data:</h4>
                <div style={{ maxHeight: '300px', overflow: 'auto', border: '1px solid #ccc' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.9rem' }}>
                    <thead>
                      <tr style={{ backgroundColor: '#f8f9fa' }}>
                        <th style={{ padding: '0.5rem', border: '1px solid #ddd' }}>Sl No</th>
                        <th style={{ padding: '0.5rem', border: '1px solid #ddd' }}>Drawing No</th>
                        <th style={{ padding: '0.5rem', border: '1px solid #ddd' }}>Rev</th>
                        <th style={{ padding: '0.5rem', border: '1px solid #ddd' }}>Item No</th>
                        <th style={{ padding: '0.5rem', border: '1px solid #ddd' }}>Qty</th>
                        <th style={{ padding: '0.5rem', border: '1px solid #ddd' }}>Remarks</th>
                      </tr>
                    </thead>
                    <tbody>
                      {importPreview.items.map((item, index) => (
                        <tr key={index}>
                          <td style={{ padding: '0.5rem', border: '1px solid #ddd' }}>
                            {importMode === 'replace' ? index + 1 : existingItemsCount + index + 1}
                          </td>
                          <td style={{ padding: '0.5rem', border: '1px solid #ddd' }}>{item.drawingNo}</td>
                          <td style={{ padding: '0.5rem', border: '1px solid #ddd' }}>{item.rev}</td>
                          <td style={{ padding: '0.5rem', border: '1px solid #ddd' }}>{item.itemNo}</td>
                          <td style={{ padding: '0.5rem', border: '1px solid #ddd' }}>{item.qty}</td>
                          <td style={{ padding: '0.5rem', border: '1px solid #ddd' }}>{item.remarks}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Actions */}
            <div style={{ display: 'flex', gap: '1rem', justifyContent: 'flex-end' }}>
              <button
                type="button"
                onClick={() => {
                  setImportPreview(null);
                  setShowPreview(false);
                }}
                className="btn btn-secondary"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleConfirmImport}
                className="btn btn-primary"
                disabled={!importPreview.isValid || importPreview.items.length === 0}
              >
                Import {importPreview.items.length} Items
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default BulkItemImport;