import React, { useState } from 'react';
import { InspectionFormData, InspectionItem, GenerationResponse } from '../types';
import { excelApi } from '../services/api';
import DownloadResults from './DownloadResults';
import BulkItemImport from './BulkItemImport';

const InspectionForm: React.FC = () => {
  const [formData, setFormData] = useState<InspectionFormData>({
    project: 'JA RODDING-BUTTS CRUSHING PLANT',
    contractNo: 'FS131',
    company: 'Emirates Global Aluminium',
    date: new Date().toISOString().split('T')[0], // Default to today
    contractor: 'Fives Solios',
    subContractor: 'Bashaer Gulf Projects and technical Services LLC',
    itpNo: 'J6P001466-TK6239-Q04-0001 Rev.02',
    itpRevNo: '3.5',
    itpReferenceClause: 'AWS D1.1/AWSD1.1M : 2025',
    certNo: 'BG-J082-850-9-QAC-SSDIR-028',
    rfiNo: 'BG-J082-850-QAC',
    itemDescription: 'Conveyor gallery-Steel structure',
    items: [
      {
        slNo: 1,
        drawingNo: 'FS131-704365-324401-324-1-B-128',
        rev: '03',
        itemNo: 'B-128',
        qty: 1,
        remarks: 'Conveyor gallery-Steel structure'
      },
      {
        slNo: 2,
        drawingNo: 'FS131-704365-324401-324-1-B-129',
        rev: '04',
        itemNo: 'B-129',
        qty: 1,
        remarks: 'Optional remarks'
      }
    ]
  });

  const [isSubmitting, setIsSubmitting] = useState(false);
  const [results, setResults] = useState<GenerationResponse | null>(null);
  const [errors, setErrors] = useState<Record<string, string>>({});

  const handleInputChange = (field: keyof InspectionFormData, value: string) => {
    setFormData(prev => ({
      ...prev,
      [field]: value
    }));
    
    // Clear error when user starts typing
    if (errors[field]) {
      setErrors(prev => {
        const newErrors = { ...prev };
        delete newErrors[field];
        return newErrors;
      });
    }
  };

  const handleItemChange = (index: number, field: keyof InspectionItem, value: string | number) => {
    setFormData(prev => ({
      ...prev,
      items: prev.items.map((item, i) => 
        i === index ? { ...item, [field]: value } : item
      )
    }));
  };

  const addItem = () => {
    setFormData(prev => ({
      ...prev,
      items: [
        ...prev.items,
        {
          slNo: prev.items.length + 1,
          drawingNo: '',
          rev: '',
          itemNo: '',
          qty: 1,
          remarks: ''
        }
      ]
    }));
  };

  const removeItem = (index: number) => {
    if (formData.items.length > 1) {
      setFormData(prev => ({
        ...prev,
        items: prev.items
          .filter((_, i) => i !== index)
          .map((item, i) => ({ ...item, slNo: i + 1 }))
      }));
    }
  };

  const handleItemsImported = (items: InspectionItem[], mode: 'replace' | 'add') => {
    setFormData(prev => {
      if (mode === 'replace') {
        // Replace all existing items with imported ones
        return {
          ...prev,
          items: items
        };
      } else {
        // Add imported items to existing ones
        const existingItems = prev.items;
        const nextSlNo = existingItems.length > 0 ? Math.max(...existingItems.map(item => item.slNo)) + 1 : 1;
        
        const itemsWithUpdatedSlNo = items.map((item, index) => ({
          ...item,
          slNo: nextSlNo + index
        }));
        
        return {
          ...prev,
          items: [...existingItems, ...itemsWithUpdatedSlNo]
        };
      }
    });
    
    // Clear any existing errors
    setErrors({});
  };

  const validateForm = (): boolean => {
    const newErrors: Record<string, string> = {};

    // Validate header fields
    const requiredFields: (keyof InspectionFormData)[] = [
      'project', 'contractNo', 'company', 'date', 'contractor', 
      'subContractor', 'itpNo', 'itpRevNo', 'itpReferenceClause', 
      'certNo', 'rfiNo', 'itemDescription'
    ];

    requiredFields.forEach(field => {
      if (!formData[field] || (typeof formData[field] === 'string' && formData[field].trim() === '')) {
        newErrors[field] = 'This field is required';
      }
    });

    // Validate items
    formData.items.forEach((item, index) => {
      if (!item.drawingNo.trim()) {
        newErrors[`item_${index}_drawingNo`] = 'Drawing number is required';
      }
      if (!item.rev.trim()) {
        newErrors[`item_${index}_rev`] = 'Revision is required';
      }
      if (!item.itemNo.trim()) {
        newErrors[`item_${index}_itemNo`] = 'Item number is required';
      }
      if (item.qty <= 0) {
        newErrors[`item_${index}_qty`] = 'Quantity must be greater than 0';
      }
    });

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!validateForm()) {
      return;
    }

    setIsSubmitting(true);
    setResults(null);

    try {
      const response = await excelApi.generateFiles(formData);
      setResults(response);
      
      if (response.success) {
        // Scroll to results
        setTimeout(() => {
          const resultsElement = document.querySelector('.results-container');
          if (resultsElement) {
            resultsElement.scrollIntoView({ behavior: 'smooth' });
          }
        }, 100);
      }
    } catch (error: any) {
      setResults({
        success: false,
        files: [],
        batchId: '',
        message: error.message
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  const resetForm = () => {
    setFormData({
      project: 'JA RODDING-BUTTS CRUSHING PLANT',
      contractNo: 'FS131',
      company: 'Emirates Global Aluminium',
      date: new Date().toISOString().split('T')[0],
      contractor: 'Fives Solios',
      subContractor: 'Bashaer Gulf Projects and technical Services LLC',
      itpNo: 'J6P001466-TK6239-Q04-0001 Rev.02',
      itpRevNo: '3.5',
      itpReferenceClause: 'AWS D1.1/AWSD1.1M : 2025',
      certNo: 'BG-J082-850-9-QAC-SSDIR-028',
      rfiNo: 'BG-J082-850-QAC',
      itemDescription: 'Conveyor gallery-Steel structure',
      items: [
        {
          slNo: 1,
          drawingNo: 'FS131-704365-324401-324-1-B-128',
          rev: '03',
          itemNo: 'B-128',
          qty: 1,
          remarks: 'Conveyor gallery-Steel structure'
        }
      ]
    });
    setResults(null);
    setErrors({});
  };

  return (
    <>
      <div className="form-container">
        <form onSubmit={handleSubmit}>
          {/* Header Information Section */}
          <div className="form-section">
            <h2>📋 Inspection Release Note Information</h2>
            
            <div className="form-row">
              <div className="form-group">
                <label>Project *</label>
                <input
                  type="text"
                  value={formData.project}
                  onChange={(e) => handleInputChange('project', e.target.value)}
                  className={errors.project ? 'error' : ''}
                  placeholder="Enter project name"
                />
                {errors.project && <span className="error-message">{errors.project}</span>}
              </div>
              
              <div className="form-group">
                <label>Contract No *</label>
                <input
                  type="text"
                  value={formData.contractNo}
                  onChange={(e) => handleInputChange('contractNo', e.target.value)}
                  className={errors.contractNo ? 'error' : ''}
                  placeholder="Enter contract number"
                />
                {errors.contractNo && <span className="error-message">{errors.contractNo}</span>}
              </div>
            </div>

            <div className="form-row">
              <div className="form-group">
                <label>Company *</label>
                <input
                  type="text"
                  value={formData.company}
                  onChange={(e) => handleInputChange('company', e.target.value)}
                  className={errors.company ? 'error' : ''}
                  placeholder="Enter company name"
                />
                {errors.company && <span className="error-message">{errors.company}</span>}
              </div>
              
              <div className="form-group">
                <label>Date *</label>
                <input
                  type="date"
                  value={formData.date}
                  onChange={(e) => handleInputChange('date', e.target.value)}
                  className={errors.date ? 'error' : ''}
                />
                {errors.date && <span className="error-message">{errors.date}</span>}
              </div>
            </div>

            <div className="form-row">
              <div className="form-group">
                <label>Contractor *</label>
                <input
                  type="text"
                  value={formData.contractor}
                  onChange={(e) => handleInputChange('contractor', e.target.value)}
                  className={errors.contractor ? 'error' : ''}
                  placeholder="Enter contractor name"
                />
                {errors.contractor && <span className="error-message">{errors.contractor}</span>}
              </div>
              
              <div className="form-group">
                <label>Sub-Contractor *</label>
                <input
                  type="text"
                  value={formData.subContractor}
                  onChange={(e) => handleInputChange('subContractor', e.target.value)}
                  className={errors.subContractor ? 'error' : ''}
                  placeholder="Enter sub-contractor name"
                />
                {errors.subContractor && <span className="error-message">{errors.subContractor}</span>}
              </div>
            </div>

            <div className="form-row">
              <div className="form-group">
                <label>ITP No *</label>
                <input
                  type="text"
                  value={formData.itpNo}
                  onChange={(e) => handleInputChange('itpNo', e.target.value)}
                  className={errors.itpNo ? 'error' : ''}
                  placeholder="Enter ITP number"
                />
                {errors.itpNo && <span className="error-message">{errors.itpNo}</span>}
              </div>
              
              <div className="form-group">
                <label>ITP Rev No *</label>
                <input
                  type="text"
                  value={formData.itpRevNo}
                  onChange={(e) => handleInputChange('itpRevNo', e.target.value)}
                  className={errors.itpRevNo ? 'error' : ''}
                  placeholder="Enter ITP revision number"
                />
                {errors.itpRevNo && <span className="error-message">{errors.itpRevNo}</span>}
              </div>
            </div>

            <div className="form-row">
              <div className="form-group">
                <label>ITP Reference Clause *</label>
                <input
                  type="text"
                  value={formData.itpReferenceClause}
                  onChange={(e) => handleInputChange('itpReferenceClause', e.target.value)}
                  className={errors.itpReferenceClause ? 'error' : ''}
                  placeholder="Enter ITP reference clause"
                />
                {errors.itpReferenceClause && <span className="error-message">{errors.itpReferenceClause}</span>}
              </div>
              
              <div className="form-group">
                <label>Cert No *</label>
                <input
                  type="text"
                  value={formData.certNo}
                  onChange={(e) => handleInputChange('certNo', e.target.value)}
                  className={errors.certNo ? 'error' : ''}
                  placeholder="Enter certificate number"
                />
                {errors.certNo && <span className="error-message">{errors.certNo}</span>}
              </div>
            </div>

            <div className="form-row">
              <div className="form-group">
                <label>RFI No *</label>
                <input
                  type="text"
                  value={formData.rfiNo}
                  onChange={(e) => handleInputChange('rfiNo', e.target.value)}
                  className={errors.rfiNo ? 'error' : ''}
                  placeholder="Enter RFI number"
                />
                {errors.rfiNo && <span className="error-message">{errors.rfiNo}</span>}
              </div>
              
              <div className="form-group">
                <label>Item Description *</label>
                <input
                  type="text"
                  value={formData.itemDescription}
                  onChange={(e) => handleInputChange('itemDescription', e.target.value)}
                  className={errors.itemDescription ? 'error' : ''}
                  placeholder="Enter item description for all sheets"
                />
                {errors.itemDescription && <span className="error-message">{errors.itemDescription}</span>}
              </div>
            </div>
          </div>

          {/* Items Section */}
          <div className="form-section">
            <h2>📊 Items for Release</h2>
            
            {/* Bulk Import Component */}
            <BulkItemImport 
              onItemsImported={handleItemsImported}
              existingItemsCount={formData.items.length}
            />
            
            <div style={{ overflowX: 'auto' }}>
              <table className="items-table">
                <thead>
                  <tr>
                    <th>Sl No</th>
                    <th>Drawing No *</th>
                    <th>Rev *</th>
                    <th>Item No *</th>
                    <th>Qty (Nos.) *</th>
                    <th>Remarks</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {formData.items.map((item, index) => (
                    <tr key={index}>
                      <td>{item.slNo}</td>
                      <td>
                        <input
                          type="text"
                          value={item.drawingNo}
                          onChange={(e) => handleItemChange(index, 'drawingNo', e.target.value)}
                          className={errors[`item_${index}_drawingNo`] ? 'error' : ''}
                          placeholder="Drawing number"
                        />
                        {errors[`item_${index}_drawingNo`] && (
                          <div className="error-message">{errors[`item_${index}_drawingNo`]}</div>
                        )}
                      </td>
                      <td>
                        <input
                          type="text"
                          value={item.rev}
                          onChange={(e) => handleItemChange(index, 'rev', e.target.value)}
                          className={errors[`item_${index}_rev`] ? 'error' : ''}
                          placeholder="Rev"
                        />
                        {errors[`item_${index}_rev`] && (
                          <div className="error-message">{errors[`item_${index}_rev`]}</div>
                        )}
                      </td>
                      <td>
                        <input
                          type="text"
                          value={item.itemNo}
                          onChange={(e) => handleItemChange(index, 'itemNo', e.target.value)}
                          className={errors[`item_${index}_itemNo`] ? 'error' : ''}
                          placeholder="Item number"
                        />
                        {errors[`item_${index}_itemNo`] && (
                          <div className="error-message">{errors[`item_${index}_itemNo`]}</div>
                        )}
                      </td>
                      <td>
                        <input
                          type="number"
                          value={item.qty}
                          onChange={(e) => handleItemChange(index, 'qty', parseInt(e.target.value) || 0)}
                          className={errors[`item_${index}_qty`] ? 'error' : ''}
                          min="1"
                        />
                        {errors[`item_${index}_qty`] && (
                          <div className="error-message">{errors[`item_${index}_qty`]}</div>
                        )}
                      </td>
                      <td>
                        <input
                          type="text"
                          value={item.remarks || ''}
                          onChange={(e) => handleItemChange(index, 'remarks', e.target.value)}
                          placeholder="Optional remarks"
                        />
                      </td>
                      <td>
                        <button
                          type="button"
                          onClick={() => removeItem(index)}
                          className="btn btn-danger"
                          disabled={formData.items.length === 1}
                          style={{ padding: '0.5rem' }}
                        >
                          ✕
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="btn-group">
              <button
                type="button"
                onClick={addItem}
                className="btn btn-secondary"
              >
                ➕ Add Item
              </button>
            </div>
          </div>

          {/* Submit Section */}
          <div className="btn-group">
            <button
              type="submit"
              disabled={isSubmitting}
              className="btn btn-primary"
            >
              {isSubmitting ? 'Generating Excel Files...' : '📊 Generate Excel Files'}
            </button>
            
            <button
              type="button"
              onClick={resetForm}
              className="btn btn-secondary"
              disabled={isSubmitting}
            >
              🔄 Reset Form
            </button>
          </div>
        </form>
      </div>

      {/* Results Section */}
      {(results || isSubmitting) && (
        <DownloadResults 
          results={results} 
          isLoading={isSubmitting} 
        />
      )}
    </>
  );
};

export default InspectionForm;