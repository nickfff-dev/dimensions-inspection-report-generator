# Excel Generation Tool

A web application for generating dimensional inspection reports from form data.

## 🚀 Features

- **Dynamic Form Interface**: Input inspection release note data with dynamic table for items
- **Excel Generation**: Clone template.xlsx and populate with form data  
- **Multiple File Output**: Generate one Excel file per inspection item
- **Download Options**: Individual file downloads or batch ZIP download
- **Auto Cleanup**: Generated files automatically cleaned up after 24 hours
- **File Naming**: Smart naming convention: `DIM-INSPECT-{ItemNo}-{SerialNumber}.xlsx`

## 📋 Form Fields

### Header Information
- Project, Contract No, Company, Date
- Contractor, Sub-Contractor  
- ITP No, ITP Rev No, ITP Reference Clause
- Cert No, RFI No

### Dynamic Items Table
- Drawing No, Rev, Item No, Qty, Remarks
- Add/remove items dynamically
- Auto-increment serial numbers

## 🛠 Technology Stack

- **Backend**: Node.js + Express + TypeScript
- **Frontend**: React + TypeScript + Vite
- **Excel Processing**: ExcelJS for template cloning and data population
- **Validation**: Zod for robust input validation
- **File Management**: Archiver for ZIP downloads
- **Package Manager**: pnpm with workspace configuration

## 📁 Project Structure

```
excel-gen/
├── backend/                    # Node.js API server
│   ├── src/
│   │   ├── controllers/        # Request handlers
│   │   ├── routes/            # API endpoint definitions
│   │   ├── services/          # Excel generation logic
│   │   ├── middleware/        # Validation & error handling
│   │   ├── types/             # TypeScript interfaces
│   │   └── validation/        # Zod schemas
│   ├── templates/
│   │   └── template.xlsx      # Excel template file
│   └── generated/             # Generated Excel files
├── frontend/                  # React web app
│   ├── src/
│   │   ├── components/        # React components
│   │   ├── services/          # API client
│   │   └── types/             # TypeScript interfaces
└── package.json               # Workspace configuration
```

## 🔄 Workflow

1. **User Input**: Fill inspection form with header data and items
2. **Validation**: Client & server-side validation using Zod schemas
3. **Generation**: Backend clones template.xlsx for each item
4. **Population**: Replace placeholders with actual form data
5. **Download**: Provide individual or batch download options
6. **Cleanup**: Auto-remove files after 24 hours

## 🚀 Getting Started

### Prerequisites
- Node.js 18+
- pnpm package manager

### Installation
```bash
# Install all dependencies
pnpm install

# Start development servers
pnpm dev
```

### Development Servers
- **Frontend**: http://localhost:3000 (React + Vite)
- **Backend**: http://localhost:3001 (Express API)

### API Endpoints
- `POST /api/generate` - Generate Excel files from form data
- `GET /api/download/:filename` - Download individual file
- `GET /api/download/batch/:batchId` - Download ZIP archive

## 📝 Template Placeholders

The Excel template supports these placeholders:
- `{{PROJECT}}`, `{{CONTRACT_NO}}`, `{{COMPANY}}`, `{{DATE}}`
- `{{CONTRACTOR}}`, `{{SUB_CONTRACTOR}}`
- `{{ITP_NO}}`, `{{ITP_REV_NO}}`, `{{ITP_REFERENCE_CLAUSE}}`
- `{{CERT_NO}}`, `{{RFI_NO}}`
- `{{DRAWING_NO}}`, `{{REV}}`, `{{ITEM_NO}}`, `{{QTY}}`, `{{REMARKS}}`, `{{SL_NO}}`

## 🔧 Configuration

### Environment Variables
- `PORT` - Backend server port (default: 3001)

### File Management
- Generated files stored in `backend/generated/`
- Automatic cleanup after 24 hours
- ZIP archives for batch downloads

## 🧪 Testing

The application includes:
- Form validation (client & server-side)
- Error handling and user feedback
- File existence checks
- Secure filename validation
- CORS configuration for development

## 🛡 Security Features

- Input validation with Zod schemas
- Filename sanitization to prevent path traversal
- CORS protection
- No sensitive data logging
- File cleanup to prevent storage bloat

## 📋 Future Enhancements

- User authentication and session management
- Template upload functionality
- Custom field mapping
- Audit logging
- Database integration for form history
- Email notifications
- Custom styling/branding options