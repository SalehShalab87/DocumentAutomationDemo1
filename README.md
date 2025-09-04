# 📄 Document Automation Demo

A powerful, enterprise-grade document automation system built with .NET 8 that generates dynamic Office documents (Word, Excel, PowerPoint) with automatic field refresh capabilities. Perfect for Angular applications requiring server-side document generation.

## ✨ Features

### 🚀 Core Capabilities
- **Multi-Format Support**: Generate Word (.docx), Excel (.xlsx), PowerPoint (.pptx) documents
- **Custom Properties Architecture**: Professional metadata-based approach using document properties
- **Automatic Field Refresh**: DOCPROPERTY fields update automatically without manual intervention
- **Template Management**: Register, store, and reuse document templates
- **Multiple Export Formats**: Word, HTML, PDF output options
- **Angular-Ready**: RESTful API design perfect for frontend integration

### 🔧 Advanced Features
- **Smart Field Detection**: Handles both simple and complex DOCPROPERTY field formats
- **Auto-Update on Open**: Documents automatically refresh fields when opened in Word
- **Field Cache Management**: Clears field caches to ensure immediate value display
- **Headers & Footers Support**: Updates fields in all document sections
- **Error Handling**: Comprehensive exception handling with detailed logging

## 🏗️ Architecture

### Project Structure
```
DocumentAutomationDemo/
├── Models/
│   └── DocumentModels.cs          # Data models and enums
├── Services/
│   ├── TemplateService.cs         # Template management and placeholders
│   └── DocumentGenerationService.cs # Document generation and field refresh
├── Templates/                     # Stored document templates
├── Output/                        # Generated documents
└── Program.cs                     # Console interface
```

### Technology Stack
- **.NET 8.0**: Modern, high-performance framework
- **DocumentFormat.OpenXml 3.0.2**: Office document manipulation
- **iTextSharp 5.5.13.3**: PDF generation capabilities
- **System.Text.Json**: Template metadata persistence

## 🚀 Quick Start

### Prerequisites
- .NET 8.0 SDK or later
- Visual Studio 2022 or Visual Studio Code

### Installation
1. Clone or download the project
2. Navigate to the project directory
3. Restore dependencies:
   ```bash
   dotnet restore
   ```
4. Build the project:
   ```bash
   dotnet build
   ```
5. Run the application:
   ```bash
   dotnet run
   ```

## 📖 Usage Guide

### 1. Template Registration
```csharp
// Register a new Word template
var templateService = new TemplateService();
string templateId = templateService.RegisterTemplate(
    "path/to/template.docx", 
    "My Template"
);
```

### 2. Document Generation
```csharp
var docService = new DocumentGenerationService(templateService);
var placeholders = new List<PlaceholderValue>
{
    new PlaceholderValue { Placeholder = "CustomerName", Value = "John Doe" },
    new PlaceholderValue { Placeholder = "PolicyNumber", Value = "P12345" }
};

string outputPath = docService.GenerateDocument(
    templateId, 
    placeholders, 
    "output-filename", 
    ExportFormat.Word
);
```

### 3. Template Placeholders

#### Supported Field Formats
- **Simple Fields**: `{ DOCPROPERTY PropertyName \* MERGEFORMAT }`
- **Complex Fields**: Field codes with Begin → Instruction → Separate → Result → End structure

#### Example Template Setup in Word
1. Insert → Quick Parts → Field
2. Choose "DocProperty" 
3. Enter property name (e.g., "CustomerName")
4. Select "Preserve formatting during updates"

## 🔌 Angular Integration

### API Endpoint Design
```typescript
// TypeScript interface for Angular
interface DocumentRequest {
  templateId: string;
  placeholders: { [key: string]: string };
  filename: string;
  format: 'word' | 'html' | 'pdf';
}

// Example Angular service call
generateDocument(request: DocumentRequest): Observable<Blob> {
  return this.http.post('/api/documents/generate', request, {
    responseType: 'blob'
  });
}
```

### Sample API Controller (to implement)
```csharp
[ApiController]
[Route("api/[controller]")]
public class DocumentsController : ControllerBase
{
    private readonly IDocumentGenerationService _docService;
    
    [HttpPost("generate")]
    public async Task<IActionResult> GenerateDocument([FromBody] DocumentRequest request)
    {
        // Implementation using DocumentGenerationService
        var outputPath = _docService.GenerateDocument(/* parameters */);
        var fileBytes = await File.ReadAllBytesAsync(outputPath);
        return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    }
}
```

## 🛠️ Configuration

### Template Management
Templates are stored in the `Templates/` directory with metadata in `templates.json`:
```json
{
  "Id": "guid",
  "Name": "Template Name",
  "FilePath": "path/to/template.docx",
  "CreatedDate": "2025-09-04T10:00:00Z",
  "Placeholders": ["CustomerName", "PolicyNumber"],
  "DocumentType": 0
}
```

### Document Types
- `0`: Word Document (.docx)
- `1`: Excel Workbook (.xlsx) 
- `2`: PowerPoint Presentation (.pptx)

## 🔍 Advanced Features

### Automatic Field Refresh
The system implements multiple strategies to ensure fields display correctly:

1. **Direct Field Updates**: Modifies field result text programmatically
2. **Auto-Update Setting**: Sets `UpdateFieldsOnOpen = true` in document settings
3. **Cache Clearing**: Marks fields as "dirty" to force recalculation

### Error Handling
- Comprehensive exception handling in all services
- Detailed logging for debugging
- Graceful degradation when field refresh fails

### Performance Optimization
- Efficient field scanning using OpenXML
- Minimal memory footprint
- Fast template processing


**Built with ❤️ for seamless document automation**
