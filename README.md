# 📄 Document Automation Demo

A powerful, enterprise-grade document automation system built with .NET 8 that generates dynamic Office documents (Word, Excel, PowerPoint) with advanced document embedding capabilities and automatic field refresh. Perfect for Angular applications requiring server-side document generation.

## ✨ Features

### 🚀 Core Capabilities
- **Multi-Format Support**: Generate Word (.docx), Excel (.xlsx), PowerPoint (.pptx) documents
- **Custom Properties Architecture**: Professional metadata-based approach using document properties
- **Automatic Field Refresh**: DOCPROPERTY fields update automatically without manual intervention
- **Template Management**: Register, store, and reuse document templates
- **Multiple Export Formats**: Word, HTML, PDF output options
- **Angular-Ready**: RESTful API design perfect for frontend integration

### 🆕 Document Embedding Feature
- **Word-to-Word Embedding**: Embed complete Word documents within other Word documents
- **Multiple Embeddings**: Support for multiple embedded documents in a single main document
- **Format Preservation**: Maintains styles, formatting, images, tables, numbering, and alignment
- **Flexible Placeholders**: Specify custom placeholders for embedding locations
- **JSON Configuration**: Support for JSON-based embedding configuration
- **Style Import**: Automatic style conflict resolution and import from embedded documents

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
│   ├── DocumentModels.cs          # Data models and embedding models
│   └── DocumentValues.cs          # JSON deserialization models
├── Services/
│   ├── TemplateService.cs         # Template management and placeholders
│   ├── DocumentGenerationService.cs # Single document generation
│   └── DocumentEmbeddingService.cs  # 🆕 Document embedding service
├── Templates/                     # Stored document templates
├── Output/                        # Generated documents
├── Examples/
│   ├── single_embed_example.json  # Single embedding example
│   ├── multi_embed_example.json   # Multiple embeddings example
│   └── test_ask_format.json       # Standard generation example
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

### Menu Options
1. **Register a new template** - Add Word/Excel/PowerPoint templates
2. **View all templates** - List registered templates with placeholders
3. **Generate document from template** - Single template document generation
4. **Generate document with embedded templates** - 🆕 Multiple document embedding
5. **Delete a template** - Remove registered templates
6. **Exit** - Close application

### 1. Template Registration
```csharp
// Register a new Word template
var templateService = new TemplateService();
string templateId = templateService.RegisterTemplate(
    "path/to/template.docx", 
    "My Template"
);
```

### 2. Standard Document Generation
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

### 3. 🆕 Document Embedding
```csharp
var embeddingService = new DocumentEmbeddingService(templateService);
var request = new DocumentEmbeddingRequest
{
    MainTemplateId = "main-template-id",
    MainTemplateValues = mainPlaceholders,
    Embeddings = new List<EmbedInfo>
    {
        new EmbedInfo
        {
            EmbedTemplateId = "embed-template-id",
            EmbedTemplateValues = embedPlaceholders,
            EmbedPlaceholder = "ICS_EmbeddedDocument"
        }
    },
    ExportFormat = ExportFormat.Word
};

string outputPath = embeddingService.GenerateDocumentWithEmbedding(request);
```

### 4. Template Placeholders

#### Supported Field Formats
- **Simple Fields**: `{ DOCPROPERTY PropertyName \* MERGEFORMAT }`
- **Complex Fields**: Field codes with Begin → Instruction → Separate → Result → End structure

#### Example Template Setup in Word
1. Insert → Quick Parts → Field
2. Choose "DocProperty" 
3. Enter property name (e.g., "CustomerName")
4. Select "Preserve formatting during updates"

## 🔌 Document Embedding

### How Document Embedding Works
The embedding feature allows you to:

1. **Main Template**: Start with a Word document containing placeholders for embedded content
2. **Embed Templates**: Select Word documents to be inserted into the main document  
3. **Placeholder Mapping**: Specify where each embedded document should be placed
4. **Format Preservation**: All formatting, styles, images, and tables are preserved
5. **Multiple Embeddings**: Support for multiple embedded documents in one main document

### What Gets Preserved During Embedding
- ✅ **Text Formatting**: Bold, italic, fonts, colors, sizes
- ✅ **Paragraph Styles**: Headings, normal text, custom styles  
- ✅ **Tables**: Structure, borders, cell formatting, alignment
- ✅ **Images**: Pictures with original positioning and sizing
- ✅ **Numbering**: Bullet points, numbered lists, multilevel lists
- ✅ **Alignment**: Left, center, right, justified alignment
- ✅ **Custom Styles**: User-defined character and paragraph styles

### JSON Configuration Examples

#### Single Embedding
```json
{
    "mainTemplateId": "main-template-id",
    "mainValues": {
        "ICS_CustomerName": "Customer Name",
        "ICS_PolicyNumber": "POL-12345"
    },
    "embeddings": [
        {
            "embedTemplateId": "embed-template-id",
            "values": {
                "ICS_BenefitName": "Health Coverage",
                "ICS_BenefitLimit": "$50,000"
            },
            "embedPlaceholder": "ICS_EmbeddedDocument"
        }
    ],
    "exportFormat": "Word"
}
```

#### Multiple Embeddings  
```json
{
    "mainTemplateId": "main-template-id",
    "mainValues": {
        "ICS_CustomerName": "Customer Name"
    },
    "embeddings": [
        {
            "embedTemplateId": "health-template-id",
            "values": { "ICS_Coverage": "Full Health Coverage" },
            "embedPlaceholder": "ICS_HealthDocument"
        },
        {
            "embedTemplateId": "dental-template-id", 
            "values": { "ICS_Coverage": "Basic Dental Coverage" },
            "embedPlaceholder": "ICS_DentalDocument"
        }
    ],
    "exportFormat": "PDF"
}
```

### Template Setup for Embedding

#### Main Template Setup
1. Create your main Word document
2. Add custom properties for main document data
3. Insert placeholders where embedded documents should appear:
   - Use text placeholders like `ICS_EmbeddedDocument`
   - Or use DOCPROPERTY fields pointing to placeholder properties

#### Embed Template Setup
1. Create Word documents with their own custom properties
2. Design content with proper formatting, styles, and images
3. These will be embedded as complete formatted documents

## 🔌 Angular Integration

### API Endpoint Design
```typescript
// TypeScript interfaces for Angular
interface DocumentRequest {
  templateId: string;
  placeholders: { [key: string]: string };
  filename: string;
  format: 'word' | 'html' | 'pdf';
}

interface DocumentEmbeddingRequest {
  mainTemplateId: string;
  mainValues: { [key: string]: string };
  embeddings: EmbedInfo[];
  exportFormat: 'word' | 'html' | 'pdf';
}

interface EmbedInfo {
  embedTemplateId: string;
  values: { [key: string]: string };
  embedPlaceholder: string;
}

// Example Angular service calls
generateDocument(request: DocumentRequest): Observable<Blob> {
  return this.http.post('/api/documents/generate', request, {
    responseType: 'blob'
  });
}

generateEmbeddedDocument(request: DocumentEmbeddingRequest): Observable<Blob> {
  return this.http.post('/api/documents/generate-embedded', request, {
    responseType: 'blob'
  });
}
```

### Sample API Controller
```csharp
[ApiController]
[Route("api/[controller]")]
public class DocumentsController : ControllerBase
{
    private readonly IDocumentGenerationService _docService;
    private readonly IDocumentEmbeddingService _embeddingService;
    
    [HttpPost("generate")]
    public async Task<IActionResult> GenerateDocument([FromBody] DocumentRequest request)
    {
        // Standard document generation
        var outputPath = _docService.GenerateDocument(/* parameters */);
        var fileBytes = await File.ReadAllBytesAsync(outputPath);
        return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    }

    [HttpPost("generate-embedded")]
    public async Task<IActionResult> GenerateEmbeddedDocument([FromBody] DocumentEmbeddingRequest request)
    {
        // Document embedding generation
        var outputPath = _embeddingService.GenerateDocumentWithEmbedding(request);
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
  "Placeholders": ["CustomerName", "PolicyNumber", "ICS_EmbeddedDocument"],
  "DocumentType": 0
}
```

### Document Types
- `0`: Word Document (.docx)
- `1`: Excel Workbook (.xlsx) 
- `2`: PowerPoint Presentation (.pptx)

### Export Formats
1. **Original/Word** - Word document (.docx)
2. **HTML** - Convert to HTML with external images (.html)
3. **PDF** - Convert to PDF format (.pdf)

## 🔍 Advanced Features

### Automatic Field Refresh
The system implements multiple strategies to ensure fields display correctly:

1. **Direct Field Updates**: Modifies field result text programmatically
2. **Auto-Update Setting**: Sets `UpdateFieldsOnOpen = true` in document settings
3. **Cache Clearing**: Marks fields as "dirty" to force recalculation

### Document Embedding Architecture
- **Style Import**: Automatically imports and resolves style conflicts
- **Content Cloning**: Deep cloning of all document elements
- **Reference Management**: Proper handling of image and numbering references
- **Memory Management**: Efficient processing with proper cleanup

### Error Handling
- Comprehensive exception handling in all services
- Detailed logging for debugging
- Graceful degradation when field refresh fails
- Fallback mechanisms for embedding failures

### Performance Optimization
- Efficient field scanning using OpenXML
- Minimal memory footprint
- Fast template processing
- Automatic temporary file cleanup
- Style deduplication to reduce file size

## 🎯 Use Cases

### Standard Document Generation
- **Invoices**: Generate invoices with dynamic customer data
- **Reports**: Create reports with calculated fields and charts
- **Letters**: Personalized correspondence with mail merge
- **Certificates**: Generate certificates with recipient details

### Document Embedding Applications
- **Insurance Policies**: Embed coverage details, terms, and conditions
- **Legal Contracts**: Include standard clauses and custom appendices  
- **Comprehensive Reports**: Combine executive summaries with detailed sections
- **Proposals**: Merge cover letters with technical specifications
- **Training Materials**: Embed modules and assessment documents

### Benefits
- **Consistency**: Maintain brand standards across all embedded content
- **Efficiency**: Reuse common document sections
- **Flexibility**: Mix and match content based on requirements
- **Quality**: Preserve professional formatting throughout

## 🔮 Future Enhancements

- Excel-to-Excel embedding support
- PowerPoint-to-PowerPoint embedding support  
- Cross-format embedding capabilities (Word into Excel/PowerPoint)
- Advanced positioning controls for embedded content
- Batch embedding operations
- Template versioning and rollback
- Real-time collaborative editing
- Advanced styling conflict resolution


**Built with ❤️ for seamless document automation**
