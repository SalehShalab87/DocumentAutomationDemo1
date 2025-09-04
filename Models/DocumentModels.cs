using System.Text.Json;

namespace DocumentAutomationDemo.Models
{
    public class DocumentTemplate
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public DateTime CreatedDate { get; set; }
        public List<string> Placeholders { get; set; } = new();
        public DocumentType DocumentType { get; set; } = DocumentType.Word;
    }

    public class PlaceholderValue
    {
        public string Placeholder { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    public class DocumentGenerationRequest
    {
        public string TemplateId { get; set; } = string.Empty;
        public List<PlaceholderValue> PlaceholderValues { get; set; } = new();
        public ExportFormat ExportFormat { get; set; } = ExportFormat.Original;
    }

    public enum ExportFormat
    {
        Original,  // Keep original format (Word/Excel/PowerPoint)
        Word,      // Convert to Word format
        Html,
        Pdf
    }

    public enum DocumentType
    {
        Word,      // .docx
        Excel,     // .xlsx  
        PowerPoint // .pptx
    }
}
