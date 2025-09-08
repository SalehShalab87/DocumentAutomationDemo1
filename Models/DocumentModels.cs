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
        Html,      // HTML with external images
        HtmlEmail, // HTML with embedded base64 images (email-friendly)
        Pdf
    }

    public enum DocumentType
    {
        Word,      // .docx
        Excel,     // .xlsx  
        PowerPoint // .pptx
    }

    public class DocumentEmbeddingRequest
    {
        public string MainTemplateId { get; set; } = string.Empty;
        public List<PlaceholderValue> MainTemplateValues { get; set; } = new();
        public List<EmbedInfo> Embeddings { get; set; } = new();
        public ExportFormat ExportFormat { get; set; } = ExportFormat.Original;
    }

    public class EmbedInfo
    {
        public string EmbedTemplateId { get; set; } = string.Empty;
        public List<PlaceholderValue> EmbedTemplateValues { get; set; } = new();
        public string EmbedPlaceholder { get; set; } = string.Empty;
    }

    public class DocumentEmbeddingValues
    {
        public string MainTemplateId { get; set; } = string.Empty;
        public Dictionary<string, string> MainValues { get; set; } = new();
        public List<EmbedValues> Embeddings { get; set; } = new();
        public ExportFormat ExportFormat { get; set; } = ExportFormat.Original;
    }

    public class EmbedValues
    {
        public string EmbedTemplateId { get; set; } = string.Empty;
        public Dictionary<string, string> Values { get; set; } = new();
        public string EmbedPlaceholder { get; set; } = string.Empty;
    }
}
