using System.Text.Json.Serialization;

namespace DocumentAutomationDemo.Models
{
    public class DocumentValues
    {
        [JsonPropertyName("templateId")]
        public string TemplateId { get; set; } = string.Empty;

        [JsonPropertyName("values")]
        public Dictionary<string, string> Values { get; set; } = new();

        [JsonPropertyName("exportFormat")]
        public ExportFormat ExportFormat { get; set; } = ExportFormat.Original;
    }
}
