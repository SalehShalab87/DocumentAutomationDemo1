using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using DocumentAutomationDemo.Models;
using System.Text.Json;

namespace DocumentAutomationDemo.Services
{
    public interface ITemplateService
    {
        string RegisterTemplate(string filePath, string templateName);
        DocumentTemplate? GetTemplate(string templateId);
        List<DocumentTemplate> GetAllTemplates();
        bool DeleteTemplate(string templateId);
        List<string> ExtractPlaceholders(string templatePath);
        Dictionary<string, string> GetCustomPropertiesWithValues(string templatePath);
    }

    public class TemplateService : ITemplateService
    {
        private readonly string _templatesDirectory;
        private readonly string _metadataFile;
        private List<DocumentTemplate> _templates = new();

        public TemplateService()
        {
            _templatesDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            _metadataFile = Path.Combine(_templatesDirectory, "templates.json");
            
            // Ensure templates directory exists
            if (!Directory.Exists(_templatesDirectory))
            {
                Directory.CreateDirectory(_templatesDirectory);
            }

            LoadTemplates();
        }

        public string RegisterTemplate(string filePath, string templateName)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Template file not found: {filePath}");

            // Detect document type from file extension
            var documentType = GetDocumentType(filePath);
            
            // Generate unique ID
            string templateId = Guid.NewGuid().ToString();
            string fileExtension = Path.GetExtension(filePath);
            string fileName = $"{templateId}{fileExtension}";
            string destinationPath = Path.Combine(_templatesDirectory, fileName);

            // Copy template to templates directory
            File.Copy(filePath, destinationPath, true);

            // Extract placeholders
            var placeholders = ExtractPlaceholders(destinationPath);

            // Create template metadata
            var template = new DocumentTemplate
            {
                Id = templateId,
                Name = templateName,
                FilePath = destinationPath,
                CreatedDate = DateTime.Now,
                Placeholders = placeholders,
                DocumentType = documentType
            };

            _templates.Add(template);
            SaveTemplates();

            return templateId;
        }

        private Models.DocumentType GetDocumentType(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".docx" => Models.DocumentType.Word,
                ".xlsx" => Models.DocumentType.Excel,
                ".pptx" => Models.DocumentType.PowerPoint,
                _ => Models.DocumentType.Word // Default to Word
            };
        }

        public DocumentTemplate? GetTemplate(string templateId)
        {
            return _templates.FirstOrDefault(t => t.Id == templateId);
        }

        public List<DocumentTemplate> GetAllTemplates()
        {
            return _templates.ToList();
        }

        public bool DeleteTemplate(string templateId)
        {
            var template = GetTemplate(templateId);
            if (template == null) return false;

            // Delete file
            if (File.Exists(template.FilePath))
            {
                File.Delete(template.FilePath);
            }

            // Remove from list
            _templates.RemoveAll(t => t.Id == templateId);
            SaveTemplates();

            return true;
        }

        public List<string> ExtractPlaceholders(string templatePath)
        {
            var placeholders = new HashSet<string>();
            var documentType = GetDocumentType(templatePath);

            try
            {
                switch (documentType)
                {
                    case Models.DocumentType.Word:
                        using (var doc = WordprocessingDocument.Open(templatePath, false))
                        {
                            var customProps = ExtractCustomPropertiesFromOpenXml(doc);
                            foreach (var prop in customProps)
                            {
                                placeholders.Add(prop);
                            }
                        }
                        break;

                    case Models.DocumentType.Excel:
                        using (var doc = SpreadsheetDocument.Open(templatePath, false))
                        {
                            var customProps = ExtractCustomPropertiesFromOpenXml(doc);
                            foreach (var prop in customProps)
                            {
                                placeholders.Add(prop);
                            }
                        }
                        break;

                    case Models.DocumentType.PowerPoint:
                        using (var doc = PresentationDocument.Open(templatePath, false))
                        {
                            var customProps = ExtractCustomPropertiesFromOpenXml(doc);
                            foreach (var prop in customProps)
                            {
                                placeholders.Add(prop);
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not extract placeholders: {ex.Message}");
            }

            return placeholders.ToList();
        }

        private List<string> ExtractCustomPropertiesFromOpenXml(OpenXmlPackage document)
        {
            var properties = new List<string>();

            try
            {
                // Get custom file properties part - this method works for all OpenXML document types
                CustomFilePropertiesPart? customPropertiesPart = null;
                
                if (document is WordprocessingDocument wordDoc)
                {
                    customPropertiesPart = wordDoc.CustomFilePropertiesPart;
                }
                else if (document is SpreadsheetDocument excelDoc)
                {
                    customPropertiesPart = excelDoc.CustomFilePropertiesPart;
                }
                else if (document is PresentationDocument pptDoc)
                {
                    customPropertiesPart = pptDoc.CustomFilePropertiesPart;
                }

                if (customPropertiesPart?.Properties != null)
                {
                    foreach (var prop in customPropertiesPart.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
                    {
                        if (prop.Name?.Value != null)
                        {
                            properties.Add(prop.Name.Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not read custom properties: {ex.Message}");
            }

            return properties;
        }


        public Dictionary<string, string> GetCustomPropertiesWithValues(string templatePath)
        {
            var properties = new Dictionary<string, string>();
            var documentType = GetDocumentType(templatePath);

            try
            {
                switch (documentType)
                {
                    case Models.DocumentType.Word:
                        using (var doc = WordprocessingDocument.Open(templatePath, false))
                        {
                            properties = GetCustomPropertiesFromOpenXml(doc);
                        }
                        break;

                    case Models.DocumentType.Excel:
                        using (var doc = SpreadsheetDocument.Open(templatePath, false))
                        {
                            properties = GetCustomPropertiesFromOpenXml(doc);
                        }
                        break;

                    case Models.DocumentType.PowerPoint:
                        using (var doc = PresentationDocument.Open(templatePath, false))
                        {
                            properties = GetCustomPropertiesFromOpenXml(doc);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not read custom properties: {ex.Message}");
            }

            return properties;
        }

        private Dictionary<string, string> GetCustomPropertiesFromOpenXml(OpenXmlPackage document)
        {
            var properties = new Dictionary<string, string>();

            try
            {
                CustomFilePropertiesPart? customPropertiesPart = null;
                
                if (document is WordprocessingDocument wordDoc)
                {
                    customPropertiesPart = wordDoc.CustomFilePropertiesPart;
                }
                else if (document is SpreadsheetDocument excelDoc)
                {
                    customPropertiesPart = excelDoc.CustomFilePropertiesPart;
                }
                else if (document is PresentationDocument pptDoc)
                {
                    customPropertiesPart = pptDoc.CustomFilePropertiesPart;
                }

                if (customPropertiesPart?.Properties != null)
                {
                    foreach (var prop in customPropertiesPart.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
                    {
                        if (prop.Name?.Value != null)
                        {
                            string name = prop.Name.Value;
                            string value = GetPropertyValue(prop);
                            properties[name] = value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not read custom properties: {ex.Message}");
            }

            return properties;
        }

        private string GetPropertyValue(DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty prop)
        {
            // Custom properties can have different types - let's get the first child element's value
            try
            {
                if (prop.VTLPWSTR != null) return prop.VTLPWSTR.Text ?? "";
                if (prop.VTFileTime != null) return prop.VTFileTime.Text ?? "";
                if (prop.VTBool != null) return prop.VTBool.Text ?? "";
                if (prop.VTInt32 != null) return prop.VTInt32.Text ?? "";

                // Generic approach - get first child's inner text
                var firstChild = prop.FirstChild;
                if (firstChild != null)
                {
                    return firstChild.InnerText ?? "";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not read property value: {ex.Message}");
            }
            
            return "";
        }

        private void LoadTemplates()
        {
            if (File.Exists(_metadataFile))
            {
                try
                {
                    var json = File.ReadAllText(_metadataFile);
                    _templates = JsonSerializer.Deserialize<List<DocumentTemplate>>(json) ?? new List<DocumentTemplate>();
                }
                catch
                {
                    _templates = new List<DocumentTemplate>();
                }
            }
            else
            {
                _templates = new List<DocumentTemplate>();
            }
        }

        private void SaveTemplates()
        {
            var json = JsonSerializer.Serialize(_templates, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_metadataFile, json);
        }
    }
}
