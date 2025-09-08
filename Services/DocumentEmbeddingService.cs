using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentAutomationDemo.Models;
using System.Text.RegularExpressions;
using Models = DocumentAutomationDemo.Models;

namespace DocumentAutomationDemo.Services
{
    public interface IDocumentEmbeddingService
    {
        string GenerateDocumentWithEmbedding(DocumentEmbeddingRequest request);
    }

    public class DocumentEmbeddingService : IDocumentEmbeddingService
    {
        private readonly ITemplateService _templateService;
        private readonly IDocumentGenerationService _documentService;
        private readonly string _outputDirectory;

        public DocumentEmbeddingService(ITemplateService templateService)
        {
            _templateService = templateService;
            _documentService = new DocumentGenerationService(templateService);
            _outputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            
            if (!Directory.Exists(_outputDirectory))
            {
                Directory.CreateDirectory(_outputDirectory);
            }
        }

        public string GenerateDocumentWithEmbedding(DocumentEmbeddingRequest request)
        {
            var mainTemplate = _templateService.GetTemplate(request.MainTemplateId);
            if (mainTemplate == null)
                throw new ArgumentException($"Main template with ID '{request.MainTemplateId}' not found");

            if (mainTemplate.DocumentType != Models.DocumentType.Word)
                throw new ArgumentException("Document embedding is only supported for Word documents");

            Console.WriteLine($"🔄 Processing document with multiple embeddings...");
            Console.WriteLine($"📋 Main template: {mainTemplate.Name}");
            Console.WriteLine($"📍 Number of embeddings: {request.Embeddings.Count}");

            // Create main document with all placeholders filled except embed placeholders
            string workingDocPath = CreateMainDocumentForEmbedding(mainTemplate, request);

            // Process each embedding
            foreach (var embedding in request.Embeddings)
            {
                ProcessSingleEmbedding(workingDocPath, embedding);
            }

            // Export to final format if needed
            string finalPath = ConvertToFinalFormat(workingDocPath, request.ExportFormat, mainTemplate.Name);

            return finalPath;
        }

        private string CreateMainDocumentForEmbedding(DocumentTemplate mainTemplate, DocumentEmbeddingRequest request)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string workingDocPath = Path.Combine(_outputDirectory, $"working_main_{timestamp}.docx");
            
            // Copy main template to working location
            File.Copy(mainTemplate.FilePath, workingDocPath, true);
            Console.WriteLine($"📝 Created working main document: {Path.GetFileName(workingDocPath)}");

            // Get embed placeholders from all embeddings
            var embedPlaceholders = request.Embeddings.Select(e => e.EmbedPlaceholder).ToHashSet();
            
            // Replace only non-embed placeholders
            var nonEmbedValues = request.MainTemplateValues
                .Where(v => !embedPlaceholders.Contains(v.Placeholder))
                .ToList();

            if (nonEmbedValues.Any())
            {
                ReplacePlaceholdersInDocument(workingDocPath, nonEmbedValues);
                Console.WriteLine($"✅ Replaced {nonEmbedValues.Count} main template placeholders");
            }

            return workingDocPath;
        }

        private void ProcessSingleEmbedding(string mainDocPath, EmbedInfo embedding)
        {
            var embedTemplate = _templateService.GetTemplate(embedding.EmbedTemplateId);
            if (embedTemplate == null)
            {
                Console.WriteLine($"❌ Embed template with ID '{embedding.EmbedTemplateId}' not found. Skipping.");
                return;
            }

            if (embedTemplate.DocumentType != Models.DocumentType.Word)
            {
                Console.WriteLine($"⚠️ Embed template '{embedTemplate.Name}' is not a Word document. Skipping.");
                return;
            }

            Console.WriteLine($"🔗 Processing embedding: {embedTemplate.Name} → {embedding.EmbedPlaceholder}");

            // Create processed embed document
            string embedDocPath = CreateProcessedEmbedDocument(embedTemplate, embedding.EmbedTemplateValues);

            try
            {
                // Embed the document at the specified placeholder
                EmbedWordDocument(mainDocPath, embedDocPath, embedding.EmbedPlaceholder);
                Console.WriteLine($"✅ Successfully embedded '{embedTemplate.Name}' at '{embedding.EmbedPlaceholder}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Failed to embed '{embedTemplate.Name}': {ex.Message}");
            }
            finally
            {
                // Clean up temporary embed document
                CleanupTempFile(embedDocPath, $"embed document '{embedTemplate.Name}'");
            }
        }

        private string CreateProcessedEmbedDocument(DocumentTemplate embedTemplate, List<PlaceholderValue> embedValues)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string random = Guid.NewGuid().ToString("N")[..8];
            string tempEmbedPath = Path.Combine(_outputDirectory, $"embed_temp_{timestamp}_{random}.docx");
            
            // Copy and process the embed template
            File.Copy(embedTemplate.FilePath, tempEmbedPath, true);
            
            // Replace placeholders in embed document
            if (embedValues.Any())
            {
                ReplacePlaceholdersInDocument(tempEmbedPath, embedValues);
                Console.WriteLine($"   ✓ Processed {embedValues.Count} placeholders in embed template");
            }
            
            return tempEmbedPath;
        }

        private void EmbedWordDocument(string mainDocPath, string embedDocPath, string embedPlaceholder)
        {
            using (var mainDoc = WordprocessingDocument.Open(mainDocPath, true))
            using (var embedDoc = WordprocessingDocument.Open(embedDocPath, false))
            {
                if (mainDoc.MainDocumentPart?.Document?.Body == null || 
                    embedDoc.MainDocumentPart?.Document?.Body == null)
                {
                    throw new InvalidOperationException("Document structure is invalid");
                }

                // Find placeholder in main document
                var placeholderElement = FindPlaceholderInWordDocument(mainDoc, embedPlaceholder);
                if (placeholderElement == null)
                {
                    Console.WriteLine($"   ⚠️ Placeholder '{embedPlaceholder}' not found in main document");
                    return;
                }

                Console.WriteLine($"   📍 Found placeholder '{embedPlaceholder}' in main document");

                // Import styles from embed document to avoid style conflicts
                ImportStyles(mainDoc, embedDoc);

                // Import numbering definitions if present
                ImportNumberingDefinitions(mainDoc, embedDoc);

                // Import images and media
                ImportImages(mainDoc, embedDoc);

                // Clone and import all content from embed document
                var sourceBody = embedDoc.MainDocumentPart.Document.Body;
                var importedElements = new List<OpenXmlElement>();

                foreach (var element in sourceBody.Elements())
                {
                    var clonedElement = element.CloneNode(true);
                    importedElements.Add(clonedElement);
                }

                // Insert imported content at placeholder location
                var parentElement = placeholderElement.Parent;
                if (parentElement != null)
                {
                    // Insert all imported elements before the placeholder
                    foreach (var element in importedElements)
                    {
                        parentElement.InsertBefore(element, placeholderElement);
                    }
                    
                    // Remove the placeholder
                    placeholderElement.Remove();
                    Console.WriteLine($"   ✓ Inserted {importedElements.Count} elements from embed document");
                }

                mainDoc.MainDocumentPart.Document.Save();
            }
        }

        private OpenXmlElement? FindPlaceholderInWordDocument(WordprocessingDocument doc, string placeholder)
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return null;

            // Look for placeholder in text content
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                var fullText = string.Join("", paragraph.Descendants<Text>().Select(t => t.Text));
                if (fullText.Contains(placeholder))
                {
                    return paragraph;
                }
            }

            // Look for placeholder in custom property fields
            foreach (var field in body.Descendants<SimpleField>())
            {
                var instruction = field.Instruction?.Value;
                if (!string.IsNullOrEmpty(instruction) && instruction.Contains(placeholder))
                {
                    return field.Parent;
                }
            }

            // Look for placeholder in complex fields (field codes)
            var fieldCodes = body.Descendants<FieldCode>();
            foreach (var fieldCode in fieldCodes)
            {
                if (fieldCode.Text != null && fieldCode.Text.Contains(placeholder))
                {
                    // Find the parent element that contains the entire field
                    var parent = fieldCode.Ancestors<Paragraph>().FirstOrDefault();
                    if (parent != null) return parent;
                }
            }

            return null;
        }

        private void ImportStyles(WordprocessingDocument mainDoc, WordprocessingDocument embedDoc)
        {
            try
            {
                if (embedDoc.MainDocumentPart?.StyleDefinitionsPart?.Styles == null) return;

                var mainStyles = mainDoc.MainDocumentPart?.StyleDefinitionsPart;
                if (mainStyles == null)
                {
                    mainStyles = mainDoc.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();
                    if (mainStyles != null)
                        mainStyles.Styles = new Styles();
                }

                if (mainStyles?.Styles == null) return;

                var embedStyles = embedDoc.MainDocumentPart.StyleDefinitionsPart.Styles;
                var existingStyleIds = mainStyles.Styles.Elements<Style>().Select(s => s.StyleId?.Value).ToHashSet();

                int importedCount = 0;
                foreach (var style in embedStyles.Elements<Style>())
                {
                    var styleId = style.StyleId?.Value;
                    if (!string.IsNullOrEmpty(styleId) && !existingStyleIds.Contains(styleId))
                    {
                        var clonedStyle = style.CloneNode(true) as Style;
                        if (clonedStyle != null)
                        {
                            mainStyles.Styles.AppendChild(clonedStyle);
                            existingStyleIds.Add(styleId);
                            importedCount++;
                        }
                    }
                }

                if (importedCount > 0)
                {
                    Console.WriteLine($"   ✓ Imported {importedCount} styles from embed document");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Could not import styles: {ex.Message}");
            }
        }

        private void ImportNumberingDefinitions(WordprocessingDocument mainDoc, WordprocessingDocument embedDoc)
        {
            try
            {
                if (embedDoc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering == null) return;

                var mainNumbering = mainDoc.MainDocumentPart?.NumberingDefinitionsPart;
                if (mainNumbering == null)
                {
                    mainNumbering = mainDoc.MainDocumentPart?.AddNewPart<NumberingDefinitionsPart>();
                    if (mainNumbering != null)
                        mainNumbering.Numbering = new Numbering();
                }

                if (mainNumbering?.Numbering == null) return;

                var embedNumbering = embedDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                int importedCount = 0;
                
                // Import abstract numbering definitions
                foreach (var abstractNum in embedNumbering.Elements<AbstractNum>())
                {
                    var cloned = abstractNum.CloneNode(true) as AbstractNum;
                    if (cloned != null)
                    {
                        mainNumbering.Numbering.AppendChild(cloned);
                        importedCount++;
                    }
                }

                // Import numbering instances
                foreach (var numInstance in embedNumbering.Elements<NumberingInstance>())
                {
                    var cloned = numInstance.CloneNode(true) as NumberingInstance;
                    if (cloned != null)
                    {
                        mainNumbering.Numbering.AppendChild(cloned);
                        importedCount++;
                    }
                }

                if (importedCount > 0)
                {
                    Console.WriteLine($"   ✓ Imported {importedCount} numbering definitions from embed document");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Could not import numbering definitions: {ex.Message}");
            }
        }

        private void ImportImages(WordprocessingDocument mainDoc, WordprocessingDocument embedDoc)
        {
            try
            {
                if (embedDoc.MainDocumentPart == null) return;

                var embedImageParts = embedDoc.MainDocumentPart.ImageParts.ToList();
                if (!embedImageParts.Any()) return;

                int importedCount = 0;
                foreach (var embedImagePart in embedImageParts)
                {
                    // Create new image part in main document
                    var newImagePart = mainDoc.MainDocumentPart?.AddImagePart(embedImagePart.ContentType);
                    if (newImagePart != null)
                    {
                        // Copy image data
                        using (var stream = embedImagePart.GetStream())
                        {
                            stream.CopyTo(newImagePart.GetStream(FileMode.Create));
                        }
                        importedCount++;
                    }
                }

                if (importedCount > 0)
                {
                    Console.WriteLine($"   ✓ Imported {importedCount} images from embed document");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Could not import images: {ex.Message}");
            }
        }

        private void ReplacePlaceholdersInDocument(string filePath, List<PlaceholderValue> placeholderValues)
        {
            try
            {
                using (var doc = WordprocessingDocument.Open(filePath, true))
                {
                    // Update custom properties
                    UpdateCustomPropertiesInDocument(doc, placeholderValues);
                    
                    // Refresh document property fields
                    RefreshWordDocPropertyFields(doc, placeholderValues);
                    
                    doc.MainDocumentPart?.Document?.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Warning: Could not update document properties: {ex.Message}");
            }
        }

        private string ConvertToFinalFormat(string workingDocPath, ExportFormat exportFormat, string templateName)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string finalFileName;
            
            switch (exportFormat)
            {
                case ExportFormat.Html:
                    finalFileName = $"{templateName}_with_embeddings_{timestamp}.html";
                    break;
                case ExportFormat.HtmlEmail:
                    finalFileName = $"{templateName}_with_embeddings_{timestamp}_email.html";
                    break;
                case ExportFormat.Pdf:
                    finalFileName = $"{templateName}_with_embeddings_{timestamp}.pdf";
                    break;
                case ExportFormat.Word:
                case ExportFormat.Original:
                default:
                    finalFileName = $"{templateName}_with_embeddings_{timestamp}.docx";
                    break;
            }
            
            string finalPath = Path.Combine(_outputDirectory, finalFileName);

            try
            {
                switch (exportFormat)
                {
                    case ExportFormat.Html:
                        return _documentService.ExportToHtml(workingDocPath, Models.DocumentType.Word);
                    case ExportFormat.HtmlEmail:
                        return _documentService.ExportToEmailHtml(workingDocPath, Models.DocumentType.Word);
                    case ExportFormat.Pdf:
                        return _documentService.ExportToPdf(workingDocPath, Models.DocumentType.Word);
                    case ExportFormat.Word:
                    case ExportFormat.Original:
                    default:
                        // Simply copy to final location with proper name
                        File.Copy(workingDocPath, finalPath, true);
                        CleanupTempFile(workingDocPath, "working document");
                        return finalPath;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Error converting to {exportFormat}: {ex.Message}");
                // Fall back to Word format
                File.Copy(workingDocPath, finalPath, true);
                CleanupTempFile(workingDocPath, "working document");
                return finalPath;
            }
        }

        private void UpdateCustomPropertiesInDocument(OpenXmlPackage package, List<PlaceholderValue> placeholderValues)
        {
            try
            {
                // Cast to WordprocessingDocument to access CustomFilePropertiesPart
                if (package is not WordprocessingDocument wordDoc) return;
                
                var customPropsPart = wordDoc.CustomFilePropertiesPart;
                if (customPropsPart?.Properties == null) return;

                foreach (var placeholderValue in placeholderValues)
                {
                    var existingProp = customPropsPart.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>()
                        .FirstOrDefault(p => p.Name?.Value == placeholderValue.Placeholder);

                    if (existingProp != null)
                    {
                        // Update existing property - use the same pattern as TemplateService
                        if (existingProp.VTLPWSTR != null)
                            existingProp.VTLPWSTR.Text = placeholderValue.Value;
                        else if (existingProp.VTFileTime != null)
                            existingProp.VTFileTime.Text = placeholderValue.Value;
                        else if (existingProp.VTBool != null)
                            existingProp.VTBool.Text = placeholderValue.Value;
                        else if (existingProp.VTInt32 != null)
                            existingProp.VTInt32.Text = placeholderValue.Value;
                    }
                }

                customPropsPart.Properties.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Warning updating custom properties: {ex.Message}");
            }
        }

        private void RefreshWordDocPropertyFields(WordprocessingDocument doc, List<PlaceholderValue> placeholderValues)
        {
            try
            {
                var body = doc.MainDocumentPart?.Document?.Body;
                if (body == null) return;

                // Update simple fields (DOCPROPERTY fields)
                foreach (var field in body.Descendants<SimpleField>())
                {
                    var instruction = field.Instruction?.Value;
                    if (string.IsNullOrEmpty(instruction)) continue;

                    var match = Regex.Match(instruction, @"DOCPROPERTY\s+""?([^""}\s]+)""?", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        var propertyName = match.Groups[1].Value;
                        var placeholderValue = placeholderValues.FirstOrDefault(pv => pv.Placeholder == propertyName);
                        if (placeholderValue != null)
                        {
                            var textElement = field.Descendants<Text>().FirstOrDefault();
                            if (textElement != null)
                            {
                                textElement.Text = placeholderValue.Value;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Warning refreshing document fields: {ex.Message}");
            }
        }

        private void CleanupTempFile(string filePath, string description)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Console.WriteLine($"🧹 Cleaned up {description}: {Path.GetFileName(filePath)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Could not clean up {description} {Path.GetFileName(filePath)}: {ex.Message}");
            }
        }
    }
}
