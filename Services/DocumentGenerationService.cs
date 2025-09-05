using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using DocumentAutomationDemo.Models;
using DocumentAutomationDemo.Services;
using System.Text;
using System.Diagnostics;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;

namespace DocumentAutomationDemo.Services
{
    public interface IDocumentGenerationService
    {
        string GenerateDocument(DocumentGenerationRequest request);
        string ExportToHtml(string templatePath, Models.DocumentType documentType);
        string ExportToEmailHtml(string templatePath, Models.DocumentType documentType);
        string ExportToPdf(string templatePath, Models.DocumentType documentType);
    }

    public class DocumentGenerationService : IDocumentGenerationService
    {
        private readonly ITemplateService _templateService;
        private readonly string _outputDirectory;
        private readonly string _libreOfficePath;

        public DocumentGenerationService(ITemplateService templateService)
        {
            _templateService = templateService;
            _outputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            
            // Try to find LibreOffice installation
            _libreOfficePath = FindLibreOfficePath();
            
            if (!Directory.Exists(_outputDirectory))
            {
                Directory.CreateDirectory(_outputDirectory);
            }
        }

        private string FindLibreOfficePath()
        {
            // Common LibreOffice installation paths on Windows
            string[] possiblePaths = {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                @"C:\Users\" + Environment.UserName + @"\AppData\Local\Programs\LibreOffice\program\soffice.exe"
            };

            foreach (string path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    Console.WriteLine($"✅ Found LibreOffice at: {path}");
                    return path;
                }
            }

            Console.WriteLine("⚠️ LibreOffice not found. PDF and HTML export will use fallback methods.");
            return string.Empty;
        }

        public string GenerateDocument(DocumentGenerationRequest request)
        {
            var template = _templateService.GetTemplate(request.TemplateId);
            if (template == null)
                throw new ArgumentException($"Template with ID '{request.TemplateId}' not found");

            // Create output filename based on requested format
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputFileName;
            string outputPath;
            
            // Determine output file extension based on export format
            switch (request.ExportFormat)
            {
                case ExportFormat.Html:
                    outputFileName = $"{template.Name}_{timestamp}.html";
                    outputPath = Path.Combine(_outputDirectory, outputFileName);
                    return ExportDirectToHtml(template, request.PlaceholderValues, outputPath);
                    
                case ExportFormat.HtmlEmail:
                    outputFileName = $"{template.Name}_{timestamp}_email.html";
                    outputPath = Path.Combine(_outputDirectory, outputFileName);
                    return ExportDirectToEmailHtml(template, request.PlaceholderValues, outputPath);
                    
                case ExportFormat.Pdf:
                    outputFileName = $"{template.Name}_{timestamp}.pdf";
                    outputPath = Path.Combine(_outputDirectory, outputFileName);
                    return ExportDirectToPdf(template, request.PlaceholderValues, outputPath);
                    
                case ExportFormat.Word:
                    outputFileName = $"{template.Name}_{timestamp}.docx";
                    break;
                    
                default: // Original format
                    string fileExtension = Path.GetExtension(template.FilePath);
                    outputFileName = $"{template.Name}_{timestamp}{fileExtension}";
                    break;
            }
            
            // For Word/Original format, create the document file
            outputPath = Path.Combine(_outputDirectory, outputFileName);
            File.Copy(template.FilePath, outputPath, true);
            ReplacePlaceholders(outputPath, request.PlaceholderValues, template.DocumentType);
            
            return outputPath;
        }

        private string ExportDirectToHtml(DocumentTemplate template, List<PlaceholderValue> placeholderValues, string outputPath)
        {
            // Create temporary document for processing
            string tempDocPath = CreateTempDocument(template, placeholderValues);
            
            try
            {
                return ConvertToHtml(tempDocPath, outputPath, template.DocumentType);
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempDocPath))
                {
                    File.Delete(tempDocPath);
                }
            }
        }

        private string ExportDirectToEmailHtml(DocumentTemplate template, List<PlaceholderValue> placeholderValues, string outputPath)
        {
            // Create temporary document for processing
            string tempDocPath = CreateTempDocument(template, placeholderValues);
            
            try
            {
                if (!string.IsNullOrEmpty(_libreOfficePath))
                {
                    Console.WriteLine("🔄 Converting directly to email-friendly HTML using LibreOffice...");
                    
                    // Convert directly using LibreOffice
                    string regularHtmlPath = ConvertUsingLibreOffice(tempDocPath, "html");
                    return ConvertHtmlToEmailFriendly(regularHtmlPath, outputPath);
                }
                else
                {
                    // Fallback method
                    return ConvertToEmailHtmlFallback(tempDocPath, outputPath, template.DocumentType);
                }
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempDocPath))
                {
                    File.Delete(tempDocPath);
                }
            }
        }

        private string ExportDirectToPdf(DocumentTemplate template, List<PlaceholderValue> placeholderValues, string outputPath)
        {
            // Create temporary document for processing
            string tempDocPath = CreateTempDocument(template, placeholderValues);
            
            try
            {
                if (!string.IsNullOrEmpty(_libreOfficePath))
                {
                    Console.WriteLine("🔄 Converting directly to PDF using LibreOffice...");
                    return ConvertUsingLibreOffice(tempDocPath, "pdf");
                }
                else
                {
                    // Fallback: HTML then PDF
                    string htmlPath = ConvertToHtml(tempDocPath, Path.ChangeExtension(outputPath, ".html"), template.DocumentType);
                    ConvertHtmlToPdfFallback(htmlPath, outputPath);
                    return outputPath;
                }
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempDocPath))
                {
                    File.Delete(tempDocPath);
                }
            }
        }

        private string CreateTempDocument(DocumentTemplate template, List<PlaceholderValue> placeholderValues)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}{Path.GetExtension(template.FilePath)}");
            File.Copy(template.FilePath, tempPath, true);
            ReplacePlaceholders(tempPath, placeholderValues, template.DocumentType);
            return tempPath;
        }

        private string ConvertToHtml(string inputPath, string outputPath, Models.DocumentType documentType)
        {
            if (!string.IsNullOrEmpty(_libreOfficePath))
            {
                try
                {
                    Console.WriteLine("🔄 Converting to HTML using LibreOffice headless...");
                    return ConvertUsingLibreOffice(inputPath, "html");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ LibreOffice conversion failed: {ex.Message}");
                    Console.WriteLine("📄 Falling back to basic HTML conversion...");
                }
            }
            
            // Fallback to basic conversion
            switch (documentType)
            {
                case Models.DocumentType.Word:
                    ExportWordToHtmlFallback(inputPath, outputPath);
                    break;
                case Models.DocumentType.Excel:
                    ExportExcelToHtmlFallback(inputPath, outputPath);
                    break;
                case Models.DocumentType.PowerPoint:
                    ExportPowerPointToHtmlFallback(inputPath, outputPath);
                    break;
            }
            return outputPath;
        }

        private string ConvertToEmailHtmlFallback(string inputPath, string outputPath, Models.DocumentType documentType)
        {
            string regularHtmlPath = ConvertToHtml(inputPath, Path.ChangeExtension(outputPath, ".temp.html"), documentType);
            return ConvertHtmlToEmailFriendly(regularHtmlPath, outputPath);
        }

        private void ReplacePlaceholders(string filePath, List<PlaceholderValue> placeholderValues, Models.DocumentType documentType)
        {
            try
            {
                switch (documentType)
                {
                    case Models.DocumentType.Word:
                        using (var doc = WordprocessingDocument.Open(filePath, true))
                        {
                            UpdateCustomPropertiesInDocument(doc, placeholderValues, documentType);
                            // NEW: Automatically refresh DOCPROPERTY field results after updating properties
                            RefreshWordDocPropertyFields(doc, placeholderValues);
                            doc.MainDocumentPart?.Document?.Save();
                        }
                        break;

                    case Models.DocumentType.Excel:
                        using (var doc = SpreadsheetDocument.Open(filePath, true))
                        {
                            UpdateCustomPropertiesInDocument(doc, placeholderValues, documentType);
                            doc.WorkbookPart?.Workbook?.Save();
                        }
                        break;

                    case Models.DocumentType.PowerPoint:
                        using (var doc = PresentationDocument.Open(filePath, true))
                        {
                            UpdateCustomPropertiesInDocument(doc, placeholderValues, documentType);
                            doc.PresentationPart?.Presentation?.Save();
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not update document properties: {ex.Message}");
            }
        }

                private void UpdateCustomPropertiesInDocument(OpenXmlPackage package, List<PlaceholderValue> placeholderValues, Models.DocumentType documentType)
        {
            try
            {
                CustomFilePropertiesPart? customPropertiesPart = null;

                // Get the custom properties part based on document type
                switch (documentType)
                {
                    case Models.DocumentType.Word:
                        var wordDoc = package as WordprocessingDocument;
                        if (wordDoc?.CustomFilePropertiesPart == null)
                        {
                            wordDoc?.AddCustomFilePropertiesPart();
                        }
                        customPropertiesPart = wordDoc?.CustomFilePropertiesPart;
                        break;
                    case Models.DocumentType.Excel:
                        var excelDoc = package as SpreadsheetDocument;
                        if (excelDoc?.CustomFilePropertiesPart == null)
                        {
                            excelDoc?.AddCustomFilePropertiesPart();
                        }
                        customPropertiesPart = excelDoc?.CustomFilePropertiesPart;
                        break;
                    case Models.DocumentType.PowerPoint:
                        var pptDoc = package as PresentationDocument;
                        if (pptDoc?.CustomFilePropertiesPart == null)
                        {
                            pptDoc?.AddCustomFilePropertiesPart();
                        }
                        customPropertiesPart = pptDoc?.CustomFilePropertiesPart;
                        break;
                }

                if (customPropertiesPart == null) return;
                if (customPropertiesPart.Properties == null)
                {
                    customPropertiesPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();
                }

                var properties = customPropertiesPart.Properties;

                foreach (var placeholderValue in placeholderValues)
                {
                    // Find existing property
                    var existingProp = properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>()
                        .FirstOrDefault(p => p.Name?.Value == placeholderValue.Placeholder);

                    if (existingProp != null)
                    {
                        // Update existing property
                        UpdatePropertyValue(existingProp, placeholderValue.Value);
                    }
                    else
                    {
                        // Create new property
                        CreateNewCustomProperty(properties, placeholderValue.Placeholder, placeholderValue.Value);
                    }
                }

                customPropertiesPart.Properties.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not update custom properties: {ex.Message}");
            }
        }

        private void UpdatePropertyValue(DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty prop, string newValue)
        {
            // Remove existing value elements
            prop.RemoveAllChildren();

            // Add new value as string
            prop.AppendChild(new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(newValue));
        }

        private void CreateNewCustomProperty(DocumentFormat.OpenXml.CustomProperties.Properties properties, string name, string value)
        {
            // Get next property ID
            int pid = 2; // Start from 2 (1 is reserved)
            var existingProps = properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>();
            if (existingProps.Any())
            {
                pid = existingProps.Max(p => p.PropertyId?.Value ?? 1) + 1;
            }

            var newProp = new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty()
            {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = pid,
                Name = name
            };

            newProp.AppendChild(new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(value));
            properties.AppendChild(newProp);
        }

        private void RefreshWordDocPropertyFields(WordprocessingDocument doc, List<PlaceholderValue> placeholderValues)
        {
            try
            {
                // Build property name -> value map for quick lookup
                var propertyValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                Console.WriteLine("📋 Starting automatic field refresh for DOCPROPERTY fields...");
                foreach (var pv in placeholderValues)
                {
                    propertyValues[pv.Placeholder] = pv.Value;
                }

                // Get all containers that might have fields (body, headers, footers)
                var containers = new List<OpenXmlElement>();
                
                if (doc.MainDocumentPart?.Document?.Body != null)
                    containers.Add(doc.MainDocumentPart.Document.Body);
                
                if (doc.MainDocumentPart != null)
                {
                    foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                        if (headerPart.Header != null) containers.Add(headerPart.Header);
                    
                    foreach (var footerPart in doc.MainDocumentPart.FooterParts)
                        if (footerPart.Footer != null) containers.Add(footerPart.Footer);
                }

                int totalFieldsUpdated = 0;
                var updatedFields = new Dictionary<string, int>();

                // Process fields in each container
                foreach (var container in containers)
                {
                    var simpleFieldsUpdated = UpdateSimpleFields(container, propertyValues);
                    var complexFieldsUpdated = UpdateComplexFields(container, propertyValues);
                    
                    // Count updated fields by property name
                    foreach (var kvp in simpleFieldsUpdated)
                    {
                        if (!updatedFields.ContainsKey(kvp.Key))
                            updatedFields[kvp.Key] = 0;
                        updatedFields[kvp.Key] += kvp.Value;
                        totalFieldsUpdated += kvp.Value;
                    }
                    
                    foreach (var kvp in complexFieldsUpdated)
                    {
                        if (!updatedFields.ContainsKey(kvp.Key))
                            updatedFields[kvp.Key] = 0;
                        updatedFields[kvp.Key] += kvp.Value;
                        totalFieldsUpdated += kvp.Value;
                    }
                }

                // CRITICAL: Force Word to update fields when document opens
                ForceFieldUpdateOnOpen(doc);
                
                // ADDITIONAL: Force immediate field calculation by clearing field caches
                ClearFieldCaches(doc);

                Console.WriteLine($"✅ Field refresh completed! Updated {totalFieldsUpdated} total field instances:");
                foreach (var kvp in updatedFields.OrderBy(x => x.Key))
                {
                    Console.WriteLine($"   • {kvp.Key}: {kvp.Value} instance(s) updated");
                }
                Console.WriteLine($"   + Added automatic field update trigger for when document opens in Word");
                Console.WriteLine($"   + Cleared field caches to force immediate recalculation");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not refresh DOCPROPERTY fields: {ex.Message}");
                Console.WriteLine($"DEBUG: Exception details: {ex.StackTrace}");
            }
        }

        // Force Word to update all fields when the document is opened
        private void ForceFieldUpdateOnOpen(WordprocessingDocument doc)
        {
            try
            {
                // Add document settings to update fields on open
                var settingsPart = doc.MainDocumentPart?.DocumentSettingsPart;
                if (settingsPart == null)
                {
                    settingsPart = doc.MainDocumentPart?.AddNewPart<DocumentSettingsPart>();
                    if (settingsPart != null)
                    {
                        settingsPart.Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
                    }
                }

                var settings = settingsPart?.Settings;
                if (settings == null) return;
                
                // Remove existing UpdateFieldsOnOpen if present
                var existingUpdateFields = settings.Elements<DocumentFormat.OpenXml.Wordprocessing.UpdateFieldsOnOpen>().ToList();
                foreach (var element in existingUpdateFields)
                {
                    element.Remove();
                }

                // Add UpdateFieldsOnOpen setting
                settings.Append(new DocumentFormat.OpenXml.Wordprocessing.UpdateFieldsOnOpen() { Val = true });
                
                Console.WriteLine("   ✓ Set document to auto-update fields when opened in Word");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Could not set auto-update fields setting: {ex.Message}");
            }
        }

        // Clear field caches to force immediate recalculation
        private void ClearFieldCaches(WordprocessingDocument doc)
        {
            try
            {
                // Get all containers that might have fields
                var containers = new List<OpenXmlElement>();
                
                if (doc.MainDocumentPart?.Document?.Body != null)
                    containers.Add(doc.MainDocumentPart.Document.Body);
                
                if (doc.MainDocumentPart != null)
                {
                    foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                        if (headerPart.Header != null) containers.Add(headerPart.Header);
                    
                    foreach (var footerPart in doc.MainDocumentPart.FooterParts)
                        if (footerPart.Footer != null) containers.Add(footerPart.Footer);
                }

                int clearedCaches = 0;
                foreach (var container in containers)
                {
                    // Mark all simple fields as dirty (need recalculation)
                    var simpleFields = container.Descendants<SimpleField>().ToList();
                    foreach (var field in simpleFields)
                    {
                        // Remove any cached field results by clearing fldCharType="separate" content
                        var instruction = field.Instruction?.Value;
                        if (!string.IsNullOrEmpty(instruction) && instruction.Contains("DOCPROPERTY"))
                        {
                            // Force field to be marked as needing update
                            field.Dirty = true;
                            clearedCaches++;
                        }
                    }

                    // Mark complex fields as dirty
                    var fieldChars = container.Descendants<FieldChar>().ToList();
                    foreach (var fieldChar in fieldChars)
                    {
                        if (fieldChar.FieldCharType?.Value == FieldCharValues.Begin)
                        {
                            fieldChar.Dirty = true;
                            clearedCaches++;
                        }
                    }
                }

                Console.WriteLine($"   ✓ Cleared {clearedCaches} field caches to force recalculation");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️  Could not clear field caches: {ex.Message}");
            }
        }

                // Handle simple field format: <w:fldSimple w:instr="DOCPROPERTY PropertyName">
        private Dictionary<string, int> UpdateSimpleFields(OpenXmlElement container, Dictionary<string, string> propertyValues)
        {
            var updatedFields = new Dictionary<string, int>();
            var simpleFields = container.Descendants<SimpleField>().ToList();
            
            foreach (var field in simpleFields)
            {
                var instruction = field.Instruction?.Value;
                if (string.IsNullOrEmpty(instruction)) continue;

                var match = System.Text.RegularExpressions.Regex.Match(instruction, @"\bDOCPROPERTY\s+(\S+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (!match.Success) continue;

                string propertyName = match.Groups[1].Value.Trim();
                if (propertyValues.TryGetValue(propertyName, out var newValue))
                {
                    var textElements = field.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                    foreach (var textElement in textElements)
                    {
                        textElement.Text = newValue;
                    }

                    if (!updatedFields.ContainsKey(propertyName))
                        updatedFields[propertyName] = 0;
                    updatedFields[propertyName]++;
                }
            }
            
            return updatedFields;
        }
        // Handle complex field format: Begin -> Instruction -> Separate -> Result -> End
        private Dictionary<string, int> UpdateComplexFields(OpenXmlElement container, Dictionary<string, string> propertyValues)
        {
            var updatedFields = new Dictionary<string, int>();
            var runs = container.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().ToList();
            
            for (int i = 0; i < runs.Count; i++)
            {
                // Look for field begin
                var fieldBegin = runs[i].Elements<FieldChar>()
                    .FirstOrDefault(fc => fc.FieldCharType?.Value == FieldCharValues.Begin);
                
                if (fieldBegin == null) continue;

                // Collect all instruction text until we find separate or end
                var instructionText = new StringBuilder();
                int instrStartIndex = i + 1;
                int separateIndex = -1;
                int endIndex = -1;

                for (int j = instrStartIndex; j < runs.Count; j++)
                {
                    var fieldChar = runs[j].Elements<FieldChar>().FirstOrDefault();
                    if (fieldChar?.FieldCharType?.Value == FieldCharValues.Separate)
                    {
                        separateIndex = j;
                        break;
                    }
                    else if (fieldChar?.FieldCharType?.Value == FieldCharValues.End)
                    {
                        endIndex = j;
                        break;
                    }
                    else
                    {
                        // Try to get instruction text from InnerText
                        if (!string.IsNullOrEmpty(runs[j].InnerText))
                        {
                            instructionText.Append(runs[j].InnerText);
                        }
                    }
                }

                // Find the end if we only found separate
                if (separateIndex != -1 && endIndex == -1)
                {
                    for (int j = separateIndex + 1; j < runs.Count; j++)
                    {
                        var fieldChar = runs[j].Elements<FieldChar>().FirstOrDefault();
                        if (fieldChar?.FieldCharType?.Value == FieldCharValues.End)
                        {
                            endIndex = j;
                            break;
                        }
                    }
                }

                if (endIndex == -1) continue; // Malformed field

                string instruction = instructionText.ToString().Trim();
                var match = System.Text.RegularExpressions.Regex.Match(instruction, @"\bDOCPROPERTY\s+(\S+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (!match.Success) continue;

                string propertyName = match.Groups[1].Value.Trim();
                if (propertyValues.TryGetValue(propertyName, out var newValue))
                {
                    // Update the result text (between separate and end, or between last instruction and end)
                    int resultStartIndex = separateIndex != -1 ? separateIndex + 1 : instrStartIndex;
                    
                    bool isFirst = true;
                    for (int k = resultStartIndex; k < endIndex; k++)
                    {
                        var textElements = runs[k].Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                        foreach (var textElement in textElements)
                        {
                            if (isFirst)
                            {
                                textElement.Text = newValue;
                                isFirst = false;
                            }
                            else
                            {
                                textElement.Text = "";
                            }
                        }
                    }

                    if (!updatedFields.ContainsKey(propertyName))
                        updatedFields[propertyName] = 0;
                    updatedFields[propertyName]++;
                }

                // Move index to after the field end
                i = endIndex;
            }
            
            return updatedFields;
        }

        public string ExportToHtml(string templatePath, Models.DocumentType documentType)
        {
            string htmlPath = Path.ChangeExtension(templatePath, ".html");
            
            if (!string.IsNullOrEmpty(_libreOfficePath))
            {
                try
                {
                    Console.WriteLine("🔄 Converting to HTML using LibreOffice headless...");
                    return ConvertUsingLibreOffice(templatePath, "html");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ LibreOffice conversion failed: {ex.Message}");
                    Console.WriteLine("📄 Falling back to basic HTML conversion...");
                }
            }
            
            // Fallback to original HTML conversion method
            try
            {
                switch (documentType)
                {
                    case Models.DocumentType.Word:
                        ExportWordToHtmlFallback(templatePath, htmlPath);
                        break;
                    case Models.DocumentType.Excel:
                        ExportExcelToHtmlFallback(templatePath, htmlPath);
                        break;
                    case Models.DocumentType.PowerPoint:
                        ExportPowerPointToHtmlFallback(templatePath, htmlPath);
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not export to HTML: {ex.Message}");
                // Create a basic HTML file with error message
                File.WriteAllText(htmlPath, $"<html><body><h1>Export Error</h1><p>Could not convert {documentType} to HTML: {ex.Message}</p></body></html>");
            }

            return htmlPath;
        }

        public string ExportToEmailHtml(string templatePath, Models.DocumentType documentType)
        {
            string emailHtmlPath = Path.ChangeExtension(templatePath, "_email.html");
            
            if (!string.IsNullOrEmpty(_libreOfficePath))
            {
                try
                {
                    Console.WriteLine("🔄 Converting to email-friendly HTML using LibreOffice...");
                    
                    // First convert using LibreOffice to get regular HTML with external images
                    string regularHtmlPath = ConvertUsingLibreOffice(templatePath, "html");
                    
                    // Then convert external images to embedded base64 images
                    return ConvertHtmlToEmailFriendly(regularHtmlPath, emailHtmlPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ LibreOffice conversion failed: {ex.Message}");
                    Console.WriteLine("📄 Falling back to basic HTML conversion...");
                }
            }
            
            // Fallback to original HTML conversion method
            try
            {
                switch (documentType)
                {
                    case Models.DocumentType.Word:
                        ExportWordToHtmlFallback(templatePath, emailHtmlPath);
                        break;
                    case Models.DocumentType.Excel:
                        ExportExcelToHtmlFallback(templatePath, emailHtmlPath);
                        break;
                    case Models.DocumentType.PowerPoint:
                        ExportPowerPointToHtmlFallback(templatePath, emailHtmlPath);
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not export to email HTML: {ex.Message}");
                // Create a basic HTML file with error message
                File.WriteAllText(emailHtmlPath, $"<html><body><h1>Export Error</h1><p>Could not convert {documentType} to email HTML: {ex.Message}</p></body></html>");
            }

            return emailHtmlPath;
        }

        public string ExportToPdf(string templatePath, Models.DocumentType documentType)
        {
            string pdfPath = Path.ChangeExtension(templatePath, ".pdf");
            
            if (!string.IsNullOrEmpty(_libreOfficePath))
            {
                try
                {
                    Console.WriteLine("🔄 Converting to PDF using LibreOffice headless...");
                    return ConvertUsingLibreOffice(templatePath, "pdf");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ LibreOffice conversion failed: {ex.Message}");
                    Console.WriteLine("📄 Falling back to basic PDF conversion...");
                }
            }
            
            // Fallback to HTML-to-PDF conversion
            try
            {
                string htmlPath = ExportToHtml(templatePath, documentType);
                ConvertHtmlToPdfFallback(htmlPath, pdfPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not export to PDF: {ex.Message}");
                // Create a basic error PDF
                CreateErrorPdf(pdfPath, $"Could not convert {documentType} to PDF: {ex.Message}");
            }

            return pdfPath;
        }

        private string ConvertUsingLibreOffice(string inputPath, string outputFormat)
        {
            string outputDir = Path.GetDirectoryName(inputPath) ?? _outputDirectory;
            string outputPath = Path.ChangeExtension(inputPath, $".{outputFormat}");
            
            // Ensure clean output directory for LibreOffice
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = _libreOfficePath,
                Arguments = $"--headless --convert-to {outputFormat} --outdir \"{outputDir}\" \"{inputPath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            Console.WriteLine($"🔧 LibreOffice command: {startInfo.FileName} {startInfo.Arguments}");

            using (var process = Process.Start(startInfo))
            {
                if (process == null)
                {
                    throw new InvalidOperationException("Failed to start LibreOffice process");
                }

                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                
                // Wait for process to complete with timeout
                if (!process.WaitForExit(30000)) // 30 second timeout
                {
                    process.Kill();
                    throw new TimeoutException("LibreOffice conversion timed out after 30 seconds");
                }

                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException($"LibreOffice conversion failed with exit code {process.ExitCode}. Error: {error}");
                }

                Console.WriteLine($"📄 LibreOffice output: {output}");
                if (!string.IsNullOrEmpty(error))
                {
                    Console.WriteLine($"⚠️ LibreOffice warnings: {error}");
                }
            }

            // Verify output file was created
            if (!File.Exists(outputPath))
            {
                throw new FileNotFoundException($"LibreOffice did not create expected output file: {outputPath}");
            }

            Console.WriteLine($"✅ Successfully converted to {outputFormat.ToUpper()}: {outputPath}");
            return outputPath;
        }

        private string ConvertHtmlToEmailFriendly(string htmlPath, string emailHtmlPath)
        {
            try
            {
                Console.WriteLine("📧 Converting HTML to email-friendly format with embedded images...");
                
                string htmlContent = File.ReadAllText(htmlPath);
                string htmlDirectory = Path.GetDirectoryName(htmlPath) ?? "";
                
                // Find all image references in the HTML
                var imageMatches = System.Text.RegularExpressions.Regex.Matches(
                    htmlContent, 
                    @"<img[^>]+src=[""']([^""']+)[""'][^>]*>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                );

                int embeddedCount = 0;
                foreach (System.Text.RegularExpressions.Match match in imageMatches)
                {
                    string imageSrc = match.Groups[1].Value;
                    string fullImagePath = Path.Combine(htmlDirectory, imageSrc);
                    
                    if (File.Exists(fullImagePath))
                    {
                        try
                        {
                            // Read image file and convert to base64
                            byte[] imageBytes = File.ReadAllBytes(fullImagePath);
                            string mimeType = GetImageMimeType(Path.GetExtension(fullImagePath));
                            string base64String = Convert.ToBase64String(imageBytes);
                            string dataUri = $"data:{mimeType};base64,{base64String}";
                            
                            // Replace the src attribute with the data URI
                            htmlContent = htmlContent.Replace($"src=\"{imageSrc}\"", $"src=\"{dataUri}\"");
                            htmlContent = htmlContent.Replace($"src='{imageSrc}'", $"src=\"{dataUri}\"");
                            
                            embeddedCount++;
                            Console.WriteLine($"   ✓ Embedded image: {imageSrc} ({FormatFileSize(imageBytes.Length)})");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"   ⚠️ Could not embed image {imageSrc}: {ex.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"   ⚠️ Image file not found: {fullImagePath}");
                    }
                }
                
                // Add email-specific CSS optimizations
                htmlContent = OptimizeHtmlForEmail(htmlContent);
                
                // Write the email-friendly HTML
                File.WriteAllText(emailHtmlPath, htmlContent, System.Text.Encoding.UTF8);
                
                Console.WriteLine($"✅ Email-friendly HTML created with {embeddedCount} embedded images");
                Console.WriteLine($"📧 Email HTML file: {emailHtmlPath}");
                
                return emailHtmlPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error converting HTML to email-friendly format: {ex.Message}");
                // Create a fallback email-friendly HTML
                CreateFallbackEmailHtml(htmlPath, emailHtmlPath);
                return emailHtmlPath;
            }
        }

        private string GetImageMimeType(string extension)
        {
            return extension.ToLowerInvariant() switch
            {
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".webp" => "image/webp",
                ".svg" => "image/svg+xml",
                _ => "image/png" // Default to PNG
            };
        }

        private string FormatFileSize(long bytes)
        {
            string[] suffixes = { "B", "KB", "MB", "GB" };
            int counter = 0;
            double number = bytes;
            while (Math.Round(number / 1024) >= 1)
            {
                number /= 1024;
                counter++;
            }
            return $"{number:n1} {suffixes[counter]}";
        }

        private string OptimizeHtmlForEmail(string htmlContent)
        {
            // Add email-specific optimizations
            var optimizedHtml = new StringBuilder();
            
            // Add email-friendly CSS and meta tags
            optimizedHtml.AppendLine("<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">");
            optimizedHtml.AppendLine("<html xmlns=\"http://www.w3.org/1999/xhtml\">");
            optimizedHtml.AppendLine("<head>");
            optimizedHtml.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />");
            optimizedHtml.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\"/>");
            optimizedHtml.AppendLine("<title>Document Export</title>");
            optimizedHtml.AppendLine("<style type=\"text/css\">");
            optimizedHtml.AppendLine("/* Email-specific CSS */");
            optimizedHtml.AppendLine("body { margin: 0; padding: 20px; font-family: Arial, sans-serif; background-color: #ffffff; }");
            optimizedHtml.AppendLine("table { border-collapse: collapse; width: 100%; max-width: 600px; margin: 0 auto; }");
            optimizedHtml.AppendLine("img { max-width: 100%; height: auto; display: block; margin: 10px auto; }");
            optimizedHtml.AppendLine("p { margin: 10px 0; line-height: 1.6; }");
            optimizedHtml.AppendLine("h1, h2, h3, h4, h5, h6 { margin: 15px 0 10px 0; }");
            optimizedHtml.AppendLine("</style>");
            optimizedHtml.AppendLine("</head>");
            
            // Extract body content from original HTML
            var bodyMatch = System.Text.RegularExpressions.Regex.Match(
                htmlContent, 
                @"<body[^>]*>(.*?)</body>", 
                System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase
            );
            
            if (bodyMatch.Success)
            {
                optimizedHtml.AppendLine("<body>");
                optimizedHtml.AppendLine(bodyMatch.Groups[1].Value);
                optimizedHtml.AppendLine("</body>");
            }
            else
            {
                // If no body found, wrap the entire content
                optimizedHtml.AppendLine("<body>");
                optimizedHtml.AppendLine(htmlContent);
                optimizedHtml.AppendLine("</body>");
            }
            
            optimizedHtml.AppendLine("</html>");
            
            return optimizedHtml.ToString();
        }

        private void CreateFallbackEmailHtml(string originalHtmlPath, string emailHtmlPath)
        {
            try
            {
                string content = File.Exists(originalHtmlPath) ? File.ReadAllText(originalHtmlPath) : "<p>Document content could not be loaded.</p>";
                var fallbackHtml = $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Document Export</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; max-width: 600px; }}
        img {{ max-width: 100%; height: auto; }}
    </style>
</head>
<body>
    <h1>Document Export</h1>
    <p><em>Note: This is a fallback email-friendly version. Some images may not display correctly.</em></p>
    {content}
</body>
</html>";
                File.WriteAllText(emailHtmlPath, fallbackHtml);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not create fallback email HTML: {ex.Message}");
            }
        }

        private void ConvertHtmlToPdfFallback(string htmlPath, string pdfPath)
        {
            try
            {
                string htmlContent = File.ReadAllText(htmlPath);

                using (var document = new iTextSharp.text.Document())
                {
                    using (var writer = PdfWriter.GetInstance(document, new FileStream(pdfPath, FileMode.Create)))
                    {
                        document.Open();
                        
                        // Parse HTML and add to PDF using basic HTML parsing
                        using (var htmlReader = new StringReader(htmlContent))
                        {
                            // Use more basic HTML parsing since HTMLWorker is obsolete
                            var paragraph = new iTextSharp.text.Paragraph(htmlContent);
                            document.Add(paragraph);
                        }
                        
                        document.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in HTML to PDF conversion: {ex.Message}");
                CreateErrorPdf(pdfPath, ex.Message);
            }
        }

        private void CreateErrorPdf(string pdfPath, string errorMessage)
        {
            try
            {
                using (var document = new iTextSharp.text.Document())
                {
                    using (var writer = PdfWriter.GetInstance(document, new FileStream(pdfPath, FileMode.Create)))
                    {
                        document.Open();
                        document.Add(new iTextSharp.text.Paragraph("PDF Conversion Error"));
                        document.Add(new iTextSharp.text.Paragraph(errorMessage));
                        document.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not create error PDF: {ex.Message}");
            }
        }

        private void ExportWordToHtmlFallback(string wordFilePath, string htmlPath)
        {
            using (var doc = WordprocessingDocument.Open(wordFilePath, false))
            {
                var body = doc.MainDocumentPart?.Document?.Body;
                if (body == null) return;

                var html = new StringBuilder();
                html.AppendLine("<!DOCTYPE html>");
                html.AppendLine("<html>");
                html.AppendLine("<head>");
                html.AppendLine("<meta charset='utf-8'>");
                html.AppendLine("<title>Generated Word Document</title>");
                html.AppendLine("<style>");
                html.AppendLine("body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }");
                html.AppendLine("p { margin: 10px 0; }");
                html.AppendLine("</style>");
                html.AppendLine("</head>");
                html.AppendLine("<body>");

                foreach (var para in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    var text = string.Join("", para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        html.AppendLine($"<p>{System.Net.WebUtility.HtmlEncode(text)}</p>");
                    }
                }

                html.AppendLine("</body>");
                html.AppendLine("</html>");

                File.WriteAllText(htmlPath, html.ToString());
            }
        }

        private void ExportExcelToHtmlFallback(string excelFilePath, string htmlPath)
        {
            using (var doc = SpreadsheetDocument.Open(excelFilePath, false))
            {
                var html = new StringBuilder();
                html.AppendLine("<!DOCTYPE html>");
                html.AppendLine("<html>");
                html.AppendLine("<head>");
                html.AppendLine("<meta charset='utf-8'>");
                html.AppendLine("<title>Generated Excel Document</title>");
                html.AppendLine("<style>");
                html.AppendLine("body { font-family: Arial, sans-serif; margin: 40px; }");
                html.AppendLine("table { border-collapse: collapse; width: 100%; }");
                html.AppendLine("th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }");
                html.AppendLine("th { background-color: #f2f2f2; }");
                html.AppendLine("</style>");
                html.AppendLine("</head>");
                html.AppendLine("<body>");
                html.AppendLine("<h1>Excel Document Content</h1>");
                
                // Basic Excel to HTML conversion (simplified)
                var workbookPart = doc.WorkbookPart;
                if (workbookPart?.Workbook?.Sheets != null)
                {
                    foreach (var sheet in workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
                    {
                        html.AppendLine($"<h2>Sheet: {sheet.Name}</h2>");
                        html.AppendLine("<p>Excel content extraction requires more complex implementation.</p>");
                    }
                }

                html.AppendLine("</body>");
                html.AppendLine("</html>");

                File.WriteAllText(htmlPath, html.ToString());
            }
        }

        private void ExportPowerPointToHtmlFallback(string pptFilePath, string htmlPath)
        {
            using (var doc = PresentationDocument.Open(pptFilePath, false))
            {
                var html = new StringBuilder();
                html.AppendLine("<!DOCTYPE html>");
                html.AppendLine("<html>");
                html.AppendLine("<head>");
                html.AppendLine("<meta charset='utf-8'>");
                html.AppendLine("<title>Generated PowerPoint Document</title>");
                html.AppendLine("<style>");
                html.AppendLine("body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }");
                html.AppendLine(".slide { margin: 20px 0; padding: 20px; border: 1px solid #ccc; }");
                html.AppendLine("h2 { color: #333; }");
                html.AppendLine("</style>");
                html.AppendLine("</head>");
                html.AppendLine("<body>");
                html.AppendLine("<h1>PowerPoint Presentation</h1>");

                // Basic PowerPoint to HTML conversion (simplified)
                var presentationPart = doc.PresentationPart;
                if (presentationPart?.Presentation?.SlideIdList != null)
                {
                    int slideNumber = 1;
                    foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<DocumentFormat.OpenXml.Presentation.SlideId>())
                    {
                        html.AppendLine($"<div class='slide'>");
                        html.AppendLine($"<h2>Slide {slideNumber}</h2>");
                        html.AppendLine("<p>PowerPoint content extraction requires more complex implementation.</p>");
                        html.AppendLine("</div>");
                        slideNumber++;
                    }
                }

                html.AppendLine("</body>");
                html.AppendLine("</html>");

                File.WriteAllText(htmlPath, html.ToString());
            }
        }
    }
}
