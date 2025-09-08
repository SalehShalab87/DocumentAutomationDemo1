using DocumentAutomation.Library.Models;
using DocumentAutomation.Library.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;

namespace DocumentAutomationDemo
{
    class Program
    {
        private static ITemplateService _templateService = null!;
        private static IDocumentGenerationService _documentService = null!;
        private static IDocumentEmbeddingService _embeddingService = null!;

        static void Main(string[] args)
        {
            Console.WriteLine("=== Document Automation Demo ===");
            Console.WriteLine("This will become a DLL for Angular integration\n");

            // Initialize services
            _templateService = new TemplateService();
            _documentService = new DocumentGenerationService(_templateService);
            _embeddingService = new DocumentEmbeddingService(_templateService);

            try
            {
                ShowMainMenu();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Details: {ex.StackTrace}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        static void ShowMainMenu()
        {
            while (true)
            {
                SafeConsoleClear();
                Console.WriteLine("=== Document Automation Demo ===\n");
                Console.WriteLine("1. Register a new template");
                Console.WriteLine("2. View all templates");
                Console.WriteLine("3. Generate document from template");
                Console.WriteLine("4. Generate document with embedded templates");
                Console.WriteLine("5. Delete a template");
                Console.WriteLine("6. Exit");
                Console.Write("\nSelect an option (1-6): ");

                var choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        RegisterTemplate();
                        break;
                    case "2":
                        ViewTemplates();
                        break;
                    case "3":
                        GenerateDocument();
                        break;
                    case "4":
                        GenerateDocumentWithEmbedding();
                        break;
                    case "5":
                        DeleteTemplate();
                        break;
                    case "6":
                        return;
                    default:
                        Console.WriteLine("Invalid option. Press any key to continue...");
                        Console.ReadKey();
                        break;
                }
            }
        }

        static void RegisterTemplate()
        {
            SafeConsoleClear();
            Console.WriteLine("=== Register New Template ===\n");

            Console.Write("Enter template name: ");
            string templateName = Console.ReadLine() ?? "";

            if (string.IsNullOrWhiteSpace(templateName))
            {
                Console.WriteLine("Template name cannot be empty.");
                Console.ReadKey();
                return;
            }

            Console.Write("Enter path to template file (.docx, .xlsx, or .pptx): ");
            string filePath = Console.ReadLine() ?? "";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("❌ File not found. Please provide a valid path to an existing template file.");
                Console.WriteLine("   Supported formats: .docx (Word), .xlsx (Excel), .pptx (PowerPoint)");
                Console.ReadKey();
                return;
            }

            // Validate file extension
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".docx" && extension != ".xlsx" && extension != ".pptx")
            {
                Console.WriteLine("❌ Unsupported file format. Only .docx, .xlsx, and .pptx files are supported.");
                Console.ReadKey();
                return;
            }

            try
            {
                string templateId = _templateService.RegisterTemplate(filePath, templateName);
                var template = _templateService.GetTemplate(templateId);

                Console.WriteLine($"\n✅ Template registered successfully!");
                Console.WriteLine($"Template ID: {templateId}");
                Console.WriteLine($"Template Name: {templateName}");
                Console.WriteLine($"Document Type: {template?.DocumentType}");
                Console.WriteLine($"Placeholders found: {string.Join(", ", template?.Placeholders ?? new List<string>())}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error registering template: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();
        }

        static void ViewTemplates()
        {
            SafeConsoleClear();
            Console.WriteLine("=== All Templates ===\n");

            var templates = _templateService.GetAllTemplates();

            if (templates.Count == 0)
            {
                Console.WriteLine("No templates found.");
            }
            else
            {
                for (int i = 0; i < templates.Count; i++)
                {
                    var template = templates[i];
                    Console.WriteLine($"{i + 1}. {template.Name} ({template.DocumentType})");
                    Console.WriteLine($"   ID: {template.Id}");
                    Console.WriteLine($"   Created: {template.CreatedDate:yyyy-MM-dd HH:mm}");
                    Console.WriteLine($"   Placeholders: {string.Join(", ", template.Placeholders)}");
                    
                    // Show actual custom properties with current values
                    Console.WriteLine("   Custom Properties:");
                    var customProps = _templateService.GetCustomPropertiesWithValues(template.FilePath);
                    if (customProps.Count > 0)
                    {
                        foreach (var prop in customProps)
                        {
                            Console.WriteLine($"     • {prop.Key}: {prop.Value}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("     • No custom properties found");
                    }
                    Console.WriteLine();
                }
            }

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void GenerateDocument()
        {
            SafeConsoleClear();
            Console.WriteLine("=== Generate Document ===\n");

            // First show available templates for reference
            var templates = _templateService.GetAllTemplates();
            if (templates.Count > 0)
            {
                Console.WriteLine("Available templates for reference:");
                foreach (var template in templates)
                {
                    Console.WriteLine($"- {template.Name} (ID: {template.Id})");
                    Console.WriteLine($"  Placeholders: {string.Join(", ", template.Placeholders)}");
                }
                Console.WriteLine();
            }

            Console.Write("Enter path to values JSON file (or press Enter for manual input): ");
            string jsonPath = Console.ReadLine() ?? "";

            if (string.IsNullOrWhiteSpace(jsonPath))
            {
                // Fall back to manual input if no JSON file specified
                GenerateDocumentManual();
                return;
            }

            if (!File.Exists(jsonPath))
            {
                Console.WriteLine("❌ JSON file not found.");
                Console.WriteLine("\nExample JSON format:");
                Console.WriteLine(@"{
    ""templateId"": ""template_1"",
    ""values"": {
        ""ICS_CustomerName"": ""John Doe"",
        ""ICS_QuotationNumber"": ""Q12345""
    },
    ""exportFormat"": ""Original""
}");
                Console.ReadKey();
                return;
            }

            try
            {
                // Read and parse JSON file
                string jsonContent = File.ReadAllText(jsonPath);
                var documentValues = System.Text.Json.JsonSerializer.Deserialize<DocumentValues>(jsonContent);

                if (documentValues == null)
                {
                    Console.WriteLine("❌ Invalid JSON format.");
                    Console.ReadKey();
                    return;
                }

                // Validate template exists
                var template = _templateService.GetTemplate(documentValues.TemplateId);
                if (template == null)
                {
                    Console.WriteLine($"❌ Template with ID '{documentValues.TemplateId}' not found.");
                    Console.ReadKey();
                    return;
                }

                // Convert dictionary to PlaceholderValues
                var placeholderValues = documentValues.Values.Select(kv => new PlaceholderValue 
                { 
                    Placeholder = kv.Key, 
                    Value = kv.Value 
                }).ToList();

                // Check if export format is specified in JSON, if not ask user
                ExportFormat finalExportFormat = documentValues.ExportFormat;
                if (documentValues.ExportFormat == ExportFormat.Original)
                {
                    Console.WriteLine($"\n📄 JSON file doesn't specify export format (or uses Original).");
                    finalExportFormat = AskForExportFormat();
                }
                else
                {
                    Console.WriteLine($"📑 Using export format from JSON: {documentValues.ExportFormat}");
                }

                var request = new DocumentGenerationRequest
                {
                    TemplateId = documentValues.TemplateId,
                    PlaceholderValues = placeholderValues,
                    ExportFormat = finalExportFormat
                };

                string outputPath = _documentService.GenerateDocument(request);
                Console.WriteLine($"\n✅ Document generated successfully!");
                Console.WriteLine($"Output file: {outputPath}");
            }
            catch (System.Text.Json.JsonException ex)
            {
                Console.WriteLine($"❌ Error parsing JSON: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error generating document: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();
        }

        static void GenerateDocumentManual()
        {
            var templates = _templateService.GetAllTemplates();
            if (templates.Count == 0)
            {
                Console.WriteLine("No templates available. Please register a template first.");
                Console.ReadKey();
                return;
            }

            // Show templates
            Console.WriteLine("Available templates:");
            for (int i = 0; i < templates.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {templates[i].Name} (ID: {templates[i].Id})");
            }

            Console.Write("\nSelect template number: ");
            if (!int.TryParse(Console.ReadLine(), out int templateIndex) || 
                templateIndex < 1 || templateIndex > templates.Count)
            {
                Console.WriteLine("Invalid template selection.");
                Console.ReadKey();
                return;
            }

            var selectedTemplate = templates[templateIndex - 1];

            // Get placeholder values
            var placeholderValues = new List<PlaceholderValue>();
            Console.WriteLine($"\nEnter values for placeholders in '{selectedTemplate.Name}':");

            foreach (var placeholder in selectedTemplate.Placeholders)
            {
                Console.Write($"{placeholder}: ");
                string value = Console.ReadLine() ?? "";
                placeholderValues.Add(new PlaceholderValue 
                { 
                    Placeholder = placeholder, 
                    Value = value 
                });
            }

            // Choose export format
            ExportFormat format = AskForExportFormat();

            try
            {
                var request = new DocumentGenerationRequest
                {
                    TemplateId = selectedTemplate.Id,
                    PlaceholderValues = placeholderValues,
                    ExportFormat = format
                };

                string outputPath = _documentService.GenerateDocument(request);
                Console.WriteLine($"\n✅ Document generated successfully!");
                Console.WriteLine($"Output file: {outputPath}");

                // Generate example JSON for future use
                var jsonExample = new DocumentValues
                {
                    TemplateId = selectedTemplate.Id,
                    Values = placeholderValues.ToDictionary(pv => pv.Placeholder, pv => pv.Value),
                    ExportFormat = format
                };

                string jsonPath = Path.Combine(
                    Path.GetDirectoryName(outputPath) ?? "",
                    Path.GetFileNameWithoutExtension(outputPath) + "_values.json"
                );

                File.WriteAllText(
                    jsonPath,
                    System.Text.Json.JsonSerializer.Serialize(jsonExample, new System.Text.Json.JsonSerializerOptions { WriteIndented = true })
                );

                Console.WriteLine($"JSON template saved to: {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error generating document: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();
        }

        static void DeleteTemplate()
        {
            SafeConsoleClear();
            Console.WriteLine("=== Delete Template ===\n");

            var templates = _templateService.GetAllTemplates();
            if (templates.Count == 0)
            {
                Console.WriteLine("No templates available.");
                Console.ReadKey();
                return;
            }

            // Show templates
            Console.WriteLine("Available templates:");
            for (int i = 0; i < templates.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {templates[i].Name} (ID: {templates[i].Id})");
            }

            Console.Write("\nSelect template number to delete: ");
            if (!int.TryParse(Console.ReadLine(), out int templateIndex) || 
                templateIndex < 1 || templateIndex > templates.Count)
            {
                Console.WriteLine("Invalid template selection.");
                Console.ReadKey();
                return;
            }

            var selectedTemplate = templates[templateIndex - 1];

            Console.Write($"Are you sure you want to delete '{selectedTemplate.Name}'? (y/N): ");
            if (Console.ReadLine()?.ToLower() == "y")
            {
                if (_templateService.DeleteTemplate(selectedTemplate.Id))
                {
                    Console.WriteLine("✅ Template deleted successfully!");
                }
                else
                {
                    Console.WriteLine("❌ Error deleting template.");
                }
            }
            else
            {
                Console.WriteLine("Delete cancelled.");
            }

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static void GenerateDocumentWithEmbedding()
        {
            SafeConsoleClear();
            Console.WriteLine("=== Generate Document with Embedded Templates ===\n");

            var templates = _templateService.GetAllTemplates();
            if (templates.Count == 0)
            {
                Console.WriteLine("❌ No templates available. Please register templates first.");
                Console.ReadKey();
                return;
            }

            // Filter to only Word documents
            var wordTemplates = templates.Where(t => t.DocumentType == DocumentType.Word).ToList();
            if (wordTemplates.Count == 0)
            {
                Console.WriteLine("❌ No Word document templates available. Document embedding only supports Word documents.");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("📋 Available Word templates:");
            for (int i = 0; i < wordTemplates.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {wordTemplates[i].Name} (ID: {wordTemplates[i].Id})");
                Console.WriteLine($"   Placeholders: {string.Join(", ", wordTemplates[i].Placeholders)}");
            }

            // Step 1: Choose main template
            Console.Write("\nSelect main template number: ");
            if (!int.TryParse(Console.ReadLine(), out int mainTemplateIndex) || 
                mainTemplateIndex < 1 || mainTemplateIndex > wordTemplates.Count)
            {
                Console.WriteLine("❌ Invalid template selection.");
                Console.ReadKey();
                return;
            }

            var mainTemplate = wordTemplates[mainTemplateIndex - 1];

            // Step 2: Get main template values (JSON or manual)
            var mainTemplateValues = GetTemplateValues(mainTemplate, "main template");
            if (mainTemplateValues == null) return;

            // Step 3: Configure embeddings (supports multiple embeddings)
            var embeddings = new List<EmbedInfo>();
            
            while (true)
            {
                Console.WriteLine($"\n📍 Current embeddings: {embeddings.Count}");
                Console.WriteLine("1. Add embedding");
                Console.WriteLine("2. Finish and generate document");
                Console.Write("Choice: ");
                
                var choice = Console.ReadLine();
                
                if (choice == "1")
                {
                    var embedding = ConfigureEmbedding(wordTemplates);
                    if (embedding != null)
                    {
                        embeddings.Add(embedding);
                        Console.WriteLine($"✅ Added embedding: {embedding.EmbedPlaceholder}");
                    }
                }
                else if (choice == "2")
                {
                    if (embeddings.Count == 0)
                    {
                        Console.WriteLine("❌ No embeddings configured. Please add at least one embedding.");
                        continue;
                    }
                    break;
                }
                else
                {
                    Console.WriteLine("Invalid choice. Please select 1 or 2.");
                }
            }

            // Step 4: Choose export format
            ExportFormat format = AskForExportFormat();

            // Step 5: Generate document with embeddings
            try
            {
                var request = new DocumentEmbeddingRequest
                {
                    MainTemplateId = mainTemplate.Id,
                    MainTemplateValues = mainTemplateValues,
                    Embeddings = embeddings,
                    ExportFormat = format
                };

                string outputPath = _embeddingService.GenerateDocumentWithEmbedding(request);
                Console.WriteLine($"\n✅ Document with {embeddings.Count} embedding(s) generated successfully!");
                Console.WriteLine($"Output file: {outputPath}");

                // Generate example JSON for future use
                var jsonExample = new DocumentEmbeddingValues
                {
                    MainTemplateId = mainTemplate.Id,
                    MainValues = mainTemplateValues.ToDictionary(pv => pv.Placeholder, pv => pv.Value),
                    Embeddings = embeddings.Select(e => new EmbedValues
                    {
                        EmbedTemplateId = e.EmbedTemplateId,
                        Values = e.EmbedTemplateValues.ToDictionary(pv => pv.Placeholder, pv => pv.Value),
                        EmbedPlaceholder = e.EmbedPlaceholder
                    }).ToList(),
                    ExportFormat = format
                };

                string jsonPath = Path.Combine(
                    Path.GetDirectoryName(outputPath) ?? "",
                    Path.GetFileNameWithoutExtension(outputPath) + "_embedding_values.json"
                );

                File.WriteAllText(
                    jsonPath,
                    System.Text.Json.JsonSerializer.Serialize(jsonExample, new System.Text.Json.JsonSerializerOptions { WriteIndented = true })
                );

                Console.WriteLine($"JSON template saved to: {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error generating document with embedding: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();
        }

        static EmbedInfo? ConfigureEmbedding(List<DocumentTemplate> wordTemplates)
        {
            Console.WriteLine("\n📄 Configure Embedding:");
            
            // Choose embed template
            Console.WriteLine("Available templates to embed:");
            for (int i = 0; i < wordTemplates.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {wordTemplates[i].Name} (ID: {wordTemplates[i].Id})");
                Console.WriteLine($"   Placeholders: {string.Join(", ", wordTemplates[i].Placeholders)}");
            }

            Console.Write("Select template to embed number: ");
            if (!int.TryParse(Console.ReadLine(), out int embedTemplateIndex) || 
                embedTemplateIndex < 1 || embedTemplateIndex > wordTemplates.Count)
            {
                Console.WriteLine("❌ Invalid template selection.");
                return null;
            }

            var embedTemplate = wordTemplates[embedTemplateIndex - 1];

            // Get placeholder for embedding location
            Console.Write($"Enter placeholder in main template where to embed '{embedTemplate.Name}': ");
            string embedPlaceholder = Console.ReadLine() ?? "";
            if (string.IsNullOrWhiteSpace(embedPlaceholder))
            {
                Console.WriteLine("❌ Embed placeholder cannot be empty.");
                return null;
            }

            // Get embed template values
            var embedTemplateValues = GetTemplateValues(embedTemplate, $"embed template '{embedTemplate.Name}'");
            if (embedTemplateValues == null) return null;

            return new EmbedInfo
            {
                EmbedTemplateId = embedTemplate.Id,
                EmbedTemplateValues = embedTemplateValues,
                EmbedPlaceholder = embedPlaceholder
            };
        }

        static List<PlaceholderValue>? GetTemplateValues(DocumentTemplate template, string templateDescription)
        {
            Console.WriteLine($"\n📄 Getting values for {templateDescription} '{template.Name}'");
            Console.Write("Enter path to values JSON file (or press Enter for manual input): ");
            string jsonPath = Console.ReadLine() ?? "";

            if (!string.IsNullOrWhiteSpace(jsonPath))
            {
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine("❌ JSON file not found.");
                    return null;
                }

                try
                {
                    string jsonContent = File.ReadAllText(jsonPath);
                    var documentValues = System.Text.Json.JsonSerializer.Deserialize<DocumentValues>(jsonContent);
                    
                    if (documentValues?.Values != null)
                    {
                        return documentValues.Values.Select(kv => new PlaceholderValue 
                        { 
                            Placeholder = kv.Key, 
                            Value = kv.Value 
                        }).ToList();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error reading JSON: {ex.Message}");
                    return null;
                }
            }

            // Manual input
            var placeholderValues = new List<PlaceholderValue>();
            Console.WriteLine($"Enter values for placeholders in '{template.Name}':");

            foreach (var placeholder in template.Placeholders)
            {
                Console.Write($"{placeholder}: ");
                string value = Console.ReadLine() ?? "";
                placeholderValues.Add(new PlaceholderValue 
                { 
                    Placeholder = placeholder, 
                    Value = value 
                });
            }

            return placeholderValues;
        }

        static ExportFormat AskForExportFormat()
        {
            Console.WriteLine("\nSelect export format:");
            Console.WriteLine("1. Keep original format (.docx/.xlsx/.pptx)");
            Console.WriteLine("2. Word (.docx)");
            Console.WriteLine("3. HTML (.html) - with external images");
            Console.WriteLine("4. HTML for Email (.html) - with embedded images");
            Console.WriteLine("5. PDF (.pdf)");
            Console.Write("Choice (1-5): ");

            var formatChoice = Console.ReadLine();
            return formatChoice switch
            {
                "1" => ExportFormat.Original,
                "2" => ExportFormat.Word,
                "3" => ExportFormat.Html,
                "4" => ExportFormat.HtmlEmail,
                "5" => ExportFormat.Pdf,
                _ => ExportFormat.Original
            };
        }

        static void SafeConsoleClear()
        {
            try
            {
                Console.Clear();
            }
            catch (IOException)
            {
                // Console.Clear() can fail when running under debugger or in some terminals
                // Just add some blank lines instead
                Console.WriteLine(new string('\n', 3));
            }
        }
    }
}
