using DocumentAutomationDemo.Models;
using DocumentAutomationDemo.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;

namespace DocumentAutomationDemo
{
    class Program
    {
        private static ITemplateService _templateService = null!;
        private static IDocumentGenerationService _documentService = null!;

        static void Main(string[] args)
        {
            Console.WriteLine("=== Document Automation Demo ===");
            Console.WriteLine("This will become a DLL for Angular integration\n");

            // Initialize services
            _templateService = new TemplateService();
            _documentService = new DocumentGenerationService(_templateService);

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
                Console.WriteLine("4. Delete a template");
                Console.WriteLine("5. Exit");
                Console.Write("\nSelect an option (1-5): ");

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
                        DeleteTemplate();
                        break;
                    case "5":
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
            Console.WriteLine("\nSelect export format:");
            Console.WriteLine("1. Keep original format");
            Console.WriteLine("2. Word (.docx)");
            Console.WriteLine("3. HTML (.html)");
            Console.WriteLine("4. PDF (.pdf)");
            Console.Write("Choice (1-4): ");

            ExportFormat format = ExportFormat.Original;
            var formatChoice = Console.ReadLine();
            switch (formatChoice)
            {
                case "1":
                    format = ExportFormat.Original;
                    break;
                case "2":
                    format = ExportFormat.Word;
                    break;
                case "3":
                    format = ExportFormat.Html;
                    break;
                case "4":
                    format = ExportFormat.Pdf;
                    break;
                default:
                    format = ExportFormat.Original;
                    break;
            }

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
