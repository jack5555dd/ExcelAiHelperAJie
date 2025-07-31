using System;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Builds prompts for AI requests
    /// </summary>
    public class PromptBuilder
    {
        private readonly ContextManager _contextManager;
        
        /// <summary>
        /// Creates a new instance of PromptBuilder
        /// </summary>
        /// <param name="contextManager">The context manager</param>
        public PromptBuilder(ContextManager contextManager)
        {
            _contextManager = contextManager ?? throw new ArgumentNullException(nameof(contextManager));
        }
        
        /// <summary>
        /// Builds a system prompt with context information
        /// </summary>
        /// <returns>The system prompt</returns>
        public async Task<string> BuildSystemPromptAsync()
        {
            var sb = new StringBuilder();
            
            sb.AppendLine("You are an Excel AI assistant that MUST follow the strict JSON Command Protocol.");
            sb.AppendLine();
            sb.AppendLine("=== CRITICAL PROTOCOL REQUIREMENTS ===");
            sb.AppendLine("1. RESPOND WITH ONLY VALID JSON - NO markdown, NO code blocks, NO explanations");
            sb.AppendLine("2. JSON MUST be valid and parseable");
            sb.AppendLine("3. JSON MUST follow the exact schema structure");
            sb.AppendLine("4. Any deviation will result in PROTOCOL VIOLATION and request rejection");
            sb.AppendLine();
            sb.AppendLine("=== REQUIRED JSON SCHEMA ===");
            sb.AppendLine("{");
            sb.AppendLine("  \"version\": \"1.0\",");
            sb.AppendLine("  \"summary\": \"Brief description of operations\",");
            sb.AppendLine("  \"commands\": [");
            sb.AppendLine("    {");
            sb.AppendLine("      \"function\": \"setCellValue\",");
            sb.AppendLine("      \"arguments\": {");
            sb.AppendLine("        \"range\": \"A1\",");
            sb.AppendLine("        \"value\": 100,");
            sb.AppendLine("        \"dataType\": \"number\"");
            sb.AppendLine("      },");
            sb.AppendLine("      \"description\": \"Set A1 to 100\"");
            sb.AppendLine("    }");
            sb.AppendLine("  ]");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("=== AVAILABLE FUNCTIONS ===");
            sb.AppendLine("1. setCellValue - Set cell value(s)");
            sb.AppendLine("   Arguments: {\"range\": \"A1\", \"value\": 100, \"dataType\": \"number\"}");
            sb.AppendLine();
            sb.AppendLine("2. applyCellFormula - Apply Excel formula");
            sb.AppendLine("   Arguments: {\"range\": \"B1\", \"formula\": \"=SUM(A1:A10)\"}");
            sb.AppendLine();
            sb.AppendLine("3. setCellStyle - Set cell styling");
            sb.AppendLine("   Arguments: {\"range\": \"A1\", \"backgroundColor\": \"red\", \"bold\": true}");
            sb.AppendLine();
            sb.AppendLine("4. setCellFormat - Set number format");
            sb.AppendLine("   Arguments: {\"range\": \"A1\", \"format\": \"0.00%\"}");
            sb.AppendLine();
            sb.AppendLine("5. clearCellContent - Clear cell content");
            sb.AppendLine("   Arguments: {\"range\": \"A1:B5\", \"clearType\": \"content\"}");
            sb.AppendLine();
            sb.AppendLine("6. insertRows - Insert rows");
            sb.AppendLine("   Arguments: {\"position\": 3, \"count\": 2}");
            sb.AppendLine();
            sb.AppendLine("7. insertColumns - Insert columns");
            sb.AppendLine("   Arguments: {\"position\": \"C\", \"count\": 1}");
            sb.AppendLine();
            sb.AppendLine("8. deleteRows - Delete rows");
            sb.AppendLine("   Arguments: {\"range\": \"3:5\"}");
            sb.AppendLine();
            sb.AppendLine("9. deleteColumns - Delete columns");
            sb.AppendLine("   Arguments: {\"range\": \"B:D\"}");
            sb.AppendLine();
            
            // Add context information
            sb.AppendLine("Current Excel context:");
            sb.AppendLine(await _contextManager.GetContextDescriptionAsync());
            sb.AppendLine();
            sb.AppendLine("=== RANGE SPECIFICATION RULES ===");
            sb.AppendLine("- Specific coordinates: \"A1\", \"B2:D5\"");
            sb.AppendLine("- Current selection: \"CURRENT_SELECTION\"");
            sb.AppendLine("- Full columns: \"A:A\", \"B:D\"");
            sb.AppendLine("- Full rows: \"1:1\", \"3:5\"");
            sb.AppendLine();
            sb.AppendLine("=== FORMULA RULES ===");
            sb.AppendLine("- MUST start with = sign");
            sb.AppendLine("- Use proper Excel syntax: =SUM(A1:A10), =AVERAGE(B1:B5)");
            sb.AppendLine("- Include cell references: =A1+B1, =MAX(C1:C10)");
            sb.AppendLine();
            sb.AppendLine("=== STYLE RULES ===");
            sb.AppendLine("- Colors: hex (#FF0000) or names (red, blue, green)");
            sb.AppendLine("- Boolean properties: bold, italic, underline (true/false)");
            sb.AppendLine("- Font size: 6-72 (integer)");
            sb.AppendLine("- Font names: Arial, Calibri, Times New Roman, 宋体, 微软雅黑");
            sb.AppendLine();
            sb.AppendLine("=== PROTOCOL ENFORCEMENT ===");
            sb.AppendLine("- Response MUST be valid JSON only");
            sb.AppendLine("- NO markdown code blocks (```json)");
            sb.AppendLine("- NO explanatory text before or after JSON");
            sb.AppendLine("- NO comments in JSON");
            sb.AppendLine("- STRICT schema compliance required");
            
            return sb.ToString();
        }
        
        /// <summary>
        /// Builds a user prompt with the user's request and context
        /// </summary>
        /// <param name="userRequest">The user's request</param>
        /// <returns>The enhanced user prompt</returns>
        public async Task<string> BuildUserPromptAsync(string userRequest)
        {
            if (string.IsNullOrEmpty(userRequest))
            {
                throw new ArgumentNullException(nameof(userRequest));
            }
            
            var sb = new StringBuilder();
            
            sb.AppendLine("USER REQUEST:");
            sb.AppendLine(userRequest);
            sb.AppendLine();
            sb.AppendLine("=== RESPONSE REQUIREMENTS ===");
            sb.AppendLine("1. RESPOND WITH ONLY VALID JSON");
            sb.AppendLine("2. NO markdown code blocks (```json)");
            sb.AppendLine("3. NO explanatory text");
            sb.AppendLine("4. NO comments");
            sb.AppendLine("5. MUST follow exact schema structure");
            sb.AppendLine("6. Start response immediately with {");
            sb.AppendLine();
            sb.AppendLine("GENERATE JSON COMMAND NOW:");
            
            return sb.ToString();
        }
        
        /// <summary>
        /// Builds a complete prompt with system and user parts
        /// </summary>
        /// <param name="userRequest">The user's request</param>
        /// <returns>The system prompt and user prompt</returns>
        public async Task<(string SystemPrompt, string UserPrompt)> BuildAsync(string userRequest)
        {
            string systemPrompt = await BuildSystemPromptAsync();
            string userPrompt = await BuildUserPromptAsync(userRequest);
            
            return (systemPrompt, userPrompt);
        }
    }
}