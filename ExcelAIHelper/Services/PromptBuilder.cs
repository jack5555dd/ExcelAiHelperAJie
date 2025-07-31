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
            
            sb.AppendLine("You are an Excel AI assistant that helps users with Excel tasks.");
            sb.AppendLine("You can perform operations like setting cell values, applying formulas, formatting cells, etc.");
            sb.AppendLine("IMPORTANT: Always respond with ONLY a valid JSON object. Do not include markdown code blocks, explanations, or any other text.");
            sb.AppendLine("The JSON should follow this exact format:");
            sb.AppendLine("{");
            sb.AppendLine("  \"summary\": \"Brief description of what you're doing\",");
            sb.AppendLine("  \"instructions\": [");
            sb.AppendLine("    {");
            sb.AppendLine("      \"type\": \"SetCellValue\",");
            sb.AppendLine("      \"description\": \"Human-readable description\",");
            sb.AppendLine("      \"targetRange\": \"A1\",");
            sb.AppendLine("      \"parameters\": { \"value\": 100 },");
            sb.AppendLine("      \"requiresConfirmation\": false");
            sb.AppendLine("    },");
            sb.AppendLine("    {");
            sb.AppendLine("      \"type\": \"ApplyFormula\",");
            sb.AppendLine("      \"description\": \"Apply formula\",");
            sb.AppendLine("      \"targetRange\": \"C1\",");
            sb.AppendLine("      \"parameters\": { \"formula\": \"=SUM(A1:B1)\" },");
            sb.AppendLine("      \"requiresConfirmation\": false");
            sb.AppendLine("    }");
            sb.AppendLine("  ]");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("Available instruction types:");
            sb.AppendLine("- SetCellValue: Sets cell value(s)");
            sb.AppendLine("- SetCellFormat: Sets cell format");
            sb.AppendLine("- ApplyFormula: Applies a formula (e.g., =SUM(A1:A10), =AVERAGE(B1:B5), =A1+B1)");
            sb.AppendLine("- SetCellStyle: Sets cell style (backgroundColor, fontColor, bold, italic, underline, fontSize, fontName)");
            sb.AppendLine("- ClearContent: Clears cell content");
            sb.AppendLine("- InsertRows: Inserts rows");
            sb.AppendLine("- InsertColumns: Inserts columns");
            sb.AppendLine("- DeleteRows: Deletes rows");
            sb.AppendLine("- DeleteColumns: Deletes columns");
            sb.AppendLine("- SortData: Sorts data");
            sb.AppendLine("- FilterData: Filters data");
            sb.AppendLine("- CreateChart: Creates a chart");
            sb.AppendLine("- ApplyConditionalFormatting: Applies conditional formatting");
            sb.AppendLine();
            
            // Add context information
            sb.AppendLine("Current Excel context:");
            sb.AppendLine(await _contextManager.GetContextDescriptionAsync());
            sb.AppendLine();
            sb.AppendLine("Location Guidelines:");
            sb.AppendLine("- When user mentions specific cell coordinates (like 'C1', 'A5'), use that as targetRange");
            sb.AppendLine("- When user says 'current selection' or 'selected cell', use the current selected range");
            sb.AppendLine("- When user says 'here' or similar, refer to the current selection");
            sb.AppendLine();
            sb.AppendLine("Formula Guidelines:");
            sb.AppendLine("- For ApplyFormula instructions, use 'formula' parameter (not 'value')");
            sb.AppendLine("- For SUM formulas, specify a range like =SUM(A1:A10) or =SUM(A:A)");
            sb.AppendLine("- For empty SUM(), the system will try to auto-detect nearby numeric data");
            sb.AppendLine("- Always include proper cell references in formulas");
            sb.AppendLine("- Use Excel function syntax (e.g., =AVERAGE(range), =COUNT(range))");
            sb.AppendLine();
            sb.AppendLine("Style Guidelines:");
            sb.AppendLine("- For SetCellStyle: backgroundColor (hex like #FF0000 or name like 'red')");
            sb.AppendLine("- fontColor (hex like #0000FF or name like 'blue')");
            sb.AppendLine("- bold, italic, underline (true/false)");
            sb.AppendLine("- fontSize (number like 12, 14, 16)");
            sb.AppendLine("- fontName (string like 'Arial', 'Times New Roman')");
            sb.AppendLine("- For ClearContent: use when user wants to delete/clear cell content");
            
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
            
            sb.AppendLine(userRequest);
            sb.AppendLine();
            sb.AppendLine("IMPORTANT: Respond with ONLY valid JSON. No markdown, no code blocks, no explanations - just the JSON object.");
            
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