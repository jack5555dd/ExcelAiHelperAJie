using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Engine for executing Excel operations
    /// </summary>
    public class ExcelOperationEngine
    {
        private readonly Excel.Application _excelApp;
        
        /// <summary>
        /// Creates a new instance of ExcelOperationEngine
        /// </summary>
        /// <param name="excelApp">The Excel application instance</param>
        public ExcelOperationEngine(Excel.Application excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
        }
        
        /// <summary>
        /// Gets the currently selected range
        /// </summary>
        /// <returns>The selected range</returns>
        public async Task<Excel.Range> GetSelectedRangeAsync()
        {
            return await Task.Run(() => {
                try
                {
                    return _excelApp.Selection as Excel.Range;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to get selected range", ex);
                }
            });
        }

        /// <summary>
        /// Gets the target range based on instruction, with intelligent fallback
        /// </summary>
        /// <param name="targetRangeSpec">The target range specification from instruction</param>
        /// <returns>The resolved target range</returns>
        private async Task<Excel.Range> GetTargetRangeAsync(string targetRangeSpec)
        {
            try
            {
                // If specific range is provided, use it
                if (!string.IsNullOrEmpty(targetRangeSpec))
                {
                    // Clean up the range specification
                    string cleanRange = targetRangeSpec.Trim().ToUpper();
                    
                    // Handle common variations
                    if (cleanRange.Contains("CURRENT") || cleanRange.Contains("SELECTED") || cleanRange.Contains("SELECTION"))
                    {
                        return await GetSelectedRangeAsync();
                    }
                    
                    // Try to parse as Excel range
                    try
                    {
                        return _excelApp.ActiveSheet.Range[cleanRange];
                    }
                    catch
                    {
                        // If parsing fails, fall back to selected range
                        System.Diagnostics.Debug.WriteLine($"Failed to parse range '{targetRangeSpec}', using selected range");
                        return await GetSelectedRangeAsync();
                    }
                }
                else
                {
                    // No target specified, use selected range
                    return await GetSelectedRangeAsync();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error resolving target range '{targetRangeSpec}': {ex.Message}");
                // Final fallback: try to get selected range
                try
                {
                    return await GetSelectedRangeAsync();
                }
                catch
                {
                    // Ultimate fallback: use A1
                    return _excelApp.ActiveSheet.Range["A1"];
                }
            }
        }
        
        /// <summary>
        /// Gets data from a range
        /// </summary>
        /// <param name="range">The range to get data from</param>
        /// <returns>A DataTable containing the range data</returns>
        public async Task<DataTable> GetRangeDataAsync(Excel.Range range)
        {
            return await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    
                    var data = new DataTable();
                    
                    // Get the values from the range
                    object rangeValue = range.Value2;
                    if (rangeValue == null) return data;
                    
                    int rows = range.Rows.Count;
                    int cols = range.Columns.Count;
                    
                    // Create columns
                    for (int c = 1; c <= cols; c++)
                    {
                        data.Columns.Add($"Column{c}");
                    }
                    
                    // Handle single cell vs multi-cell range
                    if (rows == 1 && cols == 1)
                    {
                        // Single cell
                        var row = data.NewRow();
                        row[0] = rangeValue ?? DBNull.Value;
                        data.Rows.Add(row);
                    }
                    else
                    {
                        // Multi-cell range
                        object[,] values = rangeValue as object[,];
                        if (values != null)
                        {
                            for (int r = 1; r <= rows; r++)
                            {
                                var row = data.NewRow();
                                for (int c = 1; c <= cols; c++)
                                {
                                    row[c - 1] = values[r, c] ?? DBNull.Value;
                                }
                                data.Rows.Add(row);
                            }
                        }
                    }
                    
                    return data;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to get range data", ex);
                }
            });
        }
        
        /// <summary>
        /// Sets a cell value
        /// </summary>
        /// <param name="range">The range to set</param>
        /// <param name="value">The value to set</param>
        public async Task SetCellValueAsync(Excel.Range range, object value)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    range.Value2 = value;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to set cell value", ex);
                }
            });
        }

        /// <summary>
        /// Clears the content of a range
        /// </summary>
        /// <param name="range">The range to clear</param>
        public async Task ClearRangeContentAsync(Excel.Range range)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    range.ClearContents();
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to clear range content", ex);
                }
            });
        }

        /// <summary>
        /// Sets the background color of a range
        /// </summary>
        /// <param name="range">The range to format</param>
        /// <param name="color">The color (hex format like #FF0000 or color name)</param>
        public async Task SetBackgroundColorAsync(Excel.Range range, string color)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    if (string.IsNullOrEmpty(color)) throw new ArgumentNullException(nameof(color));
                    
                    // Convert color string to RGB
                    System.Drawing.Color rgbColor = ParseColor(color);
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(rgbColor);
                }
                catch (Exception ex)
                {
                    throw new AiOperationException($"Failed to set background color '{color}'", ex);
                }
            });
        }

        /// <summary>
        /// Sets the font color of a range
        /// </summary>
        /// <param name="range">The range to format</param>
        /// <param name="color">The color (hex format like #FF0000 or color name)</param>
        public async Task SetFontColorAsync(Excel.Range range, string color)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    if (string.IsNullOrEmpty(color)) throw new ArgumentNullException(nameof(color));
                    
                    System.Drawing.Color rgbColor = ParseColor(color);
                    range.Font.Color = System.Drawing.ColorTranslator.ToOle(rgbColor);
                }
                catch (Exception ex)
                {
                    throw new AiOperationException($"Failed to set font color '{color}'", ex);
                }
            });
        }

        /// <summary>
        /// Sets the font style of a range
        /// </summary>
        /// <param name="range">The range to format</param>
        /// <param name="bold">Whether to make text bold</param>
        /// <param name="italic">Whether to make text italic</param>
        /// <param name="underline">Whether to underline text</param>
        /// <param name="fontSize">Font size (optional)</param>
        /// <param name="fontName">Font name (optional)</param>
        public async Task SetFontStyleAsync(Excel.Range range, bool? bold = null, bool? italic = null, 
            bool? underline = null, int? fontSize = null, string fontName = null)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    
                    if (bold.HasValue) range.Font.Bold = bold.Value;
                    if (italic.HasValue) range.Font.Italic = italic.Value;
                    if (underline.HasValue) range.Font.Underline = underline.Value;
                    if (fontSize.HasValue) range.Font.Size = fontSize.Value;
                    if (!string.IsNullOrEmpty(fontName)) range.Font.Name = fontName;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to set font style", ex);
                }
            });
        }

        /// <summary>
        /// Parses a color string to System.Drawing.Color
        /// </summary>
        /// <param name="colorString">Color string (hex like #FF0000 or name like "red")</param>
        /// <returns>System.Drawing.Color</returns>
        private System.Drawing.Color ParseColor(string colorString)
        {
            if (string.IsNullOrEmpty(colorString))
                return System.Drawing.Color.Black;

            colorString = colorString.Trim();

            // Handle hex colors
            if (colorString.StartsWith("#"))
            {
                return System.Drawing.ColorTranslator.FromHtml(colorString);
            }

            // Handle named colors
            try
            {
                return System.Drawing.Color.FromName(colorString);
            }
            catch
            {
                // Common color mappings
                switch (colorString.ToLower())
                {
                    case "red": return System.Drawing.Color.Red;
                    case "green": return System.Drawing.Color.Green;
                    case "blue": return System.Drawing.Color.Blue;
                    case "yellow": return System.Drawing.Color.Yellow;
                    case "orange": return System.Drawing.Color.Orange;
                    case "purple": return System.Drawing.Color.Purple;
                    case "pink": return System.Drawing.Color.Pink;
                    case "gray": case "grey": return System.Drawing.Color.Gray;
                    case "black": return System.Drawing.Color.Black;
                    case "white": return System.Drawing.Color.White;
                    default: return System.Drawing.Color.Black;
                }
            }
        }
        
        /// <summary>
        /// Applies a formula to a range
        /// </summary>
        /// <param name="range">The range to apply the formula to</param>
        /// <param name="formula">The formula to apply</param>
        public async Task ApplyFormulaAsync(Excel.Range range, string formula)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    if (string.IsNullOrEmpty(formula)) throw new ArgumentNullException(nameof(formula));
                    
                    // Validate and improve formula if needed
                    string validatedFormula = ValidateAndImproveFormula(formula, range);
                    
                    System.Diagnostics.Debug.WriteLine($"Applying formula: {validatedFormula} to range: {range.Address}");
                    
                    range.Formula = validatedFormula;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Formula application failed: {ex.Message}");
                    throw new AiOperationException($"Failed to apply formula '{formula}': {ex.Message}", ex);
                }
            });
        }

        /// <summary>
        /// Validates and improves a formula before applying it
        /// </summary>
        /// <param name="formula">The original formula</param>
        /// <param name="targetRange">The target range</param>
        /// <returns>Validated and improved formula</returns>
        private string ValidateAndImproveFormula(string formula, Excel.Range targetRange)
        {
            if (string.IsNullOrEmpty(formula))
                return formula;

            // Ensure formula starts with =
            if (!formula.StartsWith("="))
                formula = "=" + formula;

            // Handle common formula improvements
            if (formula.Equals("=SUM()", StringComparison.OrdinalIgnoreCase))
            {
                // Try to intelligently determine SUM range based on context
                string suggestedRange = GetSuggestedSumRange(targetRange);
                if (!string.IsNullOrEmpty(suggestedRange))
                {
                    formula = $"=SUM({suggestedRange})";
                    System.Diagnostics.Debug.WriteLine($"Improved SUM formula: {formula}");
                }
                else
                {
                    // Default to a simple range above the current cell
                    string cellAddress = targetRange.Address.Replace("$", "");
                    var match = System.Text.RegularExpressions.Regex.Match(cellAddress, @"([A-Z]+)(\d+)");
                    if (match.Success)
                    {
                        string column = match.Groups[1].Value;
                        int row = int.Parse(match.Groups[2].Value);
                        if (row > 1)
                        {
                            formula = $"=SUM({column}1:{column}{row - 1})";
                            System.Diagnostics.Debug.WriteLine($"Default SUM formula: {formula}");
                        }
                    }
                }
            }

            return formula;
        }

        /// <summary>
        /// Gets a suggested range for SUM formula based on nearby data
        /// </summary>
        /// <param name="targetRange">The target range</param>
        /// <returns>Suggested range string or null</returns>
        private string GetSuggestedSumRange(Excel.Range targetRange)
        {
            try
            {
                // Look for data above the current cell
                var currentRow = targetRange.Row;
                var currentColumn = targetRange.Column;
                var worksheet = targetRange.Worksheet;

                // Check if there's data in the column above
                int dataStartRow = -1;
                int dataEndRow = -1;

                for (int row = currentRow - 1; row >= 1; row--)
                {
                    var cell = worksheet.Cells[row, currentColumn];
                    if (cell.Value2 != null && IsNumeric(cell.Value2))
                    {
                        if (dataEndRow == -1) dataEndRow = row;
                        dataStartRow = row;
                    }
                    else if (dataEndRow != -1)
                    {
                        break; // Found gap, stop looking
                    }
                }

                if (dataStartRow != -1 && dataEndRow != -1)
                {
                    string columnLetter = GetColumnLetter(currentColumn);
                    return $"{columnLetter}{dataStartRow}:{columnLetter}{dataEndRow}";
                }

                // If no data above, check to the left
                for (int col = currentColumn - 1; col >= 1; col--)
                {
                    var cell = worksheet.Cells[currentRow, col];
                    if (cell.Value2 != null && IsNumeric(cell.Value2))
                    {
                        if (dataEndRow == -1) dataEndRow = col;
                        dataStartRow = col;
                    }
                    else if (dataEndRow != -1)
                    {
                        break;
                    }
                }

                if (dataStartRow != -1 && dataEndRow != -1)
                {
                    return $"{GetColumnLetter(dataStartRow)}{currentRow}:{GetColumnLetter(dataEndRow)}{currentRow}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error suggesting SUM range: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Checks if a value is numeric
        /// </summary>
        private bool IsNumeric(object value)
        {
            return value != null && (value is double || value is int || value is float || value is decimal ||
                   double.TryParse(value.ToString(), out _));
        }

        /// <summary>
        /// Converts column number to letter
        /// </summary>
        private string GetColumnLetter(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }
        
        /// <summary>
        /// Sets the number format for a range
        /// </summary>
        /// <param name="range">The range to format</param>
        /// <param name="format">The number format to apply</param>
        public async Task SetNumberFormatAsync(Excel.Range range, string format)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    if (string.IsNullOrEmpty(format)) throw new ArgumentNullException(nameof(format));
                    
                    range.NumberFormat = format;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to set number format", ex);
                }
            });
        }
        
        /// <summary>
        /// Copies data from source range to target range
        /// </summary>
        /// <param name="sourceRange">The source range to copy from</param>
        /// <param name="targetRange">The target range to copy to</param>
        /// <param name="copyType">The type of copy operation</param>
        public async Task CopyRangeAsync(Excel.Range sourceRange, Excel.Range targetRange, CopyType copyType)
        {
            await Task.Run(() => {
                try
                {
                    if (sourceRange == null) throw new ArgumentNullException(nameof(sourceRange));
                    if (targetRange == null) throw new ArgumentNullException(nameof(targetRange));
                    
                    // Copy the source range
                    sourceRange.Copy();
                    
                    // Paste based on copy type
                    switch (copyType)
                    {
                        case CopyType.All:
                            targetRange.PasteSpecial(Excel.XlPasteType.xlPasteAll);
                            break;
                        case CopyType.Values:
                            targetRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                            break;
                        case CopyType.Formulas:
                            targetRange.PasteSpecial(Excel.XlPasteType.xlPasteFormulas);
                            break;
                        case CopyType.Formats:
                            targetRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                            break;
                        default:
                            targetRange.PasteSpecial(Excel.XlPasteType.xlPasteAll);
                            break;
                    }
                    
                    // Clear clipboard
                    _excelApp.CutCopyMode = false;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to copy range", ex);
                }
            });
        }
        
        /// <summary>
        /// Finds and replaces text in a range
        /// </summary>
        /// <param name="range">The range to search in</param>
        /// <param name="findWhat">The text to find</param>
        /// <param name="replaceWith">The text to replace with</param>
        /// <param name="matchCase">Whether to match case</param>
        /// <param name="matchEntireCell">Whether to match entire cell</param>
        /// <returns>Number of replacements made</returns>
        public async Task<int> FindAndReplaceAsync(Excel.Range range, string findWhat, string replaceWith, bool matchCase = false, bool matchEntireCell = false)
        {
            return await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    if (string.IsNullOrEmpty(findWhat)) throw new ArgumentNullException(nameof(findWhat));
                    
                    int replaceCount = 0;
                    Excel.Range foundRange = range.Find(
                        What: findWhat,
                        LookIn: Excel.XlFindLookIn.xlValues,
                        LookAt: matchEntireCell ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart,
                        SearchOrder: Excel.XlSearchOrder.xlByRows,
                        SearchDirection: Excel.XlSearchDirection.xlNext,
                        MatchCase: matchCase);
                    
                    if (foundRange != null)
                    {
                        string firstAddress = foundRange.Address;
                        do
                        {
                            foundRange.Value = replaceWith;
                            replaceCount++;
                            foundRange = range.FindNext(foundRange);
                        }
                        while (foundRange != null && foundRange.Address != firstAddress);
                    }
                    
                    return replaceCount;
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to find and replace", ex);
                }
            });
        }
        
        /// <summary>
        /// Sorts a range of data
        /// </summary>
        /// <param name="range">The range to sort</param>
        /// <param name="sortColumn">The column to sort by (1-based)</param>
        /// <param name="ascending">Whether to sort in ascending order</param>
        /// <param name="hasHeaders">Whether the range has headers</param>
        public async Task SortRangeAsync(Excel.Range range, int sortColumn, bool ascending = true, bool hasHeaders = true)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    
                    Excel.Range sortKey = range.Columns[sortColumn];
                    range.Sort(
                        Key1: sortKey,
                        Order1: ascending ? Excel.XlSortOrder.xlAscending : Excel.XlSortOrder.xlDescending,
                        Header: hasHeaders ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to sort range", ex);
                }
            });
        }
        
        /// <summary>
        /// Applies auto filter to a range
        /// </summary>
        /// <param name="range">The range to filter</param>
        /// <param name="column">The column to filter by (1-based)</param>
        /// <param name="criteria">The filter criteria</param>
        public async Task ApplyAutoFilterAsync(Excel.Range range, int column, string criteria)
        {
            await Task.Run(() => {
                try
                {
                    if (range == null) throw new ArgumentNullException(nameof(range));
                    
                    // Clear existing filters
                    if (range.Worksheet.AutoFilterMode)
                    {
                        range.Worksheet.AutoFilterMode = false;
                    }
                    
                    // Apply auto filter
                    range.AutoFilter(Field: column, Criteria1: criteria);
                }
                catch (Exception ex)
                {
                    throw new AiOperationException("Failed to apply auto filter", ex);
                }
            });
        }
        
        /// <summary>
        /// Executes a set of instructions
        /// </summary>
        /// <param name="instructionSet">The instructions to execute</param>
        /// <param name="dryRun">If true, only simulate the execution</param>
        /// <returns>A summary of the execution</returns>
        public async Task<string> ExecuteInstructionsAsync(InstructionSet instructionSet, bool dryRun = false)
        {
            if (instructionSet == null || instructionSet.Instructions == null || instructionSet.Instructions.Count == 0)
            {
                return "No instructions to execute";
            }
            
            var results = new List<string>();
            
            foreach (var instruction in instructionSet.Instructions)
            {
                string result = await ExecuteInstructionAsync(instruction, dryRun);
                results.Add(result);
            }
            
            return string.Join("\n", results);
        }
        
        private async Task<string> ExecuteInstructionAsync(Instruction instruction, bool dryRun)
        {
            if (dryRun)
            {
                return $"[Simulation] Would execute: {instruction.Description}";
            }
            
            try
            {
                Excel.Range targetRange = null;
                
                // Parse target range with intelligent handling
                targetRange = await GetTargetRangeAsync(instruction.TargetRange);
                
                System.Diagnostics.Debug.WriteLine($"Target range resolved to: {targetRange.Address} for instruction: {instruction.Description}");
                
                switch (instruction.Type)
                {
                    case InstructionType.SetCellValue:
                        if (instruction.Parameters.TryGetValue("value", out object value))
                        {
                            await SetCellValueAsync(targetRange, value);
                            return $"✓ Set value '{value}' in cell {instruction.TargetRange ?? targetRange.Address}";
                        }
                        else
                        {
                            return $"✗ SetCellValue instruction missing 'value' parameter";
                        }
                        
                    case InstructionType.ApplyFormula:
                        // Try both 'formula' and 'value' parameters (AI sometimes uses different names)
                        string formula = null;
                        if (instruction.Parameters.TryGetValue("formula", out object formulaObj) && formulaObj is string f1)
                        {
                            formula = f1;
                        }
                        else if (instruction.Parameters.TryGetValue("value", out object valueObj) && valueObj is string f2)
                        {
                            formula = f2;
                        }
                        
                        if (!string.IsNullOrEmpty(formula))
                        {
                            await ApplyFormulaAsync(targetRange, formula);
                            return $"✓ Applied formula '{formula}' to {instruction.TargetRange ?? targetRange.Address}";
                        }
                        else
                        {
                            return $"✗ ApplyFormula instruction missing 'formula' or 'value' parameter";
                        }
                        
                    case InstructionType.SetCellFormat:
                        if (instruction.Parameters.TryGetValue("format", out object formatObj) && formatObj is string format)
                        {
                            await SetNumberFormatAsync(targetRange, format);
                            return $"✓ Set format '{format}' in {instruction.TargetRange ?? targetRange.Address}";
                        }
                        else
                        {
                            return $"✗ SetCellFormat instruction missing 'format' parameter";
                        }

                    case InstructionType.SetCellStyle:
                        return await HandleSetCellStyleAsync(targetRange, instruction);

                    case InstructionType.ClearContent:
                        await ClearRangeContentAsync(targetRange);
                        return $"✓ Cleared content in {instruction.TargetRange ?? targetRange.Address}";
                        
                    default:
                        return $"✗ Unsupported instruction type: {instruction.Type}";
                }
            }
            catch (Exception ex)
            {
                throw new AiOperationException($"Failed to execute instruction: {instruction.Description}", ex);
            }
        }

        /// <summary>
        /// Handles SetCellStyle instruction with various style parameters
        /// </summary>
        /// <param name="targetRange">The target range</param>
        /// <param name="instruction">The instruction with style parameters</param>
        /// <returns>Result message</returns>
        private async Task<string> HandleSetCellStyleAsync(Excel.Range targetRange, Instruction instruction)
        {
            var results = new List<string>();

            // Handle background color
            if (instruction.Parameters.TryGetValue("backgroundColor", out object bgColorObj) && bgColorObj is string bgColor)
            {
                await SetBackgroundColorAsync(targetRange, bgColor);
                results.Add($"background color to {bgColor}");
            }

            // Handle font color
            if (instruction.Parameters.TryGetValue("fontColor", out object fontColorObj) && fontColorObj is string fontColor)
            {
                await SetFontColorAsync(targetRange, fontColor);
                results.Add($"font color to {fontColor}");
            }

            // Handle font style
            bool? bold = null, italic = null, underline = null;
            int? fontSize = null;
            string fontName = null;

            if (instruction.Parameters.TryGetValue("bold", out object boldObj) && boldObj is bool b)
                bold = b;
            if (instruction.Parameters.TryGetValue("italic", out object italicObj) && italicObj is bool i)
                italic = i;
            if (instruction.Parameters.TryGetValue("underline", out object underlineObj) && underlineObj is bool u)
                underline = u;
            if (instruction.Parameters.TryGetValue("fontSize", out object fontSizeObj) && fontSizeObj is int fs)
                fontSize = fs;
            if (instruction.Parameters.TryGetValue("fontName", out object fontNameObj) && fontNameObj is string fn)
                fontName = fn;

            if (bold.HasValue || italic.HasValue || underline.HasValue || fontSize.HasValue || !string.IsNullOrEmpty(fontName))
            {
                await SetFontStyleAsync(targetRange, bold, italic, underline, fontSize, fontName);
                var styleChanges = new List<string>();
                if (bold.HasValue) styleChanges.Add($"bold: {bold.Value}");
                if (italic.HasValue) styleChanges.Add($"italic: {italic.Value}");
                if (underline.HasValue) styleChanges.Add($"underline: {underline.Value}");
                if (fontSize.HasValue) styleChanges.Add($"font size: {fontSize.Value}");
                if (!string.IsNullOrEmpty(fontName)) styleChanges.Add($"font: {fontName}");
                
                results.AddRange(styleChanges);
            }

            return $"✓ Set {string.Join(", ", results)} in {targetRange.Address}";
        }
    }
}