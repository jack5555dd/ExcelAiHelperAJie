using System;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Manages context information for AI operations
    /// </summary>
    public class ContextManager
    {
        private readonly Excel.Application _excelApp;
        private readonly ExcelOperationEngine _operationEngine;
        
        /// <summary>
        /// Creates a new instance of ContextManager
        /// </summary>
        /// <param name="excelApp">The Excel application instance</param>
        /// <param name="operationEngine">The operation engine</param>
        public ContextManager(Excel.Application excelApp, ExcelOperationEngine operationEngine)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            _operationEngine = operationEngine ?? throw new ArgumentNullException(nameof(operationEngine));
        }
        
        /// <summary>
        /// Gets the current Excel context
        /// </summary>
        /// <returns>The current context</returns>
        public async Task<ExcelContext> GetCurrentContextAsync()
        {
            try
            {
                var context = new ExcelContext();
                
                // Get active worksheet info
                context.CurrentWorksheet = new WorksheetInfo
                {
                    Name = _excelApp.ActiveSheet.Name,
                    Index = _excelApp.ActiveSheet.Index
                };
                
                // Get selected range
                var selectedRange = await _operationEngine.GetSelectedRangeAsync();
                if (selectedRange != null)
                {
                    context.SelectedRange = new RangeInfo
                    {
                        Address = selectedRange.Address,
                        FirstRow = selectedRange.Row,
                        FirstColumn = selectedRange.Column,
                        RowCount = selectedRange.Rows.Count,
                        ColumnCount = selectedRange.Columns.Count
                    };
                    
                    // Get data from selected range
                    var data = await _operationEngine.GetRangeDataAsync(selectedRange);
                    context.SelectedData = data;
                }
                
                return context;
            }
            catch (Exception ex)
            {
                throw new AiOperationException("Failed to get current context", ex);
            }
        }
        
        /// <summary>
        /// Gets a description of the current context
        /// </summary>
        /// <returns>A string describing the current context</returns>
        public async Task<string> GetContextDescriptionAsync()
        {
            var context = await GetCurrentContextAsync();
            
            var description = $"Current worksheet: {context.CurrentWorksheet.Name}\n";
            
            if (context.SelectedRange != null)
            {
                description += $"Selected range: {context.SelectedRange.Address}\n";
                description += $"Range dimensions: {context.SelectedRange.RowCount} rows x {context.SelectedRange.ColumnCount} columns\n";
                
                if (context.SelectedData != null && context.SelectedData.Rows.Count > 0)
                {
                    description += $"Sample data: {context.SelectedData.Rows.Count} rows of data\n";
                    
                    // Add sample of data (first few rows)
                    int rowsToShow = Math.Min(5, context.SelectedData.Rows.Count);
                    for (int i = 0; i < rowsToShow; i++)
                    {
                        description += "Row " + (i + 1) + ": ";
                        for (int j = 0; j < context.SelectedData.Columns.Count; j++)
                        {
                            description += context.SelectedData.Rows[i][j]?.ToString() + " | ";
                        }
                        description += "\n";
                    }
                    
                    if (context.SelectedData.Rows.Count > rowsToShow)
                    {
                        description += $"... and {context.SelectedData.Rows.Count - rowsToShow} more rows\n";
                    }
                }
            }
            else
            {
                description += "No range selected\n";
            }
            
            return description;
        }
    }
}