using System.Collections.Generic;
using System.Data;

namespace ExcelAIHelper.Models
{
    /// <summary>
    /// Information about a worksheet
    /// </summary>
    public class WorksheetInfo
    {
        /// <summary>
        /// The name of the worksheet
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// The index of the worksheet
        /// </summary>
        public int Index { get; set; }
    }
    
    /// <summary>
    /// Information about a range
    /// </summary>
    public class RangeInfo
    {
        /// <summary>
        /// The address of the range (e.g., "A1:B10")
        /// </summary>
        public string Address { get; set; }
        
        /// <summary>
        /// The first row of the range (1-based)
        /// </summary>
        public int FirstRow { get; set; }
        
        /// <summary>
        /// The first column of the range (1-based)
        /// </summary>
        public int FirstColumn { get; set; }
        
        /// <summary>
        /// The number of rows in the range
        /// </summary>
        public int RowCount { get; set; }
        
        /// <summary>
        /// The number of columns in the range
        /// </summary>
        public int ColumnCount { get; set; }
    }
    
    /// <summary>
    /// Context information for Excel operations
    /// </summary>
    public class ExcelContext
    {
        /// <summary>
        /// Information about the current worksheet
        /// </summary>
        public WorksheetInfo CurrentWorksheet { get; set; }
        
        /// <summary>
        /// Information about the selected range
        /// </summary>
        public RangeInfo SelectedRange { get; set; }
        
        /// <summary>
        /// Data from the selected range
        /// </summary>
        public DataTable SelectedData { get; set; }
        
        /// <summary>
        /// Variables for the current session
        /// </summary>
        public Dictionary<string, object> SessionVariables { get; set; } = new Dictionary<string, object>();
    }
}