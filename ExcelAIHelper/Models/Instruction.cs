using System.Collections.Generic;

namespace ExcelAIHelper.Models
{
    /// <summary>
    /// Types of instructions that can be executed
    /// </summary>
    public enum InstructionType
    {
        /// <summary>
        /// Unknown instruction type
        /// </summary>
        Unknown,
        
        /// <summary>
        /// Set cell value
        /// </summary>
        SetCellValue,
        
        /// <summary>
        /// Set cell format
        /// </summary>
        SetCellFormat,
        
        /// <summary>
        /// Apply formula
        /// </summary>
        ApplyFormula,
        
        /// <summary>
        /// Set cell style
        /// </summary>
        SetCellStyle,
        
        /// <summary>
        /// Insert rows
        /// </summary>
        InsertRows,
        
        /// <summary>
        /// Insert columns
        /// </summary>
        InsertColumns,
        
        /// <summary>
        /// Delete rows
        /// </summary>
        DeleteRows,
        
        /// <summary>
        /// Delete columns
        /// </summary>
        DeleteColumns,
        
        /// <summary>
        /// Sort data
        /// </summary>
        SortData,
        
        /// <summary>
        /// Filter data
        /// </summary>
        FilterData,
        
        /// <summary>
        /// Create chart
        /// </summary>
        CreateChart,
        
        /// <summary>
        /// Apply conditional formatting
        /// </summary>
        ApplyConditionalFormatting,
        
        /// <summary>
        /// Clear cell content
        /// </summary>
        ClearContent
    }

    /// <summary>
    /// Represents a single instruction to be executed
    /// </summary>
    public class Instruction
    {
        /// <summary>
        /// The type of instruction
        /// </summary>
        public InstructionType Type { get; set; }
        
        /// <summary>
        /// Human-readable description of the instruction
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// The target range for the instruction (e.g., "A1:B10")
        /// </summary>
        public string TargetRange { get; set; }
        
        /// <summary>
        /// Parameters for the instruction
        /// </summary>
        public Dictionary<string, object> Parameters { get; set; }
        
        /// <summary>
        /// Whether this instruction requires user confirmation
        /// </summary>
        public bool RequiresConfirmation { get; set; }
    }

    /// <summary>
    /// A set of instructions to be executed
    /// </summary>
    public class InstructionSet
    {
        /// <summary>
        /// The list of instructions
        /// </summary>
        public List<Instruction> Instructions { get; set; }
        
        /// <summary>
        /// Human-readable summary of what these instructions will do
        /// </summary>
        public string Summary { get; set; }
    }
}