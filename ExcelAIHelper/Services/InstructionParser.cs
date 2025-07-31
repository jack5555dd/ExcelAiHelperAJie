using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Parses AI responses into executable instructions using basic JSON validation
    /// </summary>
    public class InstructionParser
    {
        /// <summary>
        /// Creates a new instance of InstructionParser
        /// </summary>
        public InstructionParser()
        {
        }
        
        /// <summary>
        /// Parses the AI response into an instruction set using strict protocol validation
        /// </summary>
        /// <param name="aiResponse">The response from the AI</param>
        /// <returns>A parsed instruction set</returns>
        public async Task<InstructionSet> ParseAsync(string aiResponse)
        {
            try
            {
                // Step 1: Validate against JSON Command Protocol
                var validationResult = _validator.Validate(aiResponse);
                
                if (!validationResult.IsValid)
                {
                    // Protocol violation - throw AiFormatException
                    throw AiFormatException.CreateProtocolViolation(aiResponse, validationResult.ErrorMessage);
                }
                
                // Step 2: Convert validated JSON to InstructionSet
                var instructionSet = ConvertToInstructionSet(validationResult.ValidatedJson);
                
                return await Task.FromResult(instructionSet);
            }
            catch (AiFormatException)
            {
                // Re-throw format exceptions as-is
                throw;
            }
            catch (Exception ex)
            {
                // Wrap other exceptions as format exceptions
                throw new AiFormatException("Failed to parse AI response", aiResponse, ex);
            }
        }

        /// <summary>
        /// Converts validated JSON to InstructionSet
        /// </summary>
        /// <param name="validatedJson">JSON that has passed protocol validation</param>
        /// <returns>InstructionSet object</returns>
        private InstructionSet ConvertToInstructionSet(JObject validatedJson)
        {
            try
            {
                var instructionSet = new InstructionSet
                {
                    Summary = validatedJson["summary"]?.ToString() ?? "",
                    Instructions = new List<Instruction>()
                };
                
                var commands = validatedJson["commands"] as JArray;
                if (commands != null)
                {
                    foreach (var command in commands)
                    {
                        var instruction = ConvertCommandToInstruction(command as JObject);
                        if (instruction != null)
                        {
                            instructionSet.Instructions.Add(instruction);
                        }
                    }
                }
                
                return instructionSet;
            }
            catch (Exception ex)
            {
                throw new AiFormatException("Failed to convert validated JSON to InstructionSet", validatedJson.ToString(), ex);
            }
        }
        
        /// <summary>
        /// Converts a single command to an Instruction
        /// </summary>
        /// <param name="command">Command JSON object</param>
        /// <returns>Instruction object</returns>
        private Instruction ConvertCommandToInstruction(JObject command)
        {
            if (command == null) return null;
            
            var function = command["function"]?.ToString();
            var arguments = command["arguments"] as JObject;
            var description = command["description"]?.ToString() ?? "";
            
            // Map function names to InstructionType
            var instructionType = MapFunctionToInstructionType(function);
            
            // Convert arguments to parameters dictionary
            var parameters = new Dictionary<string, object>();
            if (arguments != null)
            {
                foreach (var prop in arguments.Properties())
                {
                    parameters[prop.Name] = prop.Value?.ToObject<object>();
                }
            }
            
            // Extract target range from arguments
            string targetRange = null;
            if (parameters.ContainsKey("range"))
            {
                targetRange = parameters["range"]?.ToString();
            }
            else if (parameters.ContainsKey("position"))
            {
                targetRange = parameters["position"]?.ToString();
            }
            
            return new Instruction
            {
                Type = instructionType,
                Description = description,
                TargetRange = targetRange,
                Parameters = parameters,
                RequiresConfirmation = ShouldRequireConfirmation(instructionType)
            };
        }
        
        /// <summary>
        /// Maps function name to InstructionType
        /// </summary>
        /// <param name="function">Function name from JSON command</param>
        /// <returns>Corresponding InstructionType</returns>
        private InstructionType MapFunctionToInstructionType(string function)
        {
            switch (function)
            {
                case "setCellValue":
                    return InstructionType.SetCellValue;
                case "applyCellFormula":
                    return InstructionType.ApplyFormula;
                case "setCellStyle":
                    return InstructionType.SetCellStyle;
                case "setCellFormat":
                    return InstructionType.SetCellFormat;
                case "clearCellContent":
                    return InstructionType.ClearContent;
                case "insertRows":
                    return InstructionType.InsertRows;
                case "insertColumns":
                    return InstructionType.InsertColumns;
                case "deleteRows":
                    return InstructionType.DeleteRows;
                case "deleteColumns":
                    return InstructionType.DeleteColumns;
                case "sortRange":
                    return InstructionType.SortData;
                case "filterRange":
                    return InstructionType.FilterData;
                case "createChart":
                    return InstructionType.CreateChart;
                default:
                    return InstructionType.Unknown;
            }
        }
        
        /// <summary>
        /// Determines if an instruction type should require confirmation
        /// </summary>
        /// <param name="type">Instruction type</param>
        /// <returns>True if confirmation is required</returns>
        private bool ShouldRequireConfirmation(InstructionType type)
        {
            // Destructive operations require confirmation
            switch (type)
            {
                case InstructionType.DeleteRows:
                case InstructionType.DeleteColumns:
                case InstructionType.ClearContent:
                    return true;
                default:
                    return false;
            }
        }
    }
}