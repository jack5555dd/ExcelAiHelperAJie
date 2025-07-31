using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;
using Newtonsoft.Json;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Parses AI responses into executable instructions
    /// </summary>
    public class InstructionParser
    {
        /// <summary>
        /// Parses the AI response into an instruction set
        /// </summary>
        /// <param name="aiResponse">The response from the AI</param>
        /// <returns>A parsed instruction set</returns>
        public async Task<InstructionSet> ParseAsync(string aiResponse)
        {
            try
            {
                // Try to parse as JSON first
                if (TryParseJson(aiResponse, out InstructionSet instructionSet))
                {
                    return await Task.FromResult(instructionSet);
                }

                // If not JSON, try to parse as natural language
                return await Task.FromResult(ParseNaturalLanguage(aiResponse));
            }
            catch (Exception ex)
            {
                throw new AiOperationException("Failed to parse AI response", ex);
            }
        }

        private bool TryParseJson(string response, out InstructionSet instructionSet)
        {
            try
            {
                // Clean the response by removing markdown code blocks
                string cleanedResponse = CleanJsonResponse(response);
                
                System.Diagnostics.Debug.WriteLine($"Original response: {response}");
                System.Diagnostics.Debug.WriteLine($"Cleaned response: {cleanedResponse}");
                
                var settings = new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    MissingMemberHandling = MissingMemberHandling.Ignore
                };
                instructionSet = JsonConvert.DeserializeObject<InstructionSet>(cleanedResponse, settings);
                return instructionSet != null && instructionSet.Instructions != null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"JSON parsing failed: {ex.Message}");
                instructionSet = null;
                return false;
            }
        }

        /// <summary>
        /// Cleans the AI response by removing markdown code blocks and extra formatting
        /// </summary>
        /// <param name="response">The raw AI response</param>
        /// <returns>Clean JSON string</returns>
        private string CleanJsonResponse(string response)
        {
            if (string.IsNullOrEmpty(response))
                return response;

            // Remove markdown code blocks (```json ... ```)
            string cleaned = response.Trim();
            
            // Check if response starts with ```json and ends with ```
            if (cleaned.StartsWith("```json", StringComparison.OrdinalIgnoreCase))
            {
                // Find the end of the opening ```json
                int startIndex = cleaned.IndexOf('\n');
                if (startIndex == -1) startIndex = 7; // length of "```json"
                else startIndex += 1; // skip the newline
                
                // Find the closing ```
                int endIndex = cleaned.LastIndexOf("```");
                if (endIndex > startIndex)
                {
                    cleaned = cleaned.Substring(startIndex, endIndex - startIndex).Trim();
                }
                else
                {
                    // No closing ```, just remove the opening
                    cleaned = cleaned.Substring(startIndex).Trim();
                }
            }
            else if (cleaned.StartsWith("```", StringComparison.OrdinalIgnoreCase))
            {
                // Generic code block
                int startIndex = cleaned.IndexOf('\n');
                if (startIndex != -1)
                {
                    int endIndex = cleaned.LastIndexOf("```");
                    if (endIndex > startIndex)
                    {
                        cleaned = cleaned.Substring(startIndex + 1, endIndex - startIndex - 1).Trim();
                    }
                    else
                    {
                        cleaned = cleaned.Substring(startIndex + 1).Trim();
                    }
                }
            }

            return cleaned;
        }

        private InstructionSet ParseNaturalLanguage(string response)
        {
            // For now, create a simple instruction set with a single text instruction
            // In a real implementation, this would use NLP techniques to extract operations
            return new InstructionSet
            {
                Instructions = new List<Instruction>
                {
                    new Instruction
                    {
                        Type = InstructionType.Unknown,
                        Description = response,
                        Parameters = new Dictionary<string, object>()
                    }
                }
            };
        }
    }
}