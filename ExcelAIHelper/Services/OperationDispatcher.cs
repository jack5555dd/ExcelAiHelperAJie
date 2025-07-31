using System;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Models;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Dispatches operations to be executed
    /// </summary>
    public class OperationDispatcher
    {
        private readonly DeepSeekClient _aiClient;
        private readonly PromptBuilder _promptBuilder;
        private readonly InstructionParser _instructionParser;
        private readonly ExcelOperationEngine _operationEngine;
        
        /// <summary>
        /// Creates a new instance of OperationDispatcher
        /// </summary>
        /// <param name="aiClient">The AI client</param>
        /// <param name="promptBuilder">The prompt builder</param>
        /// <param name="instructionParser">The instruction parser</param>
        /// <param name="operationEngine">The operation engine</param>
        public OperationDispatcher(
            DeepSeekClient aiClient,
            PromptBuilder promptBuilder,
            InstructionParser instructionParser,
            ExcelOperationEngine operationEngine)
        {
            _aiClient = aiClient ?? throw new ArgumentNullException(nameof(aiClient));
            _promptBuilder = promptBuilder ?? throw new ArgumentNullException(nameof(promptBuilder));
            _instructionParser = instructionParser ?? throw new ArgumentNullException(nameof(instructionParser));
            _operationEngine = operationEngine ?? throw new ArgumentNullException(nameof(operationEngine));
        }
        
        /// <summary>
        /// Applies a user request to Excel
        /// </summary>
        /// <param name="userRequest">The user's request</param>
        /// <param name="dryRun">If true, only simulate the execution</param>
        /// <returns>A summary of the execution</returns>
        public async Task<string> ApplyAsync(string userRequest, bool dryRun = false)
        {
            try
            {
                // Build the prompt
                var (systemPrompt, userPrompt) = await _promptBuilder.BuildAsync(userRequest);
                
                // Get AI response
                string aiResponse = await _aiClient.AskAsync(userPrompt, systemPrompt);
                
                // Parse the response into instructions
                InstructionSet instructionSet = await _instructionParser.ParseAsync(aiResponse);
                
                // Execute or simulate the instructions
                string result = await _operationEngine.ExecuteInstructionsAsync(instructionSet, dryRun);
                
                if (dryRun)
                {
                    return $"Preview of operations:\n{result}\n\nConfirm to execute.";
                }
                else
                {
                    return $"Executed operations:\n{result}";
                }
            }
            catch (Exception ex)
            {
                throw new AiOperationException("Failed to apply user request", ex);
            }
        }
    }
}