using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAIHelper.Exceptions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// JSON命令协议验证器（简化版本，不依赖Schema包）
    /// </summary>
    public class JsonCommandValidator
    {
        private static readonly Lazy<JsonCommandValidator> _instance = 
            new Lazy<JsonCommandValidator>(() => new JsonCommandValidator());
        
        /// <summary>
        /// 获取验证器单例实例
        /// </summary>
        public static JsonCommandValidator Instance => _instance.Value;
        
        /// <summary>
        /// 私有构造函数
        /// </summary>
        private JsonCommandValidator()
        {
        }
        
        /// <summary>
        /// 验证JSON命令是否符合协议（简化版本）
        /// </summary>
        /// <param name="jsonResponse">AI返回的JSON响应</param>
        /// <returns>验证结果</returns>
        public ValidationResult Validate(string jsonResponse)
        {
            if (string.IsNullOrWhiteSpace(jsonResponse))
            {
                return ValidationResult.Failure("响应内容为空");
            }
            
            try
            {
                // 1. 基础JSON格式验证
                JObject jsonObject;
                try
                {
                    jsonObject = JObject.Parse(jsonResponse);
                }
                catch (JsonReaderException ex)
                {
                    return ValidationResult.Failure($"JSON格式错误: {ex.Message}");
                }
                
                // 2. 基础结构验证
                var structureValidation = ValidateBasicStructure(jsonObject);
                if (!structureValidation.IsValid)
                {
                    return structureValidation;
                }
                
                // 3. 业务逻辑验证
                var businessValidation = ValidateBusinessRules(jsonObject);
                if (!businessValidation.IsValid)
                {
                    return businessValidation;
                }
                
                return ValidationResult.Success(jsonObject);
            }
            catch (Exception ex)
            {
                return ValidationResult.Failure($"验证过程异常: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 验证基础JSON结构
        /// </summary>
        /// <param name="jsonObject">JSON对象</param>
        /// <returns>验证结果</returns>
        private ValidationResult ValidateBasicStructure(JObject jsonObject)
        {
            // 检查必需字段
            if (jsonObject["version"] == null)
            {
                return ValidationResult.Failure("缺少必需字段: version");
            }
            
            if (jsonObject["commands"] == null)
            {
                return ValidationResult.Failure("缺少必需字段: commands");
            }
            
            var commands = jsonObject["commands"] as JArray;
            if (commands == null)
            {
                return ValidationResult.Failure("commands字段必须是数组");
            }
            
            if (commands.Count == 0)
            {
                return ValidationResult.Failure("commands数组不能为空");
            }
            
            // 验证每个命令的基础结构
            for (int i = 0; i < commands.Count; i++)
            {
                var command = commands[i] as JObject;
                if (command == null)
                {
                    return ValidationResult.Failure($"命令{i + 1}必须是对象");
                }
                
                if (command["function"] == null)
                {
                    return ValidationResult.Failure($"命令{i + 1}缺少function字段");
                }
                
                if (command["arguments"] == null)
                {
                    return ValidationResult.Failure($"命令{i + 1}缺少arguments字段");
                }
            }
            
            return ValidationResult.Success(jsonObject);
        }
        
        /// <summary>
        /// 验证业务规则
        /// </summary>
        /// <param name="jsonObject">已通过基础验证的JSON对象</param>
        /// <returns>业务验证结果</returns>
        private ValidationResult ValidateBusinessRules(JObject jsonObject)
        {
            try
            {
                var commands = jsonObject["commands"] as JArray;
                if (commands == null || commands.Count == 0)
                {
                    return ValidationResult.Failure("命令列表不能为空");
                }
                
                // 验证每个命令的业务规则
                for (int i = 0; i < commands.Count; i++)
                {
                    var command = commands[i] as JObject;
                    if (command == null) continue;
                    
                    var function = command["function"]?.ToString();
                    var arguments = command["arguments"] as JObject;
                    
                    var commandValidation = ValidateCommandBusinessRules(function, arguments, i);
                    if (!commandValidation.IsValid)
                    {
                        return commandValidation;
                    }
                }
                
                return ValidationResult.Success(jsonObject);
            }
            catch (Exception ex)
            {
                return ValidationResult.Failure($"业务规则验证异常: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 验证单个命令的业务规则
        /// </summary>
        /// <param name="function">函数名</param>
        /// <param name="arguments">参数对象</param>
        /// <param name="commandIndex">命令索引</param>
        /// <returns>验证结果</returns>
        private ValidationResult ValidateCommandBusinessRules(string function, JObject arguments, int commandIndex)
        {
            if (arguments == null)
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 参数不能为空");
            }
            
            switch (function)
            {
                case "setCellValue":
                    return ValidateSetCellValueRules(arguments, commandIndex);
                    
                case "applyCellFormula":
                    return ValidateApplyCellFormulaRules(arguments, commandIndex);
                    
                case "setCellStyle":
                    return ValidateSetCellStyleRules(arguments, commandIndex);
                    
                case "setCellFormat":
                    return ValidateSetCellFormatRules(arguments, commandIndex);
                    
                default:
                    return ValidationResult.Success();
            }
        }
        
        /// <summary>
        /// 验证setCellValue命令的业务规则
        /// </summary>
        private ValidationResult ValidateSetCellValueRules(JObject arguments, int commandIndex)
        {
            var range = arguments["range"]?.ToString();
            var value = arguments["value"];
            
            // 验证范围格式
            if (!IsValidCellRange(range))
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 无效的单元格范围 '{range}'");
            }
            
            // 验证值的合理性
            if (value == null)
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 值不能为null");
            }
            
            return ValidationResult.Success();
        }
        
        /// <summary>
        /// 验证applyCellFormula命令的业务规则
        /// </summary>
        private ValidationResult ValidateApplyCellFormulaRules(JObject arguments, int commandIndex)
        {
            var range = arguments["range"]?.ToString();
            var formula = arguments["formula"]?.ToString();
            
            if (!IsValidCellRange(range))
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 无效的单元格范围 '{range}'");
            }
            
            if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("="))
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 公式必须以=开头");
            }
            
            return ValidationResult.Success();
        }
        
        /// <summary>
        /// 验证setCellStyle命令的业务规则
        /// </summary>
        private ValidationResult ValidateSetCellStyleRules(JObject arguments, int commandIndex)
        {
            var range = arguments["range"]?.ToString();
            
            if (!IsValidCellRange(range))
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 无效的单元格范围 '{range}'");
            }
            
            // 至少需要设置一个样式属性
            var styleProperties = new[] { "backgroundColor", "fontColor", "bold", "italic", "underline", "fontSize", "fontName" };
            if (!styleProperties.Any(prop => arguments[prop] != null))
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 至少需要设置一个样式属性");
            }
            
            return ValidationResult.Success();
        }
        
        /// <summary>
        /// 验证setCellFormat命令的业务规则
        /// </summary>
        private ValidationResult ValidateSetCellFormatRules(JObject arguments, int commandIndex)
        {
            var range = arguments["range"]?.ToString();
            var format = arguments["format"]?.ToString();
            
            if (!IsValidCellRange(range))
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 无效的单元格范围 '{range}'");
            }
            
            if (string.IsNullOrWhiteSpace(format))
            {
                return ValidationResult.Failure($"命令{commandIndex + 1}: 格式不能为空");
            }
            
            return ValidationResult.Success();
        }
        
        /// <summary>
        /// 验证单元格范围格式是否有效
        /// </summary>
        private bool IsValidCellRange(string range)
        {
            if (string.IsNullOrWhiteSpace(range))
                return false;
                
            if (range == "CURRENT_SELECTION")
                return true;
                
            // 匹配A1、A1:B5、A:A、1:1等格式
            var patterns = new[]
            {
                @"^[A-Z]+\d+$",                    // A1
                @"^[A-Z]+\d+:[A-Z]+\d+$",         // A1:B5
                @"^[A-Z]+:[A-Z]+$",               // A:A
                @"^\d+:\d+$"                      // 1:1
            };
            
            return patterns.Any(pattern => System.Text.RegularExpressions.Regex.IsMatch(range, pattern));
        }
    }
    
    /// <summary>
    /// 验证结果
    /// </summary>
    public class ValidationResult
    {
        /// <summary>
        /// 是否验证成功
        /// </summary>
        public bool IsValid { get; private set; }
        
        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; private set; }
        
        /// <summary>
        /// 验证通过的JSON对象
        /// </summary>
        public JObject ValidatedJson { get; private set; }
        
        private ValidationResult(bool isValid, string errorMessage = null, JObject validatedJson = null)
        {
            IsValid = isValid;
            ErrorMessage = errorMessage;
            ValidatedJson = validatedJson;
        }
        
        /// <summary>
        /// 创建成功的验证结果
        /// </summary>
        public static ValidationResult Success(JObject validatedJson = null)
        {
            return new ValidationResult(true, null, validatedJson);
        }
        
        /// <summary>
        /// 创建失败的验证结果
        /// </summary>
        public static ValidationResult Failure(string errorMessage)
        {
            return new ValidationResult(false, errorMessage);
        }
    }
}