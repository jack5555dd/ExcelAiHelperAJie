using System;
using System.Threading.Tasks;
using ExcelAIHelper.Exceptions;
using ExcelAIHelper.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelAIHelper.Tests
{
    /// <summary>
    /// JSON命令协议测试
    /// </summary>
    [TestClass]
    public class JsonCommandProtocolTests
    {
        private InstructionParser _parser;
        
        [TestInitialize]
        public void Setup()
        {
            _parser = new InstructionParser();
        }
        
        /// <summary>
        /// 测试有效的JSON命令解析
        /// </summary>
        [TestMethod]
        public async Task ParseAsync_ValidJsonCommand_ShouldReturnInstructionSet()
        {
            // Arrange
            var validJson = @"{
                ""version"": ""1.0"",
                ""summary"": ""Set A1 to 100 and apply formula"",
                ""commands"": [
                    {
                        ""function"": ""setCellValue"",
                        ""arguments"": {
                            ""range"": ""A1"",
                            ""value"": 100,
                            ""dataType"": ""number""
                        },
                        ""description"": ""Set A1 to 100""
                    },
                    {
                        ""function"": ""applyCellFormula"",
                        ""arguments"": {
                            ""range"": ""B1"",
                            ""formula"": ""=SUM(A1:A10)""
                        },
                        ""description"": ""Apply SUM formula""
                    }
                ]
            }";
            
            // Act
            var result = await _parser.ParseAsync(validJson);
            
            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual("Set A1 to 100 and apply formula", result.Summary);
            Assert.AreEqual(2, result.Instructions.Count);
            
            // Verify first instruction
            var firstInstruction = result.Instructions[0];
            Assert.AreEqual(Models.InstructionType.SetCellValue, firstInstruction.Type);
            Assert.AreEqual("A1", firstInstruction.TargetRange);
            Assert.AreEqual(100, firstInstruction.Parameters["value"]);
            
            // Verify second instruction
            var secondInstruction = result.Instructions[1];
            Assert.AreEqual(Models.InstructionType.ApplyFormula, secondInstruction.Type);
            Assert.AreEqual("B1", secondInstruction.TargetRange);
            Assert.AreEqual("=SUM(A1:A10)", secondInstruction.Parameters["formula"]);
        }
        
        /// <summary>
        /// 测试无效JSON格式应抛出AiFormatException
        /// </summary>
        [TestMethod]
        public async Task ParseAsync_InvalidJsonFormat_ShouldThrowAiFormatException()
        {
            // Arrange
            var invalidJson = @"{
                ""version"": ""1.0"",
                ""commands"": [
                    {
                        ""function"": ""setCellValue"",
                        ""arguments"": {
                            ""range"": ""A1"",
                            ""value"": 100
                        }
                    // Missing closing brace
                ]
            ";
            
            // Act & Assert
            var exception = await Assert.ThrowsExceptionAsync<AiFormatException>(
                () => _parser.ParseAsync(invalidJson)
            );
            
            Assert.IsTrue(exception.Message.Contains("JSON格式错误"));
            Assert.AreEqual(invalidJson, exception.OriginalResponse);
            Assert.IsTrue(exception.CanRetry);
        }
        
        /// <summary>
        /// 测试协议违规应抛出AiFormatException
        /// </summary>
        [TestMethod]
        public async Task ParseAsync_ProtocolViolation_ShouldThrowAiFormatException()
        {
            // Arrange - Missing required "version" field
            var protocolViolationJson = @"{
                ""summary"": ""Test command"",
                ""commands"": [
                    {
                        ""function"": ""setCellValue"",
                        ""arguments"": {
                            ""range"": ""A1"",
                            ""value"": 100
                        }
                    }
                ]
            }";
            
            // Act & Assert
            var exception = await Assert.ThrowsExceptionAsync<AiFormatException>(
                () => _parser.ParseAsync(protocolViolationJson)
            );
            
            Assert.IsTrue(exception.Message.Contains("协议违规"));
            Assert.AreEqual(protocolViolationJson, exception.OriginalResponse);
            Assert.IsTrue(exception.CanRetry);
        }
        
        /// <summary>
        /// 测试业务规则违规应抛出AiFormatException
        /// </summary>
        [TestMethod]
        public async Task ParseAsync_BusinessRuleViolation_ShouldThrowAiFormatException()
        {
            // Arrange - Invalid cell range
            var businessRuleViolationJson = @"{
                ""version"": ""1.0"",
                ""commands"": [
                    {
                        ""function"": ""setCellValue"",
                        ""arguments"": {
                            ""range"": ""INVALID_RANGE"",
                            ""value"": 100
                        }
                    }
                ]
            }";
            
            // Act & Assert
            var exception = await Assert.ThrowsExceptionAsync<AiFormatException>(
                () => _parser.ParseAsync(businessRuleViolationJson)
            );
            
            Assert.IsTrue(exception.Message.Contains("无效的单元格范围"));
            Assert.IsTrue(exception.CanRetry);
        }
        
        /// <summary>
        /// 测试markdown代码块包装应被拒绝
        /// </summary>
        [TestMethod]
        public async Task ParseAsync_MarkdownWrappedJson_ShouldThrowAiFormatException()
        {
            // Arrange - JSON wrapped in markdown code blocks (protocol violation)
            var markdownWrappedJson = @"```json
            {
                ""version"": ""1.0"",
                ""commands"": [
                    {
                        ""function"": ""setCellValue"",
                        ""arguments"": {
                            ""range"": ""A1"",
                            ""value"": 100
                        }
                    }
                ]
            }
            ```";
            
            // Act & Assert
            var exception = await Assert.ThrowsExceptionAsync<AiFormatException>(
                () => _parser.ParseAsync(markdownWrappedJson)
            );
            
            Assert.IsTrue(exception.Message.Contains("JSON格式错误") || exception.Message.Contains("协议违规"));
            Assert.IsTrue(exception.CanRetry);
        }
        
        /// <summary>
        /// 测试空响应应抛出AiFormatException
        /// </summary>
        [TestMethod]
        public async Task ParseAsync_EmptyResponse_ShouldThrowAiFormatException()
        {
            // Arrange
            var emptyResponse = "";
            
            // Act & Assert
            var exception = await Assert.ThrowsExceptionAsync<AiFormatException>(
                () => _parser.ParseAsync(emptyResponse)
            );
            
            Assert.IsTrue(exception.Message.Contains("响应内容为空"));
            Assert.IsTrue(exception.CanRetry);
        }
        
        /// <summary>
        /// 测试复杂的有效命令
        /// </summary>
        [TestMethod]
        public async Task ParseAsync_ComplexValidCommand_ShouldParseCorrectly()
        {
            // Arrange
            var complexJson = @"{
                ""version"": ""1.0"",
                ""summary"": ""Complex styling and formula operations"",
                ""commands"": [
                    {
                        ""function"": ""setCellStyle"",
                        ""arguments"": {
                            ""range"": ""A1:C3"",
                            ""backgroundColor"": ""#FF0000"",
                            ""fontColor"": ""white"",
                            ""bold"": true,
                            ""fontSize"": 14
                        },
                        ""description"": ""Style header range""
                    },
                    {
                        ""function"": ""setCellFormat"",
                        ""arguments"": {
                            ""range"": ""D1:D10"",
                            ""format"": ""0.00%""
                        },
                        ""description"": ""Format as percentage""
                    },
                    {
                        ""function"": ""deleteRows"",
                        ""arguments"": {
                            ""range"": ""5:7""
                        },
                        ""description"": ""Delete rows 5-7""
                    }
                ]
            }";
            
            // Act
            var result = await _parser.ParseAsync(complexJson);
            
            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(3, result.Instructions.Count);
            
            // Verify style instruction
            var styleInstruction = result.Instructions[0];
            Assert.AreEqual(Models.InstructionType.SetCellStyle, styleInstruction.Type);
            Assert.AreEqual("A1:C3", styleInstruction.TargetRange);
            Assert.AreEqual("#FF0000", styleInstruction.Parameters["backgroundColor"]);
            Assert.AreEqual(true, styleInstruction.Parameters["bold"]);
            
            // Verify format instruction
            var formatInstruction = result.Instructions[1];
            Assert.AreEqual(Models.InstructionType.SetCellFormat, formatInstruction.Type);
            Assert.AreEqual("0.00%", formatInstruction.Parameters["format"]);
            
            // Verify delete instruction (should require confirmation)
            var deleteInstruction = result.Instructions[2];
            Assert.AreEqual(Models.InstructionType.DeleteRows, deleteInstruction.Type);
            Assert.IsTrue(deleteInstruction.RequiresConfirmation);
        }
    }
}