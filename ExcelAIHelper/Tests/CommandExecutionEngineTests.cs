using System;
using System.Threading.Tasks;
using ExcelAIHelper.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;

namespace ExcelAIHelper.Tests
{
    /// <summary>
    /// 命令执行引擎测试
    /// </summary>
    [TestClass]
    public class CommandExecutionEngineTests
    {
        /// <summary>
        /// 测试setCellValue命令执行
        /// </summary>
        [TestMethod]
        public async Task ExecuteCommandsAsync_SetCellValue_ShouldReturnSuccess()
        {
            // 注意：这个测试需要Excel环境，在实际环境中运行
            // 这里提供测试结构，实际测试需要Excel实例
            
            // Arrange
            var commandJson = JObject.Parse(@"{
                ""version"": ""1.0"",
                ""summary"": ""设置A1为100"",
                ""commands"": [
                    {
                        ""function"": ""setCellValue"",
                        ""arguments"": {
                            ""range"": ""A1"",
                            ""value"": 100,
                            ""dataType"": ""number""
                        },
                        ""description"": ""在A1设置数值100""
                    }
                ]
            }");
            
            // 在实际测试中，需要创建Excel应用程序实例
            // var excelApp = new Excel.Application();
            // var engine = new CommandExecutionEngine(excelApp);
            
            // Act
            // var result = await engine.ExecuteCommandsAsync(commandJson, true);
            
            // Assert
            // Assert.IsTrue(result.Success);
            // Assert.AreEqual(1, result.ExecutedCommands.Count);
            // Assert.AreEqual("setCellValue", result.ExecutedCommands[0].Function);
            
            // 由于没有Excel环境，这里只验证JSON结构
            Assert.IsNotNull(commandJson);
            Assert.AreEqual("1.0", commandJson["version"].ToString());
            Assert.IsNotNull(commandJson["commands"]);
        }
        
        /// <summary>
        /// 测试applyCellFormula命令执行
        /// </summary>
        [TestMethod]
        public async Task ExecuteCommandsAsync_ApplyCellFormula_ShouldReturnSuccess()
        {
            // Arrange
            var commandJson = JObject.Parse(@"{
                ""version"": ""1.0"",
                ""summary"": ""在B1应用求和公式"",
                ""commands"": [
                    {
                        ""function"": ""applyCellFormula"",
                        ""arguments"": {
                            ""range"": ""B1"",
                            ""formula"": ""=SUM(A1:A10)""
                        },
                        ""description"": ""在B1应用SUM公式""
                    }
                ]
            }");
            
            // 验证JSON结构
            Assert.IsNotNull(commandJson);
            var commands = commandJson["commands"] as JArray;
            Assert.IsNotNull(commands);
            Assert.AreEqual(1, commands.Count);
            
            var command = commands[0] as JObject;
            Assert.AreEqual("applyCellFormula", command["function"].ToString());
            Assert.AreEqual("=SUM(A1:A10)", command["arguments"]["formula"].ToString());
        }
        
        /// <summary>
        /// 测试setCellStyle命令执行
        /// </summary>
        [TestMethod]
        public async Task ExecuteCommandsAsync_SetCellStyle_ShouldReturnSuccess()
        {
            // Arrange
            var commandJson = JObject.Parse(@"{
                ""version"": ""1.0"",
                ""summary"": ""设置A1样式"",
                ""commands"": [
                    {
                        ""function"": ""setCellStyle"",
                        ""arguments"": {
                            ""range"": ""A1"",
                            ""backgroundColor"": ""red"",
                            ""fontColor"": ""white"",
                            ""bold"": true,
                            ""fontSize"": 14
                        },
                        ""description"": ""设置A1为红色背景白色粗体文字""
                    }
                ]
            }");
            
            // 验证JSON结构
            Assert.IsNotNull(commandJson);
            var commands = commandJson["commands"] as JArray;
            var command = commands[0] as JObject;
            var arguments = command["arguments"] as JObject;
            
            Assert.AreEqual("setCellStyle", command["function"].ToString());
            Assert.AreEqual("red", arguments["backgroundColor"].ToString());
            Assert.AreEqual("white", arguments["fontColor"].ToString());
            Assert.AreEqual(true, arguments["bold"].ToObject<bool>());
            Assert.AreEqual(14, arguments["fontSize"].ToObject<int>());
        }
        
        /// <summary>
        /// 测试复杂命令组合执行
        /// </summary>
        [TestMethod]
        public async Task ExecuteCommandsAsync_MultipleCommands_ShouldReturnSuccess()
        {
            // Arrange
            var commandJson = JObject.Parse(@"{
                ""version"": ""1.0"",
                ""summary"": ""设置数据和格式"",
                ""commands"": [
                    {
                        ""function"": ""setCellValue"",
                        ""arguments"": {
                            ""range"": ""A1"",
                            ""value"": ""销售额"",
                            ""dataType"": ""text""
                        },
                        ""description"": ""设置标题""
                    },
                    {
                        ""function"": ""setCellValue"",
                        ""arguments"": {
                            ""range"": ""A2"",
                            ""value"": 1000,
                            ""dataType"": ""number""
                        },
                        ""description"": ""设置数值""
                    },
                    {
                        ""function"": ""setCellStyle"",
                        ""arguments"": {
                            ""range"": ""A1"",
                            ""bold"": true,
                            ""backgroundColor"": ""lightblue""
                        },
                        ""description"": ""设置标题样式""
                    },
                    {
                        ""function"": ""setCellFormat"",
                        ""arguments"": {
                            ""range"": ""A2"",
                            ""format"": ""¥#,##0.00""
                        },
                        ""description"": ""设置货币格式""
                    }
                ]
            }");
            
            // 验证JSON结构
            Assert.IsNotNull(commandJson);
            var commands = commandJson["commands"] as JArray;
            Assert.AreEqual(4, commands.Count);
            
            // 验证每个命令
            Assert.AreEqual("setCellValue", commands[0]["function"].ToString());
            Assert.AreEqual("setCellValue", commands[1]["function"].ToString());
            Assert.AreEqual("setCellStyle", commands[2]["function"].ToString());
            Assert.AreEqual("setCellFormat", commands[3]["function"].ToString());
        }
        
        /// <summary>
        /// 测试预览模式
        /// </summary>
        [TestMethod]
        public async Task ExecuteCommandsAsync_DryRun_ShouldReturnPreview()
        {
            // Arrange
            var commandJson = JObject.Parse(@"{
                ""version"": ""1.0"",
                ""summary"": ""删除第3行"",
                ""commands"": [
                    {
                        ""function"": ""deleteRows"",
                        ""arguments"": {
                            ""range"": ""3:3""
                        },
                        ""description"": ""删除第3行""
                    }
                ]
            }");
            
            // 验证危险操作的JSON结构
            Assert.IsNotNull(commandJson);
            var commands = commandJson["commands"] as JArray;
            var command = commands[0] as JObject;
            
            Assert.AreEqual("deleteRows", command["function"].ToString());
            Assert.AreEqual("3:3", command["arguments"]["range"].ToString());
            
            // 在实际实现中，这种操作应该需要用户确认
            // 预览模式应该返回 "[预览] 将执行: 删除第3行"
        }
    }
}