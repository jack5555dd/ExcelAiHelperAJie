using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace ExcelAIHelper
{
    public partial class ExampleForm : Form
    {
        private Dictionary<string, ExampleInfo> examples;
        
        public ExampleForm()
        {
            InitializeComponent();
            InitializeExamples();
        }

        private void InitializeExamples()
        {
            examples = new Dictionary<string, ExampleInfo>
            {
                {
                    "智能公式助手",
                    new ExampleInfo
                    {
                        Title = "智能公式助手",
                        Description = "AI助手可以帮助您：\n\n" +
                                    "• 自动生成复杂的Excel公式\n" +
                                    "• 检查和优化现有公式\n" +
                                    "• 解释公式的含义和用法\n" +
                                    "• 提供公式使用建议\n\n" +
                                    "示例：\n" +
                                    "输入：'计算销售额的平均值'\n" +
                                    "AI会自动生成：=AVERAGE(B2:B10)\n\n" +
                                    "支持的函数类型：\n" +
                                    "- 数学函数（SUM、AVERAGE、MAX、MIN等）\n" +
                                    "- 逻辑函数（IF、AND、OR等）\n" +
                                    "- 文本函数（CONCATENATE、LEFT、RIGHT等）\n" +
                                    "- 日期函数（TODAY、DATEDIF等）",
                        Url = "https://examples.aiexcel.cn/formula-assistant"
                    }
                },
                {
                    "数据分析工具",
                    new ExampleInfo
                    {
                        Title = "数据分析工具",
                        Description = "强大的数据分析功能：\n\n" +
                                    "• 自动识别数据模式和趋势\n" +
                                    "• 生成数据透视表\n" +
                                    "• 创建图表和可视化\n" +
                                    "• 数据清洗和格式化\n\n" +
                                    "分析类型：\n" +
                                    "- 描述性统计分析\n" +
                                    "- 相关性分析\n" +
                                    "- 趋势分析\n" +
                                    "- 异常值检测\n\n" +
                                    "使用场景：\n" +
                                    "- 销售数据分析\n" +
                                    "- 财务报表分析\n" +
                                    "- 市场调研数据处理\n" +
                                    "- 运营数据监控",
                        Url = "https://examples.aiexcel.cn/data-analysis"
                    }
                },
                {
                    "智能格式化",
                    new ExampleInfo
                    {
                        Title = "智能格式化",
                        Description = "一键美化您的Excel表格：\n\n" +
                                    "• 自动应用专业样式\n" +
                                    "• 智能调整列宽和行高\n" +
                                    "• 添加条件格式\n" +
                                    "• 创建专业报表模板\n\n" +
                                    "格式化选项：\n" +
                                    "- 商务风格\n" +
                                    "- 学术风格\n" +
                                    "- 财务报表风格\n" +
                                    "- 数据仪表板风格\n\n" +
                                    "特色功能：\n" +
                                    "- 自动颜色搭配\n" +
                                    "- 智能边框设置\n" +
                                    "- 数据条和图标集\n" +
                                    "- 表格主题应用",
                        Url = "https://examples.aiexcel.cn/smart-formatting"
                    }
                },
                {
                    "聚光灯功能",
                    new ExampleInfo
                    {
                        Title = "聚光灯功能",
                        Description = "突出显示重要数据：\n\n" +
                                    "• 鼠标悬停自动高亮\n" +
                                    "• 选中区域聚光效果\n" +
                                    "• 自定义聚光颜色\n" +
                                    "• 一键开关聚光模式\n\n" +
                                    "使用方法：\n" +
                                    "1. 点击聚光灯按钮开启功能\n" +
                                    "2. 选择需要突出的单元格区域\n" +
                                    "3. 设置聚光颜色（可选）\n" +
                                    "4. 再次点击关闭聚光效果\n\n" +
                                    "适用场景：\n" +
                                    "- 数据演示和汇报\n" +
                                    "- 重要信息标记\n" +
                                    "- 数据对比分析\n" +
                                    "- 会议讨论重点",
                        Url = "https://examples.aiexcel.cn/spotlight"
                    }
                },
                {
                    "AI聊天助手",
                    new ExampleInfo
                    {
                        Title = "AI聊天助手",
                        Description = "智能对话式Excel操作：\n\n" +
                                    "• 自然语言描述需求\n" +
                                    "• AI理解并执行操作\n" +
                                    "• 实时问答和指导\n" +
                                    "• 学习Excel技巧\n\n" +
                                    "对话示例：\n" +
                                    "用户：'帮我计算这列数据的总和'\n" +
                                    "AI：'我来为您添加SUM公式...'\n\n" +
                                    "用户：'如何制作数据透视表？'\n" +
                                    "AI：'我来指导您创建数据透视表...'\n\n" +
                                    "支持的操作：\n" +
                                    "- 公式创建和修改\n" +
                                    "- 数据处理和分析\n" +
                                    "- 格式设置和美化\n" +
                                    "- Excel技能学习",
                        Url = "https://examples.aiexcel.cn/ai-chat"
                    }
                }
            };
        }

        private void ExampleForm_Load(object sender, EventArgs e)
        {
            // 加载示例列表
            foreach (var example in examples.Keys)
            {
                listExamples.Items.Add(example);
            }
            
            // 默认选择第一个示例
            if (listExamples.Items.Count > 0)
            {
                listExamples.SelectedIndex = 0;
            }
        }

        private void listExamples_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listExamples.SelectedItem != null)
            {
                string selectedExample = listExamples.SelectedItem.ToString();
                if (examples.ContainsKey(selectedExample))
                {
                    var exampleInfo = examples[selectedExample];
                    lblExampleTitle.Text = exampleInfo.Title;
                    txtExampleDescription.Text = exampleInfo.Description;
                    linkExample.Text = exampleInfo.Url;
                }
            }
        }

        private void linkExample_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenUrl(linkExample.Text);
        }

        private void btnVisitExample_Click(object sender, EventArgs e)
        {
            if (listExamples.SelectedItem != null)
            {
                string selectedExample = listExamples.SelectedItem.ToString();
                if (examples.ContainsKey(selectedExample))
                {
                    OpenUrl(examples[selectedExample].Url);
                }
            }
            else
            {
                MessageBox.Show("请先选择一个示例", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void OpenUrl(string url)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开链接: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

    public class ExampleInfo
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string Url { get; set; }
    }
}