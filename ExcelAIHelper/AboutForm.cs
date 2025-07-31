using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelAIHelper
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
        }

        private void AboutForm_Load(object sender, EventArgs e)
        {
            // 生成二维码
            GenerateQRCode();
        }

        private void GenerateQRCode()
        {
            try
            {
                // 创建一个简单的二维码图像（使用SVG格式模拟）
                // 在实际项目中，您可以使用QRCoder或ZXing.Net库来生成真实的二维码
                var qrCodeImage = CreateSimpleQRCodeImage();
                picQRCode.Image = qrCodeImage;
            }
            catch (Exception ex)
            {
                // 如果生成二维码失败，显示占位符文本
                picQRCode.BackColor = Color.LightGray;
                var font = new Font("微软雅黑", 8);
                var bitmap = new Bitmap(100, 100);
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.LightGray);
                    g.DrawString("二维码\n占位符", font, Brushes.Black, new RectangleF(10, 35, 80, 30));
                }
                picQRCode.Image = bitmap;
            }
        }

        private Image CreateSimpleQRCodeImage()
        {
            // 创建一个简单的二维码样式图像
            var bitmap = new Bitmap(100, 100);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                
                // 绘制简单的二维码模式
                var brush = new SolidBrush(Color.Black);
                var random = new Random(12345); // 使用固定种子确保一致性
                
                // 绘制定位标记（三个角）
                DrawPositionMarker(g, brush, 5, 5);
                DrawPositionMarker(g, brush, 75, 5);
                DrawPositionMarker(g, brush, 5, 75);
                
                // 绘制随机数据模块
                for (int x = 0; x < 100; x += 4)
                {
                    for (int y = 0; y < 100; y += 4)
                    {
                        // 避开定位标记区域
                        if ((x < 25 && y < 25) || (x > 70 && y < 25) || (x < 25 && y > 70))
                            continue;
                            
                        if (random.Next(2) == 0)
                        {
                            g.FillRectangle(brush, x, y, 3, 3);
                        }
                    }
                }
            }
            return bitmap;
        }

        private void DrawPositionMarker(Graphics g, SolidBrush brush, int x, int y)
        {
            // 绘制定位标记（7x7的方形图案）
            g.FillRectangle(brush, x, y, 20, 20);
            g.FillRectangle(Brushes.White, x + 3, y + 3, 14, 14);
            g.FillRectangle(brush, x + 6, y + 6, 8, 8);
        }

        private void linkWebsite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                // 打开默认浏览器访问网站
                Process.Start(new ProcessStartInfo
                {
                    FileName = linkWebsite.Text,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开网站: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}