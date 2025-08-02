using System;

namespace ExcelAIHelper.Services
{
    /// <summary>
    /// Excel应用程序通用接口 - 抽象Microsoft Excel和WPS表格的差异
    /// </summary>
    public interface IExcelApplication
    {
        /// <summary>
        /// 获取应用程序对象
        /// </summary>
        object Application { get; }

        /// <summary>
        /// 获取活动工作簿
        /// </summary>
        object ActiveWorkbook { get; }

        /// <summary>
        /// 获取活动工作表
        /// </summary>
        object ActiveSheet { get; }

        /// <summary>
        /// 获取选中的区域
        /// </summary>
        object Selection { get; }

        /// <summary>
        /// 获取指定区域
        /// </summary>
        /// <param name="address">区域地址，如"A1:B2"</param>
        /// <returns>区域对象</returns>
        object GetRange(string address);

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="range">区域对象</param>
        /// <returns>单元格值</returns>
        object GetCellValue(object range);

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="range">区域对象</param>
        /// <param name="value">要设置的值</param>
        void SetCellValue(object range, object value);

        /// <summary>
        /// 获取区域地址
        /// </summary>
        /// <param name="range">区域对象</param>
        /// <returns>区域地址字符串</returns>
        string GetRangeAddress(object range);

        /// <summary>
        /// 检查VBA访问权限
        /// </summary>
        /// <returns>是否有VBA访问权限</returns>
        bool HasVbaAccess();

        /// <summary>
        /// 执行VBA代码
        /// </summary>
        /// <param name="vbaCode">VBA代码</param>
        void ExecuteVbaCode(string vbaCode);

        /// <summary>
        /// 显示消息框
        /// </summary>
        /// <param name="message">消息内容</param>
        /// <param name="title">标题</param>
        void ShowMessage(string message, string title = "提示");
    }
}