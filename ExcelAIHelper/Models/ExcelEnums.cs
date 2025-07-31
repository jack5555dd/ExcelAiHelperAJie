namespace ExcelAIHelper.Models
{
    /// <summary>
    /// 复制类型枚举
    /// </summary>
    public enum CopyType
    {
        /// <summary>
        /// 复制所有内容（值、格式、公式）
        /// </summary>
        All,
        
        /// <summary>
        /// 仅复制值
        /// </summary>
        Values,
        
        /// <summary>
        /// 仅复制公式
        /// </summary>
        Formulas,
        
        /// <summary>
        /// 仅复制格式
        /// </summary>
        Formats
    }
    
    /// <summary>
    /// 边框类型枚举
    /// </summary>
    public enum BorderType
    {
        /// <summary>
        /// 所有边框
        /// </summary>
        All,
        
        /// <summary>
        /// 外边框
        /// </summary>
        Outline,
        
        /// <summary>
        /// 内边框
        /// </summary>
        Inside,
        
        /// <summary>
        /// 上边框
        /// </summary>
        Top,
        
        /// <summary>
        /// 下边框
        /// </summary>
        Bottom,
        
        /// <summary>
        /// 左边框
        /// </summary>
        Left,
        
        /// <summary>
        /// 右边框
        /// </summary>
        Right
    }
    
    /// <summary>
    /// 条件格式类型枚举
    /// </summary>
    public enum ConditionalFormatType
    {
        /// <summary>
        /// 背景颜色
        /// </summary>
        BackgroundColor,
        
        /// <summary>
        /// 字体颜色
        /// </summary>
        FontColor,
        
        /// <summary>
        /// 粗体
        /// </summary>
        Bold,
        
        /// <summary>
        /// 斜体
        /// </summary>
        Italic,
        
        /// <summary>
        /// 下划线
        /// </summary>
        Underline
    }
    
    /// <summary>
    /// 数据类型枚举
    /// </summary>
    public enum DataType
    {
        /// <summary>
        /// 自动检测
        /// </summary>
        Auto,
        
        /// <summary>
        /// 文本
        /// </summary>
        Text,
        
        /// <summary>
        /// 数字
        /// </summary>
        Number,
        
        /// <summary>
        /// 日期
        /// </summary>
        Date,
        
        /// <summary>
        /// 时间
        /// </summary>
        Time,
        
        /// <summary>
        /// 布尔值
        /// </summary>
        Boolean,
        
        /// <summary>
        /// 公式
        /// </summary>
        Formula
    }
    
    /// <summary>
    /// 排序方向枚举
    /// </summary>
    public enum SortDirection
    {
        /// <summary>
        /// 升序
        /// </summary>
        Ascending,
        
        /// <summary>
        /// 降序
        /// </summary>
        Descending
    }
    
    /// <summary>
    /// 对齐方式枚举
    /// </summary>
    public enum TextAlignment
    {
        /// <summary>
        /// 左对齐
        /// </summary>
        Left,
        
        /// <summary>
        /// 居中
        /// </summary>
        Center,
        
        /// <summary>
        /// 右对齐
        /// </summary>
        Right,
        
        /// <summary>
        /// 两端对齐
        /// </summary>
        Justify
    }
    
    /// <summary>
    /// 垂直对齐方式枚举
    /// </summary>
    public enum VerticalAlignment
    {
        /// <summary>
        /// 顶部对齐
        /// </summary>
        Top,
        
        /// <summary>
        /// 居中对齐
        /// </summary>
        Middle,
        
        /// <summary>
        /// 底部对齐
        /// </summary>
        Bottom
    }
    
    /// <summary>
    /// 清除类型枚举
    /// </summary>
    public enum ClearType
    {
        /// <summary>
        /// 清除内容
        /// </summary>
        Content,
        
        /// <summary>
        /// 清除格式
        /// </summary>
        Format,
        
        /// <summary>
        /// 清除所有
        /// </summary>
        All,
        
        /// <summary>
        /// 清除注释
        /// </summary>
        Comments,
        
        /// <summary>
        /// 清除超链接
        /// </summary>
        Hyperlinks
    }
}