using System;
using SimpleExcel.Styles;

namespace SimpleExcel.Attributes
{
    /// <summary>
    /// 将类的每一个实例作为Excel中的一行
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class RowAttribute : Attribute
    {
        /// <summary>
        /// 是否执行严格转化，若为True，则将实例添加至行时只会显示含有Column特性的属性
        /// </summary>
        public bool IsStrict { get; set; }

        /// <summary>
        /// 表头字体
        /// </summary>
        public string HeaderFontFamily { get; set; }

        /// <summary>
        /// 表头字号
        /// </summary>
        public short HeaderFontSize { get; set; }

        /// <summary>
        /// 表头字体颜色
        /// </summary>
        public ExcelColor HeaderFontColor { get; set; }

        /// <summary>
        /// 表头背景色
        /// </summary>
        public ExcelColor HeaderBackColor { get; set; }

        /// <summary>
        /// 表头行高
        /// </summary>
        public int HeaderHeight { get; set; }

        /// <summary>
        /// 表头以外的行高
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// 表头单元格水平对齐方式
        /// </summary>
        public HorizontalAlign HeaderHorAlign { get; set; }

        /// <summary>
        /// 表头单元格垂直对齐方式
        /// </summary>
        public VerticalAlign HeaderVerAlign { get; set; }

        /// <summary>
        /// 奇数行背景色
        /// </summary>
        public ExcelColor OddRowColor { get; set; }

        /// <summary>
        /// 偶数行背景色
        /// </summary>
        public ExcelColor EvenRowColor { get; set; }
    }
}
