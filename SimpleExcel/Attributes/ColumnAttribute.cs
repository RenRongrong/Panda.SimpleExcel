using System;
using SimpleExcel.Styles;

namespace SimpleExcel.Attributes
{
    /// <summary>
    /// 将属性作为Excel的列
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {       
        /// <summary>
        /// 显示在Excel中的表头名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 整列的样式
        /// </summary>
        internal Style Style { get; set; }

        /// <summary>
        /// 整列的字体
        /// </summary>
        public Font Font { get; set; }

        /// <summary>
        /// 字体
        /// </summary>
        public string FontFamily
        {
            get { return Font == null ? "微软雅黑" :  Font.Family; }
            set
            {
                CreateFont();
                Font.Family = value;
            }
        }

        /// <summary>
        /// 字号
        /// </summary>
        public short FontSize
        {
            get { return Font == null ? (short)0 : Font.Size; }
            set
            {
                CreateFont();
                Font.Size = value;
            }
        }

        /// <summary>
        /// 字体颜色
        /// </summary>
        public ExcelColor FontColor
        {
            get { return Font == null ? ExcelColor.None : Font.Color; }
            set
            {
                CreateFont();
                Font.Color = value;
            }
        }

        /// <summary>
        /// 列宽
        /// </summary>
        public int Width
        {
            get { return Style == null ? 0 : Style.Width; }
            set
            {
                CreateStyle();
                Style.Width = value;
            }
        }

        /// <summary>
        /// 背景颜色
        /// </summary>
        public ExcelColor BackColor
        {
            get { return Style == null ? ExcelColor.None : Style.BackColor; }
            set
            {
                CreateStyle();
                Style.BackColor = value;
            }
        }

        /// <summary>
        /// 单元格水平对齐方式
        /// </summary>
        public HorizontalAlign HorAlign
        {
            get { return Style == null ? HorizontalAlign.General : Style.HorAlign; }
            set
            {
                CreateStyle();
                Style.HorAlign = value;
            }
        }

        /// <summary>
        /// 单元格垂直对齐方式
        /// </summary>
        public VerticalAlign VerAlign
        {
            get { return Style == null ? VerticalAlign.None : Style.VerAlign; }
            set
            {
                CreateStyle();
                Style.VerAlign = value;
            }
        }

        private void CreateFont()
        {
            CreateStyle();
            if (Font == null)
            {
                Font = new Font("ColumnFont." + this.Name);
                Style.Font = Font;
            }
        }

        private void CreateStyle()
        {
            if(Style == null)
            {
                Style = new Style("ColumnStyle." + this.Name);
            }
        }

    }
}
