using System.Linq;
using NPOI.SS.UserModel;

namespace SimpleExcel.Styles
{
    /// <summary>
    /// 单元格样式
    /// </summary>
    public class Style
    {
        private WorkBook _workbook;
        private Font _font;
        /// <summary>
        /// 单元格所在的工作簿
        /// </summary>
        public WorkBook WorkBook
        {
            get { return _workbook; }
            set
            {
                _workbook = value;
                _workbook.FindOrCreateFont(this.Font);
            }
        }
        
        /// <summary>
        /// 字体
        /// </summary>
        public Font Font
        {
            get { return _font; }
            set
            {
                _font = value;
                if (_workbook != null)
                {
                    _workbook.FindOrCreateFont(value);
                }
            }
        }

        /// <summary>
        /// 样式名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public HorizontalAlign HorAlign { get; set; }

        /// <summary>
        /// 垂直对齐方式
        /// </summary>
        public VerticalAlign VerAlign { get; set; }

        /// <summary>
        /// 背景颜色
        /// </summary>
        public ExcelColor BackColor { get; set; }

        /// <summary>
        /// 列宽
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 行高
        /// </summary>
        public int Height { get; set; }

        internal IFont GetIFont()
        {
            if(_workbook == null)
            {
                return null;
            }
            return _workbook.FindFont(_font);
        }

        internal ICellStyle GetICellStyle()
        {
            if(_workbook == null)
            {
                return null;
            }
            return _workbook.FindOrCreateCellStyle(this);
        }

        /// <summary>
        /// 将多个样式合并为一个，并返回合并后的样式
        /// </summary>
        /// <param name="newStyleName">新样式的名称</param>
        /// <param name="styles">要合并的样式</param>
        /// <returns></returns>
        public static Style Merge(string newStyleName, params Style[] styles)
        {
            var sourceStyles = styles.Where(a => a != null);
            if(!sourceStyles.Any())
            {
                return null;
            }
            if(sourceStyles.Count() == 1)
            {
                return sourceStyles.First();
            }
            Style rootStyle = new Style(newStyleName, sourceStyles.First().WorkBook);
            foreach(Style style in sourceStyles)
            {
                if(style.BackColor != ExcelColor.None)
                {
                    rootStyle.BackColor = style.BackColor;
                }
                if(style.Font != null)
                {
                    rootStyle.Font = style.Font;
                }
                if(style.Height > 0)
                {
                    rootStyle.Height = style.Height;
                }
                if(style.HorAlign != HorizontalAlign.General)
                {
                    rootStyle.HorAlign = style.HorAlign;
                }
                if (style.VerAlign != VerticalAlign.None)
                {
                    rootStyle.VerAlign = style.VerAlign;
                }
                if(style.Width > 0)
                {
                    rootStyle.Width = style.Width;
                }               
            }
            return rootStyle;
        }

        /// <summary>
        /// 生成单元格样式实例
        /// </summary>
        /// <param name="styleName">样式名称</param>
        /// <param name="workbook">样式所在的工作簿</param>
        public Style(string styleName, WorkBook workbook = null)
        {
            Name = styleName;
            _workbook = workbook;
            _font = null;
            BackColor = ExcelColor.None;
        }
    }
}
