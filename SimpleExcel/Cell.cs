using NPOI.SS.UserModel;
using SimpleExcel.Styles;

namespace SimpleExcel
{
    /// <summary>
    /// Excel单元格
    /// </summary>
    public class Cell
    {
        private ICell hcell;
        /// <summary>
        /// 单元格的内容
        /// </summary>
        public string Value
        {
            get { return hcell.StringCellValue; }
            set { hcell.SetCellValue(value); }
        }
        /// <summary>
        /// 生成单元格
        /// </summary>
        /// <param name="cell">实现NPOI中的ICell接口的单元格实例</param>
        internal Cell(NPOI.SS.UserModel.ICell cell)
        {
            hcell = cell;
        }

        /// <summary>
        /// 设置单元格格式
        /// </summary>
        /// <param name="style">格式</param>
        internal void SetStyle(ICellStyle style)
        {
            if (style != null)
            {
                hcell.CellStyle = style;         
            }
        }

        internal void SetStyle(Style style)
        {
            if(style != null)
            {
                hcell.CellStyle = style.GetICellStyle();
            }
        }
    }
}
