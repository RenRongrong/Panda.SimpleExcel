using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace SimpleExcel
{
    /// <summary>
    /// Excel工作表中的行集合
    /// </summary>
    public class RowCollection
    {
        private ISheet _sheet;

        /// <summary>
        /// 获取指定行的单元格集合
        /// </summary>
        /// <param name="index">以0开始的行索引</param>
        /// <returns></returns>
        public CellCollection this[int index]
        {
            get
            {
                var row = _sheet.GetRow(index);
                if(row == null)
                {
                    row = _sheet.CreateRow(index);
                }
                return new CellCollection(row);
            }
        }

        /// <summary>
        /// 最后一行的编号（从0开始）
        /// </summary>
        public int LastRowNum
        {
            get { return _sheet.LastRowNum; }
        }

        /// <summary>
        /// 生成行集合
        /// </summary>
        /// <param name="sheet">实现NPOI中的ISheet接口的工作表实例</param>
        internal RowCollection(NPOI.SS.UserModel.ISheet sheet)
        {
            _sheet = sheet;
        }
    }
}
