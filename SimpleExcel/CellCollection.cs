using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace SimpleExcel
{
    /// <summary>
    /// Excel单元格集合
    /// </summary>
    public class CellCollection
    {
        private IRow _row;
        /// <summary>
        /// 获取指定单元格
        /// </summary>
        /// <param name="index">以0开始的单元格索引</param>
        /// <returns></returns>
        public Cell this[int index]
        {
            get
            {
                var cell = _row.GetCell(index);
                if(cell == null)
                {
                    cell = _row.CreateCell(index);
                }
                return new Cell(cell);
            }
            set
            {
                var cell = _row.GetCell(index);
                if(cell == null)
                {
                    cell = _row.CreateCell(index);
                }
                cell.SetCellValue(value.Value);
            }
        }

        /// <summary>
        /// 最后一个单元格的编号（以1开始）
        /// </summary>
        public int LastCellNum
        {
            get { return _row.LastCellNum; }
        }
        /// <summary>
        /// 生成单元格集合
        /// </summary>
        /// <param name="row">实现NPOI中的IRow接口的行实例</param>
        internal CellCollection(NPOI.SS.UserModel.IRow row)
        {
            _row = row;
        }
    }
}
