using System.Data;
using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using SimpleExcel.Attributes;
using SimpleExcel.Styles;
using NPOI.SS.UserModel;

namespace SimpleExcel
{
    /// <summary>
    /// Excel工作表
    /// </summary>
    public class Sheet
    {
        private RowCollection _rows;
        private WorkBook _workbook;
        private ISheet hsheet;
        /// <summary>
        /// 工作表的名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 工作表的行集合
        /// </summary>
        public RowCollection Rows
        {
            get { return _rows; }
        }

        /// <summary>
        /// 生成工作表
        /// </summary>
        /// <param name="sheet">实现NPOI的ISheet接口的工作表实例</param>
        /// <param name="workbook">工作表所在的工作簿</param>
        public Sheet(NPOI.SS.UserModel.ISheet sheet, WorkBook workbook)
        {
            _rows = new RowCollection(sheet);
            _workbook = workbook;
            hsheet = sheet;
            Name = sheet.SheetName;
        }

        /// <summary>
        /// 在指定行添加表头
        /// </summary>
        /// <typeparam name="T">数据的类型</typeparam>
        /// <param name="rowIndex">以0开始的行号</param>
        public void SetHeader<T>(int rowIndex) where T : class
        {
            SetHeader(rowIndex, typeof(T));
        }

        /// <summary>
        /// 在指定行添加表头
        /// </summary>
        /// <param name="rowIndex">以0开始的行号</param>
        /// <param name="type">数据的类型</param>
        public void SetHeader(int rowIndex, Type type)
        {
            short columnIndex = 0;
            bool isStrict = this.IsStrict(type);
            var properties = type.GetProperties();
            var rowStyle = CreateRowHeaderStyle(type);
            foreach (PropertyInfo info in properties)
            {
                ColumnAttribute attr = this.GetColumnAttribute(info);
                if (isStrict && (attr == null))
                {
                    continue;
                }
                string columnName = this.GetColumnName(attr);
                var cell = this.Rows[rowIndex][columnIndex];
                cell.Value = columnName == null ? info.Name : columnName;
                Style columnStyle = attr == null ? null : attr.Style;
                if (columnStyle != null)
                {
                    columnStyle.WorkBook = _workbook;
                }
                SetColumnWidth(columnIndex, attr);
                Style cellStyle = Style.Merge(type.FullName + ".RealHeader.Column" + columnIndex.ToString(), columnStyle, rowStyle);
                cell.SetStyle(cellStyle);
                columnIndex++;
            }
            SetRowHeaderHeight(rowIndex, type);
        }

        private List<Style> GetColumnStyles(Type type)
        {
            List<Style> styles = new List<Style>();
            var properties = type.GetProperties();
            foreach (PropertyInfo info in properties)
            {
                ColumnAttribute attr = this.GetColumnAttribute(info);
                styles.Add(GetColumnStyle(attr));
            }
            return styles;
        }

        private Style GetColumnStyle(ColumnAttribute attr)
        {
            Style columnStyle = attr.Style;
            if(columnStyle != null)
            {
                columnStyle.Name = this.Name + "." + columnStyle.Name;
                columnStyle.WorkBook = _workbook;
            }
            return columnStyle;
        }

        private void SetColumnWidth(int columnIndex, ColumnAttribute attr)
        {
            int width = GetColumnWidthStyle(attr);
            if (width > 0)
            {
                hsheet.SetColumnWidth(columnIndex, width * 256);
            }
        }

        private int GetColumnWidthStyle(ColumnAttribute attr)
        {
            return attr == null ? -1 : attr.Width;
        }

        /// <summary>
        /// 在指定行添加数据
        /// </summary>
        /// <typeparam name="T">数据的类型</typeparam>
        /// <param name="rowIndex">以0开始的行号</param>
        /// <param name="subject">要添加的数据实例</param>
        /// <param name="rowStyle">行格式</param>
        public void AddRow<T>(int rowIndex, T subject, Style rowStyle = null) where T : class
        {
            short columnIndex = 0;
            Type type = typeof(T);
            bool isStrict = this.IsStrict(type);
            var properties = type.GetProperties();
            string rowStyleName = rowStyle == null ? "none" : rowStyle.Name;
            foreach(PropertyInfo info in properties)
            {
                ColumnAttribute columnAttr = this.GetColumnAttribute(info);
                if(isStrict && columnAttr == null)
                {
                    continue;
                }
                var propertyValue = info.GetValue(subject, null);
                Cell cell = this.Rows[rowIndex][columnIndex];
                cell.Value = propertyValue == null ? "" : propertyValue.ToString();
                var columnStyle = columnAttr == null ? null : columnAttr.Style;
                if (columnStyle != null)
                {
                    columnStyle.WorkBook = _workbook;
                }
                Style cellStyle = Style.Merge(rowStyleName + ".Column" + columnIndex.ToString(), columnStyle, rowStyle);
                cell.SetStyle(cellStyle);
                columnIndex++;
            }
            SetRowHeight(rowIndex, type);
        }

        /// <summary>
        /// 从DataTable转换成Sheet
        /// </summary>
        /// <param name="table">要转换的DataTable</param>
        public void ConvertFromDataTable(DataTable table)
        {
            int columnIndex = 0;
            foreach (DataColumn column in table.Columns)
            {
                int rowIndex = 1;
                this.Rows[0][columnIndex].Value = column.ColumnName;
                foreach (DataRow row in table.Rows)
                {
                    var value = row[columnIndex];
                    this.Rows[rowIndex++][columnIndex].Value = value == null ? "" : value.ToString();
                }
                columnIndex++;
            }
        }

        /// <summary>
        /// 从IQueryable转换成工作表
        /// </summary>
        /// <typeparam name="T">数据的类型</typeparam>
        /// <param name="query">查询语句</param>
        /// <param name="rowIndex">作为表头的行号，默认为0</param>
        public void ConvertFromQuery<T>(IEnumerable<T> query, int rowIndex = 0) where T : class
        {
            var oddRowStyle = CreateOddRowStyle(typeof(T));
            var evenRowStyle = CreateEvenRowStyle(typeof(T));
            int rowHeight = GetRowHeightStyle(typeof(T));
            this.SetHeader(rowIndex++, typeof(T));
            foreach (T item in query)
            {
                if (rowIndex % 2 == 0)
                {
                    this.AddRow(rowIndex, item, evenRowStyle);
                }
                else
                {
                    this.AddRow(rowIndex, item, oddRowStyle);
                }
                SetRowHeight(rowIndex++, rowHeight);
            }
        }

        /// <summary>
        /// 从IQueryable转换成工作表
        /// </summary>
        /// <typeparam name="T">数据的类型</typeparam>
        /// <param name="query">查询语句</param>
        /// <param name="rowIndex">作为表头的行号，默认为0</param>
        public void ConvertFromQuery<T>(IQueryable<T> query, int rowIndex = 0) where T : class
        {
            var oddRowStyle = CreateOddRowStyle(typeof(T));
            var evenRowStyle = CreateEvenRowStyle(typeof(T));
            int rowHeight = GetRowHeightStyle(typeof(T));
            this.SetHeader(rowIndex++, typeof(T));        
            foreach (T item in query)
            {
                if (rowIndex % 2 == 0)
                {
                    this.AddRow(rowIndex, item, evenRowStyle);
                }
                else
                {
                    this.AddRow(rowIndex, item, oddRowStyle);
                }
                SetRowHeight(rowIndex++, rowHeight);
            }
        }

        /// <summary>
        /// 尝试将DataTable转换成Excel工作表
        /// </summary>
        /// <param name="table">要转换成工作表的DataTable</param>
        /// <param name="sheet">目标工作表</param>
        /// <returns></returns>
        public static bool TryParse(DataTable table, Sheet sheet)
        {
            if(table == null || table.Columns.Count < 1)
            {
                return false;
            }
            try
            {
                int columnIndex = 0;
                foreach (DataColumn column in table.Columns)
                {
                    int rowIndex = 1;
                    sheet.Rows[0][columnIndex].Value = column.ColumnName;
                    foreach(DataRow row in table.Rows)
                    {
                        sheet.Rows[rowIndex++][columnIndex].Value = row[columnIndex].ToString();
                    }
                    columnIndex++;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 根据类型的Row特性指示是否执行严格的显示控制
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns></returns>
        private bool IsStrict(Type type)
        {
            var attr = type.GetCustomAttributes(typeof(RowAttribute), true);
            if (attr.Any())
            {
                RowAttribute rowAttr = attr.First() as RowAttribute;
                return rowAttr.IsStrict;
            }
            else
            {
                return false;
            }
        }

        private ColumnAttribute GetColumnAttribute(PropertyInfo info)
        {
            var attrs = info.GetCustomAttributes(typeof(ColumnAttribute), true);
            if (!attrs.Any())
            {
                return null;
            }
            else
            {
                ColumnAttribute columnAttr = attrs.First() as ColumnAttribute;
                return columnAttr;
            }
        }

        private string GetColumnName(ColumnAttribute attr)
        {
            if (attr == null)
            {
                return null;
            }
            else
            {
                return attr.Name;
            }
        }

        private Style CreateRowHeaderStyle(Type type)
        {
            var attr = type.GetCustomAttributes(typeof(RowAttribute), true);
            if (!attr.Any())
            {
                return null;
            }
            RowAttribute rowAttr = attr.First() as RowAttribute;
            Font font = new Font(type.FullName + ".Header");
            font.Family = rowAttr.HeaderFontFamily;
            font.Size = rowAttr.HeaderFontSize;
            font.Color = rowAttr.HeaderFontColor;
            Style style = _workbook.NewStyle(type.FullName + ".Header");
            style.Font = font;
            style.BackColor = rowAttr.HeaderBackColor;
            style.Height = rowAttr.HeaderHeight;
            style.HorAlign = rowAttr.HeaderHorAlign;
            style.VerAlign = rowAttr.HeaderVerAlign;
            style.Width = 0;
            return style;
        }

        private Style CreateOddRowStyle(Type type)
        {
            var attr = type.GetCustomAttributes(typeof(RowAttribute), true);
            if (!attr.Any())
            {
                return null;
            }
            RowAttribute rowAttr = attr.First() as RowAttribute;
            if(rowAttr.OddRowColor == ExcelColor.None)
            {
                return null;
            }
            var style = _workbook.NewStyle(type.FullName + ".Odd");
            style.BackColor = rowAttr.OddRowColor;
            return style;
        }

        private Style CreateEvenRowStyle(Type type)
        {
            var attr = type.GetCustomAttributes(typeof(RowAttribute), true);
            if (!attr.Any())
            {
                return null;
            }
            RowAttribute rowAttr = attr.First() as RowAttribute;
            if (rowAttr.EvenRowColor == ExcelColor.None)
            {
                return null;
            }
            var style = _workbook.NewStyle(type.FullName + ".Even");
            style.BackColor = rowAttr.EvenRowColor;
            return style;
        }

        private void SetRowHeight(int rowIndex, int height)
        {
            if(height < 1)
            {
                return;
            }
            var row = hsheet.GetRow(rowIndex);
            row.HeightInPoints = height;
        }

        private void SetRowHeight(int rowIndex, Type type)
        {
            int height = GetRowHeightStyle(type);
            if(height > 0)
            {
                var row = hsheet.GetRow(rowIndex);
                row.HeightInPoints = height;
            }
        }

        private void SetRowHeaderHeight(int rowIndex, Type type)
        {
            var attr = type.GetCustomAttributes(typeof(RowAttribute), true);
            if(attr.Any())
            {
                RowAttribute rowAttr = attr.First() as RowAttribute;
                SetRowHeight(rowIndex, rowAttr.HeaderHeight);
            }
        }

        private int GetRowHeightStyle(Type type)
        {
            var attr = type.GetCustomAttributes(typeof(RowAttribute), true);
            if(!attr.Any())
            {
                return 0;
            }
            RowAttribute rowAttr = attr.First() as RowAttribute;
            return rowAttr.Height;
        }

    }
}
