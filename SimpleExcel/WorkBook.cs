using NPOI.HSSF.UserModel;
using System.Data;
using System.IO;
using SimpleExcel.Styles;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace SimpleExcel
{
    /// <summary>
    /// Excel工作簿
    /// </summary>
    public class WorkBook
    {
        private IWorkbook hworkbook;
        private Dictionary<string, ICellStyle> customerStyles;
        private Dictionary<string, IFont> customerFonts;

        /// <summary>
        /// 从指定的路径读取excel文件
        /// </summary>
        /// <param name="path">要读取的excel文件的路径</param>
        public WorkBook(string path)
        {
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            string extention = Path.GetExtension(path);
            if (extention == ".xls")
            {
                hworkbook = new HSSFWorkbook(file);
            }
            else if(extention == ".xlsx")
            {
                hworkbook = new XSSFWorkbook(file);
            }
            else
            {
                throw new System.Exception("无法识别的文件后缀名！");
            }
            customerStyles = new Dictionary<string, ICellStyle>();
            customerFonts = new Dictionary<string, IFont>();
            file.Close();
        }

        /// <summary>
        /// 从DataSet生成工作簿
        /// </summary>
        /// <param name="ds">源数据集</param>
        /// <param name="version">Excel的版本</param>
        public WorkBook(DataSet ds, ExcelVersion version = ExcelVersion.V2003)
        {
            switch(version)
            {
                case ExcelVersion.V2003:
                    hworkbook = new HSSFWorkbook();
                    break;
                case ExcelVersion.V2007:
                    hworkbook = new XSSFWorkbook();
                    break;
                default:
                    hworkbook = new HSSFWorkbook();
                    break;
            }
            foreach(DataTable table in ds.Tables)
            {
                var sheet = this.NewSheet(table.TableName);
                sheet.ConvertFromDataTable(table);
            }
            customerStyles = new Dictionary<string, ICellStyle>();
            customerFonts = new Dictionary<string, IFont>();
        }

        /// <summary>
        /// 新建一个Excel工作簿
        /// </summary>
        /// <param name="version">Excel的版本</param>
        public WorkBook(ExcelVersion version = ExcelVersion.V2003)
        {
            switch(version)
            {
                case ExcelVersion.V2003:
                    hworkbook = new HSSFWorkbook();
                    break;
                case ExcelVersion.V2007:
                    hworkbook = new XSSFWorkbook();
                    break;
                default:
                    hworkbook = new HSSFWorkbook();
                    break;
            }
            customerStyles = new Dictionary<string, ICellStyle>();
            customerFonts = new Dictionary<string, IFont>();
        }

        /// <summary>
        /// 将工作簿保存到指定位置，如文件已存在，则覆盖原文件
        /// </summary>
        /// <param name="path">要保存的文件路径，包括完整的文件名</param>
        public void Save(string path)
        {
            FileStream file = new FileStream(path, FileMode.Create);
            hworkbook.Write(file);
            file.Close();
        }

        /// <summary>
        /// 返回新建立的工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <returns></returns>
        public Sheet NewSheet(string sheetName)
        {
            var sheet = hworkbook.CreateSheet(sheetName);
            return new Sheet(sheet, this);
        }

        /// <summary>
        /// 根据索引获取工作表
        /// </summary>
        /// <param name="index">以0开始的工作表索引</param>
        /// <returns></returns>
        public Sheet GetSheet(int index)
        {
            return new Sheet(hworkbook.GetSheetAt(index), this);
        }

        /// <summary>
        /// 根据名称获取工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <returns></returns>
        public Sheet GetSheet(string sheetName)
        {
            return new Sheet(hworkbook.GetSheet(sheetName), this);
        }

        /// <summary>
        /// 创建新样式
        /// </summary>
        /// <param name="styleName"></param>
        /// <returns></returns>
        public Style NewStyle(string styleName)
        {
            var style = new Style(styleName, this);
            return style;
        }

        public byte[] ToBytes()
        {
            var stream = new MemoryStream();
            hworkbook.Write(stream);
            var result = stream.ToArray();
            stream.Close();
            return result;
        }

        internal ICellStyle NewCellStyle(string styleName)
        {
            var style = hworkbook.CreateCellStyle();
            customerStyles.Add(styleName, style);     
            return style;
        }

        internal ICellStyle NewCellStyle(Style customerStyle)
        {
            var style = NewCellStyle(customerStyle.Name);
            style.Alignment = (HorizontalAlignment)customerStyle.HorAlign;
            if (customerStyle.BackColor != ExcelColor.None)
            {
                style.FillForegroundColor = (short)customerStyle.BackColor;
                style.FillPattern = FillPattern.SolidForeground;
            }
            style.VerticalAlignment = (VerticalAlignment)customerStyle.VerAlign;
            var font = customerStyle.GetIFont();
            if (font != null)
            {
                style.SetFont(font);
            }
            return style;
        }

        internal ICellStyle FindCellStyle(Style customerStyle)
        {
            return FindCellStyle(customerStyle.Name);
        }

        internal ICellStyle FindCellStyle(string styleName)
        {
            if(customerStyles.ContainsKey(styleName))
            {
                return customerStyles[styleName];
            }
            else
            {
                return null;
            }
        }

        internal ICellStyle FindOrCreateCellStyle(Style customerStyle)
        {
            if(customerStyle == null)
            {
                return null;
            }
            else
            {
                var cellStyle = FindCellStyle(customerStyle);
                if(cellStyle != null)
                {
                    return cellStyle;
                }
                else
                {
                    return NewCellStyle(customerStyle);
                }
            }
        }

        internal IFont NewFont(Font font)
        {
            IFont hfont = hworkbook.CreateFont();
            if (font != null)
            {
                if (font.Family != null)
                {
                    hfont.FontName = font.Family;
                }
                if (font.Size > 0)
                {
                    hfont.FontHeightInPoints = font.Size;
                }
                if (font.Color != ExcelColor.None)
                {
                    hfont.Color = (short)font.Color;
                }
            }
            customerFonts.Add(font.Name, hfont);
            return hfont;
        }

        internal IFont FindFont(Font font)
        {
            if(font == null)
            {
                return null;
            }
            return FindFont(font.Name);
        }

        internal IFont FindFont(string fontName)
        {
            if (this.customerFonts.ContainsKey(fontName))
            {
                return this.customerFonts[fontName];
            }
            else
            {
                return null;
            }
        }

        internal IFont FindOrCreateFont(Font font)
        {
            if(font == null)
            {
                return null;
            }
            var result = FindFont(font.Name);
            if(result != null)
            {
                return result;
            }
            else
            {
                return NewFont(font);
            }
        }
    }

    /// <summary>
    /// Excel的版本
    /// </summary>
    public enum ExcelVersion
    {
        /// <summary>
        /// 2003版
        /// </summary>
        V2003,
        /// <summary>
        /// 2007版
        /// </summary>
        V2007
    }

    
}
