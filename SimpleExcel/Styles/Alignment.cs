using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace SimpleExcel.Styles
{
    /// <summary>
    /// 单元格水平对齐方式
    /// </summary>
    public enum HorizontalAlign
    {
        /// <summary>
        /// 常规
        /// </summary>
        General = 0,

        /// <summary>
        /// 靠左
        /// </summary>
        Left = 1,

        /// <summary>
        /// 居中
        /// </summary>
        Center = 2,

        /// <summary>
        /// 靠右
        /// </summary>
        Right = 3,

        /// <summary>
        /// 填充
        /// </summary>
        Fill = 4,

        /// <summary>
        /// 两端对齐
        /// </summary>
        Justify = 5,

        /// <summary>
        /// 跨列居中
        /// </summary>
        CenterSelection = 6,

        /// <summary>
        /// 分散对齐
        /// </summary>
        Distributed = 7       
    }

    /// <summary>
    /// 单元格垂直对齐方式
    /// </summary>
    public enum VerticalAlign
    {
        /// <summary>
        /// 无
        /// </summary>
        None = -1,

        /// <summary>
        /// 靠上
        /// </summary>
        Top = 0,

        /// <summary>
        /// 居中
        /// </summary>
        Center = 1,

        /// <summary>
        /// 靠下
        /// </summary>
        Bottom = 2,

        /// <summary>
        /// 两端对齐
        /// </summary>
        Justify = 3,

        /// <summary>
        /// 分散对齐
        /// </summary>
        Distributed = 4
    }
}
