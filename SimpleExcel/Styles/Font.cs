using System;

namespace SimpleExcel.Styles
{
    /// <summary>
    /// 字体
    /// </summary>
    public class Font
    {
        /// <summary>
        /// 在字典表中的键值名称
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 字体名称
        /// </summary>
        public string Family { get; set; }
        /// <summary>
        /// 字体大小
        /// </summary>
        public short Size { get; set; }
        /// <summary>
        /// 字体颜色
        /// </summary>
        public ExcelColor Color { get; set; }
        /// <summary>
        /// 生成字体的实例
        /// </summary>
        /// <param name="name">字体在字典表中的键值名称</param>
        public Font(string name)
        {
            if(string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentNullException("name", "Font的名字不能为空！");
            }
            this.Name = name;
        }
    }
}
