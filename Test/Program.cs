using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SimpleExcel;
using SimpleExcel.Attributes;
using SimpleExcel.Styles;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //var workbook = new WorkBook(ExcelVersion.V2007);
            var workbook = new WorkBook(@"F:\projects\Repos\Panda.SimpleExcel\Test\bin\Debug\test.xlsx");
            var sheet1 = workbook.GetSheet(0);
            var sheet2 = workbook.NewSheet("所有人");
            var sheet3 = workbook.NewSheet("男性");

            //直接给单元格赋值
            sheet1.Rows[0][0].Value = "Hello";

            var list = new List<Person>();
            for(int i = 0; i < 10; i++)
            {
                var person = new Person()
                {
                    Name = "测试" + i,
                    Sex = i % 2 == 0 ? "男" : "女",
                    Age = i
                };
                list.Add(person);
            }
            //将List对象添加到工作表中
            sheet2.ConvertFromQuery(list, 1);

            //将linq语句转换成工作表数据
            var p = from a in list where a.Sex == "男" select a;
            sheet3.ConvertFromQuery(p);

            string path = Environment.CurrentDirectory + @"\test.xlsx";
            workbook.Save(path);

        }
    }

    [Row(
        EvenRowColor = ExcelColor.Aqua,
        OddRowColor = ExcelColor.CornflowerBule,
        HeaderBackColor = ExcelColor.Maroon,
        HeaderFontColor = ExcelColor.White,
        HeaderHeight = 20,
        HeaderHorAlign = HorizontalAlign.Center,
        HeaderVerAlign = VerticalAlign.Center)]
    public class Person
    {
        [Column(
            BackColor = ExcelColor.Brown, 
            FontColor = ExcelColor.White, 
            FontSize = 14,
            FontFamily = "黑体",
            HorAlign = HorizontalAlign.Center,
            VerAlign = VerticalAlign.Center,
            Name = "姓名")]
        public string Name { get; set; }

        [Column(
            FontColor = ExcelColor.Red,
            HorAlign = HorizontalAlign.Left,
            VerAlign = VerticalAlign.Center,
            Name = "性别")]
        public string Sex { get; set; }

        public int Age { get; set; }
    }

}
