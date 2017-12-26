using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SimpleExcel;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new WorkBook();
            var sheet = workbook.NewSheet("sheet1");
            //直接给单元格赋值
            sheet.Rows[0][0].Value = "Hello";

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
            sheet.ConvertFromQuery(list, 1);

            //将linq语句转换成工作表数据
            var p = from a in list where a.Sex == "男" select a;
            sheet.ConvertFromQuery(p, 12);

            workbook.Save(@"D:\projects\test.xls");
        }
    }

    public class Person
    {
        public string Name { get; set; }
        public string Sex { get; set; }
        public int Age { get; set; }
    }
}
