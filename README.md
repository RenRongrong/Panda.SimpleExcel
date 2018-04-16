# Panda.SimpleExcel
简便操作excel的类库，包括读取、创建、修改等，支持直接从DataSet、IEnumerable<T>以及linq语句转换成excel工作表，并支持通过特性控制工作表样式。

## **Hello World**

    var workbook = new WorkBook();
    var sheet = workbook.NewSheet("Hello");
    sheet.Rows[0][0].Value = "Hello World";
    workbook.Save(@"F:\test.xls");

## **如何使用**
---
- 使用nuget搜索Panda.SimpleExcel并安装。
- 引入命名空间 `using SimpleExcel;`。这个命名空间下主要有两个类：`WorkBook`和`Sheet`。`WorkBook`用于对整个excel文件的操作，如创建、打开、保存，`Sheet`用于对工作表的操作，如在特定单元格中添加、修改数据，从数据源中批量导入等。可以参考以下的代码示例：

> 新建工作簿和工作表：

    var workbook = new WorkBook();
    var sheet = workbook.NewSheet("sheet1");

> 读取工作簿和工作表：

    // 根据路径读取工作簿
    var workbook = var workbook = new WorkBook(@"F:\projects\Repos\Panda.SimpleExcel\Test\bin\Debug\test.xlsx");
    
    // 根据索引读取工作表
    var sheet1 = workbook.GetSheet(0);

    // 根据名称读取工作表
    var sheet2 = workbook.GetSheet("Sheet2");

> 直接给单元格赋值：（单元格用`Sheet.Rows[rowIndex][columnIndex]`获取，并使用Value属性获取或修改它的内容）

    sheet1.Rows[0][0].Value = "Hello";

> Sheet类提供了直接从IEnumerable<T>转换数据的功能。默认情况下，它会将类型T的所有字段名作为表头，将集合中的所有对象排列出来。例如，我们先创建一个类：

    public class Person
    {
        public string Name { get; set; }
        public string Sex { get; set; }
        public int Age { get; set; }
    }

> 然后使用`Sheet.ConvertFromQuery<T>`将集合直接添加到工作表中

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
    //将List对象添加到工作表中，第一个参数是集合对象，第二个参数是起始行数，默认为0
    sheet1.ConvertFromQuery(list, 1);

> 同样，也可以直接使用linq语句将查询结果添加到工作表

    //将linq语句转换成工作表数据
    var p = from a in list where a.Sex == "男" select a;
    sheet2.ConvertFromQuery(p);

> 保存工作簿：

    workbook.Save(@"D:\projects\test.xls");

> 样式控制：可以通过特性来控制工作表的样式。使用Row特性可以控制行样式，使用Column特性可以控制列样式。

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

效果图：

![Excel效果图](https://gitee.com/pandarrr/Panda.SimpleExcel/blob/master/SimpleExcel.PNG)
