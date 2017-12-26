# Panda.SimpleExcel
简便操作excel的类库，包括读取、创建、修改等，支持直接从linq语句转换成excel工作表

## **如何使用**
---
首先引入命名空间 `using SimpleExcel;`。这个命名空间下主要有两个类：`WorkBook`和`Sheet`。`WorkBook`用于对整个excel文件的操作，如创建、打开、保存，`Sheet`用于对工作表的操作，如在特定单元格中添加、修改数据，从数据源中批量导入等。可以参考以下的代码示例：

> 新建工作簿和工作表：

    var workbook = new WorkBook();
    var sheet = workbook.NewSheet("sheet1");

> 直接给单元格赋值：（单元格用`Sheet.Rows[rowIndex][columnIndex]`获取，并使用Value属性获取或修改它的内容）

    sheet.Rows[0][0].Value = "Hello";

> Sheet类提供了直接从IEnumerable<T>转换数据的功能。默认情况下，它会将类型T的所有字段名作为表头，将集合中的所有对象排列出来。例如，我们先创建一个类：

    public class Person
    {
        public string Name { get; set; }
        public string Sex { get; set; }
        public int Age { get; set; }
    }

> 然后使用`Sheet.ConvertFromQuey<T>`将集合直接添加到工作表中

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

> 同样，也可以直接使用linq语句将查询结果添加到工作表

    //将linq语句转换成工作表数据
    var p = from a in list where a.Sex == "男" select a;
    sheet.ConvertFromQuery(p, 12);

> 保存工作簿：

    workbook.Save(@"D:\projects\test.xls");

