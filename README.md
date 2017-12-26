# Panda.SimpleExcel
简便操作excel的类库，包括读取、创建、修改等，支持直接从linq语句转换成excel工作表

新建工作簿：

    var workbook = new Workbook();
    var sheet = workbook.NewSheet("test1");
    sheet.Rows[0][0].Value = 
