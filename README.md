## C# 使用NPOI操作Excel文件

# 什么是NPOI？

> What’s NPOI
> This project is the .NET version of POI Java project at http://poi.apache.org/. POI is an open source project which can help you read/write xls, doc, ppt files. It has a wide application.
> For example, you can use it to
> a. generate a Excel report without Microsoft Office suite installed on your server and more efficient than call Microsoft Excel ActiveX at background;
> b. extract text from Office documents to help you implement full-text indexing feature (most of time this feature is used to create search engines).
> c. extract images from Office documents
> d. generate Excel sheets that contains formulas

简而言之，言而简之，NPOI是源于一个用于读取xls,doc,ppt文档的POI 项目，POI是Java项目，后面因为有.Net的市场，于是将POI移植到.Net上。

# 优势：

**在没有安装Microsoft Office Excel的机子上也可以对Excel进行操作。另外一种方法是使用.NET自带的excel API，但是这种方法需要运行环境安装微软的excel才行。**

**NPOI尤其适合在服务器端生成数据文件！因为服务器一般是不安装office这么庞大的办公软件的**

# 使用方法：

## 1.准备npoi 的 dll：

下载地址：
https://npoi.codeplex.com/releases

## 2.将NPOI的DLL导入工程中。

右键**解决方案资源管理器**里面的**引用**
![这里写图片描述](images\1)

点击**添加引用**
![这里写图片描述](images\3)

点击“浏览”->”浏览“，打开文件对话框，选择所有的NPOI的dll文件
![这里写图片描述](images\4)

## 3.引用NPOI的命名空间。

```
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.Util;1234
```

## 4.编程开发

不管是读还是写一个excel文件，都要先生成一个HSSFWorkbook对象。
NPOI里面的管理层次为：workbook->worksheet->row->cell.
类比关系型数据库就是：

| NPOI      | 关系型数据库 |
| --------- | ------------ |
| workbook  | database     |
| worksheet | table表      |
| row       | record记录   |
| cell      | field字段    |

形象一点就是：
![这里写图片描述](images\2)
具体方法为：

**以现有excel文件数据为基础，创建一个workbook对象，这种方法可以读取这个excel文件的数据内容：**

```
HSSFWorkbook wb;
FileStream file;
file = new FileStream(filepath, FileMode.Open, FileAccess.Read);
wb = new HSSFWorkbook(file);
file.Close();12345
```

可以发现是借助FileStream来读取excel文件的，其中的filepath指明excel文件的路径。

**创新一个新的excel文件的workbook对象：**

```
HSSFWorkbook wb;
wb = new HSSFWorkbook();12
```

在workbook的基础上，打开一个老的sheet,或者创建一个新的表。
**打开老的sheet: wb.GetSheet(sheet的名称)**

```
HSSFSheet sheet;
sheet=wb.GetSheet("sheet1");12
```

**创建一个新的sheet:wb.CreateSheet(sheet的名称）**

```
HSSFSheet sheet;
sheet=wb.CreateSheet("sheet1");12
```

现在就到具体操作某个行和列了。

**创建某个行：CreateRow(i)，i是行号，从0开始计数**

```
sheet.CreateRow(i）1
```

**获取某一行：GetRow(i），i是行号，从0开始计数**

**创建某一列： 需要在定位到行的基础上**

CreateCell(j)，j是列号，从0开始计数

```
sheet.GetRow(i).CreateCell(j);在i行创建第j列1
```

**获取某一行：GetCell(j),j是列号，从0开始计数**

```
sheet.GetRow(i).GetCell(j)1
```

行和列都确定了，那就是对单元格的操作啦：

**读单元格：**

```
sheet.GetRow(i).GetCell(j) 就会返回第i行j列的内容。1
```

**写单元格：**

```
sheet.GetRow(i).GetCell(j).SetCellValue(内容)1
```

**保存数据到文件中**

```
file = new FileStream(filepath, FileMode.Open, FileAccess.Write);
wb.Write(file);
file.Close();
wb.Close()1234
```

其中workbook的写入 需要借助于FileStream来打开一个文件流，在创建FileStream的时候，我们可以传入数据的保存路径和文件名。

wb.Write在实际写入数据。

最后操作完成后需要关闭资源。

# 总结：

使用NPOI操作excel很方便，关键是 workbook,sheet,row,cell的层层定位。

**另外 NPOI 使用 HSSFWorkbook 类来处理 xls，XSSFWorkbook 类来处理 xlsx，它们都继承接口 IWorkbook，因此可以通过 IWorkbook 来统一处理 xls 和 xlsx 格式的文件。**