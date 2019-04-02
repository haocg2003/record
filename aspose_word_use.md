# aspose word 使用
## 安装
从百度盘安装。
## 使用 
### 插入元素
1 插入文本的字符串：

插入文本的字符串需要通过DocumentBuilder.Write方法插入到文档。文本格式是由字体属性决定，这个对象包含不同的字体属性(字体名称,字体大小,颜色,等等)。

一些重要的字体属性也由[{ { DocumentBuilder } }]属性允许您直接访问它们。这些都是布尔属性[{{Font.Bold}}],[{{Font.Italic}}], and [{{Font.Underline}}]。

注意字符格式设置将适用于所有插入的文本。

Example
使用DocumentBuilder插入格式化文本
```
DocumentBuilder builder = new DocumentBuilder();
// Specify font formatting before adding text.
Aspose.Words.Font font = builder.Font;
font.Size = 16;
font.Bold = true;
font.Color = Color.Blue;
font.Name = "Arial";
font.Underline = Underline.Dash;
builder.Write("Sample text.");
```
2 插入一个段落

DocumentBuilder.Writeln可以插入一段文本的字符串也能添加一个段落。当前字体格式也是由DocumentBuilder所规定。字体属性和当前段落格式是由DocumentBuilder.ParagraphFormat属性所决定。

Example
如何添加一个段落到文档
```
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// Specify font formatting
Aspose.Words.Font font = builder.Font;
font.Size = 16;
font.Bold = true;
font.Color = System.Drawing.Color.Blue;
font.Name = "Arial";
font.Underline = Underline.Dash;
// Specify paragraph formatting
ParagraphFormat paragraphFormat = builder.ParagraphFormat;
paragraphFormat.FirstLineIndent = 8;
paragraphFormat.Alignment = ParagraphAlignment.Justify;
paragraphFormat.KeepTogether = true;
builder.Writeln("A whole paragraph.");
```
3 插入一张表

使用DocumentBuilder创建一个表的基本算法是非常简单的：

1.使用[{{DocumentBuilder.StartTable}}]启动表；

2.使用[{{DocumentBuilder.InsertCell}}]插入单元格，这自动生成一个新行，如果需要，使用 [{{DocumentBuilder.CellFormat}}]属性来指定单元格格式；

3.使用DocumentBuilder.methods写入单元格内容；

4.重复步骤2和3,直到行内容写完；

5.调用[{{DocumentBuilder.EndRow}}]来结束当前的行，如果需要，使用[{ { DocumentBuilder.RowFormat }}]属性来指定行格式;

6.重复步骤2 - 5直到表完成;

7.调用[{{DocumentBuilder.EndTable}}]来完成表的创建。

Example
如何创建一个2行2列的格式化表格：
```
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
Table table = builder.StartTable();
// Insert a cell
builder.InsertCell();
// Use fixed column widths.
table.AutoFit(AutoFitBehavior.FixedColumnWidths);
builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
builder.Write("This is row 1 cell 1");
// Insert a cell
builder.InsertCell();
builder.Write("This is row 1 cell 2");
builder.EndRow();
// Insert a cell
builder.InsertCell();
// Apply new row formatting
builder.RowFormat.Height = 100;
builder.RowFormat.HeightRule = HeightRule.Exactly;
builder.CellFormat.Orientation = TextOrientation.Upward;
builder.Writeln("This is row 2 cell 1");
// Insert a cell
builder.InsertCell();
builder.CellFormat.Orientation = TextOrientation.Downward;
builder.Writeln("This is row 2 cell 2");
builder.EndRow();
builder.EndTable();
```
4 插入一个间断：

如果你想开始一个新行、列、段落或者页面，调用DocumentBuilder.InsertBreak就行。

Example
在文档中插入分页符：
```
DocumentBuilder builder = new DocumentBuilder();
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
builder.Writeln("This is page 1.");
builder.InsertBreak(BreakType.PageBreak);
builder.Writeln("This is page 2.");
builder.InsertBreak(BreakType.PageBreak);
builder.Writeln("This is page 3.");
```

5 插入一个图像

DocumentBuilder提供几个({{DocumentBuilder.InsertImage}})多载集合方法，这使得能允许插入一个内联的或者浮动的图像，如果图像是一个EMF或WMF元文件,它将插入到文档的图元文件格式，所有其他的图像将以PNG格式存储。

DocumentBuilder.InsertImage方法可以使用来自不同来源的图像:

从文件或URL通过传递一串字符串参数({{DocumentBuilder.InsertImage}})
从一段流通过一个流参数({{DocumentBuilder.InsertImage}})
从一个图像对象通过一个图像参数(DocumentBuilder.InsertImage)
从一个字节数组通过一个字节数组参数({{DocumentBuilder.InsertImage}})
（1）插入内联图像
Example
如何在一个文档的光标位置插入内联图像。
```
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
builder.InsertImage(MyDir + "Watermark.png");
```
（2）插入一个浮动(绝对位置)的图像
Example
如何从文件或URL插入一个浮动图像：
```
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
builder.InsertImage(MyDir + "Watermark.png",
RelativeHorizontalPosition.Margin,
    100,
    RelativeVerticalPosition.Margin,
    100,
    200,
    100,
    WrapType.Square);
```
6 插入一个书签

插入一个书签到文档中，需要做一下几点：

调用[DocumentBuilder.StartBookmark]通过它设置想要的书签名
使用DocumentBuilder方法插入书签文本
调用[DocumentBuilder.EndBookmark]通过它设置一个与之前设置的书签相同的名字
书签可以重叠和跨越任何范围。创建一个有效的标签你需要调用DocumentBuilder.StartBookmark和DocumentBuilder书签，它们的标签名必须相同

Example
怎样使用document builder在文档中插入一个标签：
```
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
builder.StartBookmark("FineBookmark");
builder.Writeln("This is just a fine bookmark.");
builder.EndBookmark("FineBookmark");
```
7 插入一个字段：

Microsoft Word文档字段由一段字段代码和字段结果组成，这字段代码就像一个公式而字段结果就是这个公式产生的价值。字段代码也可能包括额外的指令来执行特定的操作的field switches 。

 

    你可以切换显示字段代码和使用快捷键Alt+F9得到Microsoft Word文档结果，领域代码出现在花括号({ })之间。

 

    使用[{{DocumentBuilder。InsertField}})来创建文档中的字段，需要指定一个字段类型,字段代码和字段值，如果不确定特定领域代码语法,那首先创建在Microsoft Word创建字段然后切换来看它的字段代码。

 

Example

 

使用DocumentBuilder合并一个字段到文档中：

 

C#

 
```
Document doc = new Document();

DocumentBuilder builder = new DocumentBuilder(doc);

builder.InsertField(@"MERGEFIELD MyFieldName \* MERGEFORMAT");

 
```
 

Visual Basic

 ```

Dim doc As New Document()

Dim builder As New DocumentBuilder(doc)

builder.InsertField("MERGEFIELD MyFieldName \* MERGEFORMAT")

 ```

 

8 插入一个表单字段：

表单字段是一个特殊的允许与用户交互的词字段，在Microsoft Word中表单字段包括文本框,组合框和复选框。

 

DocumentBuilder提供了特殊的方法来将每种类型的表单字段插入到文档:[{{DocumentBuilder.InsertTextInput}}]、[{{DocumentBuilder.InsertCheckBox}}]以及[{{DocumentBuilder.InsertComboBox}}]，注意,如果您为你的表单字段指定一个名称,那么会用相同的名称自动创建一个书签。

 

（1）插入文本输入：

 

使用DocumentBuilder.InsertTextInput向文档插入一个文本框

 

Example 

 

如何向文档插入一个文本输入表单字段。

 

C#

 
```
Document doc = new Document();

DocumentBuilder builder = new DocumentBuilder(doc);

builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);

``` 

Visual Basic

 ```

Dim doc As New Document()

Dim builder As New DocumentBuilder(doc)

builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0)

 ```

 

（2）插入一个复选框

 

Example

 

如何向文档插入一个复选框：

 

C#

 
```
Document doc = new Document(); DocumentBuilder builder = new DocumentBuilder(doc); builder.InsertCheckBox("CheckBox", true, 0); 

 ```

 

Visual Basic 

 

 ```

Dim doc As New Document()

Dim builder As New DocumentBuilder(doc)

builder.InsertCheckBox("CheckBox", True, 0)

 ```

 

 

（3）插入一个组合框

 

调用DocumentBuilder.InsertComboBox向文档插入一个组合框。

 

Example

 

如何将一个组合框表单字段插入文档。

 

C#

 
```
Document doc = new Document(); DocumentBuilder builder = new DocumentBuilder(doc); string[] items = { "One", "Two", "Three" }; builder.InsertComboBox("DropDown", items, 0); 
```
 

 

Visual Basic

 ```

Dim doc As New Document()

Dim builder As New DocumentBuilder(doc)

Dim items() As String = { "One", "Two", "Three" }

builder.InsertComboBox("DropDown", items, 0)
```
 

9 插入HTML

 

    你可以很容易地插入包含一个HTML片段或整个HTML文档的HTML字符串到文档里，只需要传递这字符串到DocumentBuilder.InsertHtmlmethod。

一个有用的实现方法是将一个HTML字符串存储在一个数据库,并将它插入到文档在邮件合并的格式化添加的内容,而不是构建文档构建器的使用各种方法。

 

Example

 

使用DocumentBuilder向文档添加HTML。

 

C#

 
```
Document doc = new Document();

DocumentBuilder builder = new DocumentBuilder(doc);

builder.InsertHtml(

    "<P align='right'>Paragraph right</P>" +

    "<b>Implicit paragraph left</b>" +

    "<div align='center'>Div center</div>" +

    "<h1 align='left'>Heading 1 left.</h1>");

doc.Save(MyDir + "DocumentBuilder.InsertHtml Out.doc");

 ```

Visual Basic

 ```

Dim doc As New Document()

Dim builder As New DocumentBuilder(doc)

builder.InsertHtml("<P align='right'>Paragraph right</P>" & "<b>Implicit paragraph left</b>" & "<div align='center'>Div center</div>" & "<h1 align='left'>Heading 1 left.</h1>")

doc.Save(MyDir & "DocumentBuilder.InsertHtml Out.doc")
```
