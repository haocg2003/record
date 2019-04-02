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
