# Excel-Power-Query-Extract-Cell
Extract specific cells from multiple Excel workbooks and make a list

打开一个空白Excel表格，点击“数据-获取数据-来自文件-从文件夹”，选择要获取信息的Excel表格文件所在的文件夹。

打开后点击“转换数据”，进入Power Query编辑器

= List.Transform(Folder.Files("D:\Data\**")[Content],each Table.SelectRows(Excel.Workbook(_),each [Name]="Sheet1")[Data]{0})

在“应用的步骤”栏中“源”点击右键-“插入步骤后”，添加以下一步

= List.Transform(源,each [Column1]{3})


Create multiple excel files in batch
