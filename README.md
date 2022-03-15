# Excel-Power-Query-Extract-Cell
Extract specific cells from multiple Excel workbooks and make a list

= List.Transform(Folder.Files("D:\Data\**")[Content],each Table.SelectRows(Excel.Workbook(_),each [Name]="Sheet1")[Data]{0})

= List.Transform(源,each [Column1]{3})


Create multiple excel files in batch
