# SlimExcelReader
## Function
Read a cell string value (no formula value support currently) from a specified cell, on a specified sheet, in an Excel xlsx document

## Why another library?
I needed to read cell values out of a large (>50MB) excel file from within an Azure function.
I tried 3 or 4 popular openxml based libraries
The libraries either:
- failed to open the file (parse / stream exceptions)
- Caused memory usage to soar over 1GB and, subsequently, caused Azure function timeouts

Directly using the openxml library was clunky as well

However, I had to use openxml directly in the end, to keep memory usage low, and process time snappy.

Hopefully someone else finds this wrapper helpful!

## Example

```
SlimExcelReader.ExcelReader reader = new ExcelReader("my.xlsx");
reader.OpenExcelReader();
reader.SheetName = "My Sheet";
var cellValue = reader.GetValue("A1");
```






