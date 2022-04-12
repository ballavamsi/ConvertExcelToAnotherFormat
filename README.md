# ConvertExcelToAnotherFormat

```
using ConvertExcel = ConvertExcelToAnotherFormat.Convert;
...

Process 1: Pass the file path
DataSet ds = ConvertExcel.ToDataSet("your file path");


Process 2: Pass the stream as input
Stream stream = File.OpenRead("your file path");

DataSet ds_stream = ConvertExcel.ToDataSet(stream);
```
