# CsvSchema

Schema information file generator for Windows text driver.

## Usage

Here is a simple example of this library.  

```VBS
Sub CreateSchemaIni()

    With New CsvSchema
        Call .Entry("C:\foo\bar.csv")
        Call .AddField("Field001", DataType.dtText, 10)
        Call .AddField("Field002", DataType.dtLong)
        Call .Create()
    End With

End Sub
```

And you'll get a schema file as `C:\foo\schema.ini`.  
This content are as below: 

```INI
[C:\foo\bar.csv]

Col1=Field001 Text Width 10
Col2=Field002 Long
```