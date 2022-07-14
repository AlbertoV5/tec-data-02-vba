|                                                           |                          |                          |                                     |
|---------------------------------------------------------- |------------------------- |------------------------- |------------------------------------ |
| [<<<Home](https://albertov5.github.io/tec-data/index.html) | [Lesson-1](./lesson-1.md) | [Lesson-2](./lesson-2.md) | [Challenge>>](./challenge/readme.md) |


# Introduction

Excel uses a scripting language to create its macros, it lives inside VBA and the execution is reserved to Office software. The code can be written in any other editor but Excel has a Development Environment that supports running the scripts directly inside it and they afftect the current work sheet.

Macros are `Sub` (Subroutines) that can be called by name from the spreadsheet after enabling macros. This is an example of a macro that will print out the values in Cells (1, 1) to (1, 8).

```vba
Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    Worksheets("2018").Activate
    For i = 1 To 8
        MsgBox (Cells(1, i))

    Next i

End Sub
```

We can access cells with the `Cells` keyword or using `Range`. We can always reference the existing methods and properties in Microsoft&rsquo;s website <sup><a id="fnr.1" class="footref" href="#fn.1" role="doc-backlink">1</a></sup>.


# Reference


## Subroutine Declaration

```vba
Sub MySubroutine()
    'code goes here
End Sub
```


## Variable Declaration

```vba
Dim my_number As Integer
```


## Data Types

| Type             | Size in Memory | Range of Values                                       |
|---------------- |-------------- |----------------------------------------------------- |
| Byte             | 1 Byte         | 0 to 255                                              |
| Integer          | 2 Bytes        | -32,768 to 32767                                      |
| Single           | 4 Bytes        | -3.4E38 to 3.4E38                                     |
| Long             | 8 Bytes        | -2,147,483,648 to 2,147,483,648                       |
| Date             | 8 Bytes        | January 1, 100 to December 31, 999                    |
| Currency         | 8 Bytes        | -922,337,203,685,477.5808 to 922,337,203,685,477.5807 |
| String (dynamic) | 10 Bytes       | 0 to 2 billion characters                             |
| String (fixed)   | string length  | 1 to approximately 65,400                             |
| Boolean          | 4 Bytes        | True or False                                         |
| Object           | 4 Bytes        | Object in VBA                                         |


## Object Types

The `object` type can point to data of any type, can be used as generic type for whenever you don&rsquo;t know the type of the variable it may point to.

```vba
Dim my_object As Object
```

`Object` does not contain the data value itself, as it&rsquo;s a pointer to that value. <sup><a id="fnr.2" class="footref" href="#fn.2" role="doc-backlink">2</a></sup> So it always uses the same space in memory. It&rsquo;s recommended to always define a variable by a specific type rather than trying to point to it at runtime.


# Keywords and Operators

Here is some syntax reference for the common keywords in VBA:


## Conditional

```vba
If 3 > 2 Then
    ' Code here
End If
```


## Loops

```vba
For i = 1 To 10
    ' Code here
Next i
```

## Footnotes

<sup><a id="fn.1" class="footnum" href="#fnr.1">1</a></sup> <https://docs.microsoft.com/en-us/office/vba/api/overview/excel/graph-visual-basic-reference>

<sup><a id="fn.2" class="footnum" href="#fnr.2">2</a></sup> <https://en.wikipedia.org/wiki/Pointer_(computer_programming)>
