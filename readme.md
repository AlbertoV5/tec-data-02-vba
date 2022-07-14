|                                                           |                          |                          |                                      |
|---------------------------------------------------------- |------------------------- |------------------------- |------------------------------------- |
| [<<<Home](https://albertov5.github.io/tec-data/index.html) | [Lesson-1](./lesson-1.md) | [Lesson-2](./lesson-2.md) | [Challenge>>>](./challenge/readme.md) |


# Introduction

Excel uses a scripting language to create its macros, it lives inside VBA and the execution is reserved to Office software. The code can be written in any other editor but Excel has a Development Environment that supports running the scripts directly inside it and they afftect the current work sheet.

Macros are Subroutines (`Sub`) that can be called by name from the spreadsheet after enabling macros. This is an example of a macro that will print out the values in Cells (1, 1) to (1, 8).

```java
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

*Note: I haven&rsquo;t found support for vba syntax highlight so I&rsquo;m telling the source block it&rsquo;s java code just for the colors.*

We can access cells with the `Cells` keyword or using `Range`. We can always reference the existing methods and properties in Microsoft&rsquo;s website <sup><a id="fnr.1" class="footref" href="#fn.1" role="doc-backlink">1</a></sup>.


# Reference


## Subroutine Declaration

```java
Sub MySubroutine()
    'code goes here
End Sub
```


## Variable Declaration

```java
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

```java
Dim my_object As Object
```

`Object` does not contain the data value itself, as it&rsquo;s a pointer to that value. <sup><a id="fnr.2" class="footref" href="#fn.2" role="doc-backlink">2</a></sup> So it always uses the same space in memory. It&rsquo;s recommended to always define a variable by a specific type rather than trying to point to it at runtime.


## Arrays

Arrays hold an arbitrary number of variables of the same type. Instead of creating many variables that share a type, we can create an array.

We index arrays with an integer `i`. Arrays in VBA are zero-based.

Declaring arrays.

```java
Dim tickers(11) As String
```

Accessing arrays.

```java
tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
```


## Strings

Concatenating strings.

```java
Range("A1").Value = "All Stocks (" + yearValue + ")"
```

Printing formatted strings.

```java
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
```


# Keywords and Operators

Here is some syntax reference for the common keywords in VBA:


## Conditional

If, Then.

```java
If 3 > 2 Then
    ' Code here
End If
```

Equal operator.

```java
Cells(i, 1).Value = "DQ"
```

Not equal operator.

```java
Cells(i - 1, 1).Value <> "DQ"
```

And keyword.

```java
If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
    'set starting price
End If
```

ElseIf.

```java
If Cells(4, 3) > 0 Then
    'Color the cell green
    Cells(4, 3).Interior.Color = vbGreen
ElseIf Cells(4, 3) < 0 Then
    'Color the cell red
    Cells(4, 3).Interior.Color = vbRed
Else
    'Clear the cell color
    Cells(4, 3).Interior.Color = xlNone
End If
```


## Loops

For loop.

```java
For i = 1 To 10
    ' Code here
Next i
```

Nested loops.

```java
For i = 1 To 10
    ' code here
    For j = 1 to 20
        ' code here
    Next j
Next i
```

Accessing arrays in loops.

```java
For i = 0 To 11
    ticker = tickers(i)
    ' Do stuff with ticker
Next i
```


## Visual Style Formatting

Styling.

```java
Range("A3:C3").Font.Bold = True
Columns("B").AutoFit
```

Number Formatting.

```java
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
```

Digits of Precision.

| Format              | Precision  |
|------------------- |---------- |
| &ldquo;0.0%&rdquo;  | one digit  |
| &ldquo;0.00%&rdquo; | two digits |

Colors.

```java
Cells(4, 3).Interior.Color = vbGreen
Cells(4, 3).Interior.Color = vbRed
Cells(4, 3).Interior.Color = xlNone
```


# Design Patterns

Patterns can be expressed as pseudocode, which means that we are only describing what the process looks like once we generalize it.

A good way to find them is to create a few cases by hand first and see if a pattern arises once we see them side by side. For example:

| Cases  | input | output |
|------ |----- |------ |
| Case 1 | 1     | 4      |
| Case 2 | 2     | 2      |
| Case 3 | 3     | 1.25   |
| Case 4 | 4     | 1      |

We can then assume that our output will be `output = 4/input`. So the pattern will be:

```java
Sub DivBy4()
    For i = 1 To 10
        Cells(i, 2).Value = 4 / Cells(i, 1).Value
    Next i
End Sub
```

The code will give us the result of any input we give it. So the value of cell 2 will be equal to 4 divided by the value of cell 1. This way we can operate in any number of cases. It&rsquo;s good to always test for outlier cases as our design needs to cover them too. For example, `we have to check that the input is not zero`.

Once our design is solid, we can reuse the code in other applications as long as it fits them. The more cases it can cover, the better.


# Doing Research

Another way to improve our design patterns is to look for answers that other people have come up with or simply look at the documentation for a better understanding of the features and rules of the programming language.

For example, we wouldn&rsquo;t be able to know what `xlUp` does without looking at the documentation or getting that information from a quick google search.

```java
rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
```

We can refer to Microsoft&rsquo;s documentation <sup><a id="fnr.1.100" class="footref" href="#fn.1" role="doc-backlink">1</a></sup> for more information about VBA.


# Developing a Macro

In order to create a proper algorithm/design, we will map out a plan of what we need to execute. Let&rsquo;s write down the an example of a general objective and then go step by step:

1.  Format our working sheet with headers.
2.  Initialize an array of existing items.
3.  Initialize variables based on our data.
4.  Initialize output variables with no value yet.
5.  Loop through our data.
6.  Write down the results

We have to translate that pseudocode into actual code, so now we go step by step. For the sake of simplicity, we will only visit the first step:

1.  Format our working sheet with headers:

```java
Worksheets("All Stocks Analysis").Activate
Range("A1").Value = "All Stocks (2018)"

Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"
```


# Running Macros with Buttons

Go to `Developer=>Button=>Select Area in worksheet=>Select your Macro`. Then you can edit your button name, look and feel, and then test it by clicking it.

Using user input.

```java
yearValue = InputBox("What year would you like to run the analysis on?")
```


# Debugging and Timing Code

It&rsquo;s a good idea to execute the macro everytime we complete a step and even make a commit to our repository to keep track of our progress and versions of our code.

Luckily, Excel can help us find problems in our code with its debugging tools. It&rsquo;s also not a bad idea to `google` the error messages to see possible solutions in case we don&rsquo;t see it right away.

Measuring timing.

```java
Dim startTime, endTime As Single
startTime = Timer
' Code here
endTime = Timer
totalTime = endTime - startTime
```

Performance will become much more key when working with larger datasets or performing more complex operations like machine learning. Other useful tool is profiling. <sup><a id="fnr.3" class="footref" href="#fn.3" role="doc-backlink">3</a></sup>

## Footnotes

<sup><a id="fn.1" class="footnum" href="#fnr.1">1</a></sup> <https://docs.microsoft.com/en-us/office/vba/api/overview/excel/graph-visual-basic-reference>

<sup><a id="fn.2" class="footnum" href="#fnr.2">2</a></sup> <https://en.wikipediarg/wiki/Pointer_(computer_programming)>

<sup><a id="fn.3" class="footnum" href="#fnr.3">3</a></sup> <https://en.wikipedia.org/wiki/Profiling_(computer_programming)>
