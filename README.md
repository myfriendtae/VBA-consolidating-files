# Consolidating csv. Files with VBA
VBA programme that iteratse through a set of files to load files in a folder, and collect/manipulate data, and save in a central spreadsheet. When a user clicks the 'Import Data' button, a multiple of files in the 'test' folder are imported, and then an example file will be open. After that, the user can select a column range that they want to summarise. The range must be selected without the header - for example, "$B$3:$B$6" in the screenshot below. The 'Reset' button will delete all the data in the sheet.


## Data Structure
Each file in the 'test folder' has the temperature, pH and dO<sup>2</sup> concentration. 

![screenshot](https://github.com/myfriendtae/VBA_consolidating_files/blob/master/screenshot.png?raw=true)

## Codes
```VB
Option Explicit
Option Base 1

Sub RalphieReactor()
Dim FileNames As Variant, nw As Integer
Dim ImportRange As String, UserRange As Range
Dim aWB As Workbook, tWB As Workbook
Dim time(), temp(), ph(), dO2()
Dim i As Integer

FileNames = Application.GetOpenFilename(FileFilter:="Excel Filter (*.csv), *.csv", Title:="Open File(s)", MultiSelect:=True)
Set tWB = ThisWorkbook

nw = UBound(FileNames)

Workbooks.Open FileNames(1)
Set UserRange = Application.InputBox("Select the range to import", Type:=8)
ImportRange = UserRange.Address
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = False

ReDim time(nw), temp(nw), ph(nw), dO2(nw)

For i = 1 To nw
    Workbooks.Open FileNames(i)
    Set aWB = ActiveWorkbook

    time(i) = aWB.Sheets(1).Range(ImportRange).Cells(1, 1).Text
    temp(i) = aWB.Sheets(1).Range(ImportRange).Cells(2, 1).Value
    ph(i) = aWB.Sheets(1).Range(ImportRange).Cells(3, 1).Value
    dO2(i) = aWB.Sheets(1).Range(ImportRange).Cells(4, 1).Value
    aWB.Close SaveChanges:=False

Next i

i = 1

For i = 1 To nw
    Range("A" & i) = time(i)
    Range("B" & i) = temp(i)
    Range("C" & i) = ph(i)
    Range("D" & i) = d02(i)
Next i
End Sub

Sub Reset()
Cells.Clear
End Sub

```
