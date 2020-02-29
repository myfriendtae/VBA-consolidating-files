# Consolidating Csv Files with VBA
It's a VBA subroutine that will iterate through a set of files to bring in data from those files into a central spreadsheet.

## Data Structure
Each file in the 'test folder' has the temperature, pH and dO<sup>2</sup> concentration. 

## Introduction
When the user click the 'Import Data' button, a multiple of files in the 'test' folder can be selected to import, and then an example file will be open. After that, the user can select a column range that they want to summarise. Please select the range without the header - for example, "$B$3:$B$6". The summary folder will automatically collect and analyse data in the those files. The 'Reset' button will delete all the data in the sheet.

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
