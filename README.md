# Consolidating Files with VBA
It's a VBA subroutine that will iterate through a set of files to bring in data from those files into a central spreadsheet.

# Data Structure
Each file in the 'test folder' has the temperature, pH, dO<sup>2</sup> concentration. 

# Result
When the user click the 'Import Data' button, a multiple of files in the 'test' folder can be selected to import, and then the example file will be open. After that, The user must select a column range that they want to summarise, without the header. The summary folder will automatically collect the data in the selected files and analyse them by showing the summarised data in the 'Data' sheet. The 'Reset' button will delete all the data in the sheet.

```VB
Option Explicit
Option Base 1

Sub RalphieReactor()
'Place your code here
Dim FileNames As Variant, nw As Integer
Dim ImportRange As String, UserRange As Range
Dim aWB As Workbook, tWB As Workbook
Dim time(), temp(), ph(), d02
Dim i As Integer

FileNames = Application.GetOpenFilename(FileFilter:="Excel Filter (*.csv), *.csv", Title:="Open File(s)", MultiSelect:=True)
Set tWB = ThisWorkbook

nw = UBound(FileNames)

Workbooks.Open FileNames(1)
Set UserRange = Application.InputBox("Select the range to import", Type:=8)
ImportRange = UserRange.Address
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = False

ReDim time(nw), temp(nw), ph(nw), d02(nw)

For i = 1 To nw
    Workbooks.Open FileNames(i)
    Set aWB = ActiveWorkbook

    time(i) = aWB.Sheets(1).Range(ImportRange).Cells(1, 1).Text
    temp(i) = aWB.Sheets(1).Range(ImportRange).Cells(2, 1).Value
    ph(i) = aWB.Sheets(1).Range(ImportRange).Cells(3, 1).Value
    d02(i) = aWB.Sheets(1).Range(ImportRange).Cells(4, 1).Value
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
