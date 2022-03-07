Attribute VB_Name = "testMonster"
Option Explicit

Sub SplitX()

'''''' change File_Name to the full path and name of the file "C:\adada"
'''''' Getting input from user
MkDir ActiveWorkbook.Path & "\" & "Split"
Dim vColumn As String
vColumn = InputBox("Which column? (A, B ,C ..?)")

Application.ScreenUpdating = False

Dim NewFile
''''''A) if u want choose file by yourself
''''''NewFile = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")

''''' B) if you want specify one name of file
NewFile = ActiveWorkbook.Path & "\" & "data.xlsx"


If NewFile <> False Then
Workbooks.Open (NewFile)



Dim HomeWB
Dim HomeSheet

HomeWB = ActiveWorkbook.Name
HomeSheet = ActiveSheet.Name
Dim NewWB
Dim NewSheet

Dim HelperSheet As String



Columns(vColumn).Copy


'''''' Helper sheet creation
Sheets.Add
ActiveSheet.Name = "HelperSheet"
Range("A1").PasteSpecial
Columns("A").RemoveDuplicates Columns:=1, Header:=xlYes

'''''' Counter, it can be broken?, main foor loop
Dim vCounter As Integer
vCounter = Range("A" & Rows.Count).End(xlUp).Row

Dim i As Integer
Dim vfilter
For i = 2 To vCounter
    '''''' vfilter contains unique value from specified column
    vfilter = Sheets("HelperSheet").Cells(i, 1)
    Sheets(HomeSheet).Activate
    ActiveSheet.Columns.AutoFilter field:=Columns(vColumn).Column, Criteria1:=vfilter
    ''''''Cells.Copy
    ActiveSheet.UsedRange.Copy
    Workbooks.Add
    Range("A1").PasteSpecial
    
    ''''''save
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\split\" & vfilter
    ActiveWorkbook.Close False
    ActiveSheet.AutoFilterMode = False
Next i

ActiveSheet.AutoFilterMode = False

''''''Remove helper without prompt
Application.DisplayAlerts = False
Sheets("HelperSheet").Delete
Application.DisplayAlerts = True


ActiveWorkbook.Close False
Application.ScreenUpdating = True
MsgBox "Your files are under /split directory"
End If
End Sub
