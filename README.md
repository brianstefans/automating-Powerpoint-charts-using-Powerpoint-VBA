# automating-Powerpoint-graphs-using-Powerpoint-VBA
Being able to update the chart data in PowerPoint with a click of a VBA button. The script retrieves data from an Excel workbook then updates the respective chart data
```vba
Sub update_graph()

Dim myppt As PowerPoint.Presentation
Dim mypptslide As PowerPoint.Slides
Dim mypptshape As PowerPoint.shape
Dim mypptchrt As PowerPoint.chart
Dim chart As PowerPoint.chart
Dim mypptchrtdta As PowerPoint.ChartData
Dim cTable As Object
Dim mypptApp As PowerPoint.Application
Dim wb As Excel.Workbook
Dim ws As Excel.Worksheet
Dim shape As Object
Dim slide As Object
Dim FileName As String
Dim lastrow As Range
Dim rng As Range
Dim rng2 As Range
Dim xlApp As Excel.Application
Dim starttime As Double
Dim minselapsed As String
Dim i As Integer
Dim v As Integer
Dim r As Integer
Dim j As Integer
Dim k As Integer
Dim y As Integer
Dim ii As Integer
Dim count As Integer
Dim week As Variant
Dim iterator As Variant

'Update_display (False)
'Application.Calculation = xlCalculateManual
FileName = "...\vba\test.xlsx"

'Set mypptApp = CreateObject("PowerPoint.Application")
'Set myppt = mypptApp.Presentations.Open(FileName, WithWindow:=msoFalse)
Set myppt = ActivePresentation


'set the worksheet where to retrieve the data from
Set xlApp = CreateObject("Excel.Application")
Set wb = xlApp.Workbooks.Open(FileName, False, True)
'graphs
Set ws = wb.Worksheets(1)

'initialize the counter
i = 2
count = 0
v = 0
For Each slide In myppt.Slides

    For Each shape In slide.Shapes
      
        If shape.HasChart Then
        Set chart = shape.chart
        Set mypptchrtdta = chart.ChartData
        'mypptchrtdta.Activate
        

        
        'range of the datato be pasted
        Set rng = ws.ListObjects(1).Range.Cells(i, 3)
            
        Set rng2 = ws.ListObjects(1).Range.Cells(i, 4)
            If IsEmpty(rng2.Value) And IsEmpty(rng.Value) Then 'GoTo Line1
            '(rng2 & vbNullString) > 0
            v = v + 1 'MsgBox "empty chart"
            ElseIf Not IsEmpty(rng2.Value) And Not IsEmpty(rng.Value) Then
                count = count + 1
            
                    On Error GoTo err_handling
                    mypptchrtdta.Activate
                    'pasting the data
                    With mypptchrtdta
                         '.Workbook.Application.Visible = False
                         '.Activate
                         '.Workbook.Application.WindowState = -4140
                         .Workbook.Worksheets(1).ListObjects("Table1").ListRows.Add AlwaysInsert:=True
                         .Workbook.Worksheets(1).ListObjects("Table1").Range.End(xlDown).Offset(1, 0).Value = rng.Value
                         .Workbook.Worksheets(1).ListObjects("Table1").Range.End(xlDown).Offset(0, 1).Value = rng2.Value
                         '.Workbook.Close
                    End With
                    mypptchrtdta.Workbook.Close
                         
                    'cTable.ListRows.Add AlwaysInsert:=True
                    'cTable.Range.End(xlDown).Offset(1, 0).Value = rng.Value
                    'cTable.Range.End(xlDown).Offset(0, 1).Value = rng2.Value
                    'mypptchrtdta.Workbook.Close
                    

            
'Line1:
            End If
        i = i + 1
        End If
      
    Next

Next

i = i - 1

err_handling:
                    On Error GoTo -1
                    Err.Clear
                    'Resume Line1

Set ws = Nothing
MsgBox v & " charts not updated", vbInformation
Set wb = Nothing
xlApp.Quit
End Sub

```








