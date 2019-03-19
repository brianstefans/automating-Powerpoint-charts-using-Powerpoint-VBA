# automating-Powerpoint-tables-using-Powerpoint-VBA
Being able to update the chart data in PowerPoint with a click of a VBA button. The script retrieves data from an Excel workbook then updates the respective chart data

```VBA
   Sub update_tables()
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
   'Dim lastrow As Long
   'starttime = Timer

   'Update_display (False)
   'Application.Calculation = xlCalculateManual
   FileName = "...\vba\test.xlsx"

   'Set mypptApp = CreateObject("PowerPoint.Application")
   'Set myppt = mypptApp.Presentations.Open(FileName, WithWindow:=msoFalse)
   Set myppt = ActivePresentation


   'set the worksheet where to retrieve the data from
   Set xlApp = CreateObject("Excel.Application")
   Set wb = xlApp.Workbooks.Open(FileName, False, True)
   Set ws = wb.Worksheets(5)
   week = InputBox("Input the week")
   For i = 1 To 20
      If ws.Cells(1, i) = week Then
                  Set rng3 = ws.Range(ws.Cells(1, i), ws.Cells(482, i))
      End If
   Next

   Set mypptslide = myppt.Slides
   Set test1 = mypptslide(108).Shapes(2).table

   y = 0
   For k = 108 To 145
    For Each shape In myppt.Slides(k).Shapes

      'For Each shape In slide.Shapes
       If shape.HasTable Then
         Set table = shape.table
         table.Columns.Add
         ii = table.Columns.count
         r = table.Rows.count
         table.Cell(1, ii).shape.TextFrame.TextRange.Text = week
            For j = 2 To r

                   With table
                    .Cell(j, ii).shape.TextFrame.TextRange.Text = Format(rng3.Cells(j + y, 1).Value2, "0%")
                   End With

            Next
        y = y + r
        y = y - 1
        End If
       'Next

    Next
   Next
   MsgBox v & " tables not updated", vbInformation
   'myppt.Save
   Set wb = Nothing
   xlApp.Quit
   End Sub

```








