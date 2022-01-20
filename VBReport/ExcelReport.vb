Imports Microsoft.Office.Interop.Excel

Public Class ExcelReport

    ' Properties

    Private xlApp As Application

    ' Constructor

    Public Sub New()
        xlApp = New Application()
        xlApp.Visible = True
    End Sub

    ' Public Methods

    Public Sub Create()
        Dim wb As Workbook = CreateNewWorkbook(1)
        Dim ws As Worksheet = GetWorksheet(wb, 1)
        ValueOutputTest(ws)
    End Sub

    ' Private Methods

    Private Function CreateNewWorkbook(numberOfSheets As Integer)
        Dim originalNumberOfSheets As Integer = xlApp.SheetsInNewWorkbook
        xlApp.SheetsInNewWorkbook = numberOfSheets
        Dim wb As Workbook = xlApp.Workbooks.Add()
        xlApp.SheetsInNewWorkbook = originalNumberOfSheets
        Return wb
    End Function

    Private Function GetWorksheet(wb As Workbook, sheetIndex As Integer)
        If (wb Is Nothing) Then Return Nothing
        If (sheetIndex > 0 And sheetIndex <= wb.Sheets.Count) Then
            Return wb.Sheets(sheetIndex)
        End If
        Return Nothing
    End Function

    Private Sub ValueOutputTest(ws As Worksheet)
        If (ws Is Nothing) Then Return
        ws.Name = "VB Test"

        ' I changed the ouputorder in comparison with the C# test,
        ' so that it can be seen that option 1 is working.

        ' Output Option 2 -> working
        ws.Range(ws.Cells(1, 2), ws.Cells(1, 2)).Value = "Cell R1C2"
        ' Output Option 1 -> throws System.MissingMemberException: "Public member 'Value' on type 'Range' not found."
        ws.Cells(1, 1).Value = "Cell R1C1"
    End Sub


End Class
