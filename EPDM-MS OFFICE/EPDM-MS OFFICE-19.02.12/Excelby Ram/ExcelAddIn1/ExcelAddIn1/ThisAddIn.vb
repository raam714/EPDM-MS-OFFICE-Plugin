Imports Microsoft.Office.Interop.Excel
Imports EdmLib
Public Class ThisAddIn
    Public uc As New MyTaskPane
    Public tp As Microsoft.Office.Tools.CustomTaskPane
    Public mrg As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Dim i As Integer

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        For Me.i = 1 To Me.CustomTaskPanes.Count
            Me.CustomTaskPanes.Remove(Me.CustomTaskPanes(i - 1))
        Next
        AddTaskPane()

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Public ReadOnly Property MyTP As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return tp
        End Get
    End Property
    Private Sub AddTaskPane()
        uc = New MyTaskPane
        tp = Me.CustomTaskPanes.Add(uc, "EPDM", Me.Application.ActiveWindow)
        tp.Visible = True
    End Sub
    Public ReadOnly Property wb As Excel.Workbook
        Get
            Return Me.Application.ActiveWorkbook
        End Get
    End Property

    Private Sub Application_WorkbookOpen(Wb As Microsoft.Office.Interop.Excel.Workbook) Handles Application.WorkbookOpen
        
    End Sub
End Class
