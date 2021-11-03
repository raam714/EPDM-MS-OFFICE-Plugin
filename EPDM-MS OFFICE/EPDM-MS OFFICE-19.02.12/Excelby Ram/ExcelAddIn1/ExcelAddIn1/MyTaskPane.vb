Public Class MyTaskPane

    Private Sub MyTaskPane_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Label3.Text = Globals.ThisAddIn.wb.FullName
    '    Label3.Update()
    'End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Label3.Text = Globals.ThisAddIn.wb.FullName
        Label3.Update()
    End Sub
End Class
