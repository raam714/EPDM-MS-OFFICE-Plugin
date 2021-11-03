Public NotInheritable Class AboutBox1

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Set the title of the form.
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        'Dim target As String
        'target = CType(e.Link.LinkData, String)
        'System.Diagnostics.Process.Start(target)
        Dim myprocess As New System.Diagnostics.Process

        myprocess.StartInfo = New System.Diagnostics.ProcessStartInfo("iexplore")
        myprocess.StartInfo.Arguments = "http://www.egsindia.com/"
        myprocess.Start()
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        'Dim target As String
        'target = CType(e.Link.LinkData, String)
        'System.Diagnostics.Process.Start(target)

        Dim myprocess As New System.Diagnostics.Process

        myprocess.StartInfo = New System.Diagnostics.ProcessStartInfo("iexplore")
        myprocess.StartInfo.Arguments = "http://www.egs.co.in/"
        myprocess.Start()
    End Sub
End Class
