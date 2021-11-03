Imports EdmLib
Imports System.Windows.Forms
Public Class Datacard
    Dim fname As String
    Dim xl As Excel.Application
    Dim wb As Excel.Workbook
    Dim vault As EdmVault5
    Dim vault1 As IEdmVault8
    Dim vinfo() As EdmViewInfo
    Dim fol As IEdmFolder5
    Dim dfile As IEdmFile5
    Dim findex As Integer
    Dim i As Integer
    Dim vname As String
    Dim whnd As Long
    Dim view As IEdmCardView6
    Dim wid As Integer
    Dim ht As Integer
    Dim err As String

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook

        fname = wb.FullName
        vault = New EdmVault5
        vault1 = vault
        vault1.GetVaultViews(vinfo, False)
        findex = 0
        For Me.i = 0 To UBound(vinfo)
            If fname.Contains(vinfo(i).mbsPath) = True Then
                findex = 1
                vname = vinfo(i).mbsVaultName
                Exit For
            End If
        Next
        vault.LoginAuto(vname, 0)
        dfile = vault.GetFileFromPath(fname)
        fol = vault.GetFolderFromPath(wb.Path)
        Try
            'whnd = FindWindow(vbNullString, "EPDM Data Card")
            view = fol.CreateCardView(dfile.ID, 0, 0, 0, Nothing)
            If Not view Is Nothing Then
                view.GetCardSize(wid, ht)
                Me.Width = wid + 30
                Me.Height = ht + 100
                view.ShowWindow(True)
                TableLayoutPanel1.Left = Me.Width - 124
                TableLayoutPanel1.Top = Me.Height - 98
            End If
        Catch ex As Exception
            err = ex.Message
            MsgBox(err, MsgBoxStyle.Information)
            If err.Length > 0 Then
                Exit Sub
            End If
        End Try
        
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            xl = GetObject(, "Excel.Application")
            wb = xl.ActiveWorkbook
            wb.Close(Excel.XlSaveAction.xlDoNotSaveChanges)
            If view Is Nothing Then
                Exit Sub
            Else
                view.SaveData()
            End If
            xl.Workbooks.Open(fname)
        Catch ex As Exception
            err = ex.Message
            MsgBox(err, MsgBoxStyle.Information)
            If err.Length > 0 Then
                Exit Sub
            End If
        End Try
        
    End Sub

    Private Sub Datacard_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        TableLayoutPanel1.Left = Me.Width - 124
        TableLayoutPanel1.Top = Me.Height - 98
    End Sub

    Private Sub Datacard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ToolTip1.SetToolTip(Button1, "Show")
        ToolTip2.SetToolTip(Button2, "Save")
    End Sub
End Class