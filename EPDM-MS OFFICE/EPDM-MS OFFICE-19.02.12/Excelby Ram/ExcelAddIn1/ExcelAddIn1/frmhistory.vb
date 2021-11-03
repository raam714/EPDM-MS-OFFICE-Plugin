Imports EdmLib
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmhistory
    Dim fname As String
    Dim xl As Excel.Application
    Dim wb As Excel.Workbook
    Dim vault As EdmVault5
    Dim vault1 As IEdmVault8
    Dim vinfo() As EdmViewInfo
    Dim fol As IEdmFolder5
    Dim dfile As IEdmFile5
    Dim bhis As IEdmHistory
    Dim fhis() As EdmHistoryItem
    Dim i As Integer
    Dim vname As String
    Private Sub frmhistory_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook

        fname = wb.FullName
        vault = New EdmVault5
        vault1 = vault
        vault1.GetVaultViews(vinfo, False)

        For Me.i = 0 To UBound(vinfo)
            If fname.Contains(vinfo(i).mbsPath) = True Then
                vname = vinfo(i).mbsVaultName
                Exit For
            End If
        Next

        vault.LoginAuto(vname, 0)
        fol = vault.GetFolderFromPath(wb.Path)
        dfile = vault.GetFileFromPath(fname, fol)
        bhis = vault.CreateUtility(EdmUtility.EdmUtil_History)
        bhis.AddFile(dfile.ID)
        bhis.GetHistory(fhis)

        ListView1.View = View.Details
        ListView1.GridLines = True
        ListView1.FullRowSelect = True
        ListView1.HideSelection = False
        ListView1.MultiSelect = False

        ListView1.Columns.Add("Event")
        ListView1.Columns.Add("Version")
        ListView1.Columns.Add("User")
        ListView1.Columns.Add("Date")
        ListView1.Columns.Add("Comment")
        ListView1.Columns(3).Width = 150
        ListView1.Columns(4).Width = 300

        For Me.i = 0 To UBound(fhis)
            Dim lstvi As New ListViewItem

            lstvi.Text = fhis(i).meType.ToString()
            lstvi.SubItems.Add(fhis(i).mlVersion)
            lstvi.SubItems.Add(fhis(i).mbsUserName)
            lstvi.SubItems.Add(fhis(i).moDate)
            lstvi.SubItems.Add(fhis(i).mbsComment)
            If i Mod 2 = 0 Then
                lstvi.BackColor = Color.LightYellow
            End If
            ListView1.Items.Add(lstvi)
        Next

    End Sub
End Class