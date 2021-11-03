Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports EdmLib
Public Class VersionInfo
    Dim fname As String
    Dim xl As Excel.Application
    Dim wb As Excel.Workbook
    Dim vault As EdmVault5
    Dim vault1 As IEdmVault8
    Dim vinfo() As EdmViewInfo
    Dim fol As IEdmFolder5
    Dim dfile As IEdmFile5
    Dim i As Integer
    Dim vname As String

    Private Sub VersionInfo_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
        Me.Label2.Text = dfile.CurrentVersion
        Me.Label4.Text = dfile.GetLocalVersionNo(fol.ID)

        If dfile.CurrentVersion = dfile.GetLocalVersionNo(fol.ID) Then
            PictureBox1.Visible = False
        End If

        Chart1.Series("CV").Points.Add(dfile.CurrentVersion)
        Chart1.Series("LV").Points.Add(dfile.GetLocalVersionNo(fol.ID))

    End Sub

   
End Class