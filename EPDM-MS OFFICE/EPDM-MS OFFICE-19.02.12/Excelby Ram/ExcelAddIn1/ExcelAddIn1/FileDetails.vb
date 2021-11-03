Imports EdmLib
Public Class FileDetails
    Dim fname As String
    Dim xl As Excel.Application
    Dim wb As Excel.Workbook
    Dim vault As EdmVault5
    Dim vault1 As IEdmVault8
    Dim vinfo() As EdmViewInfo
    Dim fol As IEdmFolder5
    Dim dfile As IEdmFile5
    Dim iuser As IEdmUser5
    Dim wfs As IEdmState5
    Dim lst As IEdmBatchListing
    Dim lfile() As EdmListFile
    Dim cols() As EdmListCol
    Dim i As Integer
    Dim vname As String
    Private Sub FileDetails_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
        lst = vault.CreateUtility(EdmUtility.EdmUtil_BatchList)
        lst.AddFile(fname, FileSystem.FileDateTime(fname), 0)
        lst.CreateList("", cols)
        lst.GetFiles(lfile)

        Me.Label8.Text = fname
        Me.Label9.Text = dfile.GetLocalVersionNo(fol.ID)
        Me.Label10.Text = dfile.CurrentVersion
        If dfile.IsLocked = True Then
            Me.Label11.Text = "Checked Out"
        Else
            Me.Label11.Text = "Checked In"
        End If
        iuser = dfile.LockedByUser
        If dfile.IsLocked = True Then
            Me.Label12.Text = iuser.Name
        Else
            Me.Label12.Text = "-"
        End If
        Me.Label13.Text = lfile(0).moCurrentState.mbsWorkflowName
        wfs = dfile.CurrentState
        Me.Label14.Text = wfs.Name

    End Sub
End Class