Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports EdmLib
Imports Microsoft.Win32
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Configuration
Imports System.Data
Imports System.Diagnostics

Imports System.IO
Imports System.Security.Permissions
Imports System.Threading


Public Class MyRibbon
    Dim fname As String
    Dim xl As Excel.Application
    Dim wb As Excel.Workbook
    Dim vault As EdmVault5
    Dim vault1 As IEdmVault8
    Dim vinfo() As EdmViewInfo
    Dim fol As IEdmFolder5
    Dim dfile As IEdmFile5
    Dim verenum As IEdmEnumeratorVersion5
    Dim vers As IEdmVersion5
    Dim view As IEdmCardView5
    Dim pos As IEdmPos5
    Dim statpos As IEdmPos5
    Dim wfmgr As IEdmWorkflowMgr6
    Dim wrkmrg As IEdmWorkflowMgr6
    Dim wfstst As IEdmState6
    Dim wf As IEdmWorkflow6
    Dim wfstat As IEdmState6
    Dim wftrans As IEdmTransition5
    Dim bul As IEdmBatchUnlock
    Dim bget As IEdmBatchGet
    Dim sel() As EdmSelItem
    'Dim sel(1) As Integer
    Dim cont As IEdmRefItemContainer
    Dim poitem As IEdmRefItem
    Dim btn As RibbonButton
    Dim id As Integer
    Dim vitem() As Object
    Dim folpath As String
    Dim err As String
    Dim vpath As Object
    Dim i As Integer
    Dim findex As Integer
    Dim vname As String
    Dim bret As Boolean
    Dim lst As IEdmBatchListing
    Dim lfile() As EdmListFile
    Dim cols() As EdmListCol
    Private Sub MyRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btncin_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btncin.Click
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                fol = vault.GetFolderFromPath(wb.Path)
                dfile = vault.GetFileFromPath(fname, fol)
                If dfile.IsLocked = True Then
                    Try
                        wb.Close(Excel.XlSaveAction.xlDoNotSaveChanges)
                        bul = vault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock)
                        ReDim sel(0)
                        sel(0).mlDocID = dfile.ID
                        sel(0).mlProjID = fol.ID

                        bul.AddSelection(vault, sel)

                        If bul.CreateTree(0, EdmUnlockBuildTreeFlags.Eubtf_MayUnlock) = False Then
                            Exit Sub
                        End If
                        cont = bul
                        cont.GetItems(EdmRefItemType.Edmrit_All, vitem)
                        id = LBound(vitem)
                        While id <= UBound(vitem)
                            poitem = vitem(id)
                            poitem.SetProperty(EdmRefItemProperty.Edmrip_CheckKeepLocked, False)
                            id = id + 1
                        End While
                        
                        If bul.ShowDlg(0) = True Then
                            bul.UnlockFiles(0)
                        End If

                        xl.Workbooks.Open(fname)

                    Catch ex As Exception
                        err = ex.Message
                        MsgBox(err, MsgBoxStyle.Information)
                        If err.Length > 0 Then
                            Exit Sub
                        End If
                    End Try
                Else
                    MsgBox("The file was already Checked In!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
                End If
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
        Call refresh_load()
    End Sub

    Private Sub Button16_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button16.Click
        Dim aboutform As New AboutBox1
        aboutform.ShowDialog()
    End Sub

    

    Private Sub btncout_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btncout.Click
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                fol = vault.GetFolderFromPath(wb.Path)
                dfile = vault.GetFileFromPath(fname, fol)
                If dfile.IsLocked = False Then
                    Try
                        wb.Close(Excel.XlSaveAction.xlDoNotSaveChanges)
                        bget = vault.CreateUtility(EdmUtility.EdmUtil_BatchGet)
                        ReDim sel(0)
                        sel(0).mlDocID = dfile.ID
                        sel(0).mlProjID = fol.ID

                        bget.AddSelection(vault, sel)

                        bget.CreateTree(0, EdmGetCmdFlags.Egcf_Lock)

                        If bget.ShowDlg(0) = True Then
                            bget.GetFiles(0)
                        End If

                        xl.Workbooks.Open(fname)

                    Catch ex As Exception
                        err = ex.Message
                        MsgBox(err, MsgBoxStyle.Information)
                        If err.Length > 0 Then
                            Exit Sub
                        End If
                    End Try
                Else
                    MsgBox("The file was already Checked Out!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
                End If
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
        Call refresh_load()
    End Sub

    Private Sub btnucout_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnucout.Click
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                fol = vault.GetFolderFromPath(wb.Path)
                dfile = vault.GetFileFromPath(fname, fol)
                If dfile.IsLocked = True Then
                    Try
                        wb.Close(Excel.XlSaveAction.xlDoNotSaveChanges)
                        bul = vault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock)
                        ReDim sel(0)
                        sel(0).mlDocID = dfile.ID
                        sel(0).mlProjID = fol.ID

                        bul.AddSelection(vault, sel)

                        If bul.CreateTree(0, EdmUnlockBuildTreeFlags.Eubtf_MayUndoLock) = False Then
                            Exit Sub
                        End If
                        cont = bul
                        cont.GetItems(EdmRefItemType.Edmrit_All, vitem)
                        id = LBound(vitem)
                        While id <= UBound(vitem)
                            poitem = vitem(id)
                            poitem.SetProperty(EdmRefItemProperty.Edmrip_CheckKeepLocked, False)
                            id = id + 1
                        End While

                        If bul.ShowDlg(0) = True Then
                            bul.UnlockFiles(0)
                        End If

                        xl.Workbooks.Open(fname)

                    Catch ex As Exception
                        err = ex.Message
                        MsgBox(err, MsgBoxStyle.Information)
                        If err.Length > 0 Then
                            Exit Sub
                        End If
                    End Try
                Else
                    MsgBox("The file was already Checked In!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
                End If
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
        Call refresh_load()
    End Sub

    Private Sub btnverinfo_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnverinfo.Click
        Dim verform As Form

        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                verform = New VersionInfo
                verform.ShowDialog()
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
    End Sub

    Private Sub btnlatver_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnlatver.Click
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                fol = vault.GetFolderFromPath(wb.Path)
                dfile = vault.GetFileFromPath(fname, fol)
                If dfile.IsLocked = False Then
                    Try
                        wb.Close(Excel.XlSaveAction.xlDoNotSaveChanges)
                        bget = vault.CreateUtility(EdmUtility.EdmUtil_BatchGet)
                        ReDim sel(0)
                        sel(0).mlDocID = dfile.ID
                        sel(0).mlProjID = fol.ID

                        bget.AddSelection(vault, sel)

                        bget.CreateTree(0, EdmGetCmdFlags.Egcf_Nothing)

                        If bget.ShowDlg(0) = True Then
                            bget.GetFiles(0)
                        End If

                        xl.Workbooks.Open(fname)

                    Catch ex As Exception
                        err = ex.Message
                        MsgBox(err, MsgBoxStyle.Information)
                        If err.Length > 0 Then
                            Exit Sub
                        End If
                    End Try
                Else
                    MsgBox("The file was already Checked Out!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
                End If
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
        cmbver.Items.Clear()
        Call refresh_load()
    End Sub


    Private Sub btnref_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnref.Click
     
        Call refresh_load()
    End Sub

    Private Sub cmbver_TextChanged_1(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles cmbver.TextChanged
        Dim vno As Integer

        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                fol = vault.GetFolderFromPath(wb.Path)
                dfile = vault.GetFileFromPath(fname, fol)
                Try
                    wb.Close(Excel.XlSaveAction.xlDoNotSaveChanges)
                    verenum = dfile
                    vno = Split(cmbver.Text, ".")(0)
                    vers = verenum.GetVersion(vno)
                    vers.GetFileCopy(0)
                    xl.Workbooks.Open(fname)
                Catch ex As Exception
                    err = ex.Message
                    MsgBox(err, MsgBoxStyle.Information)
                    If err.Length > 0 Then
                        Exit Sub
                    End If
                End Try
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If

    End Sub


    Private Sub btninbox_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btninbox.Click
        Dim reg As RegistryKey
        Dim kval As String
        Dim mailpath As String
        Dim vstr() As String

        kval = "SOFTWARE\\SolidWorks\\Applications\\PDMWorks Enterprise"
        reg = Registry.LocalMachine.OpenSubKey(kval)
        If Not reg Is Nothing Then
            kval = reg.GetValue("LanFileDir", 0, RegistryValueOptions.None)
            reg.Close()
            vstr = Split(kval, "\")
            mailpath = vstr(0)
            For Me.i = 1 To (UBound(vstr) - 2)
                mailpath = mailpath & "\" & vstr(i)
            Next
            mailpath = mailpath & "\Inbox.exe"
            'MsgBox(mailpath)
            System.Diagnostics.Process.Start(mailpath)
        End If
    End Sub

    Private Sub btnwflow_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnwflow.Click
        Dim wfs As Form

        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
               
                wfs = New WFStatus


                wfs.ShowDialog()
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
    End Sub

    Private Sub btndetails_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btndetails.Click
        Dim getdetail As Form

        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                getdetail = New FileDetails
                getdetail.ShowDialog()
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
    End Sub

    Private Sub btnhistory_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnhistory.Click
        Dim filehistory As Form

        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                filehistory = New frmhistory
                filehistory.ShowDialog()
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
    End Sub

    Private Sub btnbrofile_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnbrofile.Click
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
            fname = wb.FullName
            System.Diagnostics.Process.Start("explorer.exe", "/select," & fname)
        End If
    End Sub

    Private Sub btnshowcard_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnshowcard.Click
        Dim showcard As Form

        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        If wb Is Nothing Then
            MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        Else
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
            If findex = 1 Then
                vault.LoginAuto(vname, 0)
                showcard = New Datacard
                showcard.ShowDialog()
            Else
                MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            End If
        End If
    End Sub

    'Private Sub Button12_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button12.Click
    '    Diagnostics.Process.Start("D:\Program Files\SolidWorks Enterprise PDM\Search.exe")
    'End Sub

    Private Sub btnsearch_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnsearch.Click
        Dim reg As RegistryKey
        Dim kval As String
        Dim searchpath As String
        Dim vstr() As String

        kval = "SOFTWARE\\SolidWorks\\Applications\\PDMWorks Enterprise"
        reg = Registry.LocalMachine.OpenSubKey(kval)
        If Not reg Is Nothing Then
            kval = reg.GetValue("LanFileDir", 0, RegistryValueOptions.None)
            reg.Close()
            vstr = Split(kval, "\")
            searchpath = vstr(0)
            For Me.i = 1 To (UBound(vstr) - 2)
                searchpath = searchpath & "\" & vstr(i)
            Next
            searchpath = searchpath & "\Search.exe"
            'MsgBox(mailpath)
            System.Diagnostics.Process.Start(searchpath)
        End If
    End Sub

    Private Sub btn_OnClick(sender As Object, e As RibbonControlEventArgs)
        'Throw New NotImplementedException
        Dim bcstate As IEdmBatchChangeState
        Dim butn As RibbonButton
        Dim csfile As IEdmFile5
        Dim csfol As IEdmFolder5
        Dim boolstat As Boolean


        'MsgBox(butn.Label)

        bcstate = vault.CreateUtility(EdmUtility.EdmUtil_BatchChangeState)
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook
        csfile = vault.GetFileFromPath(wb.FullName)
        If csfile.IsLocked = False Then
            csfol = vault.GetFolderFromPath(wb.Path)
            bcstate.AddFile(csfile.ID, csfol.ID)
            butn = DirectCast(sender, RibbonButton)
            bcstate.CreateTree(butn.Label)
            boolstat = bcstate.ShowDlg(0)
            If boolstat = True Then
                bcstate.ChangeState(0)
            End If
        Else
            MsgBox("The file is Checked Out by " & csfile.LockedByUser.Name & "!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
        End If
        Call refresh_load()
    End Sub
    Public Sub refresh_load()
        Dim ddi As RibbonDropDownItem
        Try
            xl = GetObject(, "Excel.Application")
            wb = xl.ActiveWorkbook
            If wb Is Nothing Then
                cmbver.Items.Clear()
                MsgBox("There is no active document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
            Else
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
                If findex = 1 Then
                    vault.LoginAuto(vname, 0)
                    fol = vault.GetFolderFromPath(wb.Path)
                    dfile = vault.GetFileFromPath(fname, fol)
                    lst = vault.CreateUtility(EdmUtility.EdmUtil_BatchList)
                    lst.AddFile(fname, FileSystem.FileDateTime(fname), 0)
                    lst.CreateList("", cols)
                    lst.GetFiles(lfile)
                    a = lfile(0).moCurrentState.mbsWorkflowName
                    verenum = dfile
                    cmbver.Items.Clear()
                    For Me.i = 0 To (dfile.CurrentVersion - 1)
                        vers = verenum.GetVersion(i + 1)
                        ddi = Globals.Factory.GetRibbonFactory.CreateRibbonDropDownItem
                        If i = 0 Then
                            ddi.Label = (i + 1) & ". " & "<Created>"
                        Else
                            If vers.Comment.ToString = "" Then
                                ddi.Label = (i + 1) & ". " & "<No Comment>"
                            Else
                                ddi.Label = (i + 1) & ". " & vers.Comment.ToString
                            End If

                        End If
                        cmbver.Items.Add(ddi)
                    Next i
                    '---------------- Transition ----------'

                    wfstat = dfile.CurrentState
                    pos = wfstat.GetFirstTransitionPosition
                    Dim idx As Integer = -1
                    mnucstat.Items.Clear()
                    While Not pos.IsNull
                        wftrans = wfstat.GetNextTransition(pos)
                        'MsgBox(wftrans.Name)
                        idx = idx + 1
                        btn = Me.Factory.CreateRibbonButton
                        btn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular
                        btn.Label = wftrans.Name
                        btn.Name = "mnubtn_" & btn.Id
                        AddHandler btn.Click, AddressOf btn_OnClick
                        mnucstat.Items.Add(btn)
                    End While

                Else
                    cmbver.Items.Clear()
                    mnucstat.Items.Clear()
                    MsgBox("This is not a vault document!!", MsgBoxStyle.Information, "EPDM-MS Office Connector")
                End If
            End If
        Catch ex As Exception
            err = ex.Message
            MsgBox(err, MsgBoxStyle.Information)
            If err.Length > 0 Then
                Exit Sub
            End If
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button15.Click
        'Dim p As New Process
        'Dim psi As New ProcessStartInfo("E:\2013\EPDM-MS OfFice HELP\epdm-ms-office-connector_tmphhp\epdm-ms-office-connector.chm")
        'psi.Verb = "Open"
        'p.StartInfo = psi
        'p.Start()
        Dim reg As RegistryKey
        Dim kval As String
        Dim searchpath As String
        Dim vstr() As String

        kval = "SOFTWARE\\Microsoft\\HTMLHelp\\EGS.ExcelHelp"
        reg = Registry.LocalMachine.OpenSubKey(kval)
        If Not reg Is Nothing Then
            kval = reg.GetValue("Manifest", 0, RegistryValueOptions.None)
            reg.Close()
            vstr = Split(kval, "\")
            searchpath = vstr(0)
            For Me.i = 1 To (UBound(vstr) - 1)
                searchpath = searchpath & "\" & vstr(i)
            Next
            searchpath = searchpath & "\epdm-ms-office-connector.chm"
            'MsgBox(mailpath)
            System.Diagnostics.Process.Start(searchpath)
        End If
    End Sub
End Class
