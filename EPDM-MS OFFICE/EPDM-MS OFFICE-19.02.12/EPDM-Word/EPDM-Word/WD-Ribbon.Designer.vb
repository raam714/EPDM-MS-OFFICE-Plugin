Partial Class WD_Ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.EPDM = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.cmbver = Me.Factory.CreateRibbonComboBox
        Me.btncin = Me.Factory.CreateRibbonButton
        Me.btncout = Me.Factory.CreateRibbonButton
        Me.btnucout = Me.Factory.CreateRibbonButton
        Me.btnlatver = Me.Factory.CreateRibbonButton
        Me.btnverinfo = Me.Factory.CreateRibbonButton
        Me.mnucstat = Me.Factory.CreateRibbonMenu
        Me.btnwflow = Me.Factory.CreateRibbonButton
        Me.btninbox = Me.Factory.CreateRibbonButton
        Me.btnhistory = Me.Factory.CreateRibbonButton
        Me.btngdetails = Me.Factory.CreateRibbonButton
        Me.btndcard = Me.Factory.CreateRibbonButton
        Me.btnbtfile = Me.Factory.CreateRibbonButton
        Me.btnsearch = Me.Factory.CreateRibbonButton
        Me.btnrefresh = Me.Factory.CreateRibbonButton
        Me.btnhelp = Me.Factory.CreateRibbonButton
        Me.btnabt = Me.Factory.CreateRibbonButton
        Me.EPDM.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        '
        'EPDM
        '
        Me.EPDM.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.EPDM.Groups.Add(Me.Group1)
        Me.EPDM.Groups.Add(Me.Group2)
        Me.EPDM.Groups.Add(Me.Group3)
        Me.EPDM.Groups.Add(Me.Group4)
        Me.EPDM.Groups.Add(Me.Group5)
        Me.EPDM.Label = "EPDM"
        Me.EPDM.Name = "EPDM"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btncin)
        Me.Group1.Items.Add(Me.btncout)
        Me.Group1.Items.Add(Me.btnucout)
        Me.Group1.Label = "Vault Actions"
        Me.Group1.Name = "Group1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnlatver)
        Me.Group2.Items.Add(Me.cmbver)
        Me.Group2.Items.Add(Me.btnverinfo)
        Me.Group2.Label = "Version Management"
        Me.Group2.Name = "Group2"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.mnucstat)
        Me.Group3.Items.Add(Me.btnwflow)
        Me.Group3.Items.Add(Me.btninbox)
        Me.Group3.Label = "Work Flow Management"
        Me.Group3.Name = "Group3"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnhistory)
        Me.Group4.Items.Add(Me.btngdetails)
        Me.Group4.Items.Add(Me.btndcard)
        Me.Group4.Items.Add(Me.btnbtfile)
        Me.Group4.Label = "Document Information"
        Me.Group4.Name = "Group4"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.btnsearch)
        Me.Group5.Items.Add(Me.btnrefresh)
        Me.Group5.Items.Add(Me.btnhelp)
        Me.Group5.Items.Add(Me.btnabt)
        Me.Group5.Label = "General"
        Me.Group5.Name = "Group5"
        '
        'cmbver
        '
        Me.cmbver.Label = "Get Version"
        Me.cmbver.Name = "cmbver"
        Me.cmbver.ScreenTip = "Get Version"
        Me.cmbver.SuperTip = "Get required version of current document"
        '
        'btncin
        '
        Me.btncin.Image = Global.EPDM_Word.My.Resources.Resources.vault_checkin
        Me.btncin.Label = "Check In"
        Me.btncin.Name = "btncin"
        Me.btncin.ScreenTip = "Check In"
        Me.btncin.ShowImage = True
        Me.btncin.SuperTip = "Check in document into Vault"
        '
        'btncout
        '
        Me.btncout.Image = Global.EPDM_Word.My.Resources.Resources.vault_checkout
        Me.btncout.Label = "Check Out"
        Me.btncout.Name = "btncout"
        Me.btncout.ScreenTip = "Check Out"
        Me.btncout.ShowImage = True
        Me.btncout.SuperTip = "Check out document from vault"
        '
        'btnucout
        '
        Me.btnucout.Image = Global.EPDM_Word.My.Resources.Resources.vault_undocheckout
        Me.btnucout.Label = "Undo Check Out"
        Me.btnucout.Name = "btnucout"
        Me.btnucout.ScreenTip = "Undo Check Out"
        Me.btnucout.ShowImage = True
        Me.btnucout.SuperTip = "Undo action"
        '
        'btnlatver
        '
        Me.btnlatver.Image = Global.EPDM_Word.My.Resources.Resources.Get_Latest_Version
        Me.btnlatver.Label = "Get Latest Version"
        Me.btnlatver.Name = "btnlatver"
        Me.btnlatver.ScreenTip = "Get Latest Version"
        Me.btnlatver.ShowImage = True
        Me.btnlatver.SuperTip = "Get Latest version of current document"
        '
        'btnverinfo
        '
        Me.btnverinfo.Image = Global.EPDM_Word.My.Resources.Resources.information
        Me.btnverinfo.Label = "Version Information"
        Me.btnverinfo.Name = "btnverinfo"
        Me.btnverinfo.ScreenTip = "Version Information"
        Me.btnverinfo.ShowImage = True
        Me.btnverinfo.SuperTip = "Information about current document version"
        '
        'mnucstat
        '
        Me.mnucstat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.mnucstat.Dynamic = True
        Me.mnucstat.Image = Global.EPDM_Word.My.Resources.Resources.state_change1
        Me.mnucstat.Label = "Change State"
        Me.mnucstat.Name = "mnucstat"
        Me.mnucstat.ScreenTip = "Change State"
        Me.mnucstat.ShowImage = True
        Me.mnucstat.SuperTip = "Changing the state of current document"
        '
        'btnwflow
        '
        Me.btnwflow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnwflow.Image = Global.EPDM_Word.My.Resources.Resources.workflow
        Me.btnwflow.Label = "Work Flow Status"
        Me.btnwflow.Name = "btnwflow"
        Me.btnwflow.ScreenTip = "Work Flow Status"
        Me.btnwflow.ShowImage = True
        Me.btnwflow.SuperTip = "Shows document status in workflow"
        '
        'btninbox
        '
        Me.btninbox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btninbox.Image = Global.EPDM_Word.My.Resources.Resources.notification2
        Me.btninbox.Label = "Inbox"
        Me.btninbox.Name = "btninbox"
        Me.btninbox.ScreenTip = "Inbox"
        Me.btninbox.ShowImage = True
        Me.btninbox.SuperTip = "Access Vault Inbox"
        '
        'btnhistory
        '
        Me.btnhistory.Image = Global.EPDM_Word.My.Resources.Resources._1349433770_history
        Me.btnhistory.Label = "History"
        Me.btnhistory.Name = "btnhistory"
        Me.btnhistory.ScreenTip = "History"
        Me.btnhistory.ShowImage = True
        Me.btnhistory.SuperTip = "History of current document"
        '
        'btngdetails
        '
        Me.btngdetails.Image = Global.EPDM_Word.My.Resources.Resources.diagnostics_info
        Me.btngdetails.Label = "Get Details"
        Me.btngdetails.Name = "btngdetails"
        Me.btngdetails.ScreenTip = "Get Details"
        Me.btngdetails.ShowImage = True
        Me.btngdetails.SuperTip = "Shows various details of the document"
        '
        'btndcard
        '
        Me.btndcard.Image = Global.EPDM_Word.My.Resources.Resources.form_icon2
        Me.btndcard.Label = "Show Data Card"
        Me.btndcard.Name = "btndcard"
        Me.btndcard.ScreenTip = "Show Data Card"
        Me.btndcard.ShowImage = True
        Me.btndcard.SuperTip = "Access Vault Data Card"
        '
        'btnbtfile
        '
        Me.btnbtfile.Image = Global.EPDM_Word.My.Resources.Resources.open
        Me.btnbtfile.Label = "Browse to File"
        Me.btnbtfile.Name = "btnbtfile"
        Me.btnbtfile.ScreenTip = "Browse to File"
        Me.btnbtfile.ShowImage = True
        Me.btnbtfile.SuperTip = "Shows Current location of the document"
        '
        'btnsearch
        '
        Me.btnsearch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnsearch.Image = Global.EPDM_Word.My.Resources.Resources.images
        Me.btnsearch.Label = "Search"
        Me.btnsearch.Name = "btnsearch"
        Me.btnsearch.ScreenTip = "Search"
        Me.btnsearch.ShowImage = True
        Me.btnsearch.SuperTip = "Search contents in Vault"
        '
        'btnrefresh
        '
        Me.btnrefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnrefresh.Image = Global.EPDM_Word.My.Resources.Resources._48x48_reload
        Me.btnrefresh.Label = "Refresh"
        Me.btnrefresh.Name = "btnrefresh"
        Me.btnrefresh.ScreenTip = "Refresh"
        Me.btnrefresh.ShowImage = True
        Me.btnrefresh.SuperTip = "Refresh the document"
        '
        'btnhelp
        '
        Me.btnhelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnhelp.Image = Global.EPDM_Word.My.Resources.Resources._11954322131712176739question_mark_naught101_02_svg_hi
        Me.btnhelp.Label = "Help"
        Me.btnhelp.Name = "btnhelp"
        Me.btnhelp.ScreenTip = "Help"
        Me.btnhelp.ShowImage = True
        Me.btnhelp.SuperTip = "Get Help using EPDM"
        '
        'btnabt
        '
        Me.btnabt.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnabt.Image = Global.EPDM_Word.My.Resources.Resources.EGS_logo
        Me.btnabt.Label = "About"
        Me.btnabt.Name = "btnabt"
        Me.btnabt.ScreenTip = "About"
        Me.btnabt.ShowImage = True
        Me.btnabt.SuperTip = "About EGS Computers India Pvt.Ltd"
        '
        'WD_Ribbon
        '
        Me.Name = "WD_Ribbon"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.EPDM)
        Me.EPDM.ResumeLayout(False)
        Me.EPDM.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()

    End Sub

    Friend WithEvents EPDM As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btncin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btncout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnucout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnlatver As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents cmbver As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents btnverinfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnucstat As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnwflow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btninbox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnhistory As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btngdetails As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btndcard As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnbtfile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnsearch As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnrefresh As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnhelp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnabt As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property WD_Ribbon() As WD_Ribbon
        Get
            Return Me.GetRibbon(Of WD_Ribbon)()
        End Get
    End Property
End Class
