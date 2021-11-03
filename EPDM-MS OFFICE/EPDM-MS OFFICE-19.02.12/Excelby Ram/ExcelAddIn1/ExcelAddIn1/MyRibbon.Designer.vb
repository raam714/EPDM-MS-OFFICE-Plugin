Partial Class MyRibbon
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MyRibbon))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.btncin = Me.Factory.CreateRibbonButton
        Me.btncout = Me.Factory.CreateRibbonButton
        Me.btnucout = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btnlatver = Me.Factory.CreateRibbonButton
        Me.cmbver = Me.Factory.CreateRibbonComboBox
        Me.btnverinfo = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.mnucstat = Me.Factory.CreateRibbonMenu
        Me.btnwflow = Me.Factory.CreateRibbonButton
        Me.btninbox = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.btnhistory = Me.Factory.CreateRibbonButton
        Me.btndetails = Me.Factory.CreateRibbonButton
        Me.btnshowcard = Me.Factory.CreateRibbonButton
        Me.btnbrofile = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.btnsearch = Me.Factory.CreateRibbonButton
        Me.btnref = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "EPDM"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btncin)
        Me.Group1.Items.Add(Me.btncout)
        Me.Group1.Items.Add(Me.btnucout)
        Me.Group1.Label = "Vault Actions"
        Me.Group1.Name = "Group1"
        '
        'btncin
        '
        Me.btncin.Image = CType(resources.GetObject("btncin.Image"), System.Drawing.Image)
        Me.btncin.Label = "Check In"
        Me.btncin.Name = "btncin"
        Me.btncin.ScreenTip = "Check In"
        Me.btncin.ShowImage = True
        Me.btncin.SuperTip = "Check in document into Vault"
        '
        'btncout
        '
        Me.btncout.Image = CType(resources.GetObject("btncout.Image"), System.Drawing.Image)
        Me.btncout.Label = "Check Out"
        Me.btncout.Name = "btncout"
        Me.btncout.ScreenTip = "Check Out"
        Me.btncout.ShowImage = True
        Me.btncout.SuperTip = "Check Out document from Vault"
        '
        'btnucout
        '
        Me.btnucout.Image = CType(resources.GetObject("btnucout.Image"), System.Drawing.Image)
        Me.btnucout.Label = "Undo Check Out"
        Me.btnucout.Name = "btnucout"
        Me.btnucout.ScreenTip = "Undo Check Out"
        Me.btnucout.ShowImage = True
        Me.btnucout.SuperTip = "Undo action"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnlatver)
        Me.Group2.Items.Add(Me.cmbver)
        Me.Group2.Items.Add(Me.btnverinfo)
        Me.Group2.Label = "Version Management"
        Me.Group2.Name = "Group2"
        '
        'btnlatver
        '
        Me.btnlatver.Image = CType(resources.GetObject("btnlatver.Image"), System.Drawing.Image)
        Me.btnlatver.Label = "Get Latest Version"
        Me.btnlatver.Name = "btnlatver"
        Me.btnlatver.ScreenTip = "Get Latest Version"
        Me.btnlatver.ShowImage = True
        Me.btnlatver.SuperTip = "Get latest version of current document"
        '
        'cmbver
        '
        Me.cmbver.Label = "Get Version"
        Me.cmbver.Name = "cmbver"
        Me.cmbver.ScreenTip = "Get Version"
        Me.cmbver.SuperTip = "Get required version of the current document"
        Me.cmbver.Text = Nothing
        '
        'btnverinfo
        '
        Me.btnverinfo.Image = CType(resources.GetObject("btnverinfo.Image"), System.Drawing.Image)
        Me.btnverinfo.Label = "Version Information"
        Me.btnverinfo.Name = "btnverinfo"
        Me.btnverinfo.ScreenTip = "Version Information"
        Me.btnverinfo.ShowImage = True
        Me.btnverinfo.SuperTip = "Information about current document version"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.mnucstat)
        Me.Group3.Items.Add(Me.btnwflow)
        Me.Group3.Items.Add(Me.btninbox)
        Me.Group3.Label = "Work Flow Management"
        Me.Group3.Name = "Group3"
        '
        'mnucstat
        '
        Me.mnucstat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.mnucstat.Dynamic = True
        Me.mnucstat.Image = Global.ExcelAddIn1.My.Resources.Resources.state_change1
        Me.mnucstat.ImageName = "cstat"
        Me.mnucstat.Label = "Change State"
        Me.mnucstat.Name = "mnucstat"
        Me.mnucstat.ScreenTip = "Change State"
        Me.mnucstat.ShowImage = True
        Me.mnucstat.SuperTip = "Changing the state of current document"
        '
        'btnwflow
        '
        Me.btnwflow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnwflow.Image = CType(resources.GetObject("btnwflow.Image"), System.Drawing.Image)
        Me.btnwflow.Label = "Work Flow Status"
        Me.btnwflow.Name = "btnwflow"
        Me.btnwflow.ScreenTip = "Work Flow Status"
        Me.btnwflow.ShowImage = True
        Me.btnwflow.SuperTip = "Shows document status in work flow"
        '
        'btninbox
        '
        Me.btninbox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btninbox.Image = CType(resources.GetObject("btninbox.Image"), System.Drawing.Image)
        Me.btninbox.Label = "Inbox"
        Me.btninbox.Name = "btninbox"
        Me.btninbox.ScreenTip = "Inbox"
        Me.btninbox.ShowImage = True
        Me.btninbox.SuperTip = "Access Vault Inbox"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnhistory)
        Me.Group4.Items.Add(Me.btndetails)
        Me.Group4.Items.Add(Me.btnshowcard)
        Me.Group4.Items.Add(Me.btnbrofile)
        Me.Group4.Label = "Document Information"
        Me.Group4.Name = "Group4"
        '
        'btnhistory
        '
        Me.btnhistory.Image = Global.ExcelAddIn1.My.Resources.Resources._1349433770_history
        Me.btnhistory.Label = "History"
        Me.btnhistory.Name = "btnhistory"
        Me.btnhistory.ScreenTip = "History"
        Me.btnhistory.ShowImage = True
        Me.btnhistory.SuperTip = "History of current document"
        '
        'btndetails
        '
        Me.btndetails.Image = Global.ExcelAddIn1.My.Resources.Resources.diagnostics_info
        Me.btndetails.Label = "Get Details"
        Me.btndetails.Name = "btndetails"
        Me.btndetails.ScreenTip = "Get Details"
        Me.btndetails.ShowImage = True
        Me.btndetails.SuperTip = "Shows various details of the document "
        '
        'btnshowcard
        '
        Me.btnshowcard.Image = Global.ExcelAddIn1.My.Resources.Resources.form_icon2
        Me.btnshowcard.Label = "Show Data Card"
        Me.btnshowcard.Name = "btnshowcard"
        Me.btnshowcard.ScreenTip = "Show Data Card"
        Me.btnshowcard.ShowImage = True
        Me.btnshowcard.SuperTip = "Access Vault Data Card"
        '
        'btnbrofile
        '
        Me.btnbrofile.Image = Global.ExcelAddIn1.My.Resources.Resources.open
        Me.btnbrofile.Label = "Browse to File"
        Me.btnbrofile.Name = "btnbrofile"
        Me.btnbrofile.ScreenTip = "Browse to File"
        Me.btnbrofile.ShowImage = True
        Me.btnbrofile.SuperTip = "Shows current location of the document"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.btnsearch)
        Me.Group5.Items.Add(Me.btnref)
        Me.Group5.Items.Add(Me.Button15)
        Me.Group5.Items.Add(Me.Button16)
        Me.Group5.Label = "General"
        Me.Group5.Name = "Group5"
        '
        'btnsearch
        '
        Me.btnsearch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnsearch.Image = CType(resources.GetObject("btnsearch.Image"), System.Drawing.Image)
        Me.btnsearch.Label = "Search"
        Me.btnsearch.Name = "btnsearch"
        Me.btnsearch.ScreenTip = "Search"
        Me.btnsearch.ShowImage = True
        Me.btnsearch.SuperTip = "Search contents in Vault"
        '
        'btnref
        '
        Me.btnref.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnref.Image = Global.ExcelAddIn1.My.Resources.Resources._48x48_reload
        Me.btnref.Label = "Refresh"
        Me.btnref.Name = "btnref"
        Me.btnref.ScreenTip = "Refresh"
        Me.btnref.ShowImage = True
        Me.btnref.SuperTip = "Refresh the document"
        '
        'Button15
        '
        Me.Button15.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button15.Image = Global.ExcelAddIn1.My.Resources.Resources._11954322131712176739question_mark_naught101_02_svg_hi
        Me.Button15.Label = "Help"
        Me.Button15.Name = "Button15"
        Me.Button15.ScreenTip = "Help"
        Me.Button15.ShowImage = True
        Me.Button15.SuperTip = "Get Help using EPDM"
        '
        'Button16
        '
        Me.Button16.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button16.Image = Global.ExcelAddIn1.My.Resources.Resources.EGS_logo
        Me.Button16.Label = "About"
        Me.Button16.Name = "Button16"
        Me.Button16.ScreenTip = "About"
        Me.Button16.ShowImage = True
        Me.Button16.SuperTip = "About EGS Computers India Pvt.Ltd"
        '
        'MyRibbon
        '
        Me.Name = "MyRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
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

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btncin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btncout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnucout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnlatver As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnverinfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnwflow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btninbox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnhistory As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btndetails As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnshowcard As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnbrofile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnsearch As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnref As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents cmbver As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents mnucstat As Microsoft.Office.Tools.Ribbon.RibbonMenu
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property MyRibbon() As MyRibbon
        Get
            Return Me.GetRibbon(Of MyRibbon)()
        End Get
    End Property
End Class
