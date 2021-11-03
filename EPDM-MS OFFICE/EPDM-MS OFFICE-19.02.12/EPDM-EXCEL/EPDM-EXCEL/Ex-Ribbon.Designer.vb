Partial Class Ribbon1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.btncin = Me.Factory.CreateRibbonButton
        Me.btncout = Me.Factory.CreateRibbonButton
        Me.btnucout = Me.Factory.CreateRibbonButton
        Me.btnlatver = Me.Factory.CreateRibbonButton
        Me.Cmbver = Me.Factory.CreateRibbonComboBox
        Me.btnverinfo = Me.Factory.CreateRibbonButton
        Me.mnucstat = Me.Factory.CreateRibbonMenu
        Me.btnwflow = Me.Factory.CreateRibbonButton
        Me.btninbox = Me.Factory.CreateRibbonButton
        Me.btnhistory = Me.Factory.CreateRibbonButton
        Me.btndetails = Me.Factory.CreateRibbonButton
        Me.btnshowcard = Me.Factory.CreateRibbonButton
        Me.btnbrofile = Me.Factory.CreateRibbonButton
        Me.btnsearch = Me.Factory.CreateRibbonButton
        Me.btnref = Me.Factory.CreateRibbonButton
        Me.Btnhelp = Me.Factory.CreateRibbonButton
        Me.btnabt = Me.Factory.CreateRibbonButton
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
        Me.Group1.Label = "Vault Action"
        Me.Group1.Name = "Group1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnlatver)
        Me.Group2.Items.Add(Me.Cmbver)
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
        Me.Group4.Items.Add(Me.btndetails)
        Me.Group4.Items.Add(Me.btnshowcard)
        Me.Group4.Items.Add(Me.btnbrofile)
        Me.Group4.Label = "Document Information"
        Me.Group4.Name = "Group4"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.btnsearch)
        Me.Group5.Items.Add(Me.btnref)
        Me.Group5.Items.Add(Me.Btnhelp)
        Me.Group5.Items.Add(Me.btnabt)
        Me.Group5.Label = "General"
        Me.Group5.Name = "Group5"
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
        Me.btncout.SuperTip = "Check out document from Vault"
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
        'btnlatver
        '
        Me.btnlatver.Image = CType(resources.GetObject("btnlatver.Image"), System.Drawing.Image)
        Me.btnlatver.Label = "Get Latest Version"
        Me.btnlatver.Name = "btnlatver"
        Me.btnlatver.ScreenTip = "Get Latest Version"
        Me.btnlatver.ShowImage = True
        Me.btnlatver.SuperTip = "Get latest version of current document"
        '
        'Cmbver
        '
        Me.Cmbver.Label = "Get Version"
        Me.Cmbver.Name = "Cmbver"
        Me.Cmbver.ScreenTip = "Get Version"
        Me.Cmbver.SuperTip = "Get required version of the current document"
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
        'mnucstat
        '
        Me.mnucstat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.mnucstat.Dynamic = True
        Me.mnucstat.Image = CType(resources.GetObject("mnucstat.Image"), System.Drawing.Image)
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
        'btnhistory
        '
        Me.btnhistory.Image = CType(resources.GetObject("btnhistory.Image"), System.Drawing.Image)
        Me.btnhistory.Label = "History"
        Me.btnhistory.Name = "btnhistory"
        Me.btnhistory.ScreenTip = "History"
        Me.btnhistory.ShowImage = True
        Me.btnhistory.SuperTip = "History of current document"
        '
        'btndetails
        '
        Me.btndetails.Image = CType(resources.GetObject("btndetails.Image"), System.Drawing.Image)
        Me.btndetails.Label = "Get Details"
        Me.btndetails.Name = "btndetails"
        Me.btndetails.ScreenTip = "Get Details"
        Me.btndetails.ShowImage = True
        Me.btndetails.SuperTip = "Shows various details of the document"
        '
        'btnshowcard
        '
        Me.btnshowcard.Image = CType(resources.GetObject("btnshowcard.Image"), System.Drawing.Image)
        Me.btnshowcard.Label = "Show Data Card"
        Me.btnshowcard.Name = "btnshowcard"
        Me.btnshowcard.ScreenTip = "Show Data Card"
        Me.btnshowcard.ShowImage = True
        Me.btnshowcard.SuperTip = "Access Vault Data Card"
        '
        'btnbrofile
        '
        Me.btnbrofile.Image = CType(resources.GetObject("btnbrofile.Image"), System.Drawing.Image)
        Me.btnbrofile.Label = "Browse to File"
        Me.btnbrofile.Name = "btnbrofile"
        Me.btnbrofile.ScreenTip = "Browse to File"
        Me.btnbrofile.ShowImage = True
        Me.btnbrofile.SuperTip = "Shows current location of the document"
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
        Me.btnref.Image = CType(resources.GetObject("btnref.Image"), System.Drawing.Image)
        Me.btnref.Label = "Refresh"
        Me.btnref.Name = "btnref"
        Me.btnref.ScreenTip = "Refresh"
        Me.btnref.ShowImage = True
        Me.btnref.SuperTip = "Refresh the document"
        '
        'Btnhelp
        '
        Me.Btnhelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Btnhelp.Image = CType(resources.GetObject("Btnhelp.Image"), System.Drawing.Image)
        Me.Btnhelp.Label = "Help"
        Me.Btnhelp.Name = "Btnhelp"
        Me.Btnhelp.ScreenTip = "Help"
        Me.Btnhelp.ShowImage = True
        Me.Btnhelp.SuperTip = "Get Help using EPDM"
        '
        'btnabt
        '
        Me.btnabt.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnabt.Image = CType(resources.GetObject("btnabt.Image"), System.Drawing.Image)
        Me.btnabt.Label = "About"
        Me.btnabt.Name = "btnabt"
        Me.btnabt.ScreenTip = "About"
        Me.btnabt.ShowImage = True
        Me.btnabt.SuperTip = "About EGS Computers India Pvt.Ltd"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
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
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btncout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnucout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnlatver As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Cmbver As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents btnverinfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnucstat As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnwflow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btninbox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnhistory As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btndetails As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnshowcard As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnbrofile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnsearch As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnref As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Btnhelp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnabt As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
