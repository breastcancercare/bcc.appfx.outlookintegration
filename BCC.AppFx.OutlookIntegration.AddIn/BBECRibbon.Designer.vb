Partial Class BBECRibbon
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
        Me.tabBBEC = Factory.CreateRibbonTab()
        Me.grpInteraction = Factory.CreateRibbonGroup()
        Me.grpConnect = Factory.CreateRibbonGroup()
        Me.btnAddInteraction = Factory.CreateRibbonButton()
        Me.btnEditInteraction = Factory.CreateRibbonButton()
        Me.btnViewInteraction = Factory.CreateRibbonButton()
        Me.btnConnect = Factory.CreateRibbonButton()
        Me.tabBBEC.SuspendLayout()
        Me.grpInteraction.SuspendLayout()
        Me.grpConnect.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabBBEC
        '
        Me.tabBBEC.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.tabBBEC.Groups.Add(Me.grpInteraction)
        Me.tabBBEC.Groups.Add(Me.grpConnect)
        Me.tabBBEC.Label = "EnterpriseCRM"
        Me.tabBBEC.Name = "tabBBEC"
        '
        'grpInteraction
        '
        Me.grpInteraction.Items.Add(Me.btnAddInteraction)
        Me.grpInteraction.Items.Add(Me.btnEditInteraction)
        Me.grpInteraction.Items.Add(Me.btnViewInteraction)
        Me.grpInteraction.Label = "Interaction"
        Me.grpInteraction.Name = "grpInteraction"
        '
        'grpConnect
        '
        Me.grpConnect.Items.Add(Me.btnConnect)
        Me.grpConnect.Label = "Connecting..."
        Me.grpConnect.Name = "grpConnect"
        '
        'btnAddInteraction
        '
        Me.btnAddInteraction.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAddInteraction.Image = Global.BCC.AppFx.OutlookIntegration.AddIn.My.Resources.Resources.interactions
        Me.btnAddInteraction.Label = "Add"
        Me.btnAddInteraction.Name = "btnAddInteraction"
        Me.btnAddInteraction.ScreenTip = "Add interaction to a constituent"
        Me.btnAddInteraction.ShowImage = True
        '
        'btnEditInteraction
        '
        Me.btnEditInteraction.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnEditInteraction.Image = Global.BCC.AppFx.OutlookIntegration.AddIn.My.Resources.Resources.edit_32
        Me.btnEditInteraction.Label = "Edit"
        Me.btnEditInteraction.Name = "btnEditInteraction"
        Me.btnEditInteraction.ScreenTip = "Edit interaction (add participants)"
        Me.btnEditInteraction.ShowImage = True
        '
        'btnViewInteraction
        '
        Me.btnViewInteraction.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnViewInteraction.Image = Global.BCC.AppFx.OutlookIntegration.AddIn.My.Resources.Resources.goto_round_32
        Me.btnViewInteraction.Label = "View"
        Me.btnViewInteraction.Name = "btnViewInteraction"
        Me.btnViewInteraction.ScreenTip = "View interaction"
        Me.btnViewInteraction.ShowImage = True
        '
        'btnConnect
        '
        Me.btnConnect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnConnect.Enabled = False
        Me.btnConnect.Image = Global.BCC.AppFx.OutlookIntegration.AddIn.My.Resources.Resources.history
        Me.btnConnect.Label = "..."
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.ScreenTip = "Connecting and authorising with Enterprise"
        Me.btnConnect.ShowImage = True
        '
        'BBECRibbon
        '
        Me.Name = "BBECRibbon"
        Me.RibbonType = "Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.tabBBEC)
        Me.tabBBEC.ResumeLayout(False)
        Me.tabBBEC.PerformLayout()
        Me.grpInteraction.ResumeLayout(False)
        Me.grpInteraction.PerformLayout()
        Me.grpConnect.ResumeLayout(False)
        Me.grpConnect.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tabBBEC As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpInteraction As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnAddInteraction As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnEditInteraction As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnViewInteraction As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpConnect As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnConnect As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property BBECRibbon1() As BBECRibbon
        Get

            Return Me.GetRibbon(Of BBECRibbon)()
            
        End Get
    End Property
End Class
