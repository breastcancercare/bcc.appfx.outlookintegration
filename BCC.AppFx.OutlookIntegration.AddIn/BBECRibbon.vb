Imports Microsoft.Office.Tools.Ribbon

Public Class BBECRibbon

    Private _interactionId As Guid = Guid.Empty

    Private Sub BBECRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        Globals.ThisAddIn.ribbon = Me
        SetTabPermission()
    End Sub

    Private Sub SetButtonsEnabled(interactionPresent As Boolean)
        Me.btnAddInteraction.Enabled = Not interactionPresent
        Me.btnEditInteraction.Enabled = interactionPresent
        Me.btnViewInteraction.Enabled = interactionPresent
    End Sub

    Friend Sub SetTabPermission()
        If Not BBECHelper.addinConnected Then
            grpInteraction.Visible = False
            grpConnect.Visible = True
        Else
            If Not BBECHelper.addinPermissioned Then    ' Hides add-in tab
                grpInteraction.Visible = False
                grpConnect.Visible = False
            Else
                Dim item = GetCurrentItem()
                If item IsNot Nothing Then _interactionId = BBECHelper.GetInteractionId(item)
                SetButtonsEnabled(_interactionId <> Guid.Empty)
                grpInteraction.Visible = True
                grpConnect.Visible = False

            End If
        End If

    End Sub

    Private Function GetCurrentItem() As Outlook.MailItem

        Dim inspector = DirectCast(Me.Context, Outlook.Inspector)
        If inspector Is Nothing Then Return Nothing
        Return TryCast(inspector.CurrentItem, Outlook.MailItem)

    End Function

    Private Sub btnAddInteraction_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAddInteraction.Click

        Dim item = GetCurrentItem()
        If item Is Nothing Then Return

        Try
            _interactionId = BBECHelper.AddInteraction(item)
        Catch ex As Exception
            ' eat error
        Finally
            SetButtonsEnabled(_interactionId <> Guid.Empty)
        End Try

    End Sub

    Private Sub btnEditInteraction_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnEditInteraction.Click

        Try
            Dim item = GetCurrentItem()
            If item Is Nothing Then Return

            BBECHelper.EditInteraction(_interactionId)

        Catch ex As Exception
            BBECHelper.HandleException("There was an error viewing the tagged interaction", ex)

        End Try

    End Sub

    Private Sub btnViewInteraction_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnViewInteraction.Click

        Try
            Dim item = GetCurrentItem()
            If item Is Nothing Then Return

            BBECHelper.ViewInteraction(_interactionId, item)

        Catch ex As Exception
            BBECHelper.HandleException("There was an error showing the tagged interaction", ex)

        End Try

    End Sub

End Class