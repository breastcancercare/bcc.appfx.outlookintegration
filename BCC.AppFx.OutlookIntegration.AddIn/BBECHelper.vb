Imports Blackbaud.AppFx.WebAPI
Imports Blackbaud.AppFx.XmlTypes
Imports Blackbaud.AppFx.XmlTypes.DataForms
Imports Blackbaud.AppFx.UIModeling.DataFormWebHost
Imports Blackbaud.AppFx.WebAPI.ServiceProxy

Friend NotInheritable Class BBECHelper

#Region "BCRM connections and security"

    Private Shared _provider As AppFxWebServiceProvider
    Private Shared _permission As Boolean = False
    Private Shared _authenticated As Boolean = False

    Friend Shared Function applicationTitle() As String
        Return My.Settings.applicationTitle.ToString
    End Function

    Friend Shared Function serviceUrlBasePath() As String
        Dim s As String
        s = My.Settings.serviceUrlBasePath.ToString
        If Right(s, 1) <> "\" Then s = String.Concat(s, "\")
        Return s
    End Function

    Friend Shared Function databaseName() As String
        Return My.Settings.databaseName.ToString
    End Function

    Friend Shared Function getProvider() As AppFxWebServiceProvider

        If _provider Is Nothing Then
            _provider = New Blackbaud.AppFx.WebAPI.AppFxWebServiceProvider
            _provider.Url = String.Concat(serviceUrlBasePath, "appfxwebservice.asmx")
            _provider.Database = databaseName()
            _provider.Credentials = System.Net.CredentialCache.DefaultCredentials
        End If

        Return _provider

    End Function

    Friend Shared Property addinPermissioned As Boolean
        Get
            addinPermissioned = _permission And _authenticated
        End Get
        Set(value As Boolean)
            _permission = value
        End Set
    End Property

    Friend Shared Property addinConnected As Boolean
        Get
            addinConnected = _authenticated
        End Get
        Set(value As Boolean)
            _authenticated = value
        End Set
    End Property

    Friend Shared Sub PerformCRMLogin()
        If Not _authenticated Then
            Try
                _permission = UserHasRights(New Guid("d286b10f-2d65-4603-991e-bf322f37a9a6"), SecurityFeatureType.Form)
            Catch ex As Exception
                _permission = False
            End Try

        End If
        ' Does user have rights to the 'Outlook Integration Tag Add Data Form'

    End Sub

    Private Shared Function UserHasRights(ByVal featureID As Guid, ByVal featureType As SecurityFeatureType) As Boolean
        Dim provider = getProvider()
        Dim returnValue As Boolean = False

        Dim Request As Blackbaud.AppFx.WebAPI.ServiceProxy.SecurityUserGrantedFeatureRequest
        Request = provider.CreateRequest(Of ServiceProxy.SecurityUserGrantedFeatureRequest)()
        Request.FeatureID = featureID
        Request.FeatureType = featureType

        Dim Reply As ServiceProxy.SecurityUserGrantedFeatureReply = Nothing
        Try
            Reply = provider.Service.SecurityUserGrantedFeature(Request)
            _authenticated = True
            returnValue = Reply.Granted
        Catch Ex As Exception
            _authenticated = False
            returnValue = False
        End Try
        Return returnValue

    End Function

#End Region

#Region "Resolve lookup id values"

    Friend Shared Function GetConstituentId(item As Outlook.MailItem) As Guid

        Dim emailAddress = ResolveEmailAddress(item)
        Return GetConstituentId(emailAddress)

    End Function

    Friend Shared Function GetConstituentId(recipient As Outlook.Recipient) As Guid

        Dim emailAddress = ResolveEmailAddress(recipient)
        Return GetConstituentId(emailAddress)

    End Function

    Friend Shared Function GetConstituentId(emailAddress As String) As Guid

        Try
            Dim provider = getProvider()

            Dim form = BBECHelper.GetSearchFormWebHostDialog()
            form.SearchListId = New Guid("23C5C603-D7D8-4106-AECC-65392B563887")    'Constituent search

            'create a field value set to hold any filters we want to pass in
            Dim fvSet As New DataFormFieldValueSet

            'add a field value for each filter parameter.
            fvSet.Add(New DataFormFieldValue("EMAILADDRESS", emailAddress))
            fvSet.Add(New DataFormFieldValue("EXACTMATCHONLY", False))
            fvSet.Add(New DataFormFieldValue("ONLYPRIMARYADDRESS", False))  ' Search all email addresses
            fvSet.Add(New DataFormFieldValue("INCLUDEGROUPS", False))
            fvSet.Add(New DataFormFieldValue("INCLUDEORGANIZATIONS", False))

            'create a dataform item to contain the filter field value set
            Dim dfi As New DataFormItem
            dfi.Values = fvSet

            form.SetSearchCriteria(dfi)

            If form.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Return New Guid(form.SelectedRecordId)
            End If

        Catch ex As Exception
            HandleException("There was an error getting the constituent ID", ex)

        End Try

        Return Guid.Empty

    End Function

    Friend Shared Function ResolveEmailAddress(item As Outlook.MailItem) As String

        Try
            If String.Equals(item.SenderEmailType, "EX", StringComparison.OrdinalIgnoreCase) Then
                If item.Sender.Address = item.Session.CurrentUser.Address Then  'TODO: issue with some users where Session.CurrentUser.Address is the Exchange address not email someone@somewhere
                    Dim ns = item.Application.GetNamespace("MAPI")
                    Dim recip = ns.CreateRecipient(item.Recipients(1).Address)
                    Dim exUser = recip.AddressEntry.GetExchangeUser()
                    Return exUser.PrimarySmtpAddress
                Else
                    Dim ns = item.Application.GetNamespace("MAPI")
                    Dim recip = ns.CreateRecipient(item.SenderEmailAddress)
                    Dim exUser = recip.AddressEntry.GetExchangeUser()
                    Return exUser.PrimarySmtpAddress
                End If
            ElseIf String.Equals(item.SenderEmailType, "SMTP", StringComparison.OrdinalIgnoreCase) Then
                Return item.SenderEmailAddress
            End If
        Catch
            'eat the error
        End Try

        Return item.SenderEmailAddress

    End Function

    Friend Shared Function ResolveEmailAddress(recipient As Outlook.Recipient) As String

        Try
            Dim exUser = recipient.AddressEntry.GetExchangeUser()
            Return exUser.PrimarySmtpAddress
        Catch
            'eat the error
        End Try

        Return recipient.Address

    End Function

    Friend Shared Function ResolveMessageId(item As Outlook.MailItem) As String
        'Retrieves the RFC2822 message id - exchange EntryID will normally be unique per mailbox and changes if message moved

        Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
        Return item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID).ToString
    End Function

    Friend Shared Function GetInteractionId(item As Outlook.MailItem) As Guid

        If Not addinPermissioned Then Return Guid.Empty
        Try
            Dim provider = getProvider()

            Dim Request As Blackbaud.AppFx.WebAPI.ServiceProxy.SearchListLoadRequest
            Request = provider.CreateRequest(Of ServiceProxy.SearchListLoadRequest)()
            Request.SearchListID = New Guid("23782b71-9db9-4e07-b15d-54fbd00eb05f") ' Interaction by MessageId search
            Dim fvSet = New Blackbaud.AppFx.XmlTypes.DataForms.DataFormFieldValueSet
            fvSet.Add("MESSAGEID", ResolveMessageId(item))
            Dim dfi = New DataFormItem
            dfi.Values = fvSet

            Request.Filter = dfi
            Dim Reply As ServiceProxy.SearchListLoadReply = Nothing
            Try
                Reply = provider.Service.SearchListLoad(Request)
            Catch Ex As Exception
                'eat
            End Try

            If Reply.Output.RowCount = 1 Then
                Return New Guid(Reply.Output.Rows(0).Values(0).ToString)
            Else
                Return Guid.Empty
            End If

        Catch ex As Exception
            HandleException("There was an error getting the interaction ID", ex)
            Return Guid.Empty
        End Try

    End Function

#End Region

    Friend Shared Function AddInteraction(item As Outlook.MailItem) As Guid

        Try
            If item Is Nothing Then Return Guid.Empty

            Dim selectedConstituentId As System.Guid

            Dim provider = getProvider()

            Dim itemDate As Date
            If item.Sent Then
                itemDate = item.SentOn
            Else
                itemDate = item.ReceivedTime
            End If

            selectedConstituentId = GetConstituentId(item)

            If selectedConstituentId = Guid.Empty Then Return Guid.Empty

            Dim form = BBECHelper.GetDataFormWebHostDialog()
            form.DataFormInstanceId = New Guid("8fab24a9-ec7a-4067-a368-ccd4aec3fa1b") 'Interaction Add Form 2 (BCC) (auto-populates FUNDRAISERID in pre-load) (BCRM <= v2.94)
            'form.DataFormInstanceId = New Guid("723ad883-f995-4c40-afed-6a7914b536e3") 'Interaction Add Form 2 (BCRM >= v3.0)

            form.ContextRecordId = selectedConstituentId.ToString

            Dim dfi = New DataForms.DataFormItem

            Dim s = itemDate.ToShortDateString()
            dfi.SetValue("ACTUALDATE", s)
            dfi.SetValue("EXPECTEDDATE", s)

            s = itemDate.ToShortTimeString()
            dfi.SetValue("ACTUALSTARTTIME", s)
            dfi.SetValue("ACTUALENDTIME", s)
            dfi.SetValue("EXPECTEDSTARTTIME", s)
            dfi.SetValue("EXPECTEDENDTIME", s)

            dfi.SetValue("OBJECTIVE", item.Subject)
            dfi.SetValue("INTERACTIONTYPECODEID", New Guid("80A5448A-A96A-4B18-BAF0-92C739822226")) 'InteractionType="E-mail" (hardcoded to avoid Import of WebAPI.Constituent)
            dfi.SetValue("STATUSCODE", 2)  'Completed

            form.DefaultValues = dfi

            If form.ShowDialog() = Windows.Forms.DialogResult.OK Then

                Dim interactionId As Guid = New Guid(form.RecordId.ToString)

                Dim reqNotepad = DataFormServices.CreateDataFormSaveRequest(provider, New Guid("8f17c68a-39eb-47cb-8d67-2239670cfe73")) 'ConstituentInteractionNoteAddForm
                reqNotepad.ContextRecordID = interactionId.ToString

                Dim dfiNotepad = New DataFormItem

                dfiNotepad.SetValue("AUTHORID", selectedConstituentId)
                dfiNotepad.SetValue("NOTETYPECODEID", New Guid("431A0A3C-D519-4293-A791-32F34A37B648")) ' InteractionNoteType="E-mail"
                dfiNotepad.SetValue("TEXTNOTE", item.Body)
                'emaildata.HTMLNOTE = item.HTMLBody 'Form validation requires 'clean' HTML which this is not.
                dfiNotepad.SetValue("DATEENTERED", itemDate)
                dfiNotepad.SetValue("TITLE", Left(item.Subject, 50))    ' TITLE field has 50 character limit on form

                reqNotepad.DataFormItem = dfiNotepad

                DataFormServices.SaveData(provider, reqNotepad)

                Try

                    TagMailItem(item, interactionId)

                Catch ex As Exception
                    HandleException("There was an error tagging this email", ex)

                End Try

                Return interactionId

            End If

        Catch ex As Exception
            HandleException("There was an error creating the interaction", ex)
            Return Guid.Empty

        End Try

    End Function

    Friend Shared Sub EditInteraction(interactionId As Guid)
        Try
            Dim form = GetDataFormWebHostDialog()
            form.DataFormInstanceId = New Guid("ab3b9569-4c7c-4646-a793-347856753b60")  'Interaction Edit Form 4
            form.RecordId = interactionId.ToString()
            form.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

    Friend Shared Sub ViewInteraction(interactionId As Guid, item As Outlook.MailItem)
        Try
            Dim form = BBECHelper.GetDataFormWebHostDialog()
            form.DataFormInstanceId = New Guid("e3574968-1684-4b51-9752-3599be1b4ec4")  'Constituent Interaction View Form
            form.RecordId = interactionId.ToString
            form.ApplicationTitle = "Interaction for email from " & item.SenderName & " to " & item.To
            form.ShowDialog()
        Catch ex As Exception
            BBECHelper.HandleException("There was an error showing the View Interaction form", ex)

        End Try

    End Sub

    Friend Shared Sub TagMailItem(item As Outlook.MailItem, interactionId As Guid)

        If item Is Nothing Then Return
        If interactionId = Guid.Empty Then Return

        Try
            Dim provider = getProvider()

            Dim req = DataFormServices.CreateDataFormSaveRequest(provider, New Guid("d286b10f-2d65-4603-991e-bf322f37a9a6"))    ' Tag Add form
            req.ContextRecordID = ResolveMessageId(item)

            Dim dfi = New Blackbaud.AppFx.XmlTypes.DataForms.DataFormItem
            dfi.SetValue("INTERACTIONID", interactionId)
            req.DataFormItem = dfi

            DataFormServices.SaveData(provider, req)

        Catch ex As Exception
            HandleException("There was an error tagging this constituent", ex)

        End Try

    End Sub

#Region "General functions"

    Friend Sub MouseWaitStop()
        Windows.Forms.Cursor.Current = Windows.Forms.Cursors.Default
    End Sub

    Friend Sub MouseWait()
        Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
    End Sub

    Friend Shared Function GetDataFormWebHostDialog() As DataFormWebHostDialog
        Dim form = New DataFormWebHostDialog
        form.ServiceUrlBasePath = serviceUrlBasePath()
        form.DatabaseName = databaseName()
        form.ApplicationTitle = applicationTitle()
        form.Credentials = System.Net.CredentialCache.DefaultCredentials
        Return (form)

    End Function

    Friend Shared Function GetSearchFormWebHostDialog() As SearchFormWebHostDialog

        Dim form = New SearchFormWebHostDialog
        form.ServiceUrlBasePath = serviceUrlBasePath()
        form.DatabaseName = databaseName()
        form.ApplicationTitle = applicationTitle()
        form.Credentials = System.Net.CredentialCache.DefaultCredentials
        Return form

    End Function

    Friend Shared Sub HandleException(msg As String, ex As Exception)
        MsgBox(String.Format("{0}:  {1}", msg, ex.Message), MsgBoxStyle.Information, applicationTitle)
    End Sub

#End Region

End Class