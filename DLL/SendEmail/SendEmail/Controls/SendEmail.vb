Imports System.Net.Mail


<ComClass(SendEmail.ClassId, SendEmail.InterfaceId, SendEmail.EventsId)> _
Public Class SendEmail

#Region "VB6 Interop Code"

#If COM_INTEROP_ENABLED Then

#Region "COM Registration"

    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    Public Const ClassId As String = "0f5d290c-7f9b-44b3-ae40-b44cc1b1a21d"
    Public Const InterfaceId As String = "89795ac5-b6d1-4853-8776-ea549adf2280"
    Public Const EventsId As String = "cb5920c8-56eb-4c61-9aac-5171ca8c1217"

    'These routines perform the additional COM registration needed by ActiveX controls
    <EditorBrowsable(EditorBrowsableState.Never)> _
    <ComRegisterFunction()> _
    Private Shared Sub Register(ByVal t As Type)
        ComRegistration.RegisterControl(t, 102)
    End Sub

    <EditorBrowsable(EditorBrowsableState.Never)> _
    <ComUnregisterFunction()> _
    Private Shared Sub Unregister(ByVal t As Type)
        ComRegistration.UnregisterControl(t)
    End Sub

#End Region

#Region "VB6 Events"

    'This section shows some examples of exposing a UserControl's events to VB6.  Typically, you just
    '1) Declare the event as you want it to be shown in VB6
    '2) Raise the event in the appropriate UserControl event.

    Public Shadows Event Click() 'Event must be marked as Shadows since .NET UserControls have the same name.
    Public Event DblClick()

    Private Sub InteropUserControl_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
        RaiseEvent Click()
    End Sub

    Private Sub InteropUserControl_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DoubleClick
        RaiseEvent DblClick()
    End Sub

#End Region

#Region "VB6 Properties"

    'The following are examples of how to expose typical form properties to VB6.  
    'You can also use these as examples on how to add additional properties.

    'Must Shadow this property as it exists in Windows.Forms and is not overridable
    Public Shadows Property Visible() As Boolean
        Get
            Return MyBase.Visible
        End Get
        Set(ByVal value As Boolean)
            MyBase.Visible = value
        End Set
    End Property

    Public Shadows Property Enabled() As Boolean
        Get
            Return MyBase.Enabled
        End Get
        Set(ByVal value As Boolean)
            MyBase.Enabled = value
        End Set
    End Property

    Public Shadows Property ForegroundColor() As Integer
        Get
            Return ActiveXControlHelpers.GetOleColorFromColor(MyBase.ForeColor)
        End Get
        Set(ByVal value As Integer)
            MyBase.ForeColor = ActiveXControlHelpers.GetColorFromOleColor(value)
        End Set
    End Property

    Public Shadows Property BackgroundColor() As Integer
        Get
            Return ActiveXControlHelpers.GetOleColorFromColor(MyBase.BackColor)
        End Get
        Set(ByVal value As Integer)
            MyBase.BackColor = ActiveXControlHelpers.GetColorFromOleColor(value)
        End Set
    End Property

    Public Overrides Property BackgroundImage() As System.Drawing.Image
        Get
            Return Nothing
        End Get
        Set(ByVal value As System.Drawing.Image)
            If value IsNot Nothing Then
                MsgBox("Setting the background image of an Interop UserControl is not supported, please use a PictureBox instead.", MsgBoxStyle.Information)
            End If
            MyBase.BackgroundImage = Nothing
        End Set
    End Property

#End Region

#Region "VB6 Methods"

    Public Overrides Sub Refresh()
        MyBase.Refresh()
    End Sub

    'Ensures that tabbing across VB6 and .NET controls works as expected
    Private Sub UserControl_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
        ActiveXControlHelpers.HandleFocus(Me)
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'Raise Load event
        Me.OnCreateControl()
    End Sub

    <SecurityPermission(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.UnmanagedCode)> _
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Const WM_SETFOCUS As Integer = &H7
        Const WM_PARENTNOTIFY As Integer = &H210
        Const WM_DESTROY As Integer = &H2
        Const WM_LBUTTONDOWN As Integer = &H201
        Const WM_RBUTTONDOWN As Integer = &H204

        If m.Msg = WM_SETFOCUS Then
            'Raise Enter event
            Me.OnEnter(New System.EventArgs)

        ElseIf m.Msg = WM_PARENTNOTIFY AndAlso _
            (m.WParam.ToInt32 = WM_LBUTTONDOWN OrElse _
             m.WParam.ToInt32 = WM_RBUTTONDOWN) Then

            If Not Me.ContainsFocus Then
                'Raise Enter event
                Me.OnEnter(New System.EventArgs)
            End If

        ElseIf m.Msg = WM_DESTROY AndAlso Not Me.IsDisposed AndAlso Not Me.Disposing Then
            'Used to ensure that VB6 will cleanup control properly
            Me.Dispose()
        End If

        MyBase.WndProc(m)
    End Sub

    'This event will hook up the necessary handlers
    Private Sub InteropUserControl_ControlAdded(ByVal sender As Object, ByVal e As ControlEventArgs) Handles Me.ControlAdded
        ActiveXControlHelpers.WireUpHandlers(e.Control, AddressOf ValidationHandler)
        Me.Size = New System.Drawing.Size(24, 22)
    End Sub

    'Ensures that the Validating and Validated events fire appropriately
    Friend Sub ValidationHandler(ByVal sender As Object, ByVal e As EventArgs)

        If Me.ContainsFocus Then Return

        'Raise Leave event
        Me.OnLeave(e)

        If Me.CausesValidation Then
            Dim validationArgs As New CancelEventArgs
            Me.OnValidating(validationArgs)

            If validationArgs.Cancel AndAlso Me.ActiveControl IsNot Nothing Then
                Me.ActiveControl.Focus()
            Else
                'Raise Validated event
                Me.OnValidated(e)
            End If
        End If

    End Sub

#End Region

#End If

#End Region
    Dim tUserName As String
    Dim tPassWord As String
    Dim tSMTP As String
    Dim tPort As Integer

    Dim tFrom As String
    Dim tTo As String
    Dim tSubject As String
    Dim tBody As String
    Dim tAttachment As String

    Public Property EmailUsername() As String
        Get
            EmailUsername = tUserName
        End Get
        Set(ByVal value As String)
            tUserName = value
        End Set
    End Property
    Public Property EmailPassword() As String
        Get
            EmailPassword = tPassWord
        End Get
        Set(ByVal value As String)
            tPassWord = value
        End Set
    End Property
    Public Property EmailSMTP() As String
        Get
            EmailSMTP = tSMTP
        End Get
        Set(ByVal value As String)
            tSMTP = value
        End Set
    End Property
    Public Property EmailPORT() As Integer
        Get
            EmailPORT = tPort
        End Get
        Set(ByVal value As Integer)
            tPort = value
        End Set
    End Property

    Public Property EmailFrom() As String
        Get
            EmailFrom = tFrom
        End Get
        Set(ByVal value As String)
            tFrom = value
        End Set
    End Property
    Public Property EmailTo() As String
        Get
            EmailTo = tTo
        End Get
        Set(ByVal value As String)
            tTo = value
        End Set
    End Property
    Public Property EmailSubject() As String
        Get
            EmailSubject = tSubject
        End Get
        Set(ByVal value As String)
            tSubject = value
        End Set
    End Property
    Public Property EmailBody() As String
        Get
            EmailBody = tBody
        End Get
        Set(ByVal value As String)
            tBody = value
        End Set
    End Property
    Public Property EmailAttachment() As String
        Get
            EmailAttachment = tAttachment
        End Get
        Set(ByVal value As String)
            tAttachment = value
        End Set
    End Property

    Private Sub SetReciever(ByVal MailMessage As MailMessage, ByVal SendTo As String)
        Dim ListReceiver() As String
        ListReceiver = Split(SendTo, ";")
        If UBound(ListReceiver) >= 1 Then
            For i As Integer = LBound(ListReceiver) To UBound(ListReceiver)
                MailMessage.To.Add(ListReceiver(i))
            Next
        Else
            MailMessage.To.Add(SendTo)
        End If
    End Sub

    Private Sub SetAttachment(ByVal MailMessage As MailMessage, ByVal Attachments As String)
        Dim pAttachment As Net.Mail.Attachment
        Dim ListAttachments() As String
        ListAttachments = Split(Attachments, ";")
        If UBound(ListAttachments) >= 1 Then
            For i As Integer = 1 To UBound(ListAttachments)
                pAttachment = New Net.Mail.Attachment(ListAttachments(0) & "\" & ListAttachments(i))
                MailMessage.Attachments.Add(pAttachment)
            Next
        Else
            pAttachment = New Net.Mail.Attachment(Attachments)
            MailMessage.Attachments.Add(pAttachment)
        End If
    End Sub

    Public Sub vSendEmail()

        'Start by creating a mail message object
        Dim MyMailMessage As New MailMessage()
        'Dim pAttachment As Net.Mail.Attachment = New Net.Mail.Attachment(EmailAttachment)

        'From requires an instance of the MailAddress type
        MyMailMessage.From = New MailAddress(EmailFrom)

        'To is a collection of MailAddress types
        SetReciever(MyMailMessage, EmailTo) 'MyMailMessage.To.Add(EmailTo)

        MyMailMessage.Subject = EmailSubject
        MyMailMessage.Body = EmailBody
        SetAttachment(MyMailMessage, EmailAttachment) 'MyMailMessage.Attachments.Add(pAttachment)


        'Create the SMTPClient object and specify the SMTP GMail server
        Dim SMTPServer As New SmtpClient(EmailSMTP) '"smtp.gmail.com"
        SMTPServer.Port = EmailPORT '587
        SMTPServer.Credentials = New System.Net.NetworkCredential(EmailUsername, EmailPassword)
        SMTPServer.EnableSsl = True

        Try
            SMTPServer.Send(MyMailMessage)
            MessageBox.Show("Email Sent", "ComUnion Ticketing", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As SmtpException
            MessageBox.Show(ex.Message)
        End Try
    End Sub


End Class