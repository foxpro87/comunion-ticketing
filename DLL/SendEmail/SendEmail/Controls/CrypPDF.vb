Imports System
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf


<ComClass(CrypPDF.ClassId, CrypPDF.InterfaceId, CrypPDF.EventsId)> _
Public Class CrypPDF

#Region "VB6 Interop Code"

#If COM_INTEROP_ENABLED Then

#Region "COM Registration"

    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    Public Const ClassId As String = "dc87728f-ad3c-4216-a250-a3a6188f6b41"
    Public Const InterfaceId As String = "7a434126-f73a-49f7-bace-ef032627279f"
    Public Const EventsId As String = "4143b761-47e3-44aa-aee0-9e431c3fe05b"

    'These routines perform the additional COM registration needed by ActiveX controls
    <EditorBrowsable(EditorBrowsableState.Never)> _
    <ComRegisterFunction()> _
    Private Shared Sub Register(ByVal t As Type)
        ComRegistration.RegisterControl(t, 103)
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

    'TODO: Remember to update InteropUserControl.manifest if you are using RegFree COM

#End Region

    'Please enter any new code here, below the Interop code
    Public Sub AddPasswordToPDF(ByVal sourceFile As String, ByVal outputFile As String, ByVal password As String)
        Try
            Dim pReader As PdfReader
            pReader = New PdfReader(sourceFile)
            PdfEncryptor.Encrypt(pReader, New FileStream(outputFile, FileMode.Open), _
                                 PdfWriter.STRENGTH128BITS, password, Nothing, PdfWriter.AllowScreenReaders)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub




End Class