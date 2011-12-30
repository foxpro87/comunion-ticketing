VERSION 5.00
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmLogon 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   DrawMode        =   1  'Blackness
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2310
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "   LOG ON   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   210
      TabIndex        =   0
      Top             =   990
      Width           =   4455
      Begin VB.ComboBox CMBdept 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1770
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   2340
      End
      Begin ITGControls.ITGTextBox ITGTname 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   3855
         _ExtentX        =   6588
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         BackColor       =   12648447
         LabelBackColor  =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "sa"
         Label           =   "User ID"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   2295
      End
      Begin ITGControls.ITGTextBox ITGTpassword 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   3855
         _ExtentX        =   6588
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         BackColor       =   12648447
         LabelBackColor  =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllCaps         =   -1  'True
         Passwordchar    =   "*"
         Label           =   "Password"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   2295
      End
      Begin ITGControls.ITGTextBox ITGTextBox3 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1140
         Width           =   1335
         _ExtentX        =   2566
         _ExtentY        =   503
         LabelBackColor  =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Label           =   "Company"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   15
      End
      Begin ITGControls.ITGTextBox ITGCompany 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1155
         Width           =   3825
         _ExtentX        =   6535
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         LabelBackColor  =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllCaps         =   -1  'True
         Locked          =   -1  'True
         Label           =   ""
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   2265
         Enabled         =   0   'False
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   3
      X1              =   4890
      X2              =   4890
      Y1              =   0
      Y2              =   3180
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   4875
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   3
      X1              =   15
      X2              =   15
      Y1              =   15
      Y2              =   3210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   4860
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   30
      Picture         =   "frmLogon.frx":08CA
      Top             =   0
      Width           =   4845
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim rsUnitComp As ADODB.Recordset
Dim rsUserList As ADODB.Recordset
Dim lConnect As Boolean

Private Sub CMBdept_Click()
    On Error GoTo ErrHandler
    ITGCompany = GetValueFrTable("cCompanyName", "Company", "cCompanyID = '" & Trim(CMBdept.Text) & "'", True)

ErrHandler:
    If err.Number = 13 Then
            'ErrorLog err.Number, err.Description, Me.Name 'Error log
            Resume Next
    End If
End Sub

Private Sub CMBdept_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 110
            Unload Me
        Case vbKeyEscape
            Unload Me
        Case Else
            'ITGCompany = GetValueFrTable("cCompanyName", "Company", "cCompanyID = '" & Trim(CMBdept.Text) & "'", True)
    End Select
End Sub

Private Sub CMBdept_KeyPress(KeyAscii As Integer)
    Dim LN, FN, MI As String
    If KeyAscii = 13 Then
        If Trim(CMBdept) = "" Then
            MsgBox "Please select company.", vbExclamation, "ComUnion"
            CMBdept.SetFocus
            Exit Sub
        Else
            COID = Trim(CMBdept)
            sUserName = UCase(Trim(ITGTname))
            
            LN = GetValueFrTable("LastName", "SEC_USER", "UserID = '" & Trim(ITGTname) & "'", True)
            FN = GetValueFrTable("FirstName", "SEC_USER", "UserID = '" & Trim(ITGTname) & "'", True)
            MI = GetValueFrTable("MI", "SEC_USER", "UserID = '" & Trim(ITGTname) & "'", True)
            UserFullName = LN & " " & FN & ", " & MI
            UserRole = GetValueFrTable("RoleID", "SEC_USER", "UserID = '" & Trim(ITGTname) & "'", True)
            GetMain
            UpdateLogin
        End If
    End If
    ITGCompany = GetValueFrTable("cCompanyName", "Company", "cCompanyID = '" & Trim(CMBdept.Text) & "'", True)
End Sub

Private Sub Form_Load()
    lConnect = False
        
    If UserNameLog() = "" Then
        MsgBox "No Network Connection. " & vbCr & _
            "Please Log your Network Password.", vbCritical, "Log On"
    Else: ITGTname = LCase(UserNameLog)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsUserList = Nothing
    Set rsUnitComp = Nothing
    Set frmLogon = Nothing
End Sub

Private Sub Frame1_DblClick()
    If Not lConnect Then Exit Sub
    If Trim(ITGTname) = "" Then Exit Sub
    
    sSQL = "UPDATE SEC_USER SET PassErrCtr = 0, Locked = 0, InUse = 0 WHERE UserID = '" & Trim(ITGTname) & "' "
    cn.Execute sSQL

End Sub


Private Sub ITGTname_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 110
            Unload Me
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub ITGTpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 110
            Unload Me
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Public Sub SetConnection()
    Dim sLine As String
    Dim lPos As Long
    
    'Opens ini file
    Open App.path & "\connection.ini" For Input Access Read As #1
    Do While Not EOF(1)
      Input #1, sLine
      sLine = Trim(sLine)
        
      'Acquiring ini data
      If InStr(sLine, "HOST") Then
        lPos = InStr(sLine, "=")
        sServer = Trim(Right(sLine, Len(sLine) - lPos))
      ElseIf InStr(sLine, "DATASOURCE") Then
        lPos = InStr(sLine, "=")
        sDBname = Trim(Right(sLine, Len(sLine) - lPos))
      ElseIf InStr(sLine, "SOURCE TYPE") Then
        lPos = InStr(sLine, "=")
        sDBDriver = Trim(Right(sLine, Len(sLine) - lPos))
      ElseIf InStr(sLine, "PASSWORD") Then
        lPos = InStr(sLine, "=")
        sDBPassword = Trim(Right(sLine, Len(sLine) - lPos))
      End If
    Loop
    
    Close
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    
    'Connection for 'sa' user only [modified 08.27.2003 by {moymoy}]------
    cn.ConnectionString = "driver={" & sDBDriver & "};" & _
    "server=" & sServer & ";uid=sa;pwd=" & sDBPassword & ";database=" & sDBname

    cn.Open
    
    lConnect = True

    If Login_OK Then
        
        Set rsUnitComp = New Recordset
        rsUnitComp.Open "SELECT cCompanyID FROM SEC_COMPANYACCESS WHERE UserID = '" & Trim(ITGTname) & "'", cn, adOpenForwardOnly, adLockReadOnly

        CMBdept.Clear

        Do While Not rsUnitComp.EOF
            CMBdept.AddItem Trim(rsUnitComp!cCompanyID)
            rsUnitComp.MoveNext
        Loop
        Set rsUnitComp = Nothing
        
        CMBdept.ListIndex = 0
        CMBdept.SetFocus
        ITGCompany = IIf(GetValueFrTable("cCompanyName", "Company", "cCompanyID = '" & Trim(CMBdept) & "'", True), GetValueFrTable("cCompanyName", "Company", "cCompanyID = '" & Trim(CMBdept) & "'", True), "")
    End If
        
End Sub

Function Login_OK() As Boolean
Dim rs As New Recordset
        
    Login_OK = False
    ITGCompany = ""
    
    sSQL = "SELECT * FROM SEC_USER WHERE UserID = '" & Trim(ITGTname) & "'"
    rs.Open sSQL, cn, adOpenKeyset
    If rs.RecordCount = 0 Then
        MsgBox "User not found!", vbInformation, "ComUnion"
        ITGTname.SetFocus
    Else
        If Trim(UCase(ITGTpassword)) <> UCase(Decrypt(Trim(rs!Password))) Then
            If rs!Locked = True Then
                MsgBox "Account Locked." & Chr(13) & "Consult system administrator.", vbInformation, "ComUnion"
                ITGTname.SetFocus
            Else
                UpdateLockOut (rs!PassErrCtr)
                MsgBox "Invalid Password!", vbInformation, "ComUnion"
                ITGTpassword.SetFocus
            End If
        Else
            If rs!Locked = True Then
                MsgBox "Account Locked." & Chr(13) & "Consult system administrator.", vbInformation, "ComUnion"
                ITGTname.SetFocus
            Else
                If rs!Inuse = True Then
                    MsgBox "User currently logged in", vbInformation, "ComUnion"
                    ITGTname.SetFocus
                Else
                    SecUserID = UCase(Trim(ITGTname))
                    SecUserRole = UCase(Trim(rs!RoleID)) & ""
                    SecUserName = UCase(Trim(rs!LastName)) & ", " & UCase(Trim(rs!FirstName)) & " " & UCase(Trim(rs!MI)) & "."
                    Login_OK = True
                End If
            End If
        End If
    End If
    
    Set rs = Nothing
    
End Function

Sub UpdateLogin()
    sSQL = "UPDATE SEC_USER SET PassErrCtr = 0, InUse =1 WHERE UserID = '" & Trim(ITGTname) & "'"
    cn.Execute sSQL
End Sub

Sub UpdateLockOut(counter)
    If (counter + 1) >= 3 Then
       sSQL = "UPDATE SEC_USER SET Locked =  1 WHERE UserID = '" & Trim(ITGTname) & "'"
       cn.Execute sSQL
    End If
    
    sSQL = "UPDATE SEC_USER SET PassErrCtr = PassErrCtr + 1 WHERE UserID = '" & Trim(ITGTname) & "'"
    cn.Execute sSQL
End Sub


Private Sub ITGTpassword_LostFocus()
On Error GoTo ErrHandler
    FormWaitShow "Connecting to server . . ."
    SetConnection
ErrHandler:
    Unload frmWait
    If err.Number = -2147217843 Or err.Number = -2147467259 Then
        MsgBox "Connection Failed!", vbCritical + vbOKOnly, "ALERT"
        ITGTname.SetFocus
    ElseIf err.Number = 13 Then
            'ErrorLog err.Number, err.Description, Me.Name 'Error log
            Resume Next
    End If
End Sub

Public Function GetState(intState As Integer) As String
   Select Case intState
      Case adStateClosed
         MsgBox "Access Denied.", vbCritical + vbOKOnly, "ALERT"
         ITGTpassword.SetFocus
      Case adStateOpen
         GetMain
   End Select
End Function

Private Sub GetMain()
    frmAccountingSplash.Show
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim sPW As String
    
    lRegistered = CheckIfRegistered
    
    MsgBox "Welcome " & SecUserName & vbCrLf & _
                        "Your actual login time is: " & Time & "." & vbCrLf & _
                        "Your Local time is now synchronized with server time.", vbInformation + vbOKOnly, "ComUnion"
                    
    ITGLedgerMain.lSaleLoaded = False
    ITGLedgerMain.Show
    
    If lRegistered = False Then
        frmLicense.Show vbModal
    End If
    
    cCompany = ITGCompany.Text
    frmAccountingSplash.ZOrder
    Timer1.Enabled = False
    sPW = Trim(ITGTpassword.Text)
    dTranDate = Now
    Unload Me
    If Trim(UCase(sPW)) = "PASSWORD" Then frmSecChangePassword.Show vbModal
        
End Sub

Function UserNameLog() As String
   Dim temp As String
   temp = String(100, Chr$(0))
   GetUserName temp, 100
   UserNameLog = UCase(Left$(temp, InStr(temp, Chr$(0)) - 1))
End Function

