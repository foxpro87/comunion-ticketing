VERSION 5.00
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmSecChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security - Change Password"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   4035
      Begin ITGControls.ITGTextBox txtUserID 
         Height          =   285
         Left            =   300
         TabIndex        =   0
         Top             =   300
         Width           =   3315
         _ExtentX        =   5636
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   25
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
         TextBoxWidth    =   1755
      End
      Begin ITGControls.ITGTextBox txtOld 
         Height          =   285
         Left            =   300
         TabIndex        =   1
         Top             =   660
         Width           =   3315
         _ExtentX        =   5636
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
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
         MaxLength       =   10
         Passwordchar    =   "*"
         Label           =   "Old Password"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   1755
      End
      Begin ITGControls.ITGTextBox txtNew 
         Height          =   285
         Left            =   300
         TabIndex        =   2
         Top             =   1020
         Width           =   3315
         _ExtentX        =   5636
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
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
         MaxLength       =   10
         Passwordchar    =   "*"
         Label           =   "New Password"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   1755
      End
      Begin ITGControls.ITGTextBox txtConfirm 
         Height          =   285
         Left            =   300
         TabIndex        =   3
         Top             =   1380
         Width           =   3315
         _ExtentX        =   5636
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
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
         MaxLength       =   10
         Passwordchar    =   "*"
         Label           =   "Confirm Password"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   1755
      End
   End
   Begin ITGControls.ITGCommandButton cmdCancel 
      Height          =   345
      Left            =   2880
      TabIndex        =   5
      Top             =   2100
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Cancel"
   End
   Begin ITGControls.ITGCommandButton cmdOK 
      Height          =   345
      Left            =   1620
      TabIndex        =   4
      Top             =   2100
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
   End
   Begin VB.Label Label1 
      Caption         =   "Min Char -- 6 Max Char -- 10"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmSecChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

Private LastPasswordChange As Date
Private pword As Boolean

Private rs1 As Recordset
Private rs2 As Recordset

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
    If Trim(UCase(txtOld)) = "" Then
        MsgBox "Please type your Password .", vbInformation, "ComUnion"
        Exit Sub
    End If
    
    Set rs2 = New Recordset
    sSQL = "SELECT * FROM SEC_USER WHERE UserID = '" & UCase$(Trim(txtUserID)) & "'"
    rs2.Open sSQL, cn, adOpenKeyset
    If rs2.RecordCount <> 0 Then
        If Trim(UCase(txtOld)) <> UCase(Decrypt(Trim(rs2!Password))) Then
            MsgBox "Invalid Password.", vbInformation, "ComUnion"
            txtOld.SetFocus
            Exit Sub
        End If
        
        If Trim$(txtNew) = Trim$(txtUserID) Then
            MsgBox "New Password should not be the same as User ID!", vbInformation, "ComUnion"
            txtNew.SetFocus
            Exit Sub
        ElseIf Len(txtNew) < 6 Then
            MsgBox "New Password should be at least 6 characters!", vbInformation, "ComUnion"
            txtNew.SetFocus
            Exit Sub
        ElseIf Left(txtNew, 1) = " " Then
            MsgBox "New Password should not contain leading blank(s)!", vbInformation, "ComUnion"
            txtNew.SetFocus
            Exit Sub
        ElseIf Right(txtNew, 1) = " " Then
            MsgBox "New Password should not contain trailing blank(s)!", vbInformation, "ComUnion"
            txtNew.SetFocus
            Exit Sub
        Else
            For i = 1 To (Len(txtNew) - 2)
                If Mid(txtNew, i, 1) = Mid(txtNew, i + 1, 1) And Mid(txtNew, i, 1) = Mid(txtNew, i + 2, 1) Then
                    MsgBox "New Password should not contain more than two consecutive, identical characters!", vbInformation, "ComUnion"
                    txtNew.SetFocus
                    Exit Sub
                End If
            Next i

            If txtNew <> txtConfirm Then
                MsgBox "Confirmation rejected!", vbInformation, "ComUnion"
                txtConfirm.SetFocus
                Exit Sub
            End If
        End If

        sSQL = "UPDATE SEC_USER SET Password = '" & Encrypt(Trim$(txtNew)) & "' " & _
                 "WHERE UserID = '" & Trim$(txtUserID) & "'"
        cn.Execute sSQL

        MsgBox "Password successfully changed.", vbInformation, "ComUnion"

        pword = False

        Unload Me
    Else
        MsgBox "User Id does not exist.", vbInformation, "ComUnion"
        txtUserID.SetFocus
    End If
    rs2.Close
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set rs2 = New Recordset
 
    Set rs1 = New Recordset
    
    pword = False
    txtUserID = sUserName
    sSQL = "SELECT * FROM SEC_USER WHERE UserID = '" & Trim(sUserName) & "'"
    rs1.Open sSQL, cn, adOpenKeyset
    If rs1.RecordCount > 0 Then
        If Decrypt(Trim(rs1!Password)) = "PASSWORD" Then
            pword = True
        End If
    End If
    rs1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set frmSecChangePassword = Nothing
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
    KeyAscii = InvalidKeys(KeyAscii, "`'&")
End Sub

Private Sub txtNew_KeyPress(KeyAscii As Integer)
    KeyAscii = InvalidKeys(KeyAscii, "`'&")
End Sub

Private Sub txtOld_keypress(KeyAscii As Integer)
    KeyAscii = InvalidKeys(KeyAscii, "`'&")
End Sub


