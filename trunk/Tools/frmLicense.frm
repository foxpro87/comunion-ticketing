VERSION 5.00
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmLicense 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ComUnion Registration"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4845
   Icon            =   "frmLicense.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin ITGControls.ITGTextBox txtActivation 
      Height          =   285
      Left            =   225
      TabIndex        =   0
      Top             =   1050
      Width           =   4395
      _ExtentX        =   7541
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Label           =   "Activation Key"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   1200
      TextBoxWidth    =   3135
   End
   Begin ITGControls.ITGTextBox txtSerial 
      Height          =   285
      Left            =   225
      TabIndex        =   1
      Top             =   1455
      Width           =   4395
      _ExtentX        =   7541
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Label           =   "Serial key"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   1200
      TextBoxWidth    =   3135
   End
   Begin ITGControls.ComunionButton cmdGenSerial 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   1830
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Registered"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLicense.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   0
      Picture         =   "frmLicense.frx":0CE6
      Top             =   0
      Width           =   4845
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SerialKeyCode As eSerial

Dim cAppName As String
Dim cSection As String
Dim cKey As String

Dim HSerial As String

Private Sub cmdGenSerial_Click()
    If txtSerial = "" Then Exit Sub
    If authKey(txtSerial, txtActivation) = True Then
        'Write to registry
        cAppName = Encrypt("COMUNION")
        cSection = Encrypt("SECURITY")
        cKey = Encrypt("SERIAL")
                
        SaveSetting cAppName, cSection, cKey, StrReverse(Replace(txtSerial, " - ", ""))
        
        'Save to encrypted file
        Dim a
        a = Split(txtActivation, " - ")
        With SerialKeyCode
            .lSeparator1 = Mid(a(0), 1, 4) & vbCrLf
            .lSeparator2 = Mid(a(1), 1, 4) & vbCrLf
            .lSeparator3 = Mid(a(2), 1, 4) & vbCrLf
            .lSeparator4 = Mid(a(3), 1, 4) & vbCrLf
            .CompanyID = COID
            .HSerial = HSerial
            .ActivationKey = StrReverse(Replace(Trim(txtActivation), " - ", ""))
            .SerialKey = StrReverse(Replace(Trim(txtSerial), " - ", ""))
        End With
                        
        Open App.path & "\DLL\connection.dll" For Random As #1 Len = Len(SerialKeyCode)
        Put #1, 1, SerialKeyCode
        Close #1
        
        MsgBox "Thank you for registering ComUnion!", vbInformation + vbOKOnly, "ComUnion"
        lRegistered = True
        Unload Me
    Else
        MsgBox "Serial Number is incorrect!", vbCritical + vbOKOnly, "ComUnion"
        txtSerial.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    HSerial = Hex$(GetSerialNumber(Left(App.path, 1)))
    
    If FileExists(App.path & "\DLL\connection.dll") = True Then
        Open App.path & "\DLL\connection.dll" For Random As #1 Len = Len(SerialKeyCode)
            Get #1, 1, SerialKeyCode
        Close #1
        Dim tmpActivation As String
        tmpActivation = StrReverse(SerialKeyCode.ActivationKey)
        
        txtActivation = txtActivation & Mid(tmpActivation, 1, 5) & " - "
        txtActivation = txtActivation & Mid(tmpActivation, 6, 5) & " - "
        txtActivation = txtActivation & Mid(tmpActivation, 11, 5) & " - "
        txtActivation = txtActivation & Mid(tmpActivation, 16, 5) & " - "
        txtActivation = txtActivation & Mid(tmpActivation, 21, 5)
        
        
    Else
        txtActivation = genNumber(HSerial)
        Dim a
        a = Split(txtActivation, " - ")
        With SerialKeyCode
            .lSeparator1 = Mid(a(0), 1, 4) & vbCrLf
            .lSeparator2 = Mid(a(1), 1, 4) & vbCrLf
            .lSeparator3 = Mid(a(2), 1, 4) & vbCrLf
            .lSeparator4 = Mid(a(3), 1, 4) & vbCrLf
            .CompanyID = COID
            .HSerial = HSerial
            .ActivationKey = StrReverse(Replace(Trim(txtActivation), " - ", ""))
            .SerialKey = StrReverse(Replace(Trim(txtSerial), " - ", ""))
        End With
        
        Open App.path & "\DLL\connection.dll" For Random As #2 Len = Len(SerialKeyCode)
        Put #2, 1, SerialKeyCode
        Close #2
        
        Populate_DLL
        
    End If
ErrHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, , "License"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ITGLedgerMain.GetAccessLevel
End Sub
