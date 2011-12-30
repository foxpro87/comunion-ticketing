VERSION 5.00
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmToolReprint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ticket Reprint"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   Icon            =   "frmToolReprint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ITGControls.ComunionFrames ComunionFrames2 
      Height          =   1830
      Left            =   0
      Top             =   0
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   3228
      FrameColor      =   16777215
      FillColor       =   16777215
      RoundedCorner   =   0   'False
      Caption         =   ""
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   4
      ColorFrom       =   109785
      ColorTo         =   10874879
      Begin ITGControls.ITGTextBox txtFrom 
         Height          =   285
         Left            =   195
         TabIndex        =   0
         Top             =   630
         Width           =   2985
         _ExtentX        =   5054
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         LabelBackColor  =   16777215
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
         Label           =   "Ticket From : "
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   1280
         TextBoxWidth    =   1645
      End
      Begin ITGControls.ITGTextBox txtTo 
         Height          =   285
         Left            =   3330
         TabIndex        =   1
         Top             =   630
         Width           =   2010
         _ExtentX        =   3334
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         LabelBackColor  =   16777215
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
         Label           =   "To :"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   300
         TextBoxWidth    =   1650
      End
      Begin ITGControls.ComunionButton cmdReprint 
         Height          =   360
         Left            =   3750
         TabIndex        =   2
         Top             =   1140
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "Reprint"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         MICON           =   "frmToolReprint.frx":0CCA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmToolReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Object variables
Private oNavRec As clsNavRec
Private oFormSetup As New clsFormSetup
Private oRecordset As New clsRecordset
Private oBar As New clsToolBarMenuBit

Dim connHeader As New Connection
Dim rsHeader As New Recordset
Attribute rsHeader.VB_VarHelpID = -1

Private Sub Form_Load()
    Set connHeader = New Connection
    connHeader.CursorLocation = adUseClient
    connHeader.ConnectionString = "driver={" & sDBDriver & "};" & "server=" & sServer & ";uid=sa;pwd=" & sDBPassword & ";database=" & sDBname
    connHeader.Open
    
    
    Set FrmName = Me
    oFormSetup.FormLocking False
    oFormSetup.FormSearch True
    oBar.BitVisible ITGLedgerMain.tbrMain
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True
    
    txtFrom.Locked = False
    txtTo.Locked = False
    
    If LoadOption("COM_THEME", 4) = "1 - Blue" Then
        oFormSetup.FormTheme (1)
    Else
        oFormSetup.FormTheme (2)
    End If
    
End Sub

Private Sub cmdReprint_Click()
    Dim i As Integer
    
    If OverRideTransaction("Reprint", txtFrom.Text & " - " & txtTo.Text) = False Then Exit Sub
    
    sSQL = "select * from sales where cTranNo between '" & txtFrom & "' and '" & txtTo & "'"
    If rsHeader.State = 1 Then rsHeader.Close
    rsHeader.Open sSQL, connHeader, adOpenStatic, adLockReadOnly
    
    For i = 0 To rsHeader.RecordCount - 1
        TBPrint rsHeader!cTranNo
        rsHeader.MoveNext
    Next i
    rsHeader.Close
    
    'Audit trail
    UpdateLogFile "Reprint Ticket", txtFrom & " - " & txtTo, "Inserted"


End Sub


'Sub for Printing the transaction
Public Sub TBPrint(cTranNo As String)
    If rs.State = 1 Then rs.Close
    sSQL = "exec rsp_Tran_Sales '" & COID & "','" & cTranNo & "'"
    rs.Open sSQL, cn, adOpenStatic, adLockReadOnly
    
    PrintTicket cTranNo, sThermalPrinterName, App.path & "\Reports\TicketFormat.txt", "Tahoma"
    
    rs.Close
End Sub

Private Sub PrintTicket(cTranNo As String, cPrinterName As String, cFile As String, cFontName As String)
    Dim prnPrinter
    Dim sReprintWord As String
    
    sReprintWord = GetValueFrTable("cValue", "Parameter_User", "cType='Setting' and cParamName='Reprint'")
    
    For Each prnPrinter In Printers
        If prnPrinter.DeviceName = cPrinterName Then '"BIXOLON SRP-370" Then
            Set Printer = prnPrinter
            Exit For
        End If
    Next
    
    Printer.FontSize = 7
    Printer.Print sReprintWord '"Reprinted Copy"
    Printer.FontName = cFontName    '"Tahoma"
    
    Dim disp As String
    Dim tmp As String
    Dim attr As String
    Dim a
    
    Open cFile For Input As #1 '"c:\test.txt" For Input As #1
    While EOF(1) = 0
        Line Input #1, tmp
        
        If Mid(tmp, 1, 2) = "--" Then Printer.EndDoc: GoTo 1
        
        If Len(tmp) <> 0 Then
            If Mid(tmp, 1, 2) <> "//" Then
                attr = Mid(tmp, 2, InStrRev(tmp, "]") - 2)
                If attr = "FONTNAME" Then
                    cFontName = Mid(tmp, InStrRev(tmp, "]") + 1, Len(tmp) - (Len(attr) - 2))
                    GoTo 1
                ElseIf attr = "BARCODE" Then
                    cBarCode = Mid(tmp, InStrRev(tmp, "]") + 1, Len(tmp) - (Len(attr) - 2))
                    GoTo 1
                Else
                    disp = ReplaceCode(Mid(tmp, InStrRev(tmp, "]") + 1, Len(tmp) - (Len(attr) - 2)), cTranNo)
                End If
                a = Split(attr, ",")
                
                Printer.FontName = IIf(a(2) = 1, cBarCode, cFontName)
                Printer.FontSize = a(0)
                Printer.FontBold = IIf(a(1) = 1, True, False)
                If (a(3) = 1) Then
                    Printer.CurrentX = (Printer.ScaleWidth / 2 - Printer.TextWidth(disp) / 2)
                End If
                
                Printer.Print disp
            
            End If
        Else
            Printer.Print
1:     End If
    Wend
    Close #1
    Printer.Print sReprintWord '"Reprinted Copy"
    Printer.EndDoc
End Sub

Function ReplaceCode(cText As String, cTranNo As String) As String
    Dim sText As String
    sText = cText
    sText = Replace(sText, "<cCompName>", rs!cCompName & "")
    sText = Replace(sText, "<cTin>", rs!cTINo & "")
    sText = Replace(sText, "<cAddress>", rs!cAddress & "")
    
    sText = Replace(sText, "<cTranNo>", rs!cTranNo & "")
    sText = Replace(sText, "<nPrice>", rs!nNetPrice - (rs!nNetPrice * 0.12) & "")
    sText = Replace(sText, "<nVat>", rs!nNetPrice * 0.12 & "")
    sText = Replace(sText, "<dDateTime>", rs!dDateTime & "")
    sText = Replace(sText, "<cType>", rs!cType & "")
    sText = Replace(sText, "<nNetPrice>", Format(rs!nNetPrice, "00.00") & "")
    
    sText = Replace(sText, "<dDateTime1>", TimeValue(rs!dDateTime) & "")
    sText = Replace(sText, "<dDateTime2>", DateValue(rs!dDateTime) & "")
    
    sText = Replace(sText, "<cFirstName>", rs!FirstName & "")
    
    ReplaceCode = sText
End Function


