VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmSaleTicketing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Ticketing"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1725
      TabIndex        =   0
      Text            =   "1"
      Top             =   735
      Width           =   3090
   End
   Begin VB.CommandButton cmdShowPrinterOption 
      Caption         =   "&."
      Height          =   270
      Left            =   45
      TabIndex        =   1
      Top             =   2355
      Width           =   330
   End
   Begin VB.CheckBox chkPrinter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Printer Off"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbSales 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1605
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   7673
            MinWidth        =   4304
            Text            =   "REG:000,000 | STD:000,000 | SNR:000,000 | SPC:000,000"
            TextSave        =   "REG:000,000 | STD:000,000 | SNR:000,000 | SPC:000,000"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ITGControls.ComunionButton cmdPrint 
      Height          =   705
      Left            =   675
      TabIndex        =   3
      Top             =   735
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1244
      BTYPE           =   3
      TX              =   "&Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      MICON           =   "frmSaleTicketing.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   585
      TabIndex        =   4
      Top             =   90
      Width           =   4245
   End
End
Attribute VB_Name = "frmSaleTicketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oFormSetup As New clsFormSetup
Private connHeader As ADODB.Connection
Private TypeRS As New ADODB.Recordset
Private WithEvents rsHeader As ADODB.Recordset
Attribute rsHeader.VB_VarHelpID = -1

' DLL Function prototypes
Private Declare Function OpenUsbPort Lib "CommUSB.dll" (ByVal dwModel As Integer) As Integer
Private Declare Function WriteUSB Lib "CommUSB.dll" (ByVal byWrite As String, ByVal dwWrite As Long) As Long
Private Declare Function ReadUSB Lib "CommUSB.dll" (ByVal byWrite As String, ByVal dwWrite As Long) As Long
Private Declare Function CloseUsbPort Lib "CommUSB.dll" () As Integer

Private tType() As String
Private lEnter As Boolean

Private sHost As String
Private cTicketType As String
Private cType As String
Private Disc As String

Private nNet As Double
Private nGetPrice As Double
Private keyStr As Integer

Private Sub chkPrinter_Click()
    txtQty.SetFocus
End Sub

Private Sub cmdPrint_Click()
    txtQty_KeyDown 13, -1
End Sub

Private Sub Form_Load()
    Me.Icon = ITGLedgerMain.Icon
    Set FrmName = Me
    If LoadOption("COM_THEME", 4) = "1 - Blue" Then
        oFormSetup.FormTheme (1)
    Else
        oFormSetup.FormTheme (2)
    End If
    
    TypeInit
    
    'Open DB Connection
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT HOST_NAME() as cHost", cn, adOpenStatic, adLockReadOnly
    sHost = rs!cHost
    rs.Close
    
    Dim nInitPrice As Single
    sSQL = "SELECT DISTINCT a.cTypeID, CAST(b.nPrice -(b.nPrice * CONVERT(numeric,a.cDiscount)/100) as NUMERIC(18,2)) nPrice" & _
                 " FROM Ticket_Type a " & _
                  "  LEFT OUTER JOIN Ticket_Pricing b on a.cCompanyID = b.cCompanyID"
    rs.Open sSQL, cn, adOpenStatic, adLockReadOnly
    ReDim nTickePrice(rs.RecordCount) As Single
    ReDim sTickePrice(rs.RecordCount) As String
    For i = 0 To UBound(nTickePrice) - 1
        nTickePrice(i) = rs!nPrice
        sTickePrice(i) = rs!cTypeID
        If sTickePrice(i) = "REGULAR" Then
            nInitPrice = nTickePrice(i)
        End If
        rs.MoveNext
    Next i
    lblTotal.Caption = Format(nInitPrice * Val(txtQty), "#,##0.00")
    
    Set connHeader = New Connection
    connHeader.CursorLocation = adUseClient
    connHeader.ConnectionString = "driver={" & sDBDriver & "};" & "server=" & sServer & ";uid=sa;pwd=" & sDBPassword & ";database=" & sDBname
    connHeader.Open
    
    Set rsHeader = New Recordset
    rsHeader.Open "SELECT * FROM SALES WHERE 1=0", connHeader, adOpenStatic, adLockBatchOptimistic

End Sub

Private Sub txtQty_Change()
    lblTotal.Caption = Format(nNet * Val(txtQty), "#,##0.00")
    If txtQty.Text = "" Then
        txtQty.Text = "1"
        txtQty.SelStart = 0
        txtQty.SelLength = Len(txtQty)
    End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = 13 Then
        If Shift = -1 Then lEnter = True Else lEnter = False
        TBTransaction
        lEnter = False
    ElseIf KeyCode = 27 Then
        Unload Me
    Else
        For i = 0 To UBound(tType)
            If tType(i) = KeyCode Then
                 cTicketType = GetValueFrTable("cTypeID", "Ticket_Type", "cKeyStroke='" & KeyCode & "'", True)
                  TicketPrice cTicketType, cType, nGetPrice, nNet, Disc
                  lblTotal.Caption = Format(nNet * Val(txtQty), "#,##0.00")
                  Me.Caption = "Sales Ticketing: " & cTicketType
            End If
        Next i
    End If
End Sub

Private Sub cmdPrint_GotFocus()
    txtQty.SetFocus
End Sub

Private Sub TBTransaction()
    Dim nResult As Integer
    Dim PrintContent As String
    Dim sCmd As String
    
    Dim i As Integer
    
    On Error GoTo ErrHandler
    err.Clear
    
    If Val(txtQty) > 10 Then
        If CInt(txtQty) > 75 Then
            MsgBox "Quantity of ticket is " & txtQty.Text & " which is greater than Seventy Five(75)"
            txtQty = 1
            Exit Sub
        Else
            If MsgBox("Quantity of ticket is " & txtQty.Text & " which is greater than ten(10)" & vbCrLf & _
                "Are you sure you want to continue?", vbQuestion + vbYesNo, "Confirmation") <> vbYes Then
                    txtQty = 1
                    Exit Sub
            End If
        End If
    End If
    
    If chkPrinter.Value = 1 Then
        nResult = 1
    Else
        nResult = OpenUsbPort(2)
    End If
    
    If nResult = 1 Then
        frmSaleTicketing.Enabled = False
        txtQty.Enabled = False
        
        ReDim cTranNo(1 To txtQty.Text) As String
                
        Dim InitSeries As String
        gblQty = CInt(txtQty.Text)
        InitSeries = Generate_AutoNumber("SALESTICK", cTicketType)
        
        Do While GetValueFrTable("lLock", "Autonum", "cModuleID = 'SALESTICK' and cType = '" & cTicketType & "'", True) = True
            Me.Caption = "Sales Ticketing : Please Wait."
            DoEvents
        Loop
        Me.Caption = "Sales Ticketing"
        
        cn.Execute "UPDATE autonum SET lLock=1 WHERE cModuleID = 'SALESTICK' AND cType = '" & cTicketType & "'"
        
        'Get cType, nGetPrice, nNet, Disc
        TicketPrice cTicketType, cType, nGetPrice, nNet, Disc
        
        
        For i = 0 To Val(txtQty.Text) - 1
            cTranNo(i + 1) = Mid$(InitSeries, 1, nNumStart) & Format(CDbl(Mid(InitSeries, nNumStart + 1, nNumLen + 1)) + i, Replace(Mid$(AutonumFormat, nNumStart, nNumLen + 1), gblsNumeric, "0"))
            TBSave cTranNo(i + 1), cType, nGetPrice, nNet, Disc
        Next i
        
        cn.Execute "UPDATE autonum SET lLock=0 WHERE cModuleID = 'SALESTICK' AND cType = '" & cTicketType & "'"
        
        If chkPrinter.Value = 0 Then
            For i = 0 To Val(txtQty.Text) - 1
                TBPrint (cTranNo(i + 1))
            Next
        End If
       
        Beep
        txtQty.Enabled = True
        frmSaleTicketing.Enabled = True
        keyStr = 0
        TypeRS.Filter = "cTypeID='REGULAR'"
        cTicketType = TypeRS!cTypeID
        Me.Caption = "Sales Ticketing: " & cTicketType
        
        'Audit trail
        UpdateLogFile "Sales Ticketing", Trim(txtQty), "Inserted: " & IIf(lEnter = True, "Click", "Enter")
        
        txtQty.SetFocus
        txtQty.Text = "1"
        TicketPrice cTicketType, cType, nGetPrice, nNet, Disc
        lblTotal.Caption = Format(nNet * Val(txtQty), "#,##0.00")
        txtQty.SelStart = 0
        txtQty.SelLength = Len(txtQty)
         
        sbSales.Panels(1).Text = " | " & TicketCount
        
        If chkPrinter.Value = 0 Then
            'Optional - open the cash drawer
            sCmd = Chr(27) & Chr(112) & Chr(0) & Chr(25) & Chr(255)
            nResult = WriteUSB(sCmd, Len(sCmd))  'nResult = Written Count
            CloseUsbPort
        End If
    Else
        MsgBox "The Printer is Disconnected", vbCritical, "ComUnion - Ticketing"
    End If
        
ErrHandler:
    If err.Number <> 0 Then
        MsgBox err.Description
        
        txtQty.Enabled = True
        frmSaleTicketing.Enabled = True
        
    End If
End Sub

Public Sub TBSave(cTranNo As String, cType As String, nPrice As Double, nNetPrice As Double, cDiscount As String)
    With rsHeader
        .AddNew
        !cCompanyID = COID
        !cTerminal = sHost
        !cTranNo = cTranNo
        !dDateTime = Now
        !cType = IIf(cType = "", "REGULAR", cType)
        !nPrice = nPrice
        !nNetPrice = nNetPrice
        !cDiscount = cDiscount
        !cUserNo = sUserName
        !dTranDate = Format(Now, "mm/dd/yyyy")
    End With
    
    connHeader.BeginTrans
    rsHeader.UpdateBatch adAffectAll
    connHeader.CommitTrans
End Sub

'Sub for Printing the transaction
Public Sub TBPrint(cTranNo As String)
    If rs.State = 1 Then rs.Close
    sSQL = "EXEC rsp_Tran_Sales '" & COID & "','" & cTranNo & "'"
    rs.Open sSQL, cn, adOpenStatic, adLockReadOnly
    
    If FileExists(App.path & "\Reports\TicketFormat.txt") = True Then
        PrintTicket cTranNo, sThermalPrinterName, App.path & "\Reports\TicketFormat.txt", "Tahoma"
    Else
        MsgBox "TicketFormat.txt file does not exist"
    End If
    rs.Close
End Sub

Private Sub PrintTicket(cTranNo As String, cPrinterName As String, cFile As String, cFontName As String)
    Dim cBarCode As String
    Dim prnPrinter
    For Each prnPrinter In Printers
        If prnPrinter.DeviceName = cPrinterName Then '"BIXOLON SRP-370" Then
            Set Printer = prnPrinter
            Exit For
        End If
    Next
    
    Printer.FontSize = 9.5
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

Sub TypeInit()
    On Error GoTo ErrHandler
    Set TypeRS = New Recordset
    If TypeRS.State = 1 Then TypeRS.Close
    sSQL = "SELECT * FROM Ticket_Type"
    TypeRS.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    ReDim tType(0 To TypeRS.RecordCount - 1)
    For i = 0 To TypeRS.RecordCount - 1
        tType(i) = TypeRS!cKeyStroke
        TypeRS.MoveNext
    Next i
    cTicketType = "REGULAR"
    Me.Caption = "Sales Ticketing: " & cTicketType
    TicketPrice cTicketType, cType, nGetPrice, nNet, Disc
    
    txtQty = "1"
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty)
    
ErrHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical + vbOKOnly, "ComUnion"
    End If
End Sub

Function TicketCount() As String
    On Error GoTo ErrHandler
    Dim sTMP As String
    Dim i As Integer
    Dim TCrs As New Recordset
    Dim cmd As New Command
    With cmd
        .ActiveConnection = cn
        .CommandTimeout = 1000
        .CommandText = "TicketCount"
        .CommandType = adCmdStoredProc
        .Parameters("@dDateTime") = dTranDate
        .Parameters("@cUserID") = sUserName
        If TCrs.State = 1 Then TCrs.Close
        Set TCrs = cmd.Execute()
    End With
    
    For i = 0 To TCrs.RecordCount - 1
        DoEvents
        sTMP = sTMP & TCrs!cIssues & " | "
        TCrs.MoveNext
    Next i
    TicketCount = sTMP
ErrHandler:
    If err.Number <> 0 Then
        MsgBox err.Number & ": " & err.Description, vbCritical, "ComUnion"
    End If
End Function

Private Sub TicketPrice(cTicketType As String, ByRef cType As String, ByRef nGetPrice As Double, ByRef nNet As Double, ByRef cDisc As String)
    Dim tmpRS As New Recordset
    sSQL = "SELECT DISTINCT a.cTypeID,b.nPrice as nGetPrice,a.cDiscount as cDisc" & _
                "   ,CAST(b.nPrice -(b.nPrice * CONVERT(numeric,a.cDiscount)/100) as NUMERIC(18,2)) nNet " & _
                 " FROM Ticket_Type a " & _
                 "   LEFT OUTER JOIN Ticket_Pricing b on a.cCompanyID = b.cCompanyID " & _
                 " WHERE A.cTypeID = '" & cTicketType & "' "
    tmpRS.Open sSQL, cn, adOpenStatic, adLockReadOnly
    
    cType = cTicketType
    nGetPrice = tmpRS!nGetPrice
    nNet = tmpRS!nNet
    cDisc = tmpRS!cDisc
    
    tmpRS.Close
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case 13
        Case Else
            KeyAscii = 0
    End Select
End Sub
