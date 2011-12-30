VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolBarMenuBit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Sets the toolbar and menu status

Option Explicit

Dim sNew, sEdit, sDelete, sNewLine, sDeleteLine, sPost, sCancel, sSave, sUndo, _
    sFind, sNavigation, sPrint, sClose, sAccount, sReference, sNotify, sBit As String


Public Sub BitEnabled(ByRef frm As Form, _
        ByRef Bar As Toolbar, Optional btnNew As Boolean, _
        Optional btnEdit As Boolean, Optional btnDelete As Boolean, _
        Optional btnNewLine As Boolean, Optional btnDeleteLine As Boolean, _
        Optional btnPost As Boolean, Optional btnCancel As Boolean, _
        Optional btnSave As Boolean, Optional btnUndo As Boolean, _
        Optional btnFind As Boolean, Optional btnNavigation As Boolean, _
        Optional btnPrint As Boolean, Optional btnClose As Boolean, _
        Optional btnAccount As Boolean, Optional btnReference As Boolean, _
        Optional btnNotify As Boolean)
    
    'TOOLBAR
    Bar.Buttons("btnNew").Enabled = btnNew
    Bar.Buttons("btnEdit").Enabled = btnEdit
    Bar.Buttons("btnDelete").Enabled = btnDelete
    Bar.Buttons("btnNewLine").Enabled = btnNewLine
    Bar.Buttons("btnDeleteLine").Enabled = btnDeleteLine
    Bar.Buttons("btnPost").Enabled = btnPost
    Bar.Buttons("btnCancel").Enabled = btnCancel
    Bar.Buttons("btnSave").Enabled = btnSave
    Bar.Buttons("btnUndo").Enabled = btnUndo
    Bar.Buttons("btnFind").Enabled = btnFind
    Bar.Buttons("btnFirst").Enabled = btnNavigation
    Bar.Buttons("btnPrevious").Enabled = btnNavigation
    Bar.Buttons("btnNext").Enabled = btnNavigation
    Bar.Buttons("btnLast").Enabled = btnNavigation
    Bar.Buttons("btnPrint").Enabled = btnPrint
    Bar.Buttons("btnClose").Enabled = btnClose
    Bar.Buttons("btnAccount").Enabled = btnAccount
    Bar.Buttons("btnReference").Enabled = btnReference
    Bar.Buttons("btnNotify").Enabled = btnNotify
    
    'MENU
    ITGLedgerMain.mnuFileNew.Enabled = btnNew
    ITGLedgerMain.mnuFileEdit.Enabled = btnEdit
    ITGLedgerMain.mnuFileDelete.Enabled = btnDelete
    ITGLedgerMain.mnuFileNewLine.Enabled = btnNewLine
    ITGLedgerMain.mnuFileDeleteLine.Enabled = btnDeleteLine
'    ITGLedgerMain.mnuFilePost.Enabled = btnPost
'    ITGLedgerMain.mnuFileCancel.Enabled = btnCancel
    ITGLedgerMain.mnuFileSave.Enabled = btnSave
    ITGLedgerMain.mnuFileUndo.Enabled = btnUndo
    ITGLedgerMain.mnuFileFind.Enabled = btnFind
    'ITGLedgerMain.mnuFileFindC.Enabled = btnFind
    ITGLedgerMain.mnuFilePrint.Enabled = btnPrint
    ITGLedgerMain.mnuFileClose.Enabled = btnClose
'    ITGLedgerMain.mnuDetailAdd.Enabled = btnNewLine
'    ITGLedgerMain.mnuDetailRemove.Enabled = btnDeleteLine
'    ITGLedgerMain.mnuFileAccount.Enabled = btnAccount
'    ITGLedgerMain.mnuFileReference.Enabled = btnReference
'    ITGLedgerMain.mnuFileNotify.Enabled = btnNotify
    
    sNew = IIf(btnNew, "1", "0")
    sEdit = IIf(btnEdit, "1", "0")
    sDelete = IIf(btnDelete, "1", "0")
    sNewLine = IIf(btnNewLine, "1", "0")
    sDeleteLine = IIf(btnDeleteLine, "1", "0")
    sPost = IIf(btnPost, "1", "0")
    sCancel = IIf(btnCancel, "1", "0")
    sSave = IIf(btnSave, "1", "0")
    sUndo = IIf(btnUndo, "1", "0")
    sFind = IIf(btnFind, "1", "0")
    sNavigation = IIf(btnNavigation, "1", "0")
    sPrint = IIf(btnPrint, "1", "0")
    sClose = IIf(btnClose, "1", "0")
    sAccount = IIf(btnAccount, "1", "0")
    sReference = IIf(btnReference, "1", "0")
    sNotify = IIf(btnNotify, "1", "0")
    
    sBit = sNew & sEdit & sDelete & sNewLine & sDeleteLine & sPost & _
            sCancel & sSave & sUndo & sFind & sNavigation & sPrint & _
            sClose & sAccount & sReference & sNotify
            
    BitMemory frm, sBit
    
End Sub

Public Sub BitVisible(ByRef Bar As Toolbar, Optional Detail As Boolean, Optional PostCancel As Boolean)
        
    Bar.Buttons("btnPost").Visible = PostCancel
    Bar.Buttons("btnCancel").Visible = PostCancel
    Bar.Buttons("btnSepPostCancel").Visible = PostCancel
    Bar.Buttons("btnNewLine").Visible = Detail
    Bar.Buttons("btnDeleteLine").Visible = Detail
    Bar.Buttons("btnSepDetail").Visible = Detail
    
End Sub

Public Sub BitMemory(ByRef frm As Form, sBit As String)
    frm.sBit = sBit
End Sub

Public Sub BitReload(ByRef frm As Form, ByRef Bar As Toolbar, _
        sBit As String, Optional lNew As Boolean, _
        Optional lEdit As Boolean, Optional lDelete As Boolean, _
        Optional lNewLine As Boolean, Optional lDeleteLine As Boolean, _
        Optional lPost As Boolean, Optional lCancel As Boolean, _
        Optional lSave As Boolean, Optional lUndo As Boolean, _
        Optional lFind As Boolean, Optional lNavigation As Boolean, _
        Optional lPrint As Boolean, Optional lClose As Boolean, _
        Optional lAccount As Boolean, Optional lReference As Boolean, _
        Optional lNotify As Boolean)
    
    If Len(sBit) <> 16 Then sBit = "0000000000000000"
    
    lNew = IIf(Mid(sBit, 1, 1) = "1", True, False)
    lEdit = IIf(Mid(sBit, 2, 1) = "1", True, False)
    lDelete = IIf(Mid(sBit, 3, 1) = "1", True, False)
    lNewLine = IIf(Mid(sBit, 4, 1) = "1", True, False)
    lDeleteLine = IIf(Mid(sBit, 5, 1) = "1", True, False)
    lPost = IIf(Mid(sBit, 6, 1) = "1", True, False)
    lCancel = IIf(Mid(sBit, 7, 1) = "1", True, False)
    lSave = IIf(Mid(sBit, 8, 1) = "1", True, False)
    lUndo = IIf(Mid(sBit, 9, 1) = "1", True, False)
    lFind = IIf(Mid(sBit, 10, 1) = "1", True, False)
    lNavigation = IIf(Mid(sBit, 11, 1) = "1", True, False)
    lPrint = IIf(Mid(sBit, 12, 1) = "1", True, False)
    lClose = IIf(Mid(sBit, 13, 1) = "1", True, False)
    lAccount = IIf(Mid(sBit, 14, 1) = "1", True, False)
    lReference = IIf(Mid(sBit, 15, 1) = "1", True, False)
    lNotify = IIf(Mid(sBit, 16, 1) = "1", True, False)
    
    BitEnabled frm, Bar, lNew, lEdit, lDelete, lNewLine, lDeleteLine, lPost, lCancel, _
        lSave, lUndo, lFind, lNavigation, lPrint, lClose, lAccount, lReference, lNotify
        
    
End Sub

'Button per Access Level
Public Sub AcessBit(ByRef frm As Form, BitCode As String)
    If BitCode = "F" Then
        frm.lACNew = True
        frm.lACEdit = True
        frm.lACDelete = True
        frm.lACPost = True
        frm.lACCancel = True
        frm.lACPrint = True
    Else
        frm.lACNew = IIf(Mid(BitCode, 1, 1) = "1", True, False)
        frm.lACEdit = IIf(Mid(BitCode, 2, 1) = "1", True, False)
        frm.lACDelete = IIf(Mid(BitCode, 3, 1) = "1", True, False)
        frm.lACPost = IIf(Mid(BitCode, 4, 1) = "1", True, False)
        frm.lACCancel = IIf(Mid(BitCode, 5, 1) = "1", True, False)
        frm.lACPrint = IIf(Mid(BitCode, 6, 1) = "1", True, False)
    End If
End Sub

