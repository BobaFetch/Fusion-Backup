VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PadmPRf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Product Classes By Part Type"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PadmPRf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbTyp 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Tag             =   "8"
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Product Class"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Apply"
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Update And Apply Changes"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4800
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2340
      FormDesignWidth =   5760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "PadmPRf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions

Dim bOnLoad As Byte
Dim bGoodCode As Byte

Dim sType(8, 2) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd





Private Sub cmbCls_Click()
   bGoodCode = GetCode()
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   If Len(cmbCls) Then bGoodCode = GetCode()
   
End Sub


Private Sub cmbTyp_Click()
   On Error Resume Next
   cmbTyp.ToolTipText = sType(cmbTyp.ListIndex, 1)
   
End Sub


Private Sub cmbTyp_LostFocus()
   Dim iList As Integer
   If Val(cmbTyp) < 1 Or Val(cmbTyp) > 8 Then
      'Beep
      cmbTyp = "1"
   End If
   For iList = 0 To 7
      If Val(cmbTyp) = iList + 1 Then cmbTyp.ToolTipText = sType(iList, 1)
   Next
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   If bGoodCode Then UpdateParts
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1355
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub



Private Sub Form_Activate()
   Dim iType As Integer
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductClasses
      If cmbCls.ListCount > 0 Then
         cmbCls = cmbCls.List(0)
         bGoodCode = GetCode()
         For iType = 0 To 6
            AddComboStr cmbTyp.hwnd, sType(iType, 0)
         Next
         AddComboStr cmbTyp.hwnd, sType(iType, 0)
         cmbTyp = cmbTyp.List(0)
         cmbTyp.ToolTipText = sType(0, 1)
      End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sType(0, 0) = "1"
   sType(0, 1) = "Top Assembly Unit"
   sType(1, 0) = "2"
   sType(1, 1) = "Intermediate Assembly Unit"
   sType(2, 0) = "3"
   sType(2, 1) = "Base Manufacturing Unit"
   sType(3, 0) = "4"
   sType(3, 1) = "Raw Material Unit"
   sType(4, 0) = "5"
   sType(4, 1) = "Expense Item"
   sType(5, 0) = "6"
   sType(5, 1) = "Expense Item"
   sType(6, 0) = "7"
   sType(6, 1) = "Service Expense Item"
   sType(7, 0) = "8"
   sType(7, 1) = "Project"
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PadmPRf06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Function GetCode() As Byte
   Dim RdoCde As ADODB.Recordset
   On Error GoTo DiaErr1
   If Len(Trim(cmbCls)) > 0 Then
      sSql = "Qry_GetProductClass '" & Compress(cmbCls) & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
      If bSqlRows Then
         With RdoCde
            cmbCls = "" & Trim(!CCCODE)
            lblDsc = "" & Trim(!CCDESC)
         End With
         GetCode = 1
      Else
         lblDsc = "*** Product Class Wasn't Found ***"
         GetCode = 0
      End If
   End If
   Set RdoCde = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 7) = "*** Pro" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub


Private Sub UpdateParts()
   Dim bResponse As Byte
   Dim lRows As Long
   Dim sMsg As String
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   sPcode = Compress(cmbCls)
   sMsg = "You Are About Update All Part Types " & cmbTyp & vbCr _
          & "With Product Class " & cmbCls & ". Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdDel.Enabled = False
      
      On Error Resume Next
      sSql = "Update PartTable SET PACLASS='" _
             & sPcode & "' WHERE PALEVEL=" & Trim(cmbTyp) & " "
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      clsADOCon.ExecuteSQL sSql
      
      lRows = clsADOCon.RowsAffected
      
      MouseCursor 0
      If lRows > 0 Then
         sMsg = "You Are About To Update " & lRows & " Parts Numbers." & vbCr _
                & "Continue Updating To Product Class " & cmbCls & "."
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbYes Then
            If clsADOCon.ADOErrNum = 0 Then
               clsADOCon.CommitTrans
               MsgBox "Part Numbers Successfully Updated.", _
                  vbInformation, Caption
            Else
               clsADOCon.RollbackTrans
               MsgBox "Could Not Successfully Update Part Numbers.", _
                  vbExclamation, Caption
            End If
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            CancelTrans
         End If
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         MsgBox "The Are No Part Numbers Matching Type " & cmbTyp & ".", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   cmdDel.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "updateparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
