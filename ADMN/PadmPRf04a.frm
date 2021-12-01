VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PadmPRf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Product Class"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PadmPRf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCls 
      Height          =   288
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise Product Class (4 Char)"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "D&elete"
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Delete This Class"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   2
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
      FormDesignWidth =   5595
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1692
   End
End
Attribute VB_Name = "PadmPRf04a"
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
'8/28/06 Repaired Delete Process


Dim bOnLoad As Byte
Dim bGoodCode As Byte

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


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   If bGoodCode Then DeleteCode
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1353
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductClasses
      If cmbCls.ListCount > 0 Then
         cmbCls = cmbCls.List(0)
         bGoodCode = GetCode()
      End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PadmPRf04a = Nothing
   
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


Private Sub DeleteCode()
   Dim RdoCount As ADODB.Recordset
   Dim bResponse As Byte
   Dim lRows As Long
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sMsg = "You Are About To Permanently Delete  " & vbCr _
          & "Product Class " & cmbCls & ". Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdDel.Enabled = False
      sSql = "SELECT COUNT(PACLASS) AS RowsCounted FROM PartTable " _
             & "WHERE PACLASS='" & Compress(cmbCls) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCount, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoCount!RowsCounted) Then
            lRows = RdoCount!RowsCounted
         Else
            lRows = 0
         End If
      Else
         lRows = 0
      End If
      
      MouseCursor 0
      If lRows > 0 Then
         MsgBox "There Are (" & lRows & ") Parts Using That Class" & vbCr _
            & "Cannot Delete Product Class " & cmbCls & ".", _
            vbInformation, Caption
      Else
         On Error Resume Next
         sSql = "DELETE FROM PclsTable WHERE CCREF='" & cmbCls & "'"
         clsADOCon.ExecuteSQL sSql
         If Err > 0 Then
            MsgBox "Couldn't Delete The Product Class.", _
               vbExclamation, Caption
         Else
            MsgBox "The Product Class Was Successfully Deleted.", _
               vbInformation, Caption
            cmbCls.Clear
            FillProductClasses
            If cmbCls.ListCount > 0 Then cmbCls = cmbCls.List(0)
            bGoodCode = GetCode()
         End If
      End If
   Else
      CancelTrans
   End If
   cmdDel.Enabled = True
   Set RdoCount = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "deletecode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
