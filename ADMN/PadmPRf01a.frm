VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PadmPRf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Product Code"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
      Picture         =   "PadmPRf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Product Code"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "D&elete"
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Delete This Product Code"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
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
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "PadmPRf01a"
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
'8/28/06 Fixed Delete Function
Dim bOnLoad As Byte
Dim bGoodCode As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Click()
   bGoodCode = GetCode()
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If Len(cmbCde) Then bGoodCode = GetCode()
   
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
      OpenHelpContext 1350
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   Dim iList As Integer
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes
      If cmbCde.ListCount > 0 Then
         For iList = 0 To cmbCde.ListCount - 1
            If cmbCde.List(iList) = "BID" Then cmbCde.RemoveItem iList
         Next
         cmbCde = cmbCde.List(0)
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
   Set PadmPRf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Function GetCode() As Byte
   Dim RdoCde As ADODB.Recordset
   Dim sPcode As String
   sPcode = Compress(cmbCde)
   On Error GoTo DiaErr1
   If Len(sPcode) > 0 Then
      sSql = "Qry_GetProductCode '" & sPcode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
      If bSqlRows Then
         With RdoCde
            cmbCde = "" & Trim(!PCCODE)
            lblDsc = "" & Trim(!PCDESC)
         End With
         GetCode = True
      Else
         lblDsc = "*** Product Code Wasn't Found ***"
         GetCode = False
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
   
   If cmbCde = "BID" Or cmbCde = "TOOL" Then
      MsgBox Trim(cmbCde) & " Is Reserved And Cannot Be Removed.", _
                  vbInformation, Caption
      Exit Sub
   End If
   ' On Error GoTo DiaErr1
   On Error GoTo 0
   sMsg = "You Are About To Permanently Delete  " & vbCr _
          & "Product Code " & cmbCde & ". Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdDel.Enabled = False
      sSql = "SELECT COUNT(PAPRODCODE) AS RowsCounted FROM PartTable " _
             & "WHERE PAPRODCODE='" & Compress(cmbCde) & "'"
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
         MsgBox "There Are (" & lRows & ") Parts Using That Code" & vbCr _
            & "Cannot Delete Product Code " & cmbCde & ".", _
            vbInformation, Caption
      Else
         On Error Resume Next
         sSql = "DELETE FROM PcodTable WHERE PCREF='" & Compress(cmbCde) & "'"
         clsADOCon.ExecuteSQL sSql
         If Err > 0 Then
            MsgBox "Couldn't Delete The Product Code.", _
               vbExclamation, Caption
         Else
            MsgBox "The Product Code Was Successfully Deleted.", _
               vbInformation, Caption
            cmbCde.Clear
            FillProductCodes
            If cmbCde.ListCount > 0 Then cmbCde = cmbCde.List(0)
            bGoodCode = GetCode()
         End If
      End If
   Else
      CancelTrans
   End If
   Set RdoCount = Nothing
   cmdDel.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "deletecode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
