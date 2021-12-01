VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Discrepancy Code"
   ClientHeight    =   2385
   ClientLeft      =   2385
   ClientTop       =   1680
   ClientWidth     =   5865
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "InspRTf05a.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2385
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTf05a.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Delete This Discrepancy Code"
      Top             =   600
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5520
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2385
      FormDesignWidth =   5865
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Select Characteristic Code From List"
      Top             =   810
      Width           =   1675
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label txtCmt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code Id"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "InspRTf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim RdoCde As ADODB.Recordset
Dim bOnLoad As Byte
Dim bGoodCode As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Click()
   bGoodCode = GetCode()
   
End Sub

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 12)
   cmbCde = Trim(cmbCde)
   If Len(cmbCde) = 0 Then
      txtDsc = ""
      txtCmt = ""
      bGoodCode = 0
      Exit Sub
   Else
      bGoodCode = GetCode()
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   If bGoodCode = 0 Then
      MsgBox "Requires A Valid Discrepancy Code.", _
         vbExclamation, Caption
   Else
      DeleteCode
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6154
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombo
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoCde = Nothing
   Set InspRTf05a = Nothing
   
End Sub


Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillDescripancyCodes"
   LoadComboBox cmbCde
   If cmbCde.ListCount > 0 Then cmbCde = cmbCde.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function GetCode() As Byte
   Dim sRejCode As String
   sRejCode = Compress(cmbCde)
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "SELECT CDEREF,CDENUM,CDEDESC,CDENOTES " _
          & "FROM RjcdTable WHERE CDEREF='" & sRejCode & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
   If bSqlRows Then
      With RdoCde
         cmbCde = "" & Trim(!CDENUM)
         txtDsc = "" & Trim(!CDEDESC)
         txtCmt = "" & Trim(!CDENOTES)
      End With
      GetCode = 1
   Else
      GetCode = 0
      txtDsc = ""
      txtCmt = ""
      'RdoCde.Close
   End If
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub DeleteCode()
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sMsg = "You Are About To Pemanently Remove This" & vbCr _
          & "Discrepancy Code. Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      sSql = "SELECT DISTINCT RITCHARCODE FROM RjitTable " _
             & "WHERE RITCHARCODE='" & Compress(cmbCde) & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected = 0 Then
         sMsg = "No Inspection Reports Affected." & vbCr _
                & "Still Wish To Continue?"
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbYes Then
            On Error Resume Next
            clsADOCon.ADOErrNum = 0
            sSql = "DELETE FROM RjcdTable WHERE " _
                   & "CDEREF='" & Compress(cmbCde) & "'"
            clsADOCon.ExecuteSQL sSql
            If clsADOCon.ADOErrNum = 0 Then
               MsgBox "The Code Was Successfully Deleted.", _
                  vbInformation, Caption
               txtDsc = ""
               txtCmt = ""
               cmbCde.Clear
               FillCombo
            Else
               MsgBox "Couldn't Delete The Code.", _
                  vbExclamation, Caption
            End If
         Else
            CancelTrans
         End If
      Else
         MsgBox cmbCde & " Is Used On One Or More Inspection Reports." & vbCr _
            & "Cannot Delete This Discrepancy Code.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "deletecode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
