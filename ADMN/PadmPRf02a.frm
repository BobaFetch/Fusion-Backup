VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PadmPRf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Product Codes"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PadmPRf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4920
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Update Part Product Codes"
      Top             =   840
      Width           =   875
   End
   Begin VB.ComboBox cmbNew 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Product Code"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Product Code"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5280
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2805
      FormDesignWidth =   5925
   End
   Begin VB.Label lblNew 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Product Code"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Product Code"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "PadmPRf02a"
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
Dim bGoodOldCode As Byte
Dim bGoodNewCode As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetCode(bFirst As Byte) As Byte
   Dim RdoCde As ADODB.Recordset
   Dim sPcode As String
   
   If bFirst = 1 Then
      sPcode = Compress(cmbCde)
   Else
      sPcode = Compress(cmbNew)
   End If
   On Error GoTo DiaErr1
   If Len(sPcode) > 0 Then
      sSql = "Qry_GetProductCode '" & sPcode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
      If bSqlRows Then
         With RdoCde
            If bFirst = 1 Then
               cmbCde = "" & Trim(!PCCODE)
               lblDsc = "" & Trim(!PCDESC)
            Else
               cmbNew = "" & Trim(!PCCODE)
               lblNew = "" & Trim(!PCDESC)
            End If
         End With
         GetCode = 1
      Else
         If bFirst = 1 Then
            lblDsc = "*** Product Code Wasn't Found ***"
         Else
            lblNew = "*** Product Code Wasn't Found ***"
         End If
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

Private Sub cmbCde_Click()
   bGoodOldCode = GetCode(1)
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   bGoodOldCode = GetCode(1)
   
End Sub


Private Sub cmbNew_Click()
   bGoodNewCode = GetCode(2)
   
End Sub


Private Sub cmbNew_LostFocus()
   cmbNew = CheckLen(cmbNew, 6)
   bGoodNewCode = GetCode(2)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1351
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   If bGoodOldCode = 1 And bGoodNewCode = 1 Then
      If cmbCde = cmbNew Then
         MsgBox "The Existing And Replacement Codes Are The Same.", _
            vbExclamation, Caption
      Else
         ChangeCode
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   Dim iList As Integer
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes
      If cmbCde.ListCount > 0 Then
         For iList = 0 To cmbCde.ListCount - 1
            AddComboStr cmbNew.hwnd, cmbCde.List(iList)
         Next
         cmbCde = cmbCde.List(0)
         cmbNew = cmbCde.List(0)
         bGoodOldCode = GetCode(1)
         bGoodNewCode = GetCode(2)
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
   Set PadmPRf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub lblDsc_Change()
   If Left(lblDsc, 7) = "*** Pro" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblNew_Change()
   If Left(lblNew, 7) = "*** Pro" Then
      lblNew.ForeColor = ES_RED
   Else
      lblNew.ForeColor = vbBlack
   End If
   
End Sub


Private Sub ChangeCode()
   Dim bResponse As Byte
   Dim lRows As Long
   Dim sMsg As String
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   sPcode = Compress(cmbNew)
   sMsg = "You Are About To Replace Product Code " & cmbCde & vbCr _
          & "With Product Code " & cmbNew & ". Continue Updating?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdUpd.Enabled = False
      On Error Resume Next
      sSql = "UPDATE PartTable SET PAPRODCODE='" & sPcode _
             & "' WHERE (PAPRODCODE='" & Compress(cmbCde) & "' " _
             & "OR PAPRODCODE='" & cmbCde & "')"
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      clsADOCon.ExecuteSQL sSql
      lRows = clsADOCon.RowsAffected
      
      If lRows > 0 And clsADOCon.ADOErrNum = 0 Then
         MouseCursor 0
         sMsg = "There Are " & lRows & " Parts Using " & cmbCde & vbCr _
                & "Continiue Updating Parts Using " & cmbNew & "."
         bResponse = MsgBox(sMsg, ES_NOQUESTION + vbSystemModal, Caption)
         If bResponse = vbYes Then
            clsADOCon.CommitTrans
            MsgBox "Product Codes For Parts Where Updated.", _
               vbInformation, Caption
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            CancelTrans
         End If
      Else
         clsADOCon.RollbackTrans
         MouseCursor 0
         MsgBox "No Parts To Update The Product Codes.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   cmdUpd.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "changecode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
