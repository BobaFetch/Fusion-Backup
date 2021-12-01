VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PadmPRf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Product Classes"
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
      Picture         =   "PadmPRf05a.frx":0000
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
      ToolTipText     =   "Update Part Product Classes"
      Top             =   720
      Width           =   875
   End
   Begin VB.ComboBox cmbNew 
      Height          =   288
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Product Class"
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox cmbCls 
      Height          =   288
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Product Class"
      Top             =   840
      Width           =   855
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
      Height          =   288
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   3132
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Product Class"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Product Class"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1932
   End
End
Attribute VB_Name = "PadmPRf05a"
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
Dim bGoodOldClass As Byte
Dim bGoodNewClass As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetCode(bFirst As Byte) As Byte
   Dim RdoCde As ADODB.Recordset
   
   Dim sPClass As String
   
   If bFirst = 1 Then
      sPClass = Compress(cmbCls)
   Else
      sPClass = Compress(cmbNew)
   End If
   On Error GoTo DiaErr1
   If Len(sPClass) > 0 Then
      sSql = "Qry_GetProductClass '" & sPClass & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
      If bSqlRows Then
         With RdoCde
            If bFirst = 1 Then
               cmbCls = "" & Trim(!CCCODE)
               lblDsc = "" & Trim(!CCDESC)
            Else
               cmbNew = "" & Trim(!CCCODE)
               lblNew = "" & Trim(!CCDESC)
            End If
         End With
         GetCode = 1
      Else
         If bFirst = 1 Then
            lblDsc = "*** Product Class Wasn't Found ***"
         Else
            lblNew = "*** Product Class Wasn't Found ***"
         End If
         GetCode = False
      End If
   End If
   Set RdoCde = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getClass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function





Private Sub cmbCls_Click()
   bGoodOldClass = GetCode(1)
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   bGoodOldClass = GetCode(1)
   
End Sub


Private Sub cmbNew_Click()
   bGoodNewClass = GetCode(2)
   
End Sub


Private Sub cmbNew_LostFocus()
   cmbNew = CheckLen(cmbNew, 4)
   bGoodNewClass = GetCode(2)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1354
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   If bGoodOldClass = 1 And bGoodNewClass = 1 Then
      If cmbCls = cmbNew Then
         MsgBox "The Existing And Replacement Classes Are The Same.", _
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
      FillProductClasses
      If cmbCls.ListCount > 0 Then
         For iList = 0 To cmbCls.ListCount - 1
            AddComboStr cmbNew.hwnd, cmbCls.List(iList)
         Next
         cmbCls = cmbCls.List(0)
         cmbNew = cmbCls.List(0)
         bGoodOldClass = GetCode(1)
         bGoodNewClass = GetCode(2)
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
   Dim sPClass As String
   
   On Error GoTo DiaErr1
   sPClass = Compress(cmbNew)
   sMsg = "You Are About To Replace Product Class " & cmbCls & vbCr _
          & "With Product Class " & cmbNew & ". Continue Updating?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdUpd.Enabled = False
      On Error Resume Next
      sSql = "UPDATE PartTable SET PACLASS='" & sPClass _
             & "' WHERE (PACLASS='" & Compress(cmbCls) & "' " _
             & "OR PACLASS='" & cmbCls & "')"
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      clsADOCon.ExecuteSQL sSql
      lRows = clsADOCon.RowsAffected
      
      If lRows > 0 And clsADOCon.ADOErrNum = 0 Then
         MouseCursor 0
         sMsg = "There Are " & lRows & " Parts Using " & cmbCls & vbCr _
                & "Continiue Updating Parts Using " & cmbNew & "."
         bResponse = MsgBox(sMsg, ES_NOQUESTION + vbSystemModal, Caption)
         If bResponse = vbYes Then
            clsADOCon.CommitTrans
            MsgBox "Product Classes For Parts Where Updated.", _
               vbInformation, Caption
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            CancelTrans
         End If
      Else
         clsADOCon.RollbackTrans
         MouseCursor 0
         MsgBox "No Parts To Update The Product Classes.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   cmdUpd.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "changeClass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
