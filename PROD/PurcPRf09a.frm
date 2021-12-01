VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Purchase Order Requested By"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf09a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "Replace All Purchase Order Rows With The New Requested By"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cmbReq 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Contains Previous Table Entries (20 Char Max) Including Blanks"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "New Requested By (20 Characters/3 Min)"
      Top             =   1320
      Width           =   2052
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2220
      FormDesignWidth =   5670
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Requested By"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Contains Previous Table Entries (20 Char Max) Including Blanks"
      Top             =   1320
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Requested By"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Contains Previous Table Entries (20 Char Max) Including Blanks"
      Top             =   720
      Width           =   1572
   End
End
Attribute VB_Name = "PurcPRf09a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'5/11/06 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4358
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpd_Click()
   If txtNew = cmbReq Then Exit Sub
   If Len(Trim(txtNew)) < 3 Then
      MsgBox "The System Replacement Setting Be At Least (3) Chars.", _
         vbInformation, Caption
   Else
      ReplaceCurrentRequestBy
      
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillReqBy
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub FillReqBy()
   On Error GoTo DiaErr1
   cmbReq.Clear
   sSql = "SELECT DISTINCT POREQBY FROM PohdTable ORDER BY POREQBY"
   LoadComboBox cmbReq, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillreqby"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
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
   Set PurcPRf09a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 20)
   On Error Resume Next
   If Len(txtNew) < 4 Then txtNew = UCase$(txtNew) _
          Else txtNew = StrCase(txtNew)
   
End Sub



Private Sub ReplaceCurrentRequestBy()
   Dim RdoReq As ADODB.Recordset
   Dim bResponse As Byte
   sSql = "SELECT COUNT(POREQBY) AS ReqBy FROM PohdTable WHERE " _
          & "POREQBY='" & cmbReq & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoReq, ES_FORWARD)
   If bSqlRows Then
      bResponse = MsgBox(Trim(RdoReq!ReqBy) & " PO Rows Will Be Affected. Continue?", _
                  ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         clsADOCon.BeginTrans
         clsADOCon.ADOErrNum = 0
         
         sSql = "UPDATE PohdTable SET POREQBY='" & txtNew & "' " _
                & "WHERE POREQBY='" & cmbReq & "'"
         clsADOCon.ExecuteSQL sSql
         bResponse = MsgBox("Last Chance, Continue Applying Changes?", _
                     ES_NOQUESTION, Caption)
         If bResponse = vbYes Then
            clsADOCon.CommitTrans
            SysMsg "Matching Rows Updated.", True
            FillReqBy
            cmbReq.SetFocus
         Else
            clsADOCon.RollbackTrans
            CancelTrans
         End If
      Else
         CancelTrans
      End If
   Else
      MsgBox "No Matching Rows Were Found.", _
         vbInformation, Caption
   End If
   Set RdoReq = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "replacecurr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
