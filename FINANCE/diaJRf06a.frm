VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaJRf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GL Journals - SummaryAccount"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2730
      FormDesignWidth =   6255
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "4"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Copy Journals To Summary Account"
      Height          =   555
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Open Journal"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaJRf06a.frx":0000
      PictureDn       =   "diaJRf06a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal Start Date"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal End Date"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy GL Journals to Summary Acocunt"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "diaJRf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
' diaJRf06a - Open Journals
'
' Created (cjs)
' Revision:
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String

Public bRemote As Byte
Public bIndex As Byte ' journal type to display

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************


Private Sub cmdAdd_Click()
   Dim dBeg As Date
   Dim dEnd As Date
   Dim bResponse As Boolean
   
   On Error Resume Next
   If ((Trim(txtBeg) = "") Or (Trim(txtEnd) = "")) Then
      MsgBox "Please Select Journal Dates.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   dBeg = Format(txtBeg, "mm/dd/yy")
   dEnd = Format(txtEnd, "mm/dd/yy")
   
   If dBeg > dEnd Then
      MsgBox "There Is A Date Mismatch in Date.", _
         vbInformation, Caption
      txtBeg.SetFocus
      Exit Sub
   End If
   
   sMsg = "Do you want to copy Journal to summary Account?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      Exit Sub
   End If
   
   CopyJournals
   
End Sub

Private Sub CopyJournals()
   Dim rdoJrn As ADODB.Recordset
   Dim bResponse As Byte
   Dim strBegDt As String
   Dim strEndDt As String
   
   
   On Error GoTo DiaErr1
   
   strBegDt = Format(txtBeg, "mm/dd/yy")
   strEndDt = Format(txtEnd, "mm/dd/yy")
   
   clsADOCon.ADOErrNum = 0
   
   sSql = "GLJritTopSummary '" & strBegDt & "','" & strEndDt & "'"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   If (clsADOCon.ADOErrNum = 0) Then
      sMsg = "Successfully Copied Journal entries To Summary Account."
   Else
      sMsg = "Couldn't Copy Journal Entries to Summary Account."
   End If
   
   bResponse = MsgBox(sMsg, vbInformation, Caption)
   
   Exit Sub
   
DiaErr1:
   sProcName = "CopyJournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   txtBeg = Format(Now, "mm/01/yy")
   txtEnd = GetMonthEnd(txtBeg)
   
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaJRf06a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

