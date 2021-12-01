VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaJcbud 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Order Budgets"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select MO Part Number (Not CA or CL)"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select Run Number (Not CL or CA)"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtHrs 
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Tag             =   "1"
      Top             =   3360
      Width           =   1035
   End
   Begin VB.TextBox txtFoh 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Tag             =   "1"
      Top             =   3000
      Width           =   1035
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Tag             =   "1"
      Top             =   2640
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2280
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1920
      Width           =   1035
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   10
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
      PictureUp       =   "diaJcbud.frx":0000
      PictureDn       =   "diaJcbud.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4065
      FormDesignWidth =   5805
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblStu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Hours"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Overhead"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Expense"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Material"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Labor"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "diaJcbud"
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

Dim RdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim RdoBud As ADODB.Recordset
Dim bGoodMO As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   FindPart Me
   FillFormRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   FindPart Me
   FillFormRuns
   
End Sub


Private Sub cmbRun_Click()
   If Val(cmbRun) > 0 Then GetStatus Else _
          lblStu = ""
   
End Sub


Private Sub cmbRun_LostFocus()
   If Val(cmbRun) > 0 Then GetStatus Else _
          lblStu = ""
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS<>'CA' AND RUNSTATUS<>'CL')  "
   Set RdoQry = New ADODB.Command
   RdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   RdoQry.parameters.Append AdoParameter1
   
   bOnLoad = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoBud = Nothing
   Set AdoParameter1 = Nothing
   Set RdoQry = Nothing
   Set diaJcbud = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FillBudget()
   On Error GoTo DiaErr1
   bGoodMO = 0
   sSql = "SELECT RUNREF,RUNNO,RUNBUDLAB,RUNBUDMAT," _
          & "RUNBUDEXP,RUNBUDOH,RUNBUDHRS FROM " _
          & "RunsTable WHERE RUNREF='" & Compress(cmbPrt) & "' " _
          & "AND RUNNO=" & Val(cmbRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBud, ES_KEYSET)
   If bSqlRows Then
      With RdoBud
         txtLab = Format(!RUNBUDLAB, "#####0.000")
         txtMat = Format(!RUNBUDMAT, "#####0.000")
         txtExp = Format(!RUNBUDEXP, "#####0.000")
         txtFoh = Format(!RUNBUDOH, "#####0.000")
         txtHrs = Format(!RUNBUDHRS, "#####0.000")
         bGoodMO = 1
      End With
   Else
      bGoodMO = 0
      txtLab = ""
      txtMat = ""
      txtExp = ""
      txtFoh = ""
      txtHrs = ""
      MsgBox "No Active Manufacturing Order.", _
         vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillbudge"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub txtExp_LostFocus()
   txtExp = CheckLen(txtExp, 10)
   txtExp = Format(Abs(Val(txtExp)), "#####0.000")
   If bGoodMO = 1 Then
      On Error Resume Next
      RdoBud!RUNBUDEXP = Val(txtExp)
      RdoBud.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtFoh_LostFocus()
   txtFoh = CheckLen(txtFoh, 10)
   txtFoh = Format(Abs(Val(txtFoh)), "#####0.000")
   If bGoodMO = 1 Then
      On Error Resume Next

      RdoBud!RUNBUDOH = Val(txtFoh)
      RdoBud.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtHrs_LostFocus()
   txtHrs = CheckLen(txtHrs, 10)
   txtHrs = Format(Abs(Val(txtHrs)), "#####0.000")
   If bGoodMO = 1 Then
      On Error Resume Next

      RdoBud!RUNBUDHRS = Val(txtHrs)
      RdoBud.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 10)
   txtLab = Format(Abs(Val(txtLab)), "#####0.000")
   If bGoodMO = 1 Then
      On Error Resume Next

      RdoBud!RUNBUDLAB = Val(txtLab)
      RdoBud.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 10)
   txtMat = Format(Abs(Val(txtMat)), "#####0.000")
   If bGoodMO = 1 Then
      On Error Resume Next

      RdoBud!RUNBUDMAT = Val(txtMat)
      RdoBud.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub



Public Sub FillCombo()
   Dim RdoPcl As ADODB.Recordset
   Dim sTempPart As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF," _
          & "RUNSTATUS FROM PartTable,RunsTable WHERE " _
          & "RUNREF=PARTREF AND (RUNSTATUS<>'CA' AND RUNSTATUS<>'CL')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPcl)
   If bSqlRows Then
      With RdoPcl
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         Do Until .EOF
            If sTempPart <> Trim(!PARTNUM) Then
               'cmbPrt.AddItem "" & Trim(!PARTNUM)
               AddComboStr cmbPrt.hwnd, "" & Trim(!PARTNUM)
               sTempPart = Trim(!PARTNUM)
            End If
            .MoveNext
         Loop
      End With
      If cmbPrt.ListCount > 0 Then FillFormRuns
   Else
      MsgBox "No Matching Runs Recorded.", _
         vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoPcl = Nothing
   cmbPrt.SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub FillFormRuns()
   Dim RdoRns As ADODB.Recordset
   Dim SPartRef As String
   cmbRun.Clear
   SPartRef = Compress(cmbPrt)
   RdoQry.parameters(0).Value = SPartRef
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, RdoQry)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            'cmbRun.AddItem Format(!RUNNO, "####0")
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
      End With
   Else
   End If
   If cmbRun.ListCount > 0 Then
      cmbRun = Format(cmbRun.List(0), "####0")
      GetStatus
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillformr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetStatus()
   Dim RdoStu As ADODB.Recordset
   Dim SPartRef As String
   On Error GoTo DiaErr1
   SPartRef = Compress(cmbPrt)
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = '" & SPartRef & "' " _
          & "AND RUNNO=" & cmbRun & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStu, ES_FORWARD)
   If bSqlRows Then
      lblStu = "" & Trim(RdoStu!RUNSTATUS)
   Else
      lblStu = ""
   End If
   If Trim(lblStu) <> "CA" Or lblStu <> "CL" Then
      FillBudget
   Else
      bGoodMO = 0
      txtLab = ""
      txtMat = ""
      txtExp = ""
      txtFoh = ""
      txtHrs = ""
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getstatus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
