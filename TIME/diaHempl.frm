VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaHempl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employees"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Tag             =   "3"
      Top             =   5040
      Width           =   3855
   End
   Begin VB.CheckBox chkEngineer 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   6360
      TabIndex        =   24
      ToolTipText     =   "Is this employee an engineer?"
      Top             =   4980
      Width           =   375
   End
   Begin VB.TextBox txtComments 
      Height          =   1455
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   5880
      Width           =   5655
   End
   Begin VB.ComboBox txtPrevTerm 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox txtSta 
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      ToolTipText     =   "Add Any (1) Character Code -Includes Codes In Use"
      Top             =   3960
      Width           =   615
   End
   Begin VB.ComboBox cmbAct 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   25
      Tag             =   "3"
      Top             =   5400
      Width           =   1935
   End
   Begin VB.ComboBox txtRev 
      Height          =   315
      Left            =   5760
      TabIndex        =   15
      Tag             =   "4"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox txtRse 
      Height          =   315
      Left            =   3480
      TabIndex        =   14
      Tag             =   "4"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox txtBdy 
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Tag             =   "4"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   5760
      TabIndex        =   12
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ComboBox txtReh 
      Height          =   315
      Left            =   3480
      TabIndex        =   11
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtPhn 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   2880
      Width           =   1550
      _ExtentX        =   2725
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   12
      Mask            =   "###-###-####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   5760
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin MSMask.MaskEdBox txtSsn 
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Top             =   2880
      Width           =   1550
      _ExtentX        =   2725
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   11
      Mask            =   "###-##-####"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox optHrs 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      ToolTipText     =   "Hourly (checked) Or Salary (Unchecked)"
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txtPay 
      Height          =   285
      Left            =   4200
      TabIndex        =   21
      Tag             =   "1"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox txtDpt 
      Height          =   285
      Left            =   1200
      TabIndex        =   20
      Tag             =   "3"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.ComboBox cmbWcn 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4200
      TabIndex        =   19
      Tag             =   "8"
      ToolTipText     =   "Select Work Center From List"
      Top             =   4320
      Width           =   1775
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      TabIndex        =   18
      Tag             =   "8"
      ToolTipText     =   "Select Shop From List"
      Top             =   4320
      Width           =   1775
   End
   Begin VB.TextBox txtMar 
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Tag             =   "3"
      ToolTipText     =   "M (Married) Or S (Single)"
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtMid 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Tag             =   "3"
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtSte 
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Tag             =   "3"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtCty 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Tag             =   "2"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtAdr 
      Height          =   735
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "9"
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtLst 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox cmbEmp 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtFst 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Tag             =   "2"
      Top             =   960
      Width           =   2085
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   27
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
      PictureUp       =   "diaHempl.frx":0000
      PictureDn       =   "diaHempl.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7815
      FormDesignWidth =   6960
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   255
      Index           =   26
      Left            =   240
      TabIndex        =   59
      Top             =   5040
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Engineer?"
      Height          =   255
      Index           =   25
      Left            =   5400
      TabIndex        =   58
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   255
      Index           =   24
      Left            =   240
      TabIndex        =   57
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Prev Term"
      Height          =   255
      Left            =   4680
      TabIndex        =   55
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblActdsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   "
      Height          =   285
      Left            =   3240
      TabIndex        =   54
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label lblEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hourly"
      Height          =   255
      Index           =   23
      Left            =   5400
      TabIndex        =   52
      Top             =   4740
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Rate"
      Height          =   255
      Index           =   22
      Left            =   3120
      TabIndex        =   43
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   255
      Index           =   21
      Left            =   240
      TabIndex        =   51
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   50
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   255
      Index           =   19
      Left            =   3120
      TabIndex        =   48
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   255
      Index           =   18
      Left            =   240
      TabIndex        =   49
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status"
      Height          =   255
      Index           =   17
      Left            =   2400
      TabIndex        =   44
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   45
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Review Date"
      Height          =   255
      Index           =   15
      Left            =   4680
      TabIndex        =   42
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Raise"
      Height          =   255
      Index           =   14
      Left            =   2400
      TabIndex        =   46
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   47
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Terminated"
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   41
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rehire "
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   40
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hire Date"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   39
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SSN"
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   38
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   37
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   35
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   34
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   33
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   32
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   31
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   30
      Top             =   960
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   29
      Top             =   960
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   36
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "diaHempl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'1/13/05 Corrected Work Center fill
'Dim RdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim RdoEmp As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodEmployee As Byte
Dim sLastTermDate As String
Dim strOldRseDt As String
Dim strOldPay As String


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub chkEngineer_Click()
   If bGoodEmployee Then
      On Error Resume Next
      RdoEmp!PREMENGINEER = chkEngineer.Value
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If

End Sub

Private Sub cmbAct_Click()
   FindAccount Me
   
End Sub


Private Sub cmbAct_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   Dim sAccount As String
   cmbAct = CheckLen(cmbAct, 12)
   For iList = 0 To cmbAct.ListCount - 1
      If cmbAct = cmbAct.List(iList) Then b = True
   Next
   On Error Resume Next
   If b = 0 Then
      'Beep
      cmbAct = "" & Trim(RdoEmp!PREMACCTS)
   End If
   FindAccount Me
   sAccount = Compress(cmbAct)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMACCTS = "" & Compress(cmbAct)
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub cmbEmp_Click()
   bGoodEmployee = GetEmployee()
   
End Sub

Private Sub cmbEmp_LostFocus()
   cmbEmp = CheckLen(cmbEmp, 6)
   If Len(cmbEmp) Then
      cmbEmp = Format(cmbEmp, "000000")
      bGoodEmployee = GetEmployee()
      If Not bGoodEmployee Then AddEmployee
   Else
      ClearBoxes
   End If
   
End Sub



Private Sub cmbShp_Click()
   FillCenters 1
   
End Sub

Private Sub cmbShp_LostFocus()
   If Len(cmbShp) = 0 Then
      If cmbShp.ListCount > 0 Then cmbShp = cmbShp.List(0)
   End If
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMSHOP = "" & cmbShp
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub cmbWcn_LostFocus()
   If Len(cmbWcn) = 0 Then
      If cmbWcn.ListCount > 0 Then cmbWcn = cmbWcn.List(0)
   End If
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMCENTER = "" & cmbWcn
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   cmbEmp = ""
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs1503"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillAccounts
      FillCenters
      FillEmployees
      
      ' 10/14/2009 - Added to populate the Employeee information
      GetEmployee
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim cNewSize As Currency
   FormLoad Me
   FormatControls
   
   sSql = "SELECT * FROM EmplTable WHERE PREMNUMBER = ? "
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adInteger
   cmdObj.Parameters.Append prmObj
   
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   Set cmdObj = Nothing
   Set RdoEmp = Nothing
   Set diaHempl = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub




Private Sub lblActdsc_Change()
   If Trim(lblActdsc) = "*** Account Wasn't Found ***" Then
      lblActdsc.ForeColor = ES_RED
   Else
      lblActdsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optHrs_Click()
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      If optHrs.Value = vbChecked Then
         RdoEmp!PREMHOURLY = "H"
      Else
         RdoEmp!PREMHOURLY = "S"
      End If
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Sub optHrs_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub




Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 120)
   txtAdr = StrCase(txtAdr)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMADDR = "" & txtAdr
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtBdy_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBdy_LostFocus()
   If Len(Trim(txtBdy)) > 0 Then txtBdy = CheckDate(txtBdy)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      If Len(txtBdy) Then
         RdoEmp!PREMBIRTHDT = Format(txtBdy, "mm/dd/yy")
      Else
         RdoEmp!PREMBIRTHDT = Null
      End If
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) > 0 Then txtBeg = CheckDate(txtBeg)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      If Len(txtBeg) Then
         RdoEmp!PREMHIREDT = Format(txtBeg, "mm/dd/yy")
      Else
         RdoEmp!PREMHIREDT = Null
      End If
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub



Private Sub txtComments_LostFocus()
   txtComments = CheckLen(txtComments, 3072)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMCOMMENT = "" & txtComments
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtCty_LostFocus()
   txtCty = CheckLen(txtCty, 20)
   txtCty = StrCase(txtCty)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMCITY = "" & txtCty
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtDpt_LostFocus()
   txtDpt = CheckLen(txtDpt, 12)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMDEPT = "" & txtDpt
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Sub txtEmail_LostFocus()
   txtEmail = CheckLen(txtEmail, 60)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMEMAIL = "" & txtEmail
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub




Private Sub txtEnd_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
    Dim sTempDte As String
    
   If Len(Trim(txtEnd)) > 0 Then txtEnd = CheckDate(txtEnd)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      sTempDte = "" & RdoEmp!PREMPREVTERMDT
      If Len(txtEnd) Then
           If (txtEnd <> sTempDte) Then
                RdoEmp!PREMPREVTERMDT = sLastTermDate
                txtPrevTerm = sLastTermDate
            End If
         RdoEmp!PREMTERMDT = Format(txtEnd, "mm/dd/yy")
      Else
         RdoEmp!PREMTERMDT = Null
      End If
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtFst_LostFocus()
   txtFst = CheckLen(txtFst, 20)
   txtFst = StrCase(txtFst)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMFSTNAME = "" & txtFst
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtLst_LostFocus()
   txtLst = CheckLen(txtLst, 20)
   txtLst = StrCase(txtLst)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMLSTNAME = "" & txtLst
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtMar_LostFocus()
   txtMar = CheckLen(txtMar, 1)
   If txtMar <> "M" Then txtMar = "S"
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMMARSTAT = "" & txtMar
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtMid_LostFocus()
   txtMid = CheckLen(txtMid, 1)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMMINIT = "" & txtMid
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtPay_GotFocus()
   strOldPay = Format(txtPay, ES_QuantityDataFormat)
End Sub

Private Sub txtPay_LostFocus()
   txtPay = CheckLen(txtPay, 9)
   txtPay = Format(txtPay, ES_QuantityDataFormat)
   If Val(txtPay) > 200 Then optHrs.Value = vbChecked
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMPAYRATE = Val(txtPay)
      If optHrs.Value = vbChecked Then
         RdoEmp!PREMHOURLY = "H"
      Else
         RdoEmp!PREMHOURLY = "S"
      End If
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtPhn_LostFocus()
   txtPhn = CheckLen(txtPhn, 12)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMPHONE = "" & txtPhn
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtPrevTerm_DropDown()
    ShowCalendar Me
End Sub

Private Sub txtPrevTerm_LostFocus()
   If Len(Trim(txtPrevTerm)) > 0 Then txtPrevTerm = CheckDate(txtPrevTerm)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      If Len(txtPrevTerm) Then
         RdoEmp!PREMPREVTERMDT = Format(txtPrevTerm, "mm/dd/yy")
      Else
         RdoEmp!PREMPREVTERMDT = Null
      End If
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If

End Sub

Private Sub txtReh_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtReh_LostFocus()
   If Len(Trim(txtReh)) > 0 Then txtReh = CheckDate(txtReh)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      If Len(txtReh) Then
         RdoEmp!PREMREHIREDT = Format(txtReh, "mm/dd/yy")
      Else
         RdoEmp!PREMREHIREDT = Null
      End If
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtRev_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtRev_LostFocus()
   If Len(Trim(txtRev)) > 0 Then txtRev = CheckDate(txtRev)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      If Len(txtRev) Then
         RdoEmp!PREMREVUDT = Format(txtRev, "mm/dd/yy")
      Else
         RdoEmp!PREMREVUDT = Null
      End If
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtRse_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtRse_LostFocus()
   If Len(Trim(txtRse)) > 0 Then txtRse = CheckDate(txtRse)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      If Len(txtRse) Then
         RdoEmp!PREMLSTRAISE = Format(txtRse, "mm/dd/yy")
      Else
         RdoEmp!PREMLSTRAISE = Null
      End If
      RdoEmp.Update
      
      ' If the Date is difeerent and less then
      ' the current date update the rate for the existing records.
      If ((strOldRseDt <> txtRse) And (txtRse <> "")) Then
         
         Dim strCurdate As String
         Dim strCurPay As String
         Dim bResponse As Byte
         
         strCurdate = Format(GetServerDateTime, "mm/dd/yy")
         strCurPay = Format(txtPay, ES_QuantityDataFormat)
         
         If (CDate(txtRse) < CDate(strCurdate)) Then
            If (Trim(strCurPay) <> Trim(strOldPay)) Then
               
               bResponse = MsgBox("Do you want to change the Rate Pay to " & vbCr & _
                              Trim(strCurPay) & " as of " & _
                              Trim(txtRse) & ".", ES_YESQUESTION, Caption)
               If bResponse = vbYes Then
                  UpdateEmpRate cmbEmp, Trim(txtRse), strCurPay
                  strOldPay = strCurPay
               End If
            Else
               bResponse = MsgBox("Rate pay has not changed." & vbCr & _
                              "Do you want to change the RaiseDate?.", _
                                 ES_YESQUESTION, Caption)
               If bResponse = vbNo Then
                  txtRse = "" & Format(strOldRseDt, "mm/dd/yy")
               End If
               
            End If
         End If
         ' Reset the value
         strOldRseDt = Format(txtRse, "mm/dd/yy")
      End If
      
      
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Function UpdateEmpRate(iEmpNo As String, strRseDate As String, strRate As String)
   
   If ((strRseDate <> "") And iEmpNo <> 0) Then
      sSql = "UPDATE tcitTable SET TCRATE=" & CDbl(strRate) & "" _
         & " WHERE TCEMP = " & Val(iEmpNo) & " AND " _
         & " TCSTARTTIME >= '" & strRseDate & "'"
      clsADOCon.ExecuteSql sSql '  rdExecDirect
   End If

End Function

Private Sub txtSsn_LostFocus()
   txtSsn = CheckLen(txtSsn, 11)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMSOCSEC = "" & txtSsn
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtSta_LostFocus()
   txtSta = CheckLen(txtSta, 1)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMSTATUS = "" & txtSta
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtSte_LostFocus()
   txtSte = CheckLen(txtSte, 2)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMSTATE = "" & txtSte
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtZip_LostFocus()
   txtZip = CheckLen(txtZip, 10)
   If bGoodEmployee Then
      On Error Resume Next
      'RdoEmp.Edit
      RdoEmp!PREMZIPCD = "" & txtZip
      RdoEmp.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub



Private Sub FillCenters(Optional bSKipShops As Byte)
   On Error GoTo DiaErr1
   If bSKipShops = 0 Then
      sSql = "Qry_FillShops"
      cmbShp = "NONE"
      AddComboStr cmbShp.hwnd, "NONE"
      LoadComboBox cmbShp
   End If
   cmbWcn.Clear
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   cmbWcn = "NONE"
   AddComboStr cmbWcn.hwnd, "NONE"
   LoadComboBox cmbWcn
   Exit Sub
   
DiaErr1:
   sProcName = "fillcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillEmployees()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillEmployees"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         cmbEmp = Format(!PREMNUMBER, "000000")
         Do Until .EOF
            AddComboStr cmbEmp.hwnd, Format$(!PREMNUMBER, "000000")
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      cmbEmp = "000001"
   End If
   sSql = "SELECT DISTINCT PREMSTATUS FROM EmplTable ORDER BY PREMSTATUS"
   LoadComboBox txtSta, -1
   Set RdoCmb = Nothing
   If cmbEmp.ListCount > 0 Then bGoodEmployee = GetEmployee()
   Exit Sub
   
DiaErr1:
   sProcName = "fillemplo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetEmployee() As Byte
   On Error GoTo DiaErr1
   cmdObj.Parameters(0).Value = Val(cmbEmp)
   
   bSqlRows = clsADOCon.GetQuerySet(RdoEmp, cmdObj, ES_KEYSET, True)
   If bSqlRows Then
      With RdoEmp
         On Error Resume Next
         cmbEmp = Format(!PREMNUMBER, "000000")
         txtLst = "" & Trim(!PREMLSTNAME)
         txtFst = "" & Trim(!PREMFSTNAME)
         txtMid = "" & Trim(!PREMMINIT)
         txtAdr = "" & Trim(!PREMADDR)
         txtCty = "" & Trim(!PREMCITY)
         txtSte = "" & Trim(!PREMSTATE)
         txtZip = "" & Trim(!PREMZIPCD)
         If Len(Trim(!PREMSOCSEC)) Then txtSsn = "" & Trim(!PREMSOCSEC)
         If Len(Trim(!PREMPHONE)) Then txtPhn = "" & Trim(!PREMPHONE)
         txtBeg = "" & Format(!PREMHIREDT, "mm/dd/yy")
         txtReh = "" & Format(!PREMREHIREDT, "mm/dd/yy")
         txtEnd = "" & Format(!PREMTERMDT, "mm/dd/yy")
         txtBdy = "" & Format(!PREMBIRTHDT, "mm/dd/yy")
         txtRse = "" & Format(!PREMLSTRAISE, "mm/dd/yy")
         txtRev = "" & Format(!PREMREVUDT, "mm/dd/yy")
         txtMar = "" & Trim(!PREMMARSTAT)
         txtSta = "" & Trim(!PREMSTATUS)
         cmbShp = "" & Trim(!PREMSHOP)
         cmbWcn = "" & Trim(!PREMCENTER)
         txtDpt = "" & Trim(!PREMDEPT)
         txtEmail = "" & Trim(!PREMEMAIL)
         cmbAct = "" & Trim(!PREMACCTS)
         txtPrevTerm = "" & Format(Trim(!PREMPREVTERMDT), "mm/dd/yy")
         sLastTermDate = "" & Format(!PREMTERMDT, "mm/dd/yy")
         strOldRseDt = "" & Format(!PREMLSTRAISE, "mm/dd/yy")
         txtComments = "" & Trim(!PREMCOMMENT)
        
         If cmbAct = "" Then lblActdsc = ""
         If cmbAct.ListCount > 0 Then
            cmbAct.Enabled = True
            FindAccount Me
         End If
         txtPay = Format(0 + !PREMPAYRATE, ES_QuantityDataFormat)
         strOldPay = txtPay
         If Not IsNull(!PREMHOURLY) Then
            If !PREMHOURLY = "H" Then
               optHrs.Value = vbChecked
            Else
               optHrs.Value = vbUnchecked
            End If
         Else
            optHrs.Value = vbChecked
         End If
         chkEngineer = IIf(!PREMENGINEER, vbChecked, vbUnchecked)
      End With
      
      'lblEdit = "Editing " & cmbEmp
      ' 10/14/2009 - No need to fill the Center as the Center gets filled onLoad()
      ' FillCenters 1
      GetEmployee = True
   Else
      GetEmployee = False
      lblEdit = "No Current Employee"
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ClearBoxes()
   Dim iList As Integer
   On Error Resume Next
   RdoEmp.Close
   For iList = 0 To Controls.count - 1
      If TypeOf Controls(iList) Is TextBox Then
         Controls(iList).Text = ""
      Else
         If TypeOf Controls(iList) Is MaskEdBox Then
            Controls(iList).Mask = ""
            Controls(iList).Text = ""
         End If
      End If
   Next
   txtSsn.Mask = "###-##-####"
   txtPhn.Mask = "###-###-####"
   
End Sub

Private Sub AddEmployee()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim vDate As Variant
   
   ClearBoxes
   vDate = Format(ES_SYSDATE, "mm/dd/yy")
   On Error GoTo DiaErr1
   If Val(cmbEmp) = 0 Then
      MsgBox "Employee Number Must Be Greater Than Zero.", vbExclamation, Caption
      On Error Resume Next
      cmbEmp.SetFocus
   End If
   sMsg = "Employee " & cmbEmp & " Wasn't Found." & vbCrLf _
          & "Add The New Employee?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      sSql = "INSERT INTO EmplTable (PREMNUMBER,PREMSHOP,PREMCENTER," _
             & "PREMHIREDT,PREMREHIREDT) VALUES(" _
             & Val(cmbEmp) & ",'" _
             & cmbShp & "','" _
             & cmbWcn & "','" _
             & vDate & "'," _
             & "Null)"
      clsADOCon.ExecuteSql sSql '  rdExecDirect
      If Err = 0 Then
         SysMsg "Employee Was Added.", True
         AddComboStr cmbEmp.hwnd, cmbEmp
         bGoodEmployee = GetEmployee()
         On Error Resume Next
         txtFst.SetFocus
      Else
         MsgBox "Couldn't The Add Employee.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillAccounts()
   On Error GoTo DiaErr1
   sSql = "SELECT GLACCTREF,GLACCTNO FROM GlacTable ORDER BY GLACCTREF"
   LoadComboBox cmbAct
   If cmbAct.ListCount > 0 Then
      cmbAct = cmbAct.List(0)
      cmbAct.Enabled = True
      FindAccount Me
   Else
      cmbAct = "No Accounts."
   End If
   Exit Sub
   
DiaErr1:
   On Error GoTo 0
   
End Sub


