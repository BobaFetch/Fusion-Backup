VERSION 5.00
Begin VB.Form AdmnUperm9 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Permissions"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optVw5 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3780
      TabIndex        =   32
      Top             =   2160
      Width           =   375
   End
   Begin VB.CheckBox optFn5 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4740
      TabIndex        =   31
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox optVw4 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3780
      TabIndex        =   30
      Top             =   1800
      Width           =   375
   End
   Begin VB.CheckBox optFn4 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4740
      TabIndex        =   29
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox optEd5 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2820
      TabIndex        =   28
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox optEd4 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2820
      TabIndex        =   27
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox optGr5 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   26
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox optGr4 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   25
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox optEd3 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2820
      TabIndex        =   24
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optVw3 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3780
      TabIndex        =   23
      Top             =   1440
      Width           =   375
   End
   Begin VB.CheckBox optFn3 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4740
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optGr3 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   21
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdGrt 
      Caption         =   "&All"
      Height          =   360
      Left            =   3120
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Grant All"
      Top             =   0
      Width           =   720
   End
   Begin VB.CommandButton cmdRvk 
      Caption         =   "&None"
      Height          =   360
      Left            =   3960
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Revoke All"
      Top             =   0
      Width           =   720
   End
   Begin VB.CheckBox optGr1 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox optGr2 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox optFn2 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4740
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox optVw2 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3780
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.CheckBox optEd2 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2820
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox optFn1 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4740
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox optVw1 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3780
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox optEd1 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2820
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   720
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Certification"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Melters Log"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Heat Treatment Log"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   1860
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblSect 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Collection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Functions   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   4740
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "View         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3780
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2820
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chemical Analysis"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tensile Test"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "AdmnUperm9"
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
Dim strUserName As String
 

Private Sub FormatBoxes()
   '    Dim bControl As Byte
   '    For bControl = 0 To Controls.Count - 1
   '        If TypeOf Controls(bControl) Is CheckBox Then _
   '            Controls(bControl).Caption = "____"
   '    Next
   '
End Sub

Private Sub TestBoxes(sSwitch As Byte, bValue As Byte)
   Select Case sSwitch
      Case 1
         optEd1.Enabled = bValue
         optVw1.Enabled = bValue
         optFn1.Enabled = bValue
      Case 2
         optEd2.Enabled = bValue
         optVw2.Enabled = bValue
         optFn2.Enabled = bValue
      Case 3
         optEd3.Enabled = bValue
         optVw3.Enabled = bValue
         optFn3.Enabled = bValue
      Case 4
         optEd4.Enabled = bValue
         optVw4.Enabled = bValue
         optFn4.Enabled = bValue
      Case 5
         optEd5.Enabled = bValue
         optVw5.Enabled = bValue
         optFn5.Enabled = bValue
   End Select
   
End Sub

Private Sub CheckSection()
   'Check Section-Must be changed as added or removed
   Dim b As Byte
   If optGr1.Value = 0 And optGr2.Value = 0 Then
      b = 0
   Else
      b = 1
   End If
   Secure.UserTime = b
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub cmdCan_Click()
   'CheckSection
   Unload Me
   
End Sub



Private Sub cmdGrt_Click()
   GrantAll
   
End Sub

Private Sub cmdRvk_Click()
   RevokeAll
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      Fillboxes
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

'Grant all in use

Private Sub GrantAll()
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If TypeOf Controls(iList) Is CheckBox Then
         If Controls(iList).Visible Then Controls(iList).Value = vbChecked
      End If
   Next
   
End Sub

'Revoke all in use

Private Sub RevokeAll()
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If TypeOf Controls(iList) Is CheckBox Then
         If Controls(iList).Visible Then Controls(iList).Value = vbUnchecked
      End If
   Next
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Move AdmnUuser2.Left + 700, AdmnUuser2.Top + 1500
   FormatBoxes
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdmnUperm7 = Nothing
   
End Sub




Private Sub Fillboxes()


   strUserName = Trim(SecPw.UserLcName)
   
   CheckUserRec strUserName
   Dim RdoCls As ADODB.Recordset
   On Error Resume Next
   
   sSql = "SELECT ISNULL(FEATURENAME, '') FEATURENAME, FEATUREGRP, FEATUREEDT, FEATUREVW, FEATUREFN FROM featAccs " _
            & " WHERE USRNAME = '" & strUserName & "'"
            
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCls, ES_DYNAMIC)
   If bSqlRows Then
      With RdoCls
      Do Until .EOF
         
            If (TENILETEST = !FEATURENAME) Then
               optGr1.Value = IIf(IsNull(!FEATUREGRP), 0, !FEATUREGRP)
               optEd1.Value = IIf(IsNull(!FEATUREEDT), 0, !FEATUREEDT)
               optVw1.Value = IIf(IsNull(!FEATUREVW), 0, !FEATUREVW)
               optFn1.Value = IIf(IsNull(!FEATUREFN), 0, !FEATUREFN)
            End If
            
            If (CHEMANALYSIS = !FEATURENAME) Then
               optGr2.Value = IIf(IsNull(!FEATUREGRP), 0, !FEATUREGRP)
               optEd2.Value = IIf(IsNull(!FEATUREEDT), 0, !FEATUREEDT)
               optVw2.Value = IIf(IsNull(!FEATUREVW), 0, !FEATUREVW)
               optFn2.Value = IIf(IsNull(!FEATUREFN), 0, !FEATUREFN)
            End If
            
            If (HEATTREATMENT = !FEATURENAME) Then
               optGr3.Value = IIf(IsNull(!FEATUREGRP), 0, !FEATUREGRP)
               optEd3.Value = IIf(IsNull(!FEATUREEDT), 0, !FEATUREEDT)
               optVw3.Value = IIf(IsNull(!FEATUREVW), 0, !FEATUREVW)
               optFn3.Value = IIf(IsNull(!FEATUREFN), 0, !FEATUREFN)
            End If
            
            If (MELTERSLOG = !FEATURENAME) Then
               optGr4.Value = IIf(IsNull(!FEATUREGRP), 0, !FEATUREGRP)
               optEd4.Value = IIf(IsNull(!FEATUREEDT), 0, !FEATUREEDT)
               optVw4.Value = IIf(IsNull(!FEATUREVW), 0, !FEATUREVW)
               optFn4.Value = IIf(IsNull(!FEATUREFN), 0, !FEATUREFN)
            End If
            
            If (MATCERT = !FEATURENAME) Then
               optGr5.Value = IIf(IsNull(!FEATUREGRP), 0, !FEATUREGRP)
               optEd5.Value = IIf(IsNull(!FEATUREEDT), 0, !FEATUREEDT)
               optVw5.Value = IIf(IsNull(!FEATUREVW), 0, !FEATUREVW)
               optFn5.Value = IIf(IsNull(!FEATUREFN), 0, !FEATUREFN)
            End If
            
         .MoveNext
      Loop
      End With
      RdoCls.Close
   Else
      ' Enter Tensile Test
      sSql = "INSERT INTO featAccs (USRNAME,FEATUREID, FEATURENAME, FEATUREGRP, FEATUREEDT, " _
                  & "FEATUREVW, FEATUREFN)  " _
            & " VALUES ('" & strUserName & "', 'TNTEST', '" & TENILETEST & "', 0,0,0,0)"
      
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
      ' Enter Chem Analysis
      sSql = "INSERT INTO featAccs (USRNAME,FEATUREID, FEATURENAME, FEATUREGRP, FEATUREEDT, " _
                  & "FEATUREVW, FEATUREFN)  " _
            & " VALUES ('" & strUserName & "', 'CHEMALYSIS', '" & CHEMANALYSIS & "', 0,0,0,0)"
      
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
      ' Enter Heat Treatment
      sSql = "INSERT INTO featAccs (USRNAME,FEATUREID, FEATURENAME, FEATUREGRP, FEATUREEDT, " _
                  & "FEATUREVW, FEATUREFN)  " _
            & " VALUES ('" & strUserName & "', 'HT', '" & HEATTREATMENT & "', 0,0,0,0)"
      
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
      ' Enter Chem Analysis
      sSql = "INSERT INTO featAccs (USRNAME,FEATUREID, FEATURENAME, FEATUREGRP, FEATUREEDT, " _
                  & "FEATUREVW, FEATUREFN)  " _
            & " VALUES ('" & strUserName & "', 'MLTLOG', '" & MELTERSLOG & "', 0,0,0,0)"
      
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
      ' Enter Chem Analysis
      sSql = "INSERT INTO featAccs (USRNAME,FEATUREID, FEATURENAME, FEATUREGRP, FEATUREEDT, " _
                  & "FEATUREVW, FEATUREFN)  " _
            & " VALUES ('" & strUserName & "', 'MATCERT', '" & MATCERT & "', 0,0,0,0)"
      
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   
   
      optGr1.Value = 0
      optGr2.Value = 0
      optGr3.Value = 0
      optGr4.Value = 0
      optGr5.Value = 0
      
   End If
   ClearResultSet RdoCls
   Set RdoCls = Nothing
   
   TestBoxes 1, optGr1.Value
   TestBoxes 2, optGr2.Value
   TestBoxes 3, optGr3.Value
   TestBoxes 4, optGr4.Value
   TestBoxes 5, optGr5.Value
   
   Exit Sub
   
DiaErr1:
   sProcName = "FillBoxes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub optEd1_Click()
   
   sSql = "UPDATE featAccs SET FEATUREEDT =" _
          & optEd1.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & TENILETEST & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optEd1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd2_Click()
   sSql = "UPDATE featAccs SET FEATUREEDT =" _
          & optEd2.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & CHEMANALYSIS & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optEd2_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optEd3_Click()
   sSql = "UPDATE featAccs SET FEATUREEDT =" _
          & optEd3.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & HEATTREATMENT & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optEd3_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd4_Click()
   sSql = "UPDATE featAccs SET FEATUREEDT =" _
          & optEd4.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & MELTERSLOG & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optEd4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optEd5_Click()
   sSql = "UPDATE featAccs SET FEATUREEDT =" _
          & optEd5.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & MATCERT & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optEd5_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub optFn1_Click()
   
   sSql = "UPDATE featAccs SET FEATUREFN =" _
          & optFn1.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & TENILETEST & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optFn1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optFn2_Click()
   sSql = "UPDATE featAccs SET FEATUREFN =" _
          & optFn2.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & CHEMANALYSIS & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optFn2_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optFn3_Click()
   sSql = "UPDATE featAccs SET FEATUREFN =" _
          & optFn3.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & HEATTREATMENT & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optFn3_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optFn4_Click()
   sSql = "UPDATE featAccs SET FEATUREFN =" _
          & optFn4.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & MELTERSLOG & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optFn4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optFn5_Click()
   sSql = "UPDATE featAccs SET FEATUREFN =" _
          & optFn5.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & MATCERT & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optFn5_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optGr1_Click()
   If optGr1.Value = vbUnchecked Then
      TestBoxes 1, vbUnchecked
      optEd1.Value = vbUnchecked
      optVw1.Value = vbUnchecked
      optFn1.Value = vbUnchecked
   Else
      TestBoxes 1, vbChecked
   End If
   
   sSql = "UPDATE featAccs SET FEATUREGRP =" _
          & optGr1.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & TENILETEST & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optGr1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr2_Click()
   If optGr2.Value = vbUnchecked Then
      TestBoxes 2, vbUnchecked
      optEd2.Value = vbUnchecked
      optVw2.Value = vbUnchecked
      optFn2.Value = vbUnchecked
   Else
      TestBoxes 2, vbChecked
   End If
   
   sSql = "UPDATE featAccs SET FEATUREGRP =" _
          & optGr2.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & CHEMANALYSIS & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub


Private Sub optGr2_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr3_Click()
   If optGr3.Value = vbUnchecked Then
      TestBoxes 2, vbUnchecked
      optEd3.Value = vbUnchecked
      optVw3.Value = vbUnchecked
      optFn3.Value = vbUnchecked
   Else
      TestBoxes 3, vbChecked
   End If
   
   sSql = "UPDATE featAccs SET FEATUREGRP =" _
          & optGr3.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & HEATTREATMENT & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub


Private Sub optGr3_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
Private Sub optGr4_Click()
   If optGr4.Value = vbUnchecked Then
      TestBoxes 4, vbUnchecked
      optEd4.Value = vbUnchecked
      optVw4.Value = vbUnchecked
      optFn4.Value = vbUnchecked
   Else
      TestBoxes 4, vbChecked
   End If
   
   sSql = "UPDATE featAccs SET FEATUREGRP =" _
          & optGr4.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & MELTERSLOG & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub


Private Sub optGr4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
Private Sub optGr5_Click()
   If optGr5.Value = vbUnchecked Then
      TestBoxes 5, vbUnchecked
      optEd5.Value = vbUnchecked
      optVw5.Value = vbUnchecked
      optFn5.Value = vbUnchecked
   Else
      TestBoxes 5, vbChecked
   End If
   
   sSql = "UPDATE featAccs SET FEATUREGRP =" _
          & optGr5.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & MATCERT & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub


Private Sub optGr5_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optVw1_Click()
   sSql = "UPDATE featAccs SET FEATUREVW =" _
          & optVw1.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & TENILETEST & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optVw1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw2_Click()
   sSql = "UPDATE featAccs SET FEATUREVW =" _
          & optVw2.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & CHEMANALYSIS & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optVw2_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optVw3_Click()
   sSql = "UPDATE featAccs SET FEATUREVW =" _
          & optVw3.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & HEATTREATMENT & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optVw3_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optVw4_Click()
   sSql = "UPDATE featAccs SET FEATUREVW =" _
          & optVw4.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & MELTERSLOG & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optVw4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optVw5_Click()
   sSql = "UPDATE featAccs SET FEATUREVW =" _
          & optVw5.Value & " WHERE USRNAME ='" _
          & strUserName & "' AND FEATURENAME = '" & MATCERT & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub optVw5_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Function CheckUserRec(strUserName As String)
   
End Function
