VERSION 5.00
Begin VB.Form AddNewStatCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AddNewStatCode"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   LinkTopic       =   "AddNewStatCode"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtSCTRef 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Cancel Status code assignment"
         Top             =   600
         Width           =   915
      End
      Begin VB.ComboBox cmbStatID 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "Select or Enter Sales Order Number (List Contains Last 3 Years Up To 500 Enties)"
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton CmdAssign 
         Caption         =   "&Add"
         Height          =   315
         Left            =   4560
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Add This Sales Order Item"
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblStatType 
         Caption         =   "StatType"
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblSCTRef1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Our Item Number"
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblSCTRef2 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   600
         TabIndex        =   7
         ToolTipText     =   "Item Revision"
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "New Status Code"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblStatCd 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Status Code"
         Top             =   600
         Width           =   3015
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "AddNewStatCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbStatID_Click()
    Dim strStatID As String
    Dim RdoStatCode As ADODB.Recordset

    strStatID = cmbStatID
    sSql = "SELECT STATUS_CODE  FROM STCODETABLE " & _
                "WHERE STATUS_REF = '" & strStatID & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoStatCode, ES_FORWARD)
    If bSqlRows Then
       With RdoStatCode
          lblStatCd = "" & Trim(!STATUS_CODE)
          ClearResultSet RdoStatCode
       End With
    End If
    Set RdoStatCode = Nothing
End Sub

Private Sub CmdAssign_Click()
    Dim strStatID As String
    Dim iActStat As Integer
    
    strStatID = cmbStatID
    If strStatID = "" Then
      MsgBox "Please Select a Status Code.", _
         vbInformation, Caption
        Exit Sub
    End If
        
    StatusCode.cmbStatID = strStatID
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    StatusCode.cmbStatID = ""
    Unload Me
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      FillStatCodeFiltered
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
    FormatControls
    bOnLoad = 1
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Resize()
   Refresh
End Sub


Private Sub FillStatCodeFiltered()
    Dim RdoStCode As ADODB.Recordset
   
    Dim strStatCmdRef As String
    Dim strStatCmdRef1 As String
    Dim strStatCmdRef2 As String
    Dim strTransType As String
    Dim sSql1 As String
    
   On Error Resume Next
   On Error GoTo DiaErr1
    
    strStatCmdRef = txtSCTRef
    strStatCmdRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatCmdRef2 = lblSCTRef2
    strTransType = lblStatType
   
    sSql = "SELECT StcodeTable.STATUS_REF " & _
                " From StcodeTable " & _
            " WHERE STATUS_REF NOT IN " & _
                " (SELECT StCmtTable.STATUS_REF " & _
            " FROM StCmtTable WHERE " & _
                    " STATCODE_TYPE_REF = '" & strTransType & "' AND " & _
                    " STATUS_CMT_REF = '" & strStatCmdRef & "'" & _
                    " AND STATUS_CMT_REF1 = '" & strStatCmdRef1 & "'" & _
                    " AND STATUS_CMT_REF2 = '" & strStatCmdRef2 & "')"

   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStCode, ES_FORWARD)
   If bSqlRows Then
      With RdoStCode
         Do Until .EOF
            AddComboStr cmbStatID.hwnd, "" & Trim(!STATUS_REF)
            .MoveNext
         Loop
         ClearResultSet RdoStCode
      End With
   End If
   Set RdoStCode = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillStatCodeFiltered"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

