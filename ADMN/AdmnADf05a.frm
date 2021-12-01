VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnADf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Status Code"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbStatID 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Status ID"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdDel 
      Cancel          =   -1  'True
      Caption         =   "D&elete"
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Delete This Standard Comment"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   720
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2010
      FormDesignWidth =   6120
   End
   Begin VB.Label txtStatCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Status Code"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblListIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status ID"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Code"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "AdmnADf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbStatID_Click()
   On Error GoTo DiaErr1
   Dim RdoStc As ADODB.Recordset
   
   Dim strStatID As String
   strStatID = cmbStatID
   
   sSql = "SELECT STATUS_CODE FROM StcodeTable " & _
            "WHERE STATUS_REF = '" & strStatID & "'"
                
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStc, ES_FORWARD)
   If bSqlRows Then
      With RdoStc
        txtStatCode = Trim(!STATUS_CODE)
        ClearResultSet RdoStc
      End With
   End If
    
   Set RdoStc = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "cmbStatID"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDel_Click()
    Dim RdoStc As ADODB.Recordset
    Dim strStatID As String
    strStatID = cmbStatID
    
    If (strStatID <> "") Then
        sSql = "SELECT DISTINCT STATUS_REF FROM StCmtTable " & _
                    " WHERE STATUS_REF = '" & strStatID & "'"
                     
        bSqlRows = clsADOCon.GetDataSet(sSql, RdoStc, ES_FORWARD)
        If Not bSqlRows Then
            sSql = "DELETE FROM StcodeTable " & _
                        "WHERE STATUS_REF = '" & strStatID & "'"
            
            clsADOCon.ExecuteSQL sSql
            ' Remove the status code
            cmbStatID.Clear
            txtStatCode = ""
            FillCombo
        Else
            MsgBox "Status Code is being used.", vbInformation
            ClearResultSet RdoStc
        End If
        Set RdoStc = Nothing
        
        Exit Sub
DiaErr1:
        sProcName = "Delete StatusCode "
        CurrError.Number = Err.Number
        CurrError.Description = Err.Description
        DoModuleErrors Me
    End If
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1152
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
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
   Set AdmnADf05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   Dim RdoStc As ADODB.Recordset
   
   sSql = "SELECT STATUS_REF, STATUS_CODE " & _
                " FROM StcodeTable ORDER BY STATUS_REF"
                
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStc, ES_FORWARD)
   If bSqlRows Then
      With RdoStc
         Do Until .EOF
            cmbStatID.AddItem (Trim(!STATUS_REF))
            .MoveNext
         Loop
         ClearResultSet RdoStc
      End With
   End If
    
   Set RdoStc = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
