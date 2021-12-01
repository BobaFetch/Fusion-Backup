VERSION 5.00
Begin VB.Form AdmnADf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Fields"
   ClientHeight    =   3345
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5760
   Icon            =   "AdmnADf06a.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtColRef 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtColName 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtTblName 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtLblName 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.ComboBox cmbLabelID 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Z5 
      Caption         =   "Column Reference"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Z4 
      Caption         =   "Column Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Z3 
      Caption         =   "Table Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Z2 
      Caption         =   "Label Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Z1 
      Caption         =   "Label ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "AdmnADf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit
Dim bOnLoad As Byte

Dim lblID As String
Dim lblName As String
Dim tblName As String
Dim colName As String
Dim colRef As String

Private Sub cmbLabelID_Click()
    GetLblDetail
End Sub

Private Sub cmdApply_Click()
    UpdateReports
End Sub

Private Sub cmdCan_Click()
    Unload Me
End Sub


Private Sub Form_Load()
   FormLoad Me
   FillLabelID
   bOnLoad = 1
   MouseCursor 0

End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub FillLabelID()
   Dim RdoLbl As ADODB.Recordset
   cmbLabelID.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT LABEL_ID FROM CustomFields"
   LoadComboBox cmbLabelID, -1
   If cmbLabelID.ListCount > 0 Then
      cmbLabelID = cmbLabelID.List(0)
   End If
   Exit Sub
DiaErr1:
   sProcName = "FillLabel"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
Private Sub GetLblDetail()
   Dim RdoLbl As ADODB.Recordset
   On Error GoTo DiaErr1
   lblID = "" & Trim(cmbLabelID)
   sSql = "SELECT * from CustomFields where LABEL_ID = '" & lblID & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLbl, ES_FORWARD)
   If bSqlRows Then
      With RdoLbl
         lblName = "" & Trim(!LABEL_NAME)
         tblName = "" & Trim(!TBL_NAME)
         colName = "" & Trim(!TBL_COLNAME)
         colRef = "" & Trim(!COLREF_KEY)
         txtLblName = "" & Trim(!LABEL_NAME)
         txtTblName = "" & Trim(!TBL_NAME)
         txtColName = "" & Trim(!TBL_COLNAME)
         txtColRef = "" & Trim(!COLREF_KEY)
      End With
        ClearResultSet RdoLbl
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "GetLblDet"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   

End Sub
Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdmnADf06a = Nothing
   
End Sub

Private Sub UpdateReports()
   
   Dim RdoUpd As ADODB.Recordset
   
   'Update them
      sSql = "SELECT * FROM CustomFields WHERE " _
             & "LABEL_ID ='" & lblID & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoUpd, ES_KEYSET)
      If bSqlRows Then
         With RdoUpd
            '.Edit
            !LABEL_NAME = Compress(txtLblName)
            !TBL_NAME = Compress(txtTblName)
            !TBL_COLNAME = Compress(txtColName)
            !COLREF_KEY = Compress(txtColRef)
            .Update
         End With
      End If
   
End Sub


