VERSION 5.00
Begin VB.Form diaSfDel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Shift Code"
   ClientHeight    =   1560
   ClientLeft      =   1200
   ClientTop       =   435
   ClientWidth     =   6840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtDsc 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift Code"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "diaSfDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim bOnLoad As Byte
Dim bGoodCode As Byte
Dim RdoCde As ADODB.Recordset
'Dim RdoQry As rdoQuery
Dim cmdObj As ADODB.Command


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Click()
   bGoodCode = GetShiftCode()
   
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdDel_Click()
    Dim RdoEmp As ADODB.Recordset
    Dim strTCCard As String
    Dim strSfcode As String
    
    On Error GoTo DiaErr1
    strSfcode = cmbCde.Text
    
    sSql = "SELECT  PREMNUMBER FROM  sfempTable " & _
           "WHERE (SFREF = '" & strSfcode & "' OR SFREFSUN = '" & strSfcode & "' OR " & _
           "SFREFMON = '" & strSfcode & "'  OR SFREFTUE = '" & strSfcode & "'  OR " & _
           "SFREFWED = '" & strSfcode & "'  OR SFREFTHU = '" & strSfcode & "'  OR " & _
           "SFREFFRI = '" & strSfcode & "'  OR SFREFSAT = '" & strSfcode & "' )"
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp, ES_FORWARD)
    
    If bSqlRows Then
       MsgBox ("Shift Code is in Use")
        ClearResultSet RdoEmp
    Else
         sSql = "Delete From sfcdTable where SFREF = '" & strSfcode & "'"
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
         MsgBox ("Shift Code Deleted")
         FillCombo
    End If
    Set RdoEmp = Nothing
Exit Sub

DiaErr1:
   sProcName = "cmdDel"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub
Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT SFREF, SFCODE,SFDESC,SFSTHR," _
          & "SFENHR,SFLUNSTHR, SFLUNENHR,SFADJHR,SFADDRT FROM " _
          & "sfcdTable WHERE SFCODE= ? "
                  
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql

   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 6
   cmdObj.Parameters.Append prmObj
                  
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   bOnLoad = 1
   Show
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoCde = Nothing
   Set cmdObj = Nothing
   Set diaSfDel = Nothing
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT SFCODE FROM SfcdTable "
   LoadComboBox cmbCde, -1
   If cmbCde.ListCount > 0 Then
      cmbCde = cmbCde.List(0)
      bGoodCode = GetShiftCode
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetShiftCode() As Byte
    Dim strSftCode As String
    Dim strType As String
    
    strSftCode = Compress(cmbCde)
    On Error GoTo DiaErr1
    'RdoQry(0) = strSftCode
    'bSqlRows = GetQuerySet(RdoCde, RdoQry, ES_KEYSET)
    cmdObj.Parameters(0).Value = strSftCode
    bSqlRows = clsADOCon.GetQuerySet(RdoCde, cmdObj, ES_KEYSET)
        
    If bSqlRows Then
        With RdoCde
            cmbCde = "" & Trim(!SFCODE)
            txtDsc = "" & Trim(!SFDESC)
                        
'            txtBeg = "" & Trim(!SFSTHR)
'            txtEnd = "" & Trim(!SFENHR)
'            txtLBeg = "" & Trim(!SFLUNSTHR)
'            txtLEnd = "" & Trim(!SFLUNENHR)
            ' Get the total hours
'            CalculateLunchTime
'            txtAdj = "" & IIf(IsNull(Trim(!SFADJHR)), "0", Trim(!SFADJHR))
'            txtAddRate = "" & IIf(IsNull(Trim(!SFADDRT)), "0.00", Trim(!SFADDRT))
      End With
      GetShiftCode = True
   Else
'      txtDsc = ""
      GetShiftCode = False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "GetShiftCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

