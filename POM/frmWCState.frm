VERSION 5.00
Begin VB.Form frmWCState 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change the WC Queue Status"
   ClientHeight    =   2835
   ClientLeft      =   4410
   ClientTop       =   4980
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2160
      Picture         =   "frmWCState.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1500
   End
   Begin VB.ComboBox cboStatCode 
      Height          =   420
      Left            =   6960
      TabIndex        =   7
      Top             =   720
      Width           =   1500
   End
   Begin VB.ComboBox cboWorkCenter 
      Enabled         =   0   'False
      Height          =   420
      Left            =   3960
      TabIndex        =   1
      Top             =   720
      Width           =   1500
   End
   Begin VB.ComboBox cboShop 
      Enabled         =   0   'False
      Height          =   420
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1500
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3960
      Picture         =   "frmWCState.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Status Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label Label8 
      Caption         =   "Work Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label Label7 
      Caption         =   "Shop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Change WorkCenter Status"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   6135
   End
End
Attribute VB_Name = "frmWCState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private keysTyped As Integer    ' number of keys typed in combo
'Private keysSoFar As String     ' the keys that were typed

Private bWCStat As Boolean
Public strPart As String
Public strRun As String
Public strOp As String



Private Sub SetCenterSatus(bWCStat As Boolean)
   
   Dim status As Integer
   status = IIf(bWCStat = True, 1, 0)
   sSql = "UPDATE RnopTable SET STATUS_REF = " & status & vbCrLf _
          & " where OPREF='" & strPart & "'" & vbCrLf _
          & "and OPRUN=" & strRun & vbCrLf _
          & "and OPNO=" & strOp
   clsADOCon.ExecuteSQL sSql
   
End Sub

Private Sub FillShops()
   Dim wc As New ClassWorkCenter
   wc.PopulateShopCombo cboShop, cboWorkCenter
End Sub

Private Sub FillWorkCenters()
   Dim wc As New ClassWorkCenter
   wc.PoulateWorkCenterCombo cboShop, cboWorkCenter
End Sub

Private Sub FillStatusCode()
   
   Dim RdoStc As ADODB.Recordset
   Dim iRows As Integer
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT STATUS_REF, STATUS_CODE " & _
                " FROM StcodeTable ORDER BY STATUS_REF"
   LoadComboBox cboStatCode, -1
   DoEvents
   
   Exit Sub
   
DiaErr1:
   sProcName = "FillStatusCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdProceed_Click()
   Dim statRef As String
   statRef = Trim(cboStatCode)
   'allow empty string
   'If (statRef <> "") Then
    sSql = "UPDATE RnopTable SET STATUS_REF = '" & statRef & "'" & vbCrLf _
           & " where OPREF='" & strPart & "'" & vbCrLf _
           & "and OPRUN=" & strRun & vbCrLf _
           & "and OPNO=" & strOp
    clsADOCon.ExecuteSQL sSql
   Unload Me
   
End Sub

Private Sub Form_Load()
   CenterForm Me
    FillShops
    FillWorkCenters
    FillStatusCode
    GetWCStatus
End Sub

Private Function GetWCStatus()
   
   Dim rdoOPNum As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT ISNULL(STATUS_REF, '') STATUS_REF FROM rnopTable " _
            & "where OPREF = '" & strPart & "' " _
          & " AND OPRUN =" & strRun & " AND OPNO = '" & strOp & "'"
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdoOPNum)
   If gblnSqlRows Then
      With rdoOPNum
         cboStatCode = Trim(!STATUS_REF)
      End With
   End If
   
   Set rdoOPNum = Nothing
   
   Exit Function
   
DiaErr1:
   Set rdoOPNum = Nothing
   
End Function

