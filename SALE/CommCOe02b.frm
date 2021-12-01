VERSION 5.00
Begin VB.Form CommCOe02b 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Remove Sales Persons"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "CommCOe02b.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6135
   Begin VB.ListBox lstSlp 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox lstSel 
      Height          =   1815
      Left            =   3600
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Tag             =   "0"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdRem 
      Caption         =   "<<"
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Persons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "CommCOe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions

'*************************************************************************************
'
' CommCOe02b - Add/Remove Sales Persons
'
' Created 08/26/03 (jcw)
'
' Revisions
'   08/26/03 (nth) Revised and updated
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim lSon As Long
Dim iItem As Integer
Dim sRev As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'**************************************************************************************

Private Sub FillSalesPersons()
   Dim rdoSlp As ADODB.Recordset
   sSql = "SELECT SPNUMBER,SPFIRST,SPLAST FROM SprsTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp, ES_FORWARD)
   If bSqlRows Then
      With rdoSlp
         Do Until .EOF
            lstSlp.AddItem Trim(.Fields(0)) _
                                & " - " & Trim(.Fields(1)) _
                                & " " & Trim(.Fields(2))
            .MoveNext
         Loop
         ClearResultSet rdoSlp
      End With
   End If
   Set rdoSlp = Nothing
End Sub

Private Sub cmdAdd_Click()
   Dim iTemp As Integer
   Dim sTemp As String
   If lstSlp.ListIndex >= 0 Then
      iTemp = lstSlp.ListIndex
      sTemp = lstSlp.List(iTemp)
      lstSlp.RemoveItem iTemp
      lstSel.AddItem sTemp
      AddSalesPerson SPNumber(sTemp)
      If lstSlp.ListCount > 0 Then
         If iTemp = 0 Then
            lstSlp.ListIndex = iTemp
         Else
            lstSlp.ListIndex = iTemp - 1
         End If
      End If
   End If
End Sub

Private Sub cmdRem_Click()
   Dim iTemp As Integer
   Dim sTemp As String
   If lstSel.ListIndex >= 0 Then
      iTemp = lstSel.ListIndex
      sTemp = lstSel.List(iTemp)
      lstSel.RemoveItem iTemp
      lstSlp.AddItem sTemp
      RemoveSalesPerson SPNumber(sTemp)
      If lstSel.ListCount > 0 Then
         If iTemp = 0 Then
            lstSel.ListIndex = iTemp
         Else
            lstSel.ListIndex = iTemp - 1
         End If
      End If
   End If
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      FillSalesPersons
      FillSelected
      bOnLoad = 0
   End If
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   
End Sub

Private Sub Form_Load()
   Dim sTemp As String
   FormatControls
   lSon = CLng(CommCOe02a.cmbSon)
   sTemp = CommCOe02a.cmbItm
   If Not IsNumeric(Right(sTemp, 1)) Then
      sRev = Right(sTemp, 1)
      iItem = CInt(Left(sTemp, Len(sTemp) - 1))
   Else
      iItem = CInt(sTemp)
      sRev = ""
   End If
   bOnLoad = 1
End Sub

Private Sub Form_LostFocus()
   Unload Me
End Sub

Private Sub FillSelected()
   Dim rdoSlp As ADODB.Recordset
   Dim sTemp As String
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT SMCOSM From SpcoTable WHERE SMCOSO = " & lSon _
          & " AND SMCOSOIT = " & iItem & " AND SMCOITREV = '" & sRev _
          & "' ORDER BY SMCOSM"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp, ES_FORWARD)
   If bSqlRows Then
      With rdoSlp
         Do Until .EOF
            iList = 0
            While iList < lstSlp.ListCount
               sTemp = SPNumber(lstSlp.List(iList))
               If sTemp = Trim(.Fields(0)) Then
                  sTemp = lstSlp.List(iList)
                  lstSel.AddItem sTemp
                  lstSlp.RemoveItem iList
               End If
               iList = iList + 1
            Wend
            .MoveNext
         Loop
         ClearResultSet rdoSlp
      End With
   End If
   Set rdoSlp = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillselec"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub AddSalesPerson(sSP As String)
   On Error Resume Next
   sSql = "INSERT INTO SpcoTable (SMCOSO,SMCOSOIT,SMCOITREV,SMCOSM,SMCOUSER) " _
          & "VALUES (" & lSon & "," & iItem & ",'" & sRev & "','" & sSP & "','" _
          & sInitials & "')"
  clsADOCon.ExecuteSQL sSql 'rdExecDirect
   If Err > 0 Then
      ValidateEdit
   Else
      CommCOe02a.SetDefaultComm sSP
   End If
   
End Sub

Private Sub RemoveSalesPerson(sSP As String)
   On Error Resume Next
   sSql = "DELETE FROM SpcoTable WHERE " _
          & "SMCOSO = " & lSon & " AND SMCOSOIT = " _
          & iItem & " AND SMCOITREV = '" _
          & sRev & "' AND SMCOSM = '" _
          & sSP & "'"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   If Err > 0 Then
      ValidateEdit
   End If
   
End Sub

Private Function SPNumber(sSP As String) As String
   'returns just the sales person number from a string
   'format #### - First Name Last Name
   Dim iTemp As Integer
   iTemp = InStr(1, sSP, "-")
   SPNumber = Trim(Left(sSP, iTemp - 1))
   
End Function

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CommCOe02a.optSlp.Value = vbUnchecked
   Set CommCOe02b = Nothing
End Sub
