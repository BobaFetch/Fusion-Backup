VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form PackPSf12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer List for Shipping Manifest."
   ClientHeight    =   3990
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3990
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List (Contains Valid Customers)"
      Top             =   360
      Width           =   1555
   End
   Begin VB.ListBox lstSelCust 
      Height          =   2400
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   915
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Cancel Selected Invoice"
      Top             =   1440
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   5535
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   3
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
      PictureUp       =   "PackPSf12.frx":0000
      PictureDn       =   "PackPSf12.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Work Centers"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1065
   End
End
Attribute VB_Name = "PackPSf12"
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

'*********************************************************************************
' PackPSf12 - Assign Customer Payers
'
' Notes:
'
' Created:
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Public bRemote As Byte
Dim sRpt As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   ' fill customer combo box
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME,CUCUTOFF,SOCUST" & vbCrLf _
          & "FROM CustTable cust" & vbCrLf _
          & "join SohdTable so on cust.CUREF = so.SOCUST" & vbCrLf _
          & "join SoitTable item on item.ITSO = so.SONUMBER" & vbCrLf _
          & "where SOCANCELED = 0" & vbCrLf _
          & "AND ITACTUAL IS NULL AND ITQTY>0 AND ITPSNUMBER='' AND ITINVOICE=0 AND ITCANCELED=0" & vbCrLf _
          & "ORDER BY CUREF"
   
   LoadComboBox cmbCst
   If cmbCst.ListCount = 0 Then
      '    cmbCst = cmbCst.List(0)
      'Else
      MsgBox "No Customers With Open Sales Orders Found.", vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmdAdd_Click()
   Dim sItem As String
   
   Dim i As Integer
   Dim strCust As String
   Dim strASN As String
   Dim strShpFrom As String
   strShpFrom = "208120"
   On Error Resume Next
   
   strCust = Compress(cmbCst)
   
   If (CheckIfCustExists(strCust) <> "") Then
      MsgBox "The Customer already exists in the List - " & strCust & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   
   strASN = GetCurrentASN
   
   If (strASN = "") Then
      MsgBox "The ASN number is Empty. Check the Admin for more inforamtion.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   sSql = "INSERT INTO ASNInfoTable(SHPFRMIDCODE, CUREF, SHPREF, LASTASNNUM, PACCARDPART," _
               & "TRUCKPLANT, BOEINGPART, POLETTERREF) VALUES (" _
            & "'" & strShpFrom & "','" & strCust & "','BOE','" _
            & strASN & "',0,0,1,'VMS')"
   
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   lstSelCust.AddItem strCust
   
   Exit Sub
DiaErr1:
   sProcName = "cmdAdd_Click"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdDel_Click()
   Dim sItem As String
   Dim i As Integer
   With lstSelCust
      i = .ListIndex
      If i > -1 Then
         sItem = .List(i)
         On Error Resume Next
         sSql = "DELETE FROM ASNInfoTable WHERE CUREF = '" & sItem & "'"
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
         .RemoveItem (i)
         If i = .ListCount Then
            i = i - 1
         End If
         .ListIndex = i
      End If
   End With
   
   Exit Sub
DiaErr1:
   sProcName = "cmdDel_Click"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      If Not bRemote Then
         FillCombo
      End If
      FillSelCust
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not bRemote Then
      FormUnload
      SaveCurrentSelections
   End If
   Set PackPSf12 = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Function CheckIfCustExists(strCust As String) As String
   Dim RdoCust As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT CUREF FROM ASNInfoTable WHERE CUREF LIKE '" & strCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCust, ES_FORWARD)
   If bSqlRows Then
      With RdoCust
         CheckIfCustExists = "" & Trim(!CUREF)
         ClearResultSet RdoCust
      End With
   Else
      CheckIfCustExists = ""

   End If
   Set RdoCust = Nothing
   Exit Function

modErr1:
   sProcName = "CheckIfCustExists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Function


Private Function GetCurrentASN() As String
   Dim RdoASN As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT DISTINCT LASTASNNUM FROM ASNInfoTable WHERE BOEINGPART = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoASN, ES_FORWARD)
   If bSqlRows Then
      With RdoASN
         GetCurrentASN = Trim(!LASTASNNUM)
         ClearResultSet RdoASN
      End With
   Else
      GetCurrentASN = ""

   End If
   Set RdoASN = Nothing
   Exit Function

modErr1:
   sProcName = "GetCurrentASN"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Function

Private Sub FillSelCust()
   Dim RdoSelCust As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT DISTINCT CUREF FROM ASNInfoTable WHERE BOEINGPART = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSelCust, ES_FORWARD)
   If bSqlRows Then
      With RdoSelCust
         Do Until .EOF
            lstSelCust.AddItem "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet RdoSelCust
      End With
   End If
   Set RdoSelCust = Nothing
   Exit Sub

modErr1:
   sProcName = "FillSelCust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Sub

Private Sub lstSelWC_Click()

'   Dim i As Integer
'   Dim sItem As String
'
'   i = lstAva.ListIndex
'   If i > -1 Then
'      sItem = .List(i)
'      txtPrt.Text = sItem
'   End If
   
End Sub
