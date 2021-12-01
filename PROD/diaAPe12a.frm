VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPe12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Part Prefix to SINC Manufacturing Release"
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
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Enter Leading Char(s) Or Blank (200 Max Selected)"
      Top             =   840
      Width           =   2775
   End
   Begin VB.ListBox lstSel 
      Height          =   2010
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox lstAva 
      Height          =   2010
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   915
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Cancel Selected Invoice"
      Top             =   1800
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   5535
   End
   Begin VB.ComboBox cmbBU 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Contains Customers With Invoices"
      Top             =   360
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
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
      PictureUp       =   "diaAPe12a.frx":0000
      PictureDn       =   "diaAPe12a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Prefix's"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Prefix Categories"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Business Unit"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1065
   End
End
Attribute VB_Name = "diaAPe12a"
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
' diaAPe12a - Assign Customer Payers
'
' Notes:
'
' Created: (nth) 07/12/04
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

Private Sub cmbBU_Click()
   sRpt = Compress(cmbBU)
   GetPartialParts sRpt

End Sub

Private Sub cmbBU_LostFocus()
   If Not bCancel Then
      FindBU cmbBU
      sRpt = Compress(cmbBU)
      GetPartialParts sRpt
   End If
End Sub


Private Sub cmdAdd_Click()
   Dim sItem As String
   
   Dim I As Integer
   Dim strPreFixPart As String
   
   On Error Resume Next
   
   strPreFixPart = Compress(txtPrt.Text)
   
   If (CheckIfPartExists(strPreFixPart) <> "") Then
      MsgBox "The Part Prefix Exists - " & strPreFixPart & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   
   ' Insert the part
   sSql = "INSERT INTO sincbuPCTable (BUCODE,BUPARTREFCODE) VALUES('" & sRpt _
          & "','" & strPreFixPart & "')"
   'RdoCon.Execute sSql, rdExecDirect
   clsADOCon.ExecuteSQL sSql
   
   lstAva.AddItem strPreFixPart
   txtPrt.Text = ""
   
   Exit Sub
DiaErr1:
   sProcName = "cmdAdd_Click"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdDel_Click()
   Dim sItem As String
   Dim I As Integer
   With lstAva
      I = .ListIndex
      If I > -1 Then
         sItem = .List(I)
         On Error Resume Next
         sSql = "DELETE FROM sincbuPCTable WHERE BUPARTREFCODE = '" & sItem _
                & "' AND BUCODE = '" & sRpt & "'"
         'RdoCon.Execute sSql, rdExecDirect
          clsADOCon.ExecuteSQL sSql
         .RemoveItem (I)
         If I = .ListCount Then
            I = I - 1
         End If
         .ListIndex = I
      End If
   End With
   
   Exit Sub
DiaErr1:
   sProcName = "removevendors"
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
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      If Not bRemote Then
         FillCombo
      End If
      'FindCustomer Me, cmbBU
      sRpt = Compress(cmbBU)
      GetPartialParts sRpt
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
      cUR.CurrentVendor = cmbBU
      SaveCurrentSelections
   End If
   Set diaAPe12a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   FillSINCBU
   Exit Sub
DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function CheckIfPartExists(strPreFixPart As String) As String
   Dim RdoSinc As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT BUPARTREFCODE FROM sincbuPCTable WHERE BUPARTREFCODE LIKE '" & strPreFixPart & "'"
'   bSqlRows = GetDataSet(RdoSinc, ES_FORWARD)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSinc, ES_FORWARD)
   If bSqlRows Then
      With RdoSinc
         CheckIfPartExists = "" & Trim(!BUPARTREFCODE)
         ClearResultSet RdoSinc
      End With
   Else
      CheckIfPartExists = ""

   End If
   Set RdoSinc = Nothing
   Exit Function

modErr1:
   sProcName = "FillSINCBU"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Function

Private Function FindBU(strBU As String)
   Dim RdoSinc As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT SINC_BUCODE FROM SinchdTable WHERE SINC_BUCODE = '" & strBU & "'"
   'bSqlRows = GetDataSet(RdoSinc, ES_FORWARD)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSinc, ES_FORWARD)
   If bSqlRows Then
      With RdoSinc
         ClearResultSet RdoSinc
      End With
   Else
      
      Dim bResponse As Byte
      Dim sMsg As String
      
      sMsg = "Business Unit couldn't be found. Do you want to add as new BU?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      
      If bResponse = vbYes Then
         sSql = "INSERT INTO SinchdTable (SINC_BUCODE) VALUES ('" & strBU & "')"
         'RdoCon.Execute sSql, rdExecDirect
         clsADOCon.ExecuteSQL sSql
      End If
      
      
   End If
   Set RdoSinc = Nothing
   Exit Function

modErr1:
   sProcName = "FillSINCBU"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Function

Private Sub FillSINCBU()
   Dim RdoSinc As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT DISTINCT SINC_BUCODE FROM SinchdTable"
   'bSqlRows = GetDataSet(RdoSinc, ES_FORWARD)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSinc, ES_FORWARD)
   If bSqlRows Then
      With RdoSinc
         cmbBU = "" & Trim(!SINC_BUCODE)
         Do Until .EOF
            If Trim(!SINC_BUCODE) <> "NONE" Then _
                    cmbBU.AddItem "" & Trim(!SINC_BUCODE)
            .MoveNext
         Loop
         ClearResultSet RdoSinc
      End With
   End If
   Set RdoSinc = Nothing
   Exit Sub

modErr1:
   sProcName = "FillSINCBU"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Sub

Private Sub GetPartialParts(sBUCode As String)
   Dim RdoCst As ADODB.Recordset
   Dim sPayers() As String
   Dim I As Integer
   Dim sIn As String
   
   lstAva.Clear
   lstSel.Clear
   On Error GoTo DiaErr1
   
   sSql = "SELECT BUPARTREFCODE FROM sincbuPCTable WHERE BUCODE = '" & sBUCode & "'"
   'bSqlRows = GetDataSet(RdoCst)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst)
   
   If bSqlRows Then
      With RdoCst
         While Not .EOF
            lstAva.AddItem "" & Trim(.Fields(0))
            .MoveNext
         Wend
         .Cancel
      End With
      Set RdoCst = Nothing
   End If
   Exit Sub
DiaErr1:
   sProcName = "GetPartialParts"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub lstAva_Click()

'   Dim i As Integer
'   Dim sItem As String
'
'   i = lstAva.ListIndex
'   If i > -1 Then
'      sItem = .List(i)
'      txtPrt.Text = sItem
'   End If
   
End Sub
