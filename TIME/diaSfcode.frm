VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSfcode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shift Codes"
   ClientHeight    =   3405
   ClientLeft      =   1200
   ClientTop       =   855
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAddRate 
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtLBeg 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   " :"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtLEnd 
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Text            =   " :"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtHrs 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtAdj 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Text            =   " :"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtBeg 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Text            =   " :"
      Top             =   1440
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4800
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3405
      FormDesignWidth =   5520
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise Shift Code (2 char)"
      Top             =   480
      Width           =   660
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   12
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
      PictureUp       =   "diaSfcode.frx":0000
      PictureDn       =   "diaSfcode.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "in minutes"
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   20
      Top             =   2930
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Rate"
      Height          =   255
      Index           =   8
      Left            =   2640
      TabIndex        =   19
      ToolTipText     =   "Additional rate"
      Top             =   2445
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "LunchStart Time"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "LunchEnd Time"
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Hours"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2445
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Adjustment"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2925
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Time"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "diaSfcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'See the UpdateTables procedure for database revisions
Option Explicit
'Dim RdoCde As rdoResultset
'Dim RdoQry As rdoQuery
Dim RdoCde As ADODB.Recordset
Dim cmdObj As ADODB.Command
Dim prmObj As ADODB.Parameter
   
Dim bOnLoad As Byte
Dim bGoodCode As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Click()
   bGoodCode = GetShiftCode()
   
End Sub

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 2)
   If Len(cmbCde) Then
      bGoodCode = GetShiftCode()
      If Not bGoodCode Then AddShiftCode
   Else
      bGoodCode = False
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbCde = ""
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs1504"
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
   FormLoad Me
   FormatControls
   
   sSql = "SELECT SFREF, SFCODE,SFDESC,SFSTHR," _
          & "SFENHR,SFLUNSTHR, SFLUNENHR,SFADJHR,SFADDRT FROM " _
          & "sfcdTable WHERE SFCODE= ? "
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 6
   cmdObj.Parameters.Append prmObj
   
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
   Set prmObj = Nothing
   
   Set diaSfcode = Nothing
   
End Sub




Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub txtAddRate_LostFocus()
    
    Dim cAddRt As Currency
    If bGoodCode Then
       'RdoCde.Edit
       If (Trim(txtAddRate) = "") Then
          cAddRt = 0#
       Else
          cAddRt = CDbl(txtAddRate)
       End If
       RdoCde!SFADDRT = cAddRt
       RdoCde.Update
       If Err > 0 Then ValidateEdit Me
    End If
End Sub


Private Sub txtAdj_LostFocus()
    Dim iHrs As Integer
    If bGoodCode Then
       'RdoCde.Edit
       If (Trim(txtAdj) = "") Then
          iHrs = 0
       Else
          iHrs = CInt(txtAdj)
       End If
       RdoCde!SFADJHR = iHrs
       RdoCde.Update
       If Err > 0 Then ValidateEdit Me
    End If

End Sub



Private Sub txtLBeg_LostFocus()
    Dim tc As New ClassTimeCharge
    
    txtLBeg = tc.GetTime(txtLBeg)     'returns blank if invalid
    txtLBeg = Format(txtLBeg, "hh:nna/p")
    If bGoodCode Then
       'RdoCde.Edit
       RdoCde!SFLUNSTHR = "" & txtLBeg
       RdoCde.Update
       If Err > 0 Then ValidateEdit Me
    End If
    CalculateLunchTime
    
End Sub

Private Sub txtLEnd_LostFocus()
    Dim tc As New ClassTimeCharge
    
    txtLEnd = tc.GetTime(txtLEnd)     'returns blank if invalid
    txtLEnd = Format(txtLEnd, "hh:nna/p")
    If bGoodCode Then
       'RdoCde.Edit
       RdoCde!SFLUNENHR = "" & txtLEnd
       RdoCde.Update
       If Err > 0 Then ValidateEdit Me
    End If
    CalculateLunchTime
End Sub

Private Sub txtBeg_LostFocus()
    Dim tc As New ClassTimeCharge
    Dim minutes As Integer
    
    txtBeg = tc.GetTime(txtBeg)     'returns blank if invalid
    txtBeg = Format(txtBeg, "hh:nna/p")
    
    If (tc.IsValidTime(txtBeg)) Then
        If bGoodCode Then
           'RdoCde.Edit
           RdoCde!SFSTHR = "" & txtBeg
           RdoCde.Update
           If Err > 0 Then ValidateEdit Me
        End If
        
        CalculateLunchTime
        
    Else
        Dim strMsg As String
        strMsg = "Shift wrong format."
        MsgBox strMsg, vbInformation, Caption
        
        txtBeg = ""
        txtHrs = ""
    End If
End Sub

Private Sub txtEnd_LostFocus()
    Dim tc As New ClassTimeCharge
    Dim minutes As Integer
    txtEnd = tc.GetTime(txtEnd)
    txtEnd = Format(txtEnd, "hh:nna/p")
        
    If (tc.IsValidTime(txtBeg)) Then
       If bGoodCode Then
          'RdoCde.Edit
          RdoCde!SFENHR = txtEnd
          RdoCde.Update
          If Err > 0 Then ValidateEdit Me
       End If
        CalculateLunchTime
        
    Else
        Dim strMsg As String
        strMsg = "Shift time wrong format."
        MsgBox strMsg, vbInformation, Caption
        
        txtHrs = ""
        txtEnd = ""
    End If
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   On Error Resume Next
   If bGoodCode Then
      'RdoCde.Edit
      RdoCde!SFDESC = "" & txtDsc
      RdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Function GetShiftCode() As Byte
    Dim strSftCode As String
    Dim strType As String
    
    strSftCode = Compress(cmbCde)
    On Error GoTo DiaErr1
    'RdoQry(0) = strSftCode
    cmdObj.Parameters(0).Value = strSftCode
    bSqlRows = clsADOCon.GetQuerySet(RdoCde, cmdObj, ES_KEYSET, True)
    If bSqlRows Then
        With RdoCde
            cmbCde = "" & Trim(!SFCODE)
            txtDsc = "" & Trim(!SFDESC)
                        
            txtBeg = "" & Trim(!SFSTHR)
            txtEnd = "" & Trim(!SFENHR)
            txtLBeg = "" & Trim(!SFLUNSTHR)
            txtLEnd = "" & Trim(!SFLUNENHR)
            ' Get the total hours
            CalculateLunchTime
            txtAdj = "" & IIf(IsNull(Trim(!SFADJHR)), "0", Trim(!SFADJHR))
            txtAddRate = "" & IIf(IsNull(Trim(!SFADDRT)), "0.00", Trim(!SFADDRT))
      End With
      GetShiftCode = True
   Else
      txtDsc = ""
      GetShiftCode = False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "GetShiftCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddShiftCode()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim strShiftCode As String
   
   strShiftCode = Format(Compress(cmbCde), "00")
   
   sMsg = strShiftCode & " Wasn't Found. Add The Shift Code?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO SfcdTable (SFREF, SFCODE) " _
             & "VALUES('" & strShiftCode & "','" & strShiftCode & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected Then
         SysMsg "Shift Code Added.", True
         txtEnd = ""
         txtBeg = ""
         txtDsc = ""
         txtHrs = ""
         txtAdj = ""
         txtAddRate = ""
         cmbCde = strShiftCode
         AddComboStr cmbCde.hwnd, strShiftCode
         bGoodCode = GetShiftCode()
         On Error Resume Next
         txtDsc.SetFocus
      Else
         MsgBox "Couldn't The Add Shift Code.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "AddShiftCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
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

Private Sub CalculateLunchTime()
   On Error GoTo DiaErr1
   Dim lHrs As Currency
   Dim lLunHrs As Currency
   Dim minutes As Integer
   Dim bResponse As Byte
   Dim sMsg As String
   Dim strBegDate As String
   Dim strEndDate As String
   '
   
   If (Trim(txtBeg) <> "" And Trim(txtEnd) <> "") Then
      If (CDate(txtBeg) > CDate(txtEnd)) Then
         'sMsg = "Do you like to have shift overlap to next day ?"
         'bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         'If bResponse = vbYes Then
            strBegDate = Format(Now, "mm/dd/yy ") & txtBeg
            strEndDate = Format(DateAdd("d", 1, Now), "mm/dd/yy ") & txtEnd
            
            minutes = DateDiff("n", CDate(strBegDate), CDate(strEndDate))
            lHrs = Format(minutes / 60, "##0.00")
                     
         'End If
      Else
         minutes = DateDiff("n", txtBeg, txtEnd)
         lHrs = Format(minutes / 60, "##0.00")
      End If
   Else
       lHrs = 0
   End If
   
   If (Trim(txtLBeg) <> "" And Trim(txtLEnd) <> "") Then
       minutes = DateDiff("n", txtLBeg, txtLEnd)
       lLunHrs = Format(minutes / 60, "##0.00")
   Else
       lLunHrs = 0
   End If
   ' Get the lunch hours
   txtHrs = Format((lHrs - lLunHrs), "##0.00")

   Exit Sub
DiaErr1:
   sProcName = "CalculateLunchTime"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
