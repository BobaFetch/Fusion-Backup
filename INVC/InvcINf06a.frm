VERSION 5.00
Begin VB.Form InvcINf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Obsolete Part Number"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter Part Number or Select From List (Tools Are Not Included) "
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Delete this Part Number"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.PictureBox ReSize1 
      Height          =   480
      Left            =   5400
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   11
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obsolete Part"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Close All Other Sections Before Proceeding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   7
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "InvcINf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/4/04 Omit Tools
'10/21/05 Added Null Join to query (FillCombo)
Option Explicit
Dim bOnLoad As Byte
Dim bOkToDelete As Byte
Dim bEstiTable As Byte

Dim iInvActive As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'7/9/01 See if EsiTable is here

Private Function CheckEstTable() As Byte
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT BIDREF FROM EstiTable where BIDREF>0"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then CheckEstTable = 1 Else CheckEstTable = 0
   
End Function

'2/21/02 Double check activity
'10/14/03 added escape

Private Function CheckInventory() As Integer
   Dim AdoInv As ADODB.Recordset
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   sSql = "SELECT COUNT(INPART) AS TOTALINV FROM InvaTable WHERE INPART='" _
          & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoInv, ES_FORWARD)
   If bSqlRows Then
      With AdoInv
         iList = Val("" & !TOTALINV)
         ClearResultSet AdoInv
      End With
   End If
   CheckInventory = iList
   Set AdoInv = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkinventory"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbPrt_Click()
   FindPart cmbPrt
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   FindPart cmbPrt
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   iInvActive = CheckInventory()
   If iInvActive > 1 Then
      MsgBox "This Part Has " & str$(iInvActive) & " Inventory Activities.   " & vbCr _
         & "Can't Delete The Part Number.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Len(cmbPrt) Then bOkToDelete = CheckPart()
   If bOkToDelete = 1 Then
      DeletePart
   Else
      MsgBox "Part Is In Use Or Has Been Used Can't Be Deleted...   ", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5150"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      b = CheckWindows()
      FillObsCombo
      bEstiTable = CheckEstTable()
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   lblWrn(0).ForeColor = ES_RED
   lblWrn(1).ForeColor = ES_RED
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set InvcINf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Function CheckPart() As Byte
   Dim sPart As String
   
   CheckPart = 0
   sPart = Compress(cmbPrt)
   If sPart = "" Then Exit Function
   On Error GoTo DiaErr1
   
   'Estimates
   If bEstiTable = 1 Then
      sSql = "SELECT DISTINCT BIDPART FROM EstiTable " _
             & "WHERE BIDPART='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         MsgBox "This Part Is Used On An Estimate. " _
            & vbCr & "Cannot Delete This Part Number.", _
            vbInformation, Caption
         CheckPart = 0
         Exit Function
      End If
   End If
   'Parts list
   sSql = "SELECT DISTINCT BMPARTREF FROM BmplTable " _
          & "WHERE BMPARTREF='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Parts List. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'Picks
   sSql = "SELECT DISTINCT PKMOPART FROM MopkTable " _
          & "WHERE PKMOPART='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Pick List. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   sSql = "SELECT DISTINCT PKPARTREF FROM MopkTable " _
          & "WHERE PKPARTREF='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Has A Pick List. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'PO's
   sSql = "SELECT DISTINCT PIPART FROM PoitTable " _
          & "WHERE PIPART='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Purchase Order. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'Pack Slip
   sSql = "SELECT DISTINCT PIPART FROM PsitTable " _
          & "WHERE PIPART='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Packing Slip. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'Rejections
   sSql = "SELECT DISTINCT REJPART FROM RjhdTable " _
          & "WHERE REJPART='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Rejection Tag. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'Runops
   sSql = "SELECT DISTINCT OPSERVPART FROM RnopTable " _
          & "WHERE OPSERVPART='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used As A Service Part. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'Routings
   sSql = "SELECT DISTINCT OPSERVPART FROM RtopTable " _
          & "WHERE OPSERVPART='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On An MO Service OP. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'Runs
   sSql = "SELECT DISTINCT RUNREF FROM RunsTable " _
          & "WHERE RUNREF='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Manufacturing Order. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'SO's
   sSql = "SELECT DISTINCT ITPART FROM SoitTable " _
          & "WHERE ITPART='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Sales Order. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   
   'Invoice
   sSql = "SELECT DISTINCT VITMO FROM ViitTable " _
          & "WHERE VITMO='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Purchase Order Receipt. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'Documents
   sSql = "SELECT DISTINCT DLSREF FROM DlstTable " _
          & "WHERE DLSREF='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Has A Document List. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'SPC Keys
   sSql = "SELECT DISTINCT KEYREF FROM RjkyTable " _
          & "WHERE KEYREF='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Has A Referenced SPC Key. " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   
   'Down V 2/21/02
   'Price Books
   sSql = "SELECT DISTINCT PBIPARTREF FROM PbitTable " _
          & "WHERE PBIPARTREF='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is On A Price Book... " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   
   'Alias Books
   sSql = "SELECT DISTINCT ALPARTREF FROM PaalTable " _
          & "WHERE ALPARTREF='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Has An Alias Structure... " _
         & vbCr & "Cannot Delete This Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   'Tools
   sSql = "SELECT DISTINCT TOOLLISTIT_TOOLREF FROM TlitTable " _
          & "WHERE TOOLLISTIT_TOOLREF='" & sPart & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      MsgBox "This Part Is Used On A Tool List... " _
         & vbCr & "Cannot Delete This (Tool) Part Number.", _
         vbInformation, Caption
      CheckPart = 0
      Exit Function
   End If
   CheckPart = 1
   Exit Function
   
DiaErr1:
   sProcName = "checkpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillObsCombo()
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbPrt.Clear
   sSql = "SELECT PartTable.PARTREF FROM PartTable " _
          & "WHERE PAOBSOLETE = 1" _
          & "ORDER BY PartTable.PARTREF"
   LoadComboBox cmbPrt, -1
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      FindPart cmbPrt
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Lots 3/14/02


Private Sub DeletePart()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sPart As String
   
   On Error GoTo DiaErr1
   sPart = Compress(cmbPrt)
   sMsg = "It Is Not A Good Idea To Delete A Part Number " & vbCr _
          & "If There Is Any Chance That It Is In Use Right Now."
   MsgBox sMsg, vbInformation, Caption
   
   sMsg = "This Operation Will Remove All Records Of The Part." & vbCr _
          & "Are You Certain That You Want To Delete " & cmbPrt & "."
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      '            sSql = "ALTER TABLE PartTable NOCHECK Constraint PK_PartTable_PARTREF"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE BmhdTable NOCHECK Constraint FK_BmhdTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE LohdTable NOCHECK Constraint FK_LohdTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE LoitTable NOCHECK Constraint FK_LoitTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      
      
        sSql = "ALTER TABLE MopkTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE PartTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE RnopTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE RunsTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE BmhdTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE BmplTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE PsitTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE LoitTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE LohdTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
      
      Err.Clear
      sSql = "DELETE FROM InvaTable WHERE INPART='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM BmhdTable WHERE BMHREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM LoitTable WHERE LOIPARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      'Lots
      sSql = "DELETE FROM LohdTable WHERE LOTPARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      'Tools
      sSql = "DELETE FROM TohdTable WHERE TOOL_PARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM BuypTable WHERE BYPARTNUMBER='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM CcltTable WHERE CLPARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM CcitTable WHERE CIPARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM MopkTable WHERE PKPARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM MopkTable WHERE PKMOPART='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM MrplTable WHERE MRP_PARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM PaalTable WHERE ALPARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM PsitTable WHERE PIPART='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM RndlTable WHERE RUNDLSRUNREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM RnopTable WHERE OPREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      
      sSql = "DELETE FROM BmplTable WHERE BMASSYPART='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM RnalTable WHERE RAREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM VnapTable WHERE AVPARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      ' NOT for now
      'sSql = "DELETE FROM RunsTable WHERE RUNREF='" & sPart & "' "
      'clsADOCon.ExecuteSQL sSql
      
      ' Delete part record.
      sSql = "DELETE FROM PartTable WHERE PARTREF='" & sPart & "' "
      clsADOCon.ExecuteSQL sSql
      
      
      If cUR.CurrentPart = Trim(cmbPrt) Then
         cUR.CurrentPart = ""
         SaveSetting "Esi2000", "Current", "Part", cUR.CurrentPart
      End If
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "The Current Part Number Was Deleted.", _
            vbInformation, Caption
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Could Not Delete The Current Part Number.", _
            vbInformation, Caption
      End If
      
        sSql = "ALTER TABLE MopkTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE PartTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE RnopTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE RunsTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE BmhdTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE BmplTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE PsitTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE LoitTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "ALTER TABLE LohdTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSQL sSql
        
        
      '            sSql = "ALTER TABLE PartTable CHECK Constraint PK_PartTable_PARTREF"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE BmhdTable CHECK Constraint FK_BmhdTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE LohdTable CHECK Constraint FK_LohdTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE LoitTable CHECK Constraint FK_LoitTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      cmbPrt = ""
      FillObsCombo
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "deletepar"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub


Private Function CheckWindows() As Byte
   Dim b As Byte
   b = Val(GetSetting("Esi2000", "Sections", "admn", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "prod", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "engr", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "sale", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "fina", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "qual", 0))
   If b > 0 Then
      lblWrn(0) = sSysCaption & " Has Determined " & b & " Other Open Section(s)"
      lblWrn(0).Visible = True
      lblWrn(1).Visible = True
      cmdDel.Enabled = False
   End If
   CheckWindows = b
   
End Function
