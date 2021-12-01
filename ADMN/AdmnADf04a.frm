VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form AdmnADf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Logo"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog fileDialog 
      Left            =   7200
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.*"
   End
   Begin VB.CheckBox chkCompLogo 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Update database"
      Height          =   435
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Delete This Standard Comment"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Tag             =   "2"
      ToolTipText     =   "40 Characters Max"
      Top             =   600
      Width           =   3495
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3930
      FormDesignWidth =   7890
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use as Company Logo in &Reports"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   1200
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Logo"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "AdmnADf04a"
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
Dim bGoodComment As Byte
Dim bNewImage As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Const LOGO_ID As Integer = 1

'Private Declare Function CreateDirectory Lib "kernel32" (ByVal lpNewDirectory As String, lpSecurityAttributes As Any) As Long
'Private Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As Any) As Long


Private Sub chkCompLogo_Click()
    UpdateUseLogo (chkCompLogo.Value)
End Sub

Private Sub cmdAdd_Click()
    Dim bRet As Boolean
    Dim strFileName As String
    strFileName = Trim(txtFileName.Text)
    If (strFileName <> "") Then
        ' save to the database
        If (SavePictureToDB(strFileName, LOGO_ID) = True) Then
            ' Display the image
            bRet = ReadImageFromDB(strFileName, LOGO_ID)
            bNewImage = False
            MsgBox "New image saved"
        Else
            MsgBox "Picture could not be stored to the database.", _
                        vbInformation, Caption
        End If
        
    Else
         MsgBox "Select a picture File as Company Logo.", _
            vbInformation, Caption
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim strFileName
    
    ' Show file dialog
    fileDialog.ShowOpen
    If (fileDialog.FileName <> "") Then
        strFileName = fileDialog.FileName
        txtFileName.Text = strFileName
        ' Display the image
        imgLogo.Picture = LoadPicture(strFileName)
        ' Set the dirty bit update the database.
        bNewImage = True
    
    End If
    
End Sub

Private Sub Form_Activate()
    MDISect.lblBotPanel = Caption
    If bOnLoad Then
        FillLogoImage
        ' Set the check box from the company table
        GetUseLogo
        txtFileName.Text = ""
        bNewImage = False
        bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   'txtFileName.SetFocus = True
   txtFileName.Text = ""
   
   
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
   
    If (bNewImage = True) Then
        If (MsgBox("You have selected new Image for the company logo." & _
                    "Do you want save the Image.", _
                    vbYesNo, "Save Image") = vbYes) Then
            
            ' Save the picture to the database
            Dim strFileName As String
            strFileName = Trim(txtFileName.Text)
            If (strFileName <> "") Then
                SavePictureToDB strFileName, LOGO_ID
            End If
        End If
                    
    End If
    Set AdmnADf04a = Nothing
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub
Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1152
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub
Private Sub FillLogoImage()
    
    Dim sFileName As String
    Dim sFilePath As String
    Dim sFullPath As String
    Dim bRet As Boolean
    
    sFilePath = "C:\Program Files\ES2000\Temp\"
    If (CreateDirectory(sFilePath) >= 0) Then
        sFileName = "picture1.jpg"
        sFullPath = sFilePath + sFileName
        bRet = ReadImageFromDB(sFullPath, LOGO_ID)
        If (bRet = True) Then
            imgLogo.Picture = LoadPicture(sFullPath)
        End If
    End If
End Sub

Private Sub UpdateUseLogo(iChecked As Integer)
    ' Update the Use Company logo field in Company table
   sSql = "UPDATE ComnTable SET COLUSELOGO = '" & CStr(iChecked) & "'"
   clsADOCon.ExecuteSQL sSql
End Sub

Private Sub GetUseLogo()
    Dim RdoLogo As ADODB.Recordset
    Dim bRows As Boolean
    ' Assumed that COMREF is 1 all the time
    sSql = "SELECT ISNULL(COLUSELOGO, 0) as COLUSELOGO FROM ComnTable WHERE COREF = 1"
    bRows = clsADOCon.GetDataSet(sSql, RdoLogo, ES_FORWARD)

    If bRows Then
        With RdoLogo
            chkCompLogo.Value = !COLUSELOGO
        End With
        'RdoLogo.Close
        ClearResultSet RdoLogo
    End If
    Set RdoLogo = Nothing
End Sub
