VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form SSRSViewer 
   Caption         =   "SSRS Viewer"
   ClientHeight    =   11505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16770
   LinkTopic       =   "Form1"
   ScaleHeight     =   11505
   ScaleWidth      =   16770
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   40000
      ExtentX         =   70556
      ExtentY         =   20135
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "SSRSViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Path As String
Public report As String
Public URL As String

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    Dim rdo As ADODB.Recordset
    Dim SQL As String
    SQL = "SELECT top 1 SSRSFolderUrl from SSRSInfo"
    bSqlRows = clsADOCon.GetDataSet(SQL, rdo, ES_FORWARD)
    If bSqlRows Then
        Path = rdo(0)
        If Mid(Path, Len(Path), 1) <> "/" Then
            Path = Path & "/"
        End If
        'WebBrowser1.Navigate Path + report & "&rs:Command=Render&User=" & sInitials
        If InStr(1, report, "MRP Actions by Part Number", vbTextCompare) > 0 Then
            WebBrowser1.Navigate Path + report & "&User=" & sInitials
         Else
            WebBrowser1.Navigate Path + report
         End If
         
'         Do Until WebBrowser1.readyState = READYSTATE_COMPLETE
'            DoEvents
'         Loop
'
'        WebBrowser1.Document.body.Scroll = "no"
    Else
        MsgBox "SSRS Server does not exist or has not been defined."
    End If
    
    Set rdo = Nothing
End Sub

Private Sub Form_Resize()
    WebBrowser1.Top = 0
    WebBrowser1.Left = 0
    WebBrowser1.Width = Me.Width - 100    'leave room for scrollbars
    WebBrowser1.Height = Me.Height - 500
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
        'WebBrowser1.Height = WebBrowser1.Height - 5000
End Sub

