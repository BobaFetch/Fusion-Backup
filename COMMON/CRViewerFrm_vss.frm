VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Begin VB.Form CRViewerFrm 
   Caption         =   "Crystal Report Viewer"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   10965
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRViewerObj 
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      _cx             =   18441
      _cy             =   12091
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
   End
End
Attribute VB_Name = "CRViewerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CRViewerObj_CloseButtonClicked(UseDefault As Boolean)
    Unload Me
End Sub

Private Sub CRViewerObj_PrintButtonClicked(UseDefault As Boolean)
    Dim intCopies As Integer
    
    UseDefault = False
    intCopies = 3
    'report.PrintOut False, intCopies
End Sub


Private Sub Form_Load()
'Set the Report source for the Report Viewer to the Report
End Sub

Public Sub ShowReport(report As CRAXDRT.report)
    CRViewerObj.ReportSource = report
    CRViewerObj.EnableGroupTree = True
    CRViewerObj.Zoom 100
    CRViewerObj.EnableCloseButton = True
    CRViewerObj.ViewReport
End Sub
Private Sub Form_Resize()
    CRViewerObj.top = 0
    CRViewerObj.Left = 0
    CRViewerObj.Height = ScaleHeight
    CRViewerObj.Width = ScaleWidth
End Sub


