Attribute VB_Name = "TerryCrystal"
Option Explicit

'Public Sub GenerateCrystalReport(frm As Form, Crw As CrystalReport)
'   'Allow Crystal to catch up. Especially to refresh DAO.
'   'A bug in Crystal that does not propertly refresh data until the
'   'second hit.
'    If Crw.DataFiles(1) <> "" Then Sleep 1000
'
'    'set report size
'
'    'set report size
'    Crw.WindowState = crptMaximized
'
'    'other settings
'    Crw.WindowBorderStyle = crptSizable
'    Crw.WindowControlBox = True
'    Crw.WindowMaxButton = True
'    Crw.WindowMinButton = True
'    Crw.WindowShowCancelBtn = True
'    Crw.WindowShowCloseBtn = True
'    Crw.WindowShowExportBtn = True
'    Crw.WindowShowGroupTree = False
'    Crw.WindowShowNavigationCtls = True
'    Crw.WindowShowPrintBtn = True
'    Crw.WindowShowPrintSetupBtn = True
'    Crw.WindowShowRefreshBtn = True
'    Crw.WindowShowZoomCtl = True
'    Crw.WindowShowSearchBtn = False
'
'    'set connection
'    Crw.Connect = "uid=" & gstrSaAdmin & ";pwd=" & gstrSaPassword & ";driver={SQL Server};" _
'                & "server=" & gstrServer & ";database=" & gstrDatabase & ";"
'
'    Crw.ReportTitle = frm.Caption
'    Crw.WindowTitle = frm.Caption
'    If frm.optPrn.Value = True Then
'        Crw.Destination = crptToPrinter
'    Else
'        Crw.Destination = crptToWindow
'    End If
'    On Error Resume Next
'    Err.Clear
'    Crw.Action = 1
'    If Err Then
'        MsgBox Err.Description
'    End If
'
'End Sub
'
'
