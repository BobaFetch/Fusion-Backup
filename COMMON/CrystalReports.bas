Attribute VB_Name = "CrystalReports"
Option Explicit

'Public bBold As Byte
'Public sProcName As String
'Public sReportPath As String
'Public bNoCrystal As Boolean
'Public iZoomLevel As Integer
'Public bUserAction As Boolean
'Public iBarOnTop As Byte


Function CrystalParameterString(CallingForm As Form, Optional Crw As CrystalReport)
   'construct a string to display or print crystal parameters
   
   Dim cr As CrystalReport
   If Crw Is Nothing Then
      Set cr = MDISect.Crw
   Else
      Set cr = Crw
   End If
   
   Dim s As String
   s = "Crystal Reports parameters for report " & cr.ReportFileName & " called from form " & CallingForm.Name & vbCrLf
   Dim j As Integer
   For j = 0 To 20
      If Len(cr.Formulas(j)) > 0 Then
         s = s & "(" & j & ") " & cr.Formulas(j) & vbCrLf
      End If
   Next
   
   'now add SectionFormats
   Dim formatsFound As Boolean
   For j = 0 To 10
      If cr.SectionFormat(j) <> "" Then
         If Not formatsFound Then
            formatsFound = True
            s = s & "Section Formats:" & vbCrLf
         End If
         s = s & "(" & j & ") " & cr.SectionFormat(j) & vbCrLf
      End If
   Next
   
   'now add StoredProcParam
   Dim spParamsFound As Boolean
   For j = 0 To 10
      If cr.StoredProcParam(j) <> "" Then
         If Not spParamsFound Then
            spParamsFound = True
            s = s & "Stored Procedure Parameters:" & vbCrLf
         End If
         s = s & "(" & j & ") " & cr.StoredProcParam(j) & vbCrLf
      End If
   Next
   
   s = s & "SQL: " & cr.SelectionFormula & vbCrLf
   CrystalParameterString = s
End Function

Public Sub SetCrystalAction(frm As Form)
   'fires crystal and sets zoom level
   'if user has selected one
   Dim b As Byte
   Dim bInstr As Byte
   Dim FormDriver As String
   Dim FormPort As String
   Dim FormPrinter As String
   
   MouseCursor ccHourglass
   
   On Error GoTo modErr1
   
   ' if debugging, show info
   If Debugging() Then
      Select Case MsgBox(CrystalParameterString(frm) & vbCrLf & "Proceed?", vbYesNo + vbQuestion)
      Case vbNo
         On Error Resume Next          'there are not always buttons with these names
         frm.optPrn.Enabled = True
         frm.optDis.Enabled = True
         MouseCursor ccDefault
         MDISect.Crw.Reset
         Exit Sub
      End Select
   End If
   
   '1/13/04 Allow Crystal to catch up. Especially to refresh DAO.
   'A bug in Crystal that does not propertly refresh data until the
   'second hit.
   If MDISect.Crw.DataFiles(1) <> "" Then Sleep 500
   MDISect.Crw.ReportTitle = frm.Caption
   MDISect.Crw.WindowTitle = frm.Caption
   If frm.optPrn.value = True Or frm.optPrn.value = vbChecked Then
      On Error Resume Next
      FormPrinter = Trim(frm.lblPrinter)
      If Err > 0 Then FormPrinter = ""
      If FormPrinter = "Default Printer" Then FormPrinter = ""
      If Not bBold Then
         MDISect.Crw.SectionFont(0) = "ALL;;;;N"
      Else
         MDISect.Crw.SectionFont(0) = "ALL;;;;Y"
      End If
      If Len(Trim(FormPrinter)) > 0 Then
         b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
      Else
         FormPrinter = ""
         FormDriver = ""
         FormPort = ""
      End If
      MDISect.Crw.PrinterName = FormPrinter
      MDISect.Crw.PrinterDriver = FormDriver
      MDISect.Crw.PrinterPort = FormPort
      MDISect.Crw.Destination = crptToPrinter
   Else
      MDISect.Crw.Destination = crptToWindow
   End If
   On Error Resume Next
   Err.Clear
   MDISect.Crw.Action = 1
   sDsn = "ESI2000" 'MM TODO'RegisterSqlDsn("ESI2000")
   If Err = 20599 Then
      'Crystal didn't like the DSN, try again
      'using the default (make one if necessary).
      'If it still fails then something else is amiss.
      MDISect.Crw.Connect = "DSN=" & sDsn & ";UID=" & sSaAdmin & ";PWD=" _
                            & sSaPassword & ";DSQ=" & sDataBase & " "
      'SaveSetting "Esi2000", "System", "SqlDsn", sDsn
      SaveUserSetting USERSETTING_SqlDsn, sDsn
      MDISect.Crw.Action = 1
   Else
      If Err > 0 Then
         'any other errors
         '8/19/05 find the report and do it twice
         CurrError.Number = Err.Number
         CurrError.Description = Err.Description
         sProcName = MDISect.Crw.ReportFileName
         sProcName = Left(sProcName, Len(sProcName) - 4)
         bInstr = InStr(4, sProcName, "\")
         sProcName = Right$(sProcName, Len(sProcName) - bInstr)
         bInstr = InStr(4, sProcName, "\")
         sProcName = Right$(sProcName, Len(sProcName) - bInstr)
         GoTo modErr2
      End If
   End If
   If frm.optPrn.value = False Then
      'Allow for the bug in Crystal that shows the first
      'form full screen
      If bNoCrystal Then
         SendKeys "% R", True
         bNoCrystal = False
      End If
      If iZoomLevel > 0 Then MDISect.Crw.PageZoom (iZoomLevel)
   End If
   MDISect.Crw.DataFiles(1) = ""
   MDISect.Crw.Reset                '@@@ added 2/13/08
   frm.optPrn.Enabled = True
   frm.optDis.Enabled = True
   MouseCursor ccDefault
   Exit Sub
   
modErr1:
   sProcName = "SetCrystalAction"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
modErr2:
   DoModuleErrors frm
   On Error Resume Next 'following buttons may not exist
   frm.optPrn.Enabled = True
   frm.optDis.Enabled = True
   MouseCursor ccDefault
   MDISect.Crw.Reset
End Sub



'Allows selection of printers for individual reports
'Crystal requires this stuff

Public Function GetPrinterPort(devPrinter As String, devDriver As String, devPort As String) As Byte
   Dim SysPrinter As Printer
   For Each SysPrinter In Printers
      If Trim(SysPrinter.DeviceName) = devPrinter Then
         devDriver = SysPrinter.DriverName
         devPort = SysPrinter.Port
         Exit For
      End If
   Next
   
End Function


'3/7/05 Add DiscardSavedData = True

Public Sub SetMdiReportsize(frm As Form, Optional bAllowDrillDown As Boolean)
   'Sets report size based on monitor size
   'requires as large as possible for Win9x generic
   'monitor
   Dim bWindowSize As Byte
   Dim A As Integer
   Dim b As Integer
   
   On Error Resume Next
   sProcName = "printreport"
   frm.optPrn.Enabled = False
   frm.optDis.Enabled = False
   MDISect.Crw.Reset
   GetCrystalConnect
   sSql = ""
   'clear any report variables
   'Resolve a bug in Crystal that doesn't clear a report
   For b = 0 To 60
      MDISect.Crw.Formulas(b) = ""
      MDISect.Crw.SectionFormat(b) = ""
      MDISect.Crw.SectionFont(b) = ""
   Next
   A = Screen.TwipsPerPixelX
   b = Screen.TwipsPerPixelY
   bUserAction = True
   MDISect.Crw.WindowAllowDrillDown = bAllowDrillDown
   MDISect.Crw.DiscardSavedData = True
   bWindowSize = GetSetting("Esi2000", "System", "ReportMax", bWindowSize)
   If bWindowSize = 0 Then
      '10/24/06
      If iBarOnTop = False Then
         MDISect.Crw.WindowState = 0
         MDISect.Crw.WindowTop = 650 / b
         MDISect.Crw.WindowHeight = (MDISect.Height / b) - (1100 / b)
         MDISect.Crw.WindowLeft = 2000 / A
         MDISect.Crw.WindowWidth = (MDISect.Width / A) - (2330 / A)
      Else
         MDISect.Crw.WindowTop = 1280 / b
         MDISect.Crw.WindowHeight = (MDISect.Height / b) - (1750 / b)
         MDISect.Crw.WindowLeft = 220 / A
         MDISect.Crw.WindowWidth = (MDISect.Width / A) - (750 / A)
      End If
   Else
      MDISect.Crw.WindowState = 2
      MDISect.Crw.WindowTop = 0
      MDISect.Crw.WindowHeight = Screen.Height
      MDISect.Crw.WindowLeft = 0
      MDISect.Crw.WindowWidth = Screen.Width
   End If
   
End Sub


'Use the default ESI2000 DSN if one hasn't been registered in ES/2000

Public Sub GetCrystalConnect()
   'MdiSect.Crw.Connect = "DSN=" & sDsn & ";UID=" & sSaAdmin & ";PWD=" _
   '                      & sSaPassword & ";DSQ=" & sDataBase
   MDISect.Crw.Connect = "DSN=ESI2000;UID=" & sSaAdmin & ";PWD=" _
                         & sSaPassword & ";DSQ=" & sDataBase
   MDISect.Crw.WindowBorderStyle = crptSizable
   MDISect.Crw.WindowControlBox = True
   MDISect.Crw.WindowMaxButton = True
   MDISect.Crw.WindowMinButton = True
   MDISect.Crw.WindowShowCancelBtn = True
   MDISect.Crw.WindowShowCloseBtn = True
   MDISect.Crw.WindowShowExportBtn = True
   MDISect.Crw.WindowShowGroupTree = False
   MDISect.Crw.WindowShowNavigationCtls = True
   MDISect.Crw.WindowShowPrintBtn = True
   MDISect.Crw.WindowShowPrintSetupBtn = True
   MDISect.Crw.WindowShowRefreshBtn = True
   MDISect.Crw.WindowShowZoomCtl = True
   MDISect.Crw.WindowShowSearchBtn = False
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub



Sub GetCrystalDSN()
   'sDsn = GetSetting("Esi2000", "System", "SqlDsn", sDsn)
   sDsn = GetUserSetting(USERSETTING_SqlDsn)
   If Trim(sDsn) = "" Then
      MsgBox "GetCrystalDSN - This should never happen."
      sDsn = "ESI2000" 'RegisterSqlDsn("ESI2000")
      If Trim(sDsn) <> "" Then
         'SaveSetting "Esi2000", "System", "SqlDsn", sDsn
         SaveUserSetting USERSETTING_SqlDsn, sDsn
      End If
   End If
End Sub

Public Function CrystalDate(dt As Variant) As String
   'returns date in format suitable for Crystal Reports SQL
   CrystalDate = " Date(" & Format(dt, "yyyy,mm,dd") & ") "
End Function
