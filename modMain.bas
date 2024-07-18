Attribute VB_Name = "modMain"
'Proposed Branch Assigment of Employees....
'C001-Mobitek Dagupan
'GCC1-Guanzon Corporate Center
'GCO1-Guanzon Central Office
'GMO1-Guanzon Manila Office
'M001-GMC Dagupan - Honda
'M029-UEMI Roosevelt - Suzuki

Option Explicit

Private Const pxeMODULENAME = "modMain"

Public oApp As clsAppDriver
Public oReport As CRAXDRT.Report
Public oRepApp As New CRAXDRT.Application

Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Enum adReport
   ViewReport = 0
   PrintReport = 1
End Enum

Private Sub Main()
'   Dim lsCommand As String
'   Dim lasParam() As String
'
'   On Error GoTo errProc
'
'   lsCommand = Command()
'   lasParam = Split(lsCommand)
'
'   Set oApp = New clsAppDriver
'   If oApp.LoadEnv(lasParam(0), lasParam(1)) = False Then Exit Sub
'
'   Set oApp.mdiMain = mdiMain
'   mdiMain.Caption = oApp.ProductName
'   mdiMain.Show
'
'endProc:
'   Exit Sub
'errProc:
'   MsgBox "Line No:" & Erl & vbCrLf & Err.Description, vbCritical, "Error"
'   End

   Set oApp = New clsAppDriver

   If oApp.LoadEnv("PetMgr") = False Then
      Exit Sub
   End If

   If oApp.LogIn("PetMgr") = False Then
      Exit Sub
   End If

   Set oApp.mdiMain = mdiMain
   mdiMain.Caption = oApp.ProductName
   mdiMain.Show
End Sub

Public Sub SetNextFocus()
   keybd_event &H9, 0, 0, 0
   keybd_event &H9, 0, &H2, 0
End Sub

Public Sub SetPreviousFocus()
   keybd_event &H10, 0, 0, 0
   keybd_event &H9, 0, 0, 0
   keybd_event &H10, 0, &H2, 0
End Sub

Public Sub CenterChildForm(frmMDIForm As MDIForm, frmChild As Form)
   Dim lbX As Long, lbY As Long
   
   lbX = frmMDIForm.ScaleWidth
   lbY = frmMDIForm.ScaleHeight
   
   frmChild.Left = CLng((lbX - frmChild.Width) / 2)
   frmChild.Top = CLng((lbY - frmChild.Height) / 2)
   
   If frmChild.Left < 0 Then frmChild.Left = 0
   If frmChild.Top < 0 Then frmChild.Top = 0
End Sub

Public Sub setGrayText(ByVal lnColor As Long)
   SetSysColors 1, 17, lnColor
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
         End
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub

Public Function TransStat(nStat As Integer) As String
   Select Case nStat
   Case 0
      TransStat = "OPEN"
   Case 1
      TransStat = "CLOSED"
   Case 2
      TransStat = "POSTED"
   Case 3
      TransStat = "CANCELLED"
   Case 4
      TransStat = "UNKNOWN"
   End Select
End Function

Public Function getCTime(ByVal sTime) As String
   Const sALLOWEDCHAR As String = "0123456789"
   
   Dim sAllowed As String
   Dim sChar As String
   Dim retVal As String
   Dim sTTemp As String
   Dim sHH As String
   Dim sMM As String
   Dim sExt As String
   
   Dim lnLen As Integer
   Dim lnExt As Integer
   Dim lnCtr As Integer

   If sTime = "" Then GoTo endProc
   'add brakets to string for using LIKE
   sAllowed = "[" & sALLOWEDCHAR & "]"
   'get time
   sTime = LCase(Replace(sTime, " ", ""))
   'get length
   lnLen = Len(sTime)
   'check the length; maximum 7
   If lnLen < 4 Or lnLen > 7 Then GoTo endProc
   
   'get extension position
   If InStr(sTime, "a") > 4 Then
      lnExt = InStr(sTime, "a") - 1
   ElseIf InStr(sTime, "p") > 4 Then
      lnExt = InStr(sTime, "p") - 1
   Else
      lnExt = lnLen
   End If
   
   'set time to temp
   sTTemp = Left(sTime, lnExt)
   'get extension
   sExt = Right(sTime, lnLen - lnExt)
   
   'get the minutes
   sMM = Right(sTTemp, 2)
   If Not IsNumeric(sMM) Then Exit Function
'   sMM = Mid(sTTemp, 4, 2)
   
   'set the hour to temp
   sTTemp = Left(sTTemp, Len(sTTemp) - 2)
'   sTTemp = Left(sTTemp, 2)
   'Now loop through all characters in the string removing all unwanted charaters
   For lnCtr = 1 To Len(sTTemp)
       sChar = Mid$(sTTemp, lnCtr, 1)
       If sChar Like sAllowed Then
           retVal = retVal & sChar
       End If
   Next
   'set hour
   sHH = retVal
   If Not IsNumeric(sHH) Then Exit Function
   
   getCTime = Format(sHH & ":" & sMM & sExt, "HH:MM AM/PM")
   
endProc:
   Exit Function
End Function

Function getEmpType() As String
   Dim loFrm As frmEmpTypeMod
   
   Set loFrm = New frmEmpTypeMod
   Set loFrm.AppDriver = oApp
   
   loFrm.Show 1
   
   If Not loFrm.Cancelled Then
      If loFrm.optField(0).Value = True Then
         getEmpType = "T" & loFrm.chkField.Value
      ElseIf loFrm.optField(1).Value = True Then
         getEmpType = "R" & loFrm.chkField.Value
      Else
         getEmpType = "A" & loFrm.chkField.Value
      End If
   End If
   
   Set loFrm = Nothing
End Function

'MAC(01-26-12)
Function strShortDate(ByVal Value As String) As String
   strShortDate = Format(Value, "MM-DD-YYYY")
End Function

Function strLongDate(ByVal Value As String) As String
   strLongDate = Format(Value, "MMMM DD, YYYY")
End Function

Function strCurrency(ByVal Value As String) As String
   strCurrency = Format(Value, "#,##0.00")
End Function

Sub SetGridRowColor(ByVal loGrid As MSFlexGrid, _
                     ByVal lnMode As Integer, _
                     ByVal lnCol As Integer, _
                     Optional ByVal lnRow As Integer = 0)
   Dim lnCtr As Integer

   Select Case lnMode
      Case 0 ' full
         With loGrid
            For lnCtr = 1 To .Rows - 1
               If lnCtr Mod 2 = 0 Then
                  .FillStyle = flexFillRepeat
                  .Row = lnCtr
                  .RowSel = lnCtr
                  .Col = lnCol
                  .ColSel = .Cols - 1
                  .CellBackColor = &HFFC0FF
               End If
            Next
            .Row = .Rows - 1
         End With
      Case 1 ' single
         With loGrid
            If IsMissing(lnRow) Then Exit Sub
            If lnRow = 0 Or lnRow Mod 2 = 1 Then Exit Sub

            .FillStyle = flexFillRepeat
            .Row = lnRow
            .RowSel = lnRow
            .Col = lnCol
            .ColSel = .Cols - 1
            .CellBackColor = &HFFC0FF
            .Row = .Rows - 1
         End With
   End Select
End Sub

Function WhoIs(ByVal fsID As String, Optional ByVal fbCypher As Boolean = False) As String
   Dim lsSQL As String
   Dim lors As Recordset

   If fbCypher Then
      fsID = Decrypt(fsID)
   End If

   lsSQL = "SELECT sUserName" & _
          " FROM xxxSysUser" & _
          " WHERE sUserIDxx = " & strParm(fsID)
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)

   If lors.EOF Then
      WhoIs = "N-O-N-E"
   Else
      WhoIs = Decrypt(lors("sUserName"), oApp.Machinex)
   End If

   Set lors = Nothing
End Function

Function testBatchShift()
   Dim loTrans As clsLogManual2
   Set loTrans = New clsLogManual2
   Set loTrans.AppDriver = oApp
   loTrans.Branch = oApp.BranchCode
   
   If Not loTrans.InitTransaction Then
      MsgBox "Unable to initialize manual log."
      End
   End If
   
   MsgBox "Ola"
End Function

Public Sub sendMacroPayrollValidation(fsEmpTypID As String, fdPeriodFr As Date)
   Dim lsSQL As String
   Dim lors As Recordset
   
   Dim loMac As ggcPETValidator.clsPayrollMacro
   Set loMac = New ggcPETValidator.clsPayrollMacro
   Set loMac.AppDriver = oApp
   
   If Not loMac.InitTransaction Then
      MsgBox "Error Initializing Payroll Macro Validation"
      Exit Sub
   End If
      
   Set lors = loMac.getDivisions()
   
   Do Until lors.EOF
      If Not loMac.NewTransaction(fsEmpTypID, lors("cPayDivCD"), lors("cDivision"), fdPeriodFr) Then
         MsgBox "Error Creating Payroll Macro Validation for " & lors("sDivsnDsc")
         Exit Sub
      End If
      
      If Not loMac.SaveTransaction Then
         MsgBox "Error Saving Payroll Macro Validation for " & lors("sDivsnDsc")
         Exit Sub
      End If
         
      If loMac.send2Validator() Then
         MsgBox "Macro Validation for " & lors("sDivsnDsc") & " was sent successfully!"
      Else
         MsgBox "Macro Validation for " & lors("sDivsnDsc") & " was not sent!"
      End If
      
      lors.MoveNext
   Loop
End Sub
