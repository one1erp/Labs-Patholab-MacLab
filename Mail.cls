VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private mstComputer As String

Public SendPrg As String
Public Message As String
Public inifile As String

Public Enum SeverityTypes
    sevNormal = 0
    sevHigh = 1
    sevLow = 2
End Enum

Public Severity As SeverityTypes

Public Function Send() As Boolean
    
    Dim lMailMsg As MailMessage
    Dim lIni As IniMngr
    Dim lCon As ADODB.Connection
    Dim lRs As ADODB.Recordset
    Dim lMailList As Dictionary

    Dim lstOraCon As String
    Dim SQL As String
    
    Dim lstFromEmail As String
    Dim lstFromName As String
    Dim lstMailServer As String
    Dim lstSubject  As String
    Dim lstMessage As String
    Dim lstToEmail As String
    Dim lstToName As String
    Dim lstSection As String
    Dim lstIniFile As String
    
    Dim i As Integer
    
    On Error GoTo ErrHnd
    
    Set lMailList = New Dictionary
    Set lIni = New IniMngr
    lstIniFile = IIf(inifile <> "", inifile, CurDir & "\Mail.ini")
    lstFromEmail = lIni.GetIniValue(lstIniFile, "Mail_Details", "FromEmail", "limserver@mac.org.il")
    lstFromName = lIni.GetIniValue(lstIniFile, "Mail_Details", "FromName", "limserver")
    lstMailServer = lIni.GetIniValue(lstIniFile, "Mail_Details", "MailServer", "10.74.1.223")
    lstOraCon = lIni.GetIniValue(lstIniFile, "Mail_Details", "OraCon", "")
    
    lstSubject = GetMsg("SUBJECT")
    lstMessage = GetMsg("MESSAGE")
    
    Set lMailMsg = New MailMessage
    With lMailMsg
        .FromEmail = lstFromEmail
        .FromName = lstFromName
        .MailServer = lstMailServer
        .Subject = lstSubject
        .Message = lstMessage
    End With
    
    On Error GoTo NoOracle
    Set lCon = New ADODB.Connection
    lCon.Open lstOraCon
    SQL = "select PE.PHRASE_NAME Name, PE.PHRASE_DESCRIPTION Email" & vbCrLf & _
          "from  PHRASE_ENTRY PE, PHRASE_HEADER PH" & vbCrLf & _
          "where PH.PHRASE_ID = PE.PHRASE_ID" & vbCrLf & _
          "  and upper(PH.NAME) = 'MAIL ADDRESS'"
    Set lRs = lCon.Execute(SQL)
    
    While Not lRs.EOF
        lMailList.Add lRs("Name"), lRs("Email")
        lRs.MoveNext
    Wend
    
    lRs.Close
    Set lRs = Nothing
    lCon.Close
    Set lCon = Nothing
GoTo Send
NoOracle:
    On Error GoTo ErrHnd
    i = 1
    lstSection = "SendTo " & i
    lstToName = lIni.GetIniValue(lstIniFile, lstSection, "Name", "")
    While lstToName <> ""
        lstToEmail = lIni.GetIniValue(lstIniFile, lstSection, "Email", "")
        lMailList.Add lstToName, lstToEmail
        i = i + 1
        lstSection = "SendTo " & i
        lstToName = lIni.GetIniValue(lstIniFile, lstSection, "Name", "")
    Wend
Send:
    On Error GoTo ErrHnd
    For i = 0 To lMailList.Count - 1
        lstToName = lMailList.Keys(i)
        lstToEmail = lMailList(lstToName)
'        MsgBox "Sending to " & lstToEmail
        With lMailMsg
            .ToEMail = lstToEmail
            .ToName = lstToName
            .SendEmail
        End With
    Next
        
    Set lMailList = Nothing
    Set lMailMsg = Nothing
Exit Function
ErrHnd:
    Dim lLog As New LogMngr
    Dim lstErr As String
    lstErr = Err.Description
    lLog.LogPath = "C:\Mail.Log"
    lLog.WriteLog "Error sending mail, " & lstErr
End Function

Private Function GetMsg(pType As String) As String
    
    Dim lstComputerName As String
    Dim lloStrLen As Long
    
    Dim lstSubject As String
    Dim lstMessage As String
    On Error GoTo ErrHnd
    lloStrLen = 50
    lstComputerName = Space(50)
    Call GetComputerName(lstComputerName, lloStrLen)
    lstComputerName = UCase(Left(lstComputerName, lloStrLen))

    lstMessage = "Hello," & vbCrLf & vbCrLf
    
    Select Case Severity
    Case Severity = sevHigh
        lstMessage = lstMessage & "Severity : HIGH" & vbCrLf
        lstSubject = "HIGH SEVERITY !!! "
    Case Severity = sevNormal
        lstMessage = lstMessage & "Severity : NORMAL" & vbCrLf
    Case Severity = sevLow
        lstMessage = lstMessage & "Severity : LOW" & vbCrLf
        lstSubject = "Low Severity. "
    End Select

    lstMessage = lstMessage & "Time     : " & Format(Now, "dd/mm/yyyy hh:MM") & vbCrLf
    lstMessage = lstMessage & "Computer : " & lstComputerName & vbCrLf
    lstMessage = lstMessage & "Program  : " & SendPrg & vbCrLf
    lstMessage = lstMessage & String(20, "*") & vbCrLf
    lstMessage = lstMessage & "Message  : " & vbCrLf & Message & vbCrLf & vbCrLf
    lstMessage = lstMessage & String(20, "*") & vbCrLf
    lstMessage = lstMessage & "Thank you for your support !"
    lstSubject = lstSubject & SendPrg & " at " & lstComputerName & " : " & Left(Message, 15) & "..."
    
    Select Case UCase(pType)
    Case "SUBJECT"
        GetMsg = lstSubject
    Case Else
        GetMsg = lstMessage
    End Select
Exit Function
ErrHnd:
    GetMsg = "Error in Building Mail Msg : " & vbCrLf & Err.Description
End Function

Private Sub Class_Initialize()
    Severity = sevNormal
End Sub
