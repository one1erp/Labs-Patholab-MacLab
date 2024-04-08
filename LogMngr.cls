VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogMngr"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mloLogFile As Long
Private mstLogPath As String
Private mstLogFile As String
Private mstLogFolder As String
Private mboSyncLogFile As Boolean
Attribute mboSyncLogFile.VB_VarDescription = "Synchronize log file. True - open and close file every time. False - open at the begining and close at the end. (True)"

Public Property Let LogFile(pLogFile As String)
Attribute LogFile.VB_Description = "The name of the log file"
    mstLogFile = pLogFile
    mstLogPath = mstLogFolder & mstLogFile
End Property

Public Property Get LogFile() As String
    LogFile = mstLogFile
End Property

Public Property Let LogFolder(pLogFolder As String)
Attribute LogFolder.VB_Description = "The folder for log. created if not exists."
    mstLogFolder = pLogFolder
    If Right(mstLogFile, 1) <> "\" Then mstLogFile = mstLogFile + "\"
    mstLogPath = mstLogFolder & mstLogFile
    CreateLogFolder mstLogFolder
End Property

Public Property Get LogFolder() As String
    LogFolder = mstLogFolder
End Property

Public Property Let LogPath(pLogPath As String)
Attribute LogPath.VB_Description = "Full path of log file"
    mstLogFile = Mid(pLogPath, Len(pLogPath) - InStr(1, StrReverse(pLogPath), "\") + 2)
    mstLogFolder = Mid(pLogPath, 1, Len(pLogPath) - InStr(1, StrReverse(pLogPath), "\") + 1)
    
    If mstLogFile = "" Then
        mstLogFile = pLogPath
        mstLogFolder = "C:\"
    End If
    CreateLogFolder mstLogFolder
    mstLogPath = pLogPath
End Property

Public Property Get LogPath() As String
    LogPath = mstLogPath
End Property

Public Property Let SyncLogFile(pSyncOpt As Boolean)
Attribute SyncLogFile.VB_Description = "Synchronization of log file. True - File saved every log (Default), False - File saved at the end."
    mboSyncLogFile = pSyncOpt
    If pSyncOpt Then
        CloseFile
    Else
        OpenFile
    End If
End Property

Public Property Get SyncLogFile() As Boolean
    SyncLogFile = mboSyncLogFile
End Property

Private Sub CreateLogFolder(pFolder As String)
    On Error Resume Next
    Dim lFSO As New FileSystemObject
    If Not lFSO.FolderExists(pFolder) Then
        MkDir pFolder
    End If
End Sub

Private Sub OpenFile()
    On Error Resume Next
    If mloLogFile = 0 And mstLogFile <> "" Then
        mloLogFile = FreeFile
        Open mstLogPath For Append As #mloLogFile
    End If
End Sub

Private Sub CloseFile()
    On Error Resume Next
    If mloLogFile <> 0 Then
        Close #mloLogFile
        mloLogFile = 0
    End If
End Sub

Public Sub WriteLog(pMsg As String, Optional pPrintDateTime As Boolean = True)
Attribute WriteLog.VB_Description = "כותב את ההודעה ללוג"
    Dim lstDateTime As String
    On Error Resume Next
    If mboSyncLogFile Then OpenFile
    If mloLogFile = 0 Then Exit Sub
    
    lstDateTime = IIf(pPrintDateTime, Format(Now, "DD/MM/YY HH:MM:SS "), "")
    Print #mloLogFile, lstDateTime & pMsg
    
    If mboSyncLogFile Then CloseFile
End Sub

Private Sub Class_Initialize()
    mboSyncLogFile = True
    mloLogFile = 0
    mstLogFolder = "C:\"

End Sub

Private Sub Class_Terminate()
    CloseFile
End Sub
