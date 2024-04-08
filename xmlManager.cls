VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmlManager"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const REPLACE_CHAR = "_"

Private pXmlFolder As String
Private file_machine As New FileSystemObject

Public Sub SaveXmlFile(xmlDoc As DOMDocument, xmlFileName As String)
    On Error GoTo ErrEnd
    Dim fName As String
    Dim fCount As Integer

    If UCase(Right(xmlFileName, 4)) = ".XML" Then
        xmlFileName = Left(xmlFileName, Len(xmlFileName) - 4)
    End If

    xmlFileName = Replace(xmlFileName, "\", REPLACE_CHAR)
    xmlFileName = Replace(xmlFileName, "/", REPLACE_CHAR)
    xmlFileName = Replace(xmlFileName, ":", REPLACE_CHAR)
    xmlFileName = Replace(xmlFileName, "*", REPLACE_CHAR)
    xmlFileName = Replace(xmlFileName, "?", REPLACE_CHAR)
    xmlFileName = Replace(xmlFileName, """", REPLACE_CHAR)
    xmlFileName = Replace(xmlFileName, "<", REPLACE_CHAR)
    xmlFileName = Replace(xmlFileName, ">", REPLACE_CHAR)
    xmlFileName = Replace(xmlFileName, "|", REPLACE_CHAR)

    fName = pXmlFolder & _
            "\" & _
            xmlFileName & _
            "_" & _
            Format(Now, "yyyymmddHhNnSs") & _
            ".xml"

    fCount = 1
    While file_machine.FileExists(fName)
        fName = pXmlFolder & _
                "\" & _
                xmlFileName & _
                "_" & _
                Format(Now, "yyyymmddHhNnSs") & _
                "_" & _
                fCount & _
                ".xml"
        fCount = fCount + 1
    Wend
    xmlDoc.save (fName)
    Exit Sub
ErrEnd:
End Sub

Private Sub Class_Initialize()
    pXmlFolder = "C:\"
End Sub

Public Property Let XmlFolder(FolderName As String)
    On Error GoTo ErrEnd
    pXmlFolder = FolderName
    If Not file_machine.FolderExists(FolderName) Then
        file_machine.CreateFolder (pXmlFolder)
    End If
ErrEnd:
End Property

Public Function GetDefaultFolderFromWorkStation(WorkStationID As String, _
                                                Con As ADODB.Connection) As String
    On Error GoTo ErrEnd
    Dim RsWorkStation As ADODB.Recordset
    Dim strSQL As String

    GetDefaultFolderFromWorkStation = ""
    strSQL = "select description " & _
             "from lims_sys.workstation " & _
             "where workstation_id = '" & WorkStationID & "'"
    Set RsWorkStation = Con.Execute(strSQL)
    If Not RsWorkStation.EOF Then
        GetDefaultFolderFromWorkStation = Trim(nte(RsWorkStation(0)))
    End If
    Exit Function
ErrEnd:
End Function

Private Function nte(e As Variant) As Variant
    nte = IIf(IsNull(e), "", e)
End Function

