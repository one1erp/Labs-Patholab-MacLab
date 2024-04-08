VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniMngr"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As Any, _
   ByVal lpDefault As String, _
   ByVal lpReturnedString As String, _
   ByVal nSize As Long, _
   ByVal lpFileName As String) As Long

Public Function GetIniValue(pstIniFile As String, _
                            pstSectionName As String, _
                            pstKeyName As String, _
                            Optional pstDefaultValue As String = "") As String

    GetIniValue = ProfileGetItem(pstSectionName, _
                                 pstKeyName, _
                                 pstDefaultValue, _
                                 pstIniFile)

End Function

Private Function ProfileGetItem(lpSectionName As String, _
                                lpKeyName As String, _
                                DefaultValue As String, _
                                inifile As String) As String

        
   Dim success As Long
   Dim nSize As Long
   Dim ret As String
  
   ret = Space$(2048)
   nSize = Len(ret)
   success = GetPrivateProfileString(lpSectionName, _
                                     lpKeyName, _
                                     DefaultValue, _
                                     ret, _
                                     nSize, _
                                     inifile)
   
   If success Then
      ProfileGetItem = Left$(ret, success)
   End If
   
End Function
