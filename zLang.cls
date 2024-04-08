VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zLang"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const KLF_REORDER = &H8
Private Const lang_Hebrew = 67961869
Private Const lang_English = 67699721

Private loCurrLan As Long

Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long

Public Sub English()
    ' שינוי השפה הפעילה לאנגלית
    Call ActivateKeyboardLayout(lang_English, KLF_REORDER)
End Sub

Public Sub Hebrew()
    ' שינוי השפה הפעילה לעברית
    Call ActivateKeyboardLayout(lang_Hebrew, KLF_REORDER)
End Sub

Public Sub OtherLang(ploLang As Long)
    ' שינוי לכל שפה אחרת
    Call ActivateKeyboardLayout(ploLang, KLF_REORDER)
End Sub

Public Sub SaveOrigLang()
    loCurrLan = GetKeyboardLayout(0)
End Sub

Public Sub SetOrigLang()
    Call OtherLang(loCurrLan)
End Sub

Private Sub Class_Initialize()
    loCurrLan = GetKeyboardLayout(0)
End Sub

Private Sub Class_Terminate()
    Call OtherLang(loCurrLan)
End Sub
