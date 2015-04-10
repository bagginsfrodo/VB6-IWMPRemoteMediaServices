VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CdbgWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class CdbgWrite
'©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
'License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc

'Little Debugging class hastily whipped up
Option Explicit

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const SBBOTTOM As Long = 7
Private Const WMVSCROLL As Long = &H115


Public MyTextBox As TextBox
Public PrintDebug As Boolean

Public Sub Wr(msg, Optional addVBCRLF As Boolean = False)
    If Not MyTextBox Is Nothing Then
        MyTextBox.Text = MyTextBox.Text & msg & IIf(addVBCRLF, vbCrLf, vbNullString)
        PostMessage MyTextBox.hwnd, WMVSCROLL, SBBOTTOM, 0
    Else
        Debug.Print "---------------->Someone forgot to set my textbox....."
    End If

    If PrintDebug = True Then
        Debug.Print msg
    End If
End Sub

Public Sub WrLn(ParamArray msg() As Variant) 'msg As String)
Dim x As Long
    For x = LBound(msg) To UBound(msg)
        Wr msg(x), True
    Next x
End Sub



