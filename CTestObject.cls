VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTestObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class CTestObject
'©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
'License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc

Option Explicit

Private m_Parent As Object

Public Function alert(msg) 'Provide functionality from this object..
    alert = MsgBox(msg)
    Dbg.WrLn "Function ""alert"" called from Test Object with message:" & vbCrLf & vbTab & msg
End Function

Public Function Throw(msg As String) 'Just a catalyst to show we can access script
    m_Parent.Event_TestObj msg
End Function

Public Property Set Parent(p_Parent As Object) 'We need a reference to the script to do stuff
    Set m_Parent = p_Parent
End Property

Public Property Get Parent() As Object
    Set Parent = m_Parent
End Property
