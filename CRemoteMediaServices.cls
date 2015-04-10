VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRemoteMediaServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class CRemoteMediaServices
'©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
'License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc

Option Explicit

'https://msdn.microsoft.com/en-us/library/windows/desktop/dd563634%28v=vs.85%29.aspx
Implements RemoteWMP.IWMPRemoteMediaServices


Private m_SkinFile As String                    'Our skin file we are telling it to load, prepare to handle errors if something is wrong with it !
Private m_ScriptableObject As CScriptableObject 'The object we are passing to the skin as a scriptable object


Private Sub IWMPRemoteMediaServices_GetApplicationName(pbstrName As String)
    'We don't need this here. read the docs
    Err.Raise E_NOTIMPL
End Sub

Private Sub IWMPRemoteMediaServices_GetCustomUIMode(pbstrFile As String)
    'A WTF moment. The skin file needs to have "file://" in front, took a while to figure that one out, still don't know where I found it
    pbstrFile = "file://" & m_SkinFile
End Sub

Private Sub IWMPRemoteMediaServices_GetScriptableObject(pbstrName As String, ppDispatch As Object)
    pbstrName = "MyScriptableObject"  'This is the name your object will be referred to in your skin
    Set ppDispatch = m_ScriptableObject 'This is the actual object you will be supplying
End Sub

Private Sub IWMPRemoteMediaServices_GetServiceType(pbstrType As String)
            'pbstrType = "Remote" 'If you want to read the docs, this stuff works too, although I won't be showing it here
    
    
    pbstrType = "Local" 'We tell it we are local, we are not interested in Remoting the control in this project
End Sub




'The object we are passing to the skin as a scriptable object
Public Property Set ScriptObject(scrObj As CScriptableObject)
    Set m_ScriptableObject = scrObj
End Property
'The object we are passing to the skin as a scriptable object
Public Property Get ScriptObject() As CScriptableObject
    Set ScriptObject = m_ScriptableObject
End Property


'Our skin file we are telling it to load, prepare to handle errors if something is wrong with it !
Public Property Let SkinFile(pSkin As String)
    m_SkinFile = pSkin
End Property
'Our skin file we are telling it to load, prepare to handle errors if something is wrong with it !
Public Property Get SkinFile() As String
    SkinFile = m_SkinFile
End Property

