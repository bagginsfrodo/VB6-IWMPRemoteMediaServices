VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRemoteHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class CRemostHost
'©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
'License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc


Option Explicit


Implements olelib2.IServiceProvider 'https://msdn.microsoft.com/en-us/library/windows/desktop/dd564264%28v=vs.85%29.aspx
                                    'To enable remote embedding of the Windows Media Player control,
                                    'your program must implement the IServiceProvider and IWMPRemoteMediaServices
                                    'interfaces. IServiceProvider is a standard Component Object Model (COM) interface
                                    'with a single method called QueryService. Windows Media Player calls this method
                                    'to retrieve a pointer to an IWMPRemoteMediaServices interface.


Implements olelib.IOleClientSite 'We also have to be a client site, we just forward to our main form




Private Const E_NOTIMPL As Long = -2147467263 '&H80004001
Private Const E_NOINTERFACE As Long = -2147467262 '&H80004002


Private m_Container As olelib.IOleContainer
Private m_OldClientSite As olelib.IOleClientSite

Private m_Remote As RemoteWMP.IWMPRemoteMediaServices 'https://msdn.microsoft.com/en-us/library/windows/desktop/dd563634%28v=vs.85%29.aspx

Private m_WMP As WindowsMediaPlayer


'IServiceProvider Members
Private Sub IServiceProvider_QueryService(guidService As olelib.UUID, RequestedInterfaceIID As olelib.UUID, RemoteObjectPointer As Long)
    Dim tmpGUIDString As String
    tmpGUIDString = GetUUIDString(RequestedInterfaceIID)

    
    
    If tmpGUIDString = RemoteWMP.IID_IWMPRemoteMediaServices Then
        Dbg.WrLn "QueryService Success, found Interface IWMPRemoteMediaServices: " & vbCrLf & vbTab & tmpGUIDString
        
        'Pass our CRemoteMediaServices Class (IWMPRemoteMediaServices)
        RemoteObjectPointer = ObjPtr(m_Remote) 'Weak reference, may need to look at this not sure, we can always change the interface
    Else
        Dim tStr As String
        If tmpGUIDString = "{BFB377E5-C594-4369-A970-DE896D5ECE74}" Then
            tStr = "IWMPGraphCreation" 'We don't need this, just showing it's there, you can do neat stuff with it though...
        End If
        
        Dbg.WrLn "IServiceProvider QueryService Request for Interface: " & tStr & vbCrLf & vbTab & tmpGUIDString
        
        Err.Raise E_NOINTERFACE
    End If
End Sub

'IOleClientSite Members
'We are just forwarding back here, may or may not be necessary, safe anyway so why not...
Private Function IOleClientSite_GetContainer() As olelib.IOleContainer
  Set IOleClientSite_GetContainer = m_Container 'our form in this case
End Function

Private Function IOleClientSite_GetMoniker(ByVal dwAssign As olelib.OLEGETMONIKER, ByVal dwWhichMoniker As olelib.OLEWHICHMK) As olelib.IMoniker
    Set IOleClientSite_GetMoniker = m_OldClientSite.GetMoniker(dwAssign, dwWhichMoniker)
End Function

Private Sub IOleClientSite_OnShowWindow(ByVal fShow As olelib.BOOL)
    m_OldClientSite.OnShowWindow fShow
End Sub

Private Sub IOleClientSite_RequestNewObjectLayout()
    m_OldClientSite.RequestNewObjectLayout
End Sub

Private Sub IOleClientSite_SaveObject()
    m_OldClientSite.SaveObject
End Sub

Private Sub IOleClientSite_ShowObject()
    m_OldClientSite.ShowObject
End Sub




Private Sub Class_Initialize()
'Nothing to do I don't think
End Sub

Public Sub Init()
On Error GoTo errhandler
    'Fail if this isn't set up right
    If Not m_Container Is Nothing Then
        If Not m_Remote Is Nothing Then
            If Not m_WMP Is Nothing Then
                Dim tOleObj As olelib.IOleObject
              
                Set tOleObj = m_WMP.Object
                Set m_OldClientSite = tOleObj.GetClientSite
                
                tOleObj.SetClientSite Me 'Where the magic happens, sets this class as a client site, telling wmp to ask if we implement IWMPRemoteMediaServices
                m_WMP.uiMode = "custom"  'Where the extra special magic happens, we can now set a skin and pass wmp a a scriptable object...
                Exit Sub 'No Error
            End If
        End If
    End If
errhandler:
    Dbg.WrLn Err.Description
    Dbg.WrLn "*********************Probably something wrong with the skin!"
    Err.Raise E_FAIL
End Sub

Public Sub CloseOut()
On Error GoTo errhandler
    Dim tOleObj As olelib.IOleObject
    
    If Not m_WMP Is Nothing Then
        Set tOleObj = m_WMP.Object
        If Not m_OldClientSite Is Nothing Then
            tOleObj.SetClientSite m_OldClientSite 'Restore back for proper teardown
        End If
    End If
    Exit Sub
errhandler:
    Debug.Print Err.Description
    Err.Raise E_FAIL
End Sub

'Cleanup
Private Sub Class_Terminate()
    Set m_WMP = Nothing
    Set m_Remote = Nothing
    Set m_Container = Nothing
    Set m_OldClientSite = Nothing
End Sub



'CRemoteMediaServices - My implementation of IWMPRemoteMediaServices
Public Property Set RemoteObject(ByRef remObj As RemoteWMP.IWMPRemoteMediaServices)
    Set m_Remote = remObj
End Property
'CRemoteMediaServices - My implementation of IWMPRemoteMediaServices
Public Property Get RemoteObject() As RemoteWMP.IWMPRemoteMediaServices
    Set RemoteObject = m_Remote
End Property

'The form/host we came from
Public Property Set Container(ByRef contObj As IOleContainer)
    Set m_Container = contObj
End Property
'The form/host we came from
Public Property Get Container() As IOleContainer
    Set Container = m_Container
End Property



'Reference to our control on the form/host
Public Property Set WMP(ByRef wmpObj As WindowsMediaPlayer)
    Set m_WMP = wmpObj
End Property
'Reference to our control on the form/host
Public Property Get WMP() As WindowsMediaPlayer
    Set WMP = m_WMP
End Property


