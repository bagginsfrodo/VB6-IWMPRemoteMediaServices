VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScriptableObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class CScriptableObject
'©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
'License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc

Option Explicit

'This are just events we are providing back to the main app for ease of use
Public Event WootReceived()

Public Event SkinScriptObjectReceived(ByRef SkinScriptObject As Object)
Public Event PlayerObjectFromSkinReceived(ByRef PlayerObject As Object)

Public Event VisualEffectsFromSkinReceived(ByRef Effects As WMPEffects, ByRef EffectsDispatchObject As Object)
Public Event EqualizerFromSkinReceived(ByRef eqsettings As WMPEqualizerSettingsCtrl)
Public Event VideoFromSkinReceived(ByRef vidCtrl As WMPVideoCtrl)
Public Event VideoSettingsFromSkinReceived(ByRef vidSettingsCtrl As WMPVideoSettingsCtrl)
Public Event SettingsFromSkinReceived(ByRef PlayerSettings As IWMPSettings2)
Public Event ViewFromSkinReceived(ByRef SkinView As IWMPLayoutView)
Public Event ThemeFromSkinReceived(ByRef SkinTheme As IWMPTheme)


Private m_SkinScript As Object 'The actual skin script object
Private WithEvents m_PlayerFromSkinObject As WMPCore 'the player object from the skin
Attribute m_PlayerFromSkinObject.VB_VarHelpID = -1

Private m_VisualEffects As Object 'We're going to leave this as object to get multiple interfaces...
Private m_Equalizer As WMPEqualizerSettingsCtrl
Private WithEvents m_Video As WMPVideoCtrl 'WithEvents just to show that it works for this object
Attribute m_Video.VB_VarHelpID = -1
Private m_VideoSettings As WMPVideoSettingsCtrl
Private m_Settings As IWMPSettings2

Private m_View As IWMPLayoutView 'The "VIEW" object from the skin
Private m_Theme As IWMPTheme 'The "THEME" object from the skin

Private m_TestObj As New CTestObject 'An object used just to show how we can add some functionality to skins, like asking for objects..

Private Sub Class_Initialize()
'
End Sub

'this is called from the skin script file to pass objects back to us
Public Sub InitScript(ScriptFromSkin As Object)
On Error GoTo errhandler

#If False Then 'Using a conditional compile to work around a bug in the IDE that keeps messing with the casing on object names!!!!!!
            Dim settings
            Dim s_viseffects
            Dim s_eqsettings
            Dim s_vid
            Dim s_vidsettings
            Dim s_myview
            Dim theme
            Dim view
#End If


'Called From Skin Script
'Check the casing here! everything is supposed to be lower case from the skin (.player, .settings, .theme)
'Our stuff is defined lower case too

    If Not ScriptFromSkin Is Nothing Then
        Set m_SkinScript = ScriptFromSkin
        With m_SkinScript
            Set m_VisualEffects = .s_viseffects
            Set m_Equalizer = .s_eqsettings
            Set m_Video = .s_vid
            Set m_VideoSettings = .s_vidsettings
            Set m_PlayerFromSkinObject = .player
            Set m_Settings = .settings
            Set m_View = .s_myview
            Set m_Theme = .theme
            Set m_TestObj.Parent = m_SkinScript
        End With
    End If
        
    'Pass Objects to anyone consuming us
    RaiseEvent PlayerObjectFromSkinReceived(m_PlayerFromSkinObject)
    RaiseEvent SkinScriptObjectReceived(m_SkinScript)
    
    RaiseEvent VisualEffectsFromSkinReceived(m_VisualEffects, m_VisualEffects)
    RaiseEvent EqualizerFromSkinReceived(m_Equalizer)
    RaiseEvent VideoFromSkinReceived(m_Video)
    RaiseEvent VideoSettingsFromSkinReceived(m_VideoSettings)
    RaiseEvent SettingsFromSkinReceived(m_Settings)
    RaiseEvent ViewFromSkinReceived(m_View)
    RaiseEvent ThemeFromSkinReceived(m_Theme)

Exit Sub
errhandler:
    Debug.Print Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub Class_Terminate()
'
End Sub

Public Sub Woot() 'Script is Calling us !
    RaiseEvent WootReceived
End Sub

Public Function alert(msg) 'Provide some functionality to skin..
    alert = MsgBox(msg)
    Dbg.WrLn "Function ""alert"" called from Scriptable Object with message:" & vbCrLf & vbTab & msg
End Function

Public Function GetTestObj() As Object 'Provide a Test object to skin.. which means the script asked us for it..
    Set GetTestObj = m_TestObj
    Dbg.WrLn "Skin Script asked for the test object."
End Function

Public Function ThrowEventFromTestObj() 'Just a demo
    m_TestObj.Throw "TestObj Sends Love"
End Function

Public Function ThrowEventFromMyScriptableObject() 'Just a demo
    m_SkinScript.Event_MyScriptableObject "MyScriptableObject checking in.."
End Function




Public Property Get SkinScript() As Object 'Provide means to get a reference to the skin script
    Set SkinScript = m_SkinScript
End Property

Public Property Get PlayerObject() As Object 'Provide means to get a reference to the skin script
    Set PlayerObject = m_PlayerFromSkinObject
End Property

Public Property Get VisualEffects() As WMPEffects 'Provide means to get a reference to the VisualEffects object
    Set VisualEffects = m_VisualEffects
End Property

Public Property Get VisualEffectsDispatchObject() As Object 'Provide means to get a reference to the VisualEffects IDispatch object
    Set VisualEffectsDispatchObject = m_VisualEffects
End Property


Public Property Get Equalizer() As WMPEqualizerSettingsCtrl 'Provide means to get a reference to the Equalizer object
    Set Equalizer = m_Equalizer
End Property

Public Property Get Video() As WMPVideoCtrl 'Provide means to get a reference to the Video object
    Set Video = m_Video
End Property
Public Property Get VideoSettings() As WMPVideoSettingsCtrl 'Provide means to get a reference to the Video Settings object
    Set VideoSettings = m_VideoSettings
End Property

Public Property Get PlayerSettings() As IWMPSettings2 'Provide means to get a reference to the Player Settings object
    Set PlayerSettings = m_Settings
End Property

Public Property Get SkinView() As IWMPLayoutView 'Provide means to get a reference to the "VIEW" object
    Set SkinView = m_View
End Property

Public Property Get SkinTheme() As IWMPTheme 'Provide means to get a reference to the "THEME" object
    Set SkinTheme = m_Theme
End Property



'Events from Video Control
'Just showing this because it's there if you need it
Private Sub m_Video_onvideoend()
'
End Sub

Private Sub m_Video_onvideostart()
'
End Sub
