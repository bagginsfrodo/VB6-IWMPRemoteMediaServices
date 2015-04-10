VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FTestHost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remoting Windows Media Player..."
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18180
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   18180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlayerControl 
      Caption         =   "Next"
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   37
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdPlayerControl 
      Caption         =   "Play"
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   36
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdPlayerControl 
      Caption         =   "Prev"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   35
      Top             =   4800
      Width           =   975
   End
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   3600
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClearFiles 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1680
      TabIndex        =   34
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddFiles 
      Caption         =   "AddFiles"
      Height          =   495
      Left            =   240
      TabIndex        =   33
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox lstPlaylist 
      Height          =   3435
      Left            =   240
      TabIndex        =   32
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Timer tmrPosition 
      Interval        =   150
      Left            =   11400
      Top             =   4440
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   0
      Max             =   100
      TabIndex        =   31
      Top             =   0
      Width           =   18135
   End
   Begin VB.CheckBox chkEnhancedAudio 
      Caption         =   "Enhanced Audio"
      Height          =   255
      Left            =   7920
      TabIndex        =   29
      Top             =   7560
      Width           =   1815
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1695
      Index           =   12
      LargeChange     =   5
      Left            =   11280
      Max             =   0
      Min             =   100
      MousePointer    =   2  'Cross
      TabIndex        =   28
      Top             =   2040
      Value           =   50
      Width           =   375
   End
   Begin VB.CommandButton cmdResetEQ 
      Caption         =   "Reset Default"
      Height          =   495
      Left            =   15480
      TabIndex        =   27
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CheckBox chkSplineTension 
      Caption         =   "Spline Tension"
      Height          =   375
      Left            =   15480
      TabIndex        =   26
      Top             =   6120
      Width           =   2655
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   11
      LargeChange     =   5
      Left            =   9120
      Max             =   0
      Min             =   100
      MousePointer    =   2  'Cross
      TabIndex        =   24
      Top             =   4800
      Value           =   50
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   10
      LargeChange     =   5
      Left            =   8280
      Max             =   0
      Min             =   100
      MousePointer    =   2  'Cross
      TabIndex        =   23
      Top             =   4800
      Value           =   50
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   0
      LargeChange     =   5
      Left            =   10320
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   21
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   1
      LargeChange     =   5
      Left            =   10815
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   20
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   2
      LargeChange     =   5
      Left            =   11310
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   19
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   3
      LargeChange     =   5
      Left            =   11805
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   18
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   4
      LargeChange     =   5
      Left            =   12300
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   17
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   5
      LargeChange     =   5
      Left            =   12795
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   16
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   6
      LargeChange     =   5
      Left            =   13290
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   15
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   7
      LargeChange     =   5
      Left            =   13785
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   14
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   8
      LargeChange     =   5
      Left            =   14280
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   13
      Top             =   4800
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Index           =   9
      LargeChange     =   5
      Left            =   14775
      Max             =   20
      Min             =   -20
      MousePointer    =   2  'Cross
      TabIndex        =   12
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton MyShortcutButtons 
      Caption         =   "Purple Back"
      Height          =   615
      Index           =   2
      Left            =   2640
      TabIndex        =   10
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton MyShortcutButtons 
      Caption         =   "Load Video"
      Height          =   615
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdVisEffects 
      Caption         =   "Next Preset"
      Height          =   615
      Index           =   3
      Left            =   16800
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdVisEffects 
      Caption         =   "Previous Preset"
      Height          =   615
      Index           =   2
      Left            =   15720
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdVisEffects 
      Caption         =   "Next Effect"
      Height          =   615
      Index           =   1
      Left            =   16800
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdVisEffects 
      Caption         =   "Previous Effect"
      Height          =   615
      Index           =   0
      Left            =   15720
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdThrowEvent 
      Caption         =   "Throw Event From Direct Script Access"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   8040
      TabIndex        =   4
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton cmdThrowEvent 
      Caption         =   "Throw Event From TestObj"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   8040
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdThrowEvent 
      Caption         =   "Throw Event From MyScriptableObject"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   8040
      TabIndex        =   2
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   8160
      Width           =   17775
   End
   Begin VB.CommandButton MyShortcutButtons 
      Caption         =   "Load Audio"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Test Buttons with preloaded paths"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   6720
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   4440
      Picture         =   "FTestHost.frx":0000
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2880
   End
   Begin VB.Label Label3 
      Caption         =   "Volume"
      Height          =   255
      Left            =   11160
      TabIndex        =   30
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "  Tru Bass  SRS WOW"
      Height          =   255
      Left            =   7800
      TabIndex        =   25
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "31Hz  62  125  250  500 1khz 2k   4k   8k  16k"
      Height          =   375
      Left            =   10320
      TabIndex        =   22
      Top             =   7200
      Width           =   4935
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   4095
      Left            =   12000
      TabIndex        =   11
      Top             =   360
      Width           =   5895
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   10398
      _cy             =   7223
   End
End
Attribute VB_Name = "FTestHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form FTestHost
'©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
'License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc

'Some/All of this is NOT production quality code, but you get the gist, improve it in your own programs

Option Explicit

Private WithEvents myScriptableObject As CScriptableObject  'Object I have made and provided to the skin script
Attribute myScriptableObject.VB_VarHelpID = -1
Private myRemote As CRemoteMediaServices                    'IWMPRemoteMediaServices Implementation
Private myRemoteHost As CRemoteHost                         'IServiceProvider, IOleClientSite Implementation


Private mySkinScriptObject As Object                        'The skin/script object FROM Windows Media Player
Private myPlayerFromSkin As WMPCore                         'The "player" global object from inside the skin, we don't necessarily need this, just showing we can get it

Private myTheme As IWMPTheme                                'The "Theme" (<THEME/>) object from the skin, in my skin
Private myView As IWMPLayoutView                            'The "View" (<VIEW/>) object from the skin, in my skin: id="s_myview"

Private myVisEffects As WMPEffects                          'The Visual Effects (WMPEffects) (<EFFECTS/>) object from the skin, in my skin: id="s_viseffects"
Private myVisEffectsDispatch As Object                      'The Visual Effects (WMPEffects) (<EFFECTS/>) object from the skin, in my skin: id="s_viseffects", as an IDispatch object, so we can get at the other
                                                            'interface of this object, which has width, height, etc.. (the properties of the actual skinnable object (<EFFECTS/>)

Private myEQ As WMPEqualizerSettingsCtrl                    'The Equalizer (WMPEqualizerSettingsCtrl) (<EQUALIZERSETTINGS/>) object from the skin, in my skin: id="s_eqsettings"

Private myVideo As WMPVideoCtrl                             'The Video (WMPVideoCtrl) (<VIDEO/>) object from the skin, in my skin: id="s_vid"
Private myVideoSettings As WMPVideoSettingsCtrl             'The VideoSettings (WMPVideoSettingsCtrl) (<VIDEOSETTINGS/>) object from the skin, in my skin: id="s_vidsettings"

Private myPlayerSettings As IWMPSettings2                   'The "player.settings" (IWMPSettings2) object from the skin, in my skin



Private m_InitEQ As Boolean
Private m_NoChangePos As Boolean
Private m_LastMedia As String 'For a simple debounce routine

'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Begin Load/Unload
Private Sub Form_Load()
   Init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseOut
End Sub

Private Sub Init()
On Error GoTo errhandler

    'Setup my simple debugging display object
    With Dbg
        .PrintDebug = True
        Set .MyTextBox = Text1
    End With
    
    Set myScriptableObject = New CScriptableObject  'What we pass to the skin/script
    
    
    Set myRemote = New CRemoteMediaServices         'IWMPRemoteMediaServices
    Set myRemoteHost = New CRemoteHost              'IServiceProvider, IOleClientSite
    
    With myRemote
        Set .ScriptObject = myScriptableObject
        .SkinFile = SmartAppPath & "Skin\RTestSkin.wms"
    End With
        
    With myRemoteHost                               'Give the class some local references for glue code
        Set .Container = Me
        Set .RemoteObject = myRemote
        Set .WMP = wmp1
        .Init   'Let's Roll
    End With
       
    Me.Caption = "Remoting Windows Media Player In Local Mode With Custom Skin, Visualization Access, And Equalizer"
    
    'Playing with the skin, look in the skin file
    With mySkinScriptObject.s_text1
        .scrollingDelay = 100
        .scrollingAmount = 3
        .Value = Me.Caption & Space$(10)
    End With
    
    'Set up EQ
    DoInitEQ True 'Reset on firstrun
    
    With myTheme
        Dbg.WrLn "Skin Info", vbTab & .Title, vbTab & .author, vbTab & .copyright
    End With
    
Exit Sub
errhandler:
    Debug.Print Err.Description
    Unload Me
End Sub

Private Sub CloseOut()
    m_InitEQ = False
    myRemoteHost.CloseOut 'Let's Ensure Proper Teardown...
End Sub
'End Load/Unload
'------------------------------------------------------------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Begin Simple Load/Play Etc
Private Sub MyShortcutButtons_Click(Index As Integer)
    Select Case Index
        Case 0
            'Play "M:\Music\3 Doors Down\the better life\3 Doors Down - Duck and Run.mp3"
            Dbg.WrLn "This button is commented out in code..."
        Case 1
            'Play "E:\User Profile Files\Videos\Personal\Lana Del Rey - Summertime Sadness.mp4"
            Dbg.WrLn "This button is commented out in code..."
        Case 2
            myView.backgroundColor = "purple"
    End Select
End Sub

Private Sub Play(url As String) 'Oh the ways we could call it...
    'mySkinScriptObject.player.URL = url
    'myPlayerFromSkin.url = url
    wmp1.url = url
End Sub

'Next/Previous/Play/Pause
Private Sub cmdPlayerControl_Click(Index As Integer)
    Select Case Index
        Case 0
            If lstPlaylist.ListCount > 0 Then
                If lstPlaylist.ListIndex <> 0 Then
                    lstPlaylist.ListIndex = lstPlaylist.ListIndex - 1
                Else
                    lstPlaylist.ListIndex = lstPlaylist.ListCount - 1
                End If
                Play lstPlaylist.List(lstPlaylist.ListIndex)
            End If
        Case 1
            If wmp1.playState = wmppsPlaying Then
                wmp1.Controls.pause
            ElseIf wmp1.playState = wmppsPaused Then
                wmp1.Controls.Play
            End If
        Case 2
            If lstPlaylist.ListCount > 0 Then
                If lstPlaylist.ListIndex <> lstPlaylist.ListCount - 1 Then
                    lstPlaylist.ListIndex = lstPlaylist.ListIndex + 1
                Else
                    lstPlaylist.ListIndex = 0
                End If
                Play lstPlaylist.List(lstPlaylist.ListIndex)
            End If
    End Select
End Sub

'Add files to playlist
Private Sub cmdAddFiles_Click()
On Error GoTo errhandler
    With comDialog
        .FileName = vbNullString
        .InitDir = "C:\"
        .Filter = "Media Files (*.mp3)|*.mp3|*.mp4)|*.mp4|*.m4a)|*.m4a|All files (*.*)|*.*"
        .DefaultExt = "mp3"
        .DialogTitle = "Select Files..."
        .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
        .MaxFileSize = 32767
        .ShowOpen
    End With
    
    Dim tFileArray() As String
        tFileArray = Split(comDialog.FileName, vbNullChar)
    
    If UBound(tFileArray) > 0 Then
        Dim tFilePath As String
            tFilePath = tFileArray(0) & "\"
        
        Dim x As Long
        For x = 1 To UBound(tFileArray)
            lstPlaylist.AddItem tFilePath & tFileArray(x)
        Next x
    Else
        lstPlaylist.AddItem tFileArray(0)
    End If

Exit Sub
errhandler:
    Dbg.WrLn Err.Description
End Sub

'Clear playlist
Private Sub cmdClearFiles_Click()
    lstPlaylist.Clear
End Sub

'Play from playlist
Private Sub lstPlaylist_DblClick()
On Error GoTo errhandler
    Play lstPlaylist.List(lstPlaylist.ListIndex)
    Dbg.WrLn "Playing: """ & lstPlaylist.List(lstPlaylist.ListIndex) & """"

Exit Sub
errhandler:
    Dbg.WrLn Err.Description
End Sub

'End Simple Load/Play Etc
'------------------------------------------------------------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Begin Communicating with skin/script
Private Sub cmdThrowEvent_Click(Index As Integer)
    'Demos Calling Script Inside Media Player Skin
    Select Case Index
        Case 0
            myScriptableObject.ThrowEventFromMyScriptableObject
        Case 1
            myScriptableObject.ThrowEventFromTestObj
        Case 2
            mySkinScriptObject.Event_DirectScriptAccess "Event From Direct Script Access"
    End Select
End Sub
'End Communicating with skin/script
'------------------------------------------------------------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Begin Visual Effects Change
Private Sub cmdVisEffects_Click(Index As Integer)
    With myVisEffects
        Select Case Index
        Case 0
            .previousEffect
        Case 1
            .nextEffect
        Case 2
            .previousPreset
        Case 3
            .nextPreset
        End Select
    End With
End Sub
'End Visual Effects Change
'------------------------------------------------------------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Begin Track Positioning
Private Sub tmrPosition_Timer()
On Error GoTo errhandler
    m_NoChangePos = True
    If Not wmp1.currentMedia Is Nothing Then
        HScroll1.Value = (wmp1.Controls.currentPosition / wmp1.currentMedia.duration) * 100
    End If
    
    'Using the timer here for Play/Pause as well, sorry it's out of place lol
    If wmp1.playState = wmppsPaused Then
        cmdPlayerControl(1).Caption = "Play"
    ElseIf wmp1.playState = wmppsPlaying Then
        cmdPlayerControl(1).Caption = "Pause"
    End If
    
    
errhandler:
    m_NoChangePos = False
End Sub

Private Sub HScroll1_Change()
On Error GoTo errhandler
    If m_NoChangePos = False Then
        If Not wmp1.currentMedia Is Nothing Then
            wmp1.Controls.currentPosition = (HScroll1.Value / 100) * wmp1.currentMedia.duration
        End If
    End If
errhandler:
    tmrPosition.Enabled = True
End Sub

Private Sub HScroll1_Scroll()
    tmrPosition.Enabled = False
End Sub
'End Track Positioning
'------------------------------------------------------------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Begin myScriptableObject Events
Private Sub myScriptableObject_SkinScriptObjectReceived(SkinScriptObject As Object)
    Set mySkinScriptObject = SkinScriptObject
    Dbg.WrLn "Object Received From Skin: " & "SkinScript"
End Sub

Private Sub myScriptableObject_PlayerObjectFromSkinReceived(PlayerObject As Object)
    Set myPlayerFromSkin = PlayerObject
    Dbg.WrLn "Object Received From Skin: " & "Player"
End Sub

Private Sub myScriptableObject_VisualEffectsFromSkinReceived(Effects As WMPLibCtl.WMPEffects, EffectsDispatchObject As Object)
    Set myVisEffects = Effects
    Set myVisEffectsDispatch = EffectsDispatchObject
    Dbg.WrLn "Object Received From Skin: " & "Visual Effects"
    Dbg.WrLn "Object Received From Skin: " & "Visual Effects Dispatch"
End Sub

Private Sub myScriptableObject_EqualizerFromSkinReceived(eqsettings As WMPLibCtl.WMPEqualizerSettingsCtrl)
    Set myEQ = eqsettings
    Dbg.WrLn "Object Received From Skin: " & "Equalizer"
End Sub

Private Sub myScriptableObject_VideoFromSkinReceived(vidCtrl As WMPLibCtl.WMPVideoCtrl)
    Set myVideo = vidCtrl
    Dbg.WrLn "Object Received From Skin: " & "Video"
End Sub

Private Sub myScriptableObject_VideoSettingsFromSkinReceived(vidSettingsCtrl As WMPLibCtl.WMPVideoSettingsCtrl)
    Set myVideoSettings = vidSettingsCtrl
    Dbg.WrLn "Object Received From Skin: " & "Video Settings"
End Sub

Private Sub myScriptableObject_SettingsFromSkinReceived(PlayerSettings As WMPLibCtl.IWMPSettings2)
    Set myPlayerSettings = PlayerSettings
    Dbg.WrLn "Object Received From Skin: " & "Player Settings"
End Sub

Private Sub myScriptableObject_ViewFromSkinReceived(SkinView As WMPLibCtl.IWMPLayoutView)
    Set myView = SkinView
    Dbg.WrLn "Object Received From Skin: " & "View"
End Sub

Private Sub myScriptableObject_ThemeFromSkinReceived(SkinTheme As WMPLibCtl.IWMPTheme)
    Set myTheme = SkinTheme
    Dbg.WrLn "Object Received From Skin: " & "Theme"
End Sub


Private Sub myScriptableObject_WootReceived()
    Dbg.WrLn "Woot! from Skin Script Received"
End Sub
'End myScriptableObject Events
'------------------------------------------------------------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Using this here to hide/show visual effects depending on media type
Private Sub wmp1_MediaChange(ByVal Item As Object)
    'This fires multiple times, haven't spent the time to research if there is a reason, or if it's a bug,
    'either way, for our purposes we only want to fire it once, so there is a simple debouncing routine here
    'to only fire it if the actual media has changed, although here it wouldn't hurt anyway

    If Item.sourceURL <> m_LastMedia Then
        myVisEffectsDispatch.Visible = IsMediaAudio(Item) 'Can't use myVisEffects.Visible, because interface doesn't support visible, but the actual object does
    End If
    m_LastMedia = Item.sourceURL
End Sub

Private Function IsMediaAudio(ByRef tMedia As WMPLibCtl.IWMPMedia3)
    'Might want to look at this in general, I'm not sure if IWMPMedia3.getItemInfo("MediaClassPrimaryID") is reliable
    
    'Dbg.WrLn tMedia.getItemInfo("MediaClassPrimaryID")
    
    Select Case tMedia.getItemInfo("MediaClassPrimaryID")
        Case "{D1607DBC-E323-4BE2-86A1-48A42A28441E}", "{01CD0F29-DA4E-4157-897B-6275D50C4F11}", ""
            IsMediaAudio = True
            Dbg.WrLn "Media is Audio"
        Case Else
            IsMediaAudio = False
            Dbg.WrLn "Media is Video"
    End Select

End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Begin EQ Stuff
Private Sub DoInitEQ(Optional reset As Boolean = False)
    
    With myEQ
    
        Dim x As Integer
        
        If reset = True Then
            For x = 1 To 10
                .gainLevels(x) = 0
            Next x
            
            .enhancedAudio = False
            .splineTension = False
            .truBassLevel = 50
            .wowLevel = 50
            myPlayerSettings.volume = 50
        End If
        
        
        
        For x = 1 To 10
            VScroll1(x - 1).Value = -.gainLevels(x)
        Next x
            
        chkEnhancedAudio = IIf(.enhancedAudio = True, vbChecked, vbUnchecked)
        chkSplineTension = IIf(.enableSplineTension = True, vbChecked, vbUnchecked)
        
        VScroll1(10).Value = .truBassLevel
        VScroll1(11).Value = .wowLevel
        VScroll1(12).Value = myPlayerSettings.volume
    End With
    
    m_InitEQ = True

End Sub

Private Sub chkEnhancedAudio_Click()
    If m_InitEQ = True Then
        If chkEnhancedAudio = vbChecked Then
            myEQ.enhancedAudio = True
        Else
            myEQ.enhancedAudio = False
        End If
    End If
    
    DoInitEQ
End Sub

Private Sub chkSplineTension_Click()
    If m_InitEQ = True Then
        myEQ.splineTension = 0.3
        If chkSplineTension.Value = vbChecked Then
            myEQ.enableSplineTension = True
        Else
            myEQ.enableSplineTension = False
        End If
    End If
End Sub

Private Sub cmdResetEQ_Click()
    DoInitEQ True
End Sub

Private Sub VScroll1_Change(Index As Integer)
    If m_InitEQ = True Then
        Select Case Index
            Case Is < 10
                myEQ.gainLevels(Index + 1) = -VScroll1(Index).Value
            Case 10
                myEQ.truBassLevel = VScroll1(Index).Value
            Case 11
                myEQ.wowLevel = VScroll1(Index).Value
            Case 12
                wmp1.settings.volume = VScroll1(Index).Value
        End Select
    End If

    DoInitEQ
End Sub

Private Sub VScroll1_Scroll(Index As Integer)
    VScroll1_Change Index
End Sub
'End EQ Stuff
'------------------------------------------------------------------------------------------------------------------------------------------------------------------




