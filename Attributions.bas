Attribute VB_Name = "Attributions"
'Module Attributions
'Acknowledgements, Etc

'Eric Gunnerson     Eric Gunnerson's Compendium
'   Where I really started my research
'
'      http://blogs.msdn.com/b/ericgu/
'      http://blogs.msdn.com/b/ericgu/archive/2005/06/22/431783.aspx

'Jonathan Dibble
'   Wrote the code to do do some of this in C#, unfortunatly his code is hard to find
'   now. Fortunatly, some google foo let me save the relevant parts in the References folder
'   included with this project. (JonathanDibble_RemotedWindowsMediaPlayer_C#.zip)


'Chuck Holbrook AKA godofcpu/god_of_cpu     http://www.mp3car.com/members/god_of_cpu.html   https://www.youtube.com/user/godofcpu
'   Chuck provided the hint on how to procede after the remoting/skinning part was done
'
'      http://www.tech-archive.net/Archive/Media/microsoft.public.windowsmedia.sdk/2004-10/0228.html
'      http://www.tech-archive.net/Archive/Media/microsoft.public.windowsmedia.sdk/2004-10/0035.html


'Eduardo A. Morcillo    AKA Edanmo
'   Provided excellent Ole Type Libraries where so I didn't have to massage them myself
'
'   Namespace Edanmo:
'      http://www.mvps.org/emorcillo/en/index.shtml
'      http://www.mvps.org/emorcillo/en/code/vb6/index.shtml
'      http://www.mvps.org/emorcillo/download/vb6/tl_ole.zip    'Ole Type Libraries




'MSDN
'   Using Skins with the Windows Media Player Control
'       https://msdn.microsoft.com/en-us/library/windows/desktop/dd564572%28v=vs.85%29.aspx
'
'   IWMPRemoteMediaServices
'       https://msdn.microsoft.com/en-us/library/windows/desktop/dd563634%28v=vs.85%29.aspx
'
'   Skin Programming Reference
'       https://msdn.microsoft.com/en-us/library/windows/desktop/dd564359%28v=vs.85%29.aspx
'       https://msdn.microsoft.com/en-us/library/windows/desktop/dd564952%28v=vs.85%29.aspx



'----------------------------------------------------------------------------------------------------------------------------------------------
'Original Source - ©2000 Gus Molina
'
'Private Type GUID
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data4(7) As Byte
'End Type
'
'Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
'
'Public Function GetGUID() As String
''(c) 2000 Gus Molina
'
'    Dim udtGUID As GUID
'
'    If (CoCreateGuid(udtGUID) = 0) Then
'        GetGUID = _
'            String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
'            String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
'            String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
'            IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
'            IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
'            IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
'            IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
'            IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
'            IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
'            IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
'            IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
'    End If
'End Function
'----------------------------------------------------------------------------------------------------------------------------------------------


