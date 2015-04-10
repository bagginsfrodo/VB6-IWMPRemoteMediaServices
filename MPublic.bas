Attribute VB_Name = "MPublic"
'Module MPublic
'©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
'License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc


Option Explicit

Public Dbg As New CdbgWrite

Public Function GetUUIDString(udtGUID As olelib.UUID) As String
'Originally ©2000 Gus Molina, Modified, See Attributions Module
    GetUUIDString = UCase("{" & _
    String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & "-" & _
    String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & "-" & _
    String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & "-" & _
    IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
    IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & "-" & _
    IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
    IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
    IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
    IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
    IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
    IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7)) & "}")
End Function


Public Function SmartAppPath() As String
    SmartAppPath = App.Path
    If Right$(SmartAppPath, 1) <> "\" Then
        SmartAppPath = SmartAppPath & "\"
    End If
End Function



