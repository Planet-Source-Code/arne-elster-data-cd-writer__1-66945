Attribute VB_Name = "modMSF"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    dest As Any, _
    source As Any, _
    ByVal dlen As Long _
)

Public Type msf
    M   As Byte
    s   As Byte
    F   As Byte
End Type

' MSF structure to string => "**:**:**"
Public Function MSF2STR(fmt As msf) As String
    MSF2STR = Format(fmt.M, "00") & ":" & _
              Format(fmt.s, "00") & ":" & _
              Format(fmt.F, "00")
End Function

' MSF to LBA
Public Function MSF2LBA(fmt As msf, _
                Optional pos As Boolean) As Long

    With fmt
        MSF2LBA = CLng(.M) * 60 * 75 + (.s * 75) + .F
    End With

    If fmt.M < 90 Or pos Then
        MSF2LBA = MSF2LBA - 150
    Else
        MSF2LBA = MSF2LBA - 450150
    End If
End Function

' LBA to MSF
Public Function LBA2MSF(ByVal LBA As Long) As msf
    Dim M As Long, s As Long, F As Long, Start As Long

    Start = Choose(Abs(CBool(LBA >= -150)) + 1, 450150, 150)

    With LBA2MSF
        .M = Fix((LBA + Start) / (60 * 75))
        .s = Fix((LBA + Start - M * 60 * 75) / 75)
        .F = Fix(LBA + Start - M * 60 * 75 - s * 75)
    End With
End Function

' copy byte arrays or string to a long
Public Function VarToLBA(ParamArray fmt() As Variant) As Long
    Dim btLng()     As Byte
    Dim lng         As Long

    If TypeName(fmt(0)) = "String" Then
        If Len(fmt(0)) = 4 Then
            btLng = StrConv(StrReverse(fmt(0)), vbFromUnicode)
            CopyMemory lng, btLng(0), 4
            VarToLBA = lng
        Else
            VarToLBA = Val(fmt(0))
        End If
    ElseIf UBound(fmt) = 3 Then
        ReDim btLng(3) As Byte
        btLng(3) = fmt(0)
        btLng(2) = fmt(1)
        btLng(1) = fmt(2)
        btLng(0) = fmt(3)
        CopyMemory lng, btLng(0), 4
        VarToLBA = lng
    ElseIf UBound(fmt(0)) = 3 Then
        ReDim btLng(3) As Byte
        btLng(3) = fmt(0)(0)
        btLng(2) = fmt(0)(1)
        btLng(1) = fmt(0)(2)
        btLng(0) = fmt(0)(3)
        CopyMemory lng, btLng(0), 4
        VarToLBA = lng
    End If
End Function

' convert MSF strings or byte arrays to a MSF structure
Public Function VarToMSF(ParamArray fmt() As Variant) As msf
    If TypeName(fmt(0)) = "String" Then
        VarToMSF.M = Left$(fmt(0), InStr(fmt(0), ":") - 1)
        fmt(0) = Mid$(fmt(0), InStr(fmt(0), ":") + 1)
        VarToMSF.s = Left$(fmt(0), InStr(fmt(0), ":") - 1)
        fmt(0) = Val(Mid$(fmt(0), InStr(fmt(0), ":") + 1))
        VarToMSF.F = fmt(0)
    ElseIf UBound(fmt) = 2 Then
        VarToMSF.M = fmt(0)
        VarToMSF.s = fmt(1)
        VarToMSF.F = fmt(2)
    ElseIf UBound(fmt) = 3 Then
        VarToMSF.M = fmt(1)
        VarToMSF.s = fmt(2)
        VarToMSF.F = fmt(3)
    ElseIf UBound(fmt(0)) = 2 Then
        VarToMSF.M = fmt(0)(0)
        VarToMSF.s = fmt(0)(1)
        VarToMSF.F = fmt(0)(2)
    ElseIf UBound(fmt(0)) = 3 Then
        VarToMSF.M = fmt(0)(1)
        VarToMSF.s = fmt(0)(2)
        VarToMSF.F = fmt(0)(3)
    End If
End Function
