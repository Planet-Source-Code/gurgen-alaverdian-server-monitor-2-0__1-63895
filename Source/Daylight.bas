Attribute VB_Name = "Daylight"
Public Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type


Public Enum cstDSTType
    cstDSTUnknown = 0
    cstDSTStandard = 1
    cstDSTDaylight = 2
End Enum


Public TZInfo As TIME_ZONE_INFORMATION
Public TZType As Long

