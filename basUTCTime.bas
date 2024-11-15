Attribute VB_Name = "basUTCTime"
Option Explicit
'UTC/Local Time Conversion
'Adapted from code by Tim Hall published at https://github.com/VBA-tools/VBA-UtcConverter

'PUBLIC FUNCTIONS:
'    - UTCtoLocal(utc_UtcDate As Date) As Date     converts UTC datetimes to local
'    - LocalToUTC(utc_LocalDate As Date) As Date   converts local DateTime to UTC
'    - TimestampToLocal(st As String) As Date      converts epoch timestamp to Local Time
'    - LocalToTimestamp(dt as date) as String      converts Local Time to timestamp
'Accuracy confirmed for several variations of time zones & DST rules. (ashleedawg)
'===============================================================================
Private Type utc_SYSTEMTIME
    utc_wYear As Integer: utc_wMonth As Integer: utc_wDayOfWeek As Integer: utc_wDay As Integer
    utc_wHour As Integer: utc_wMinute As Integer: utc_wSecond As Integer: utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long: utc_StandardName(0 To 31) As Integer: utc_StandardDate As utc_SYSTEMTIME: utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer: utc_DaylightDate As utc_SYSTEMTIME: utc_DaylightBias As Long
End Type

'http://msdn.microsoft.com/library/windows/desktop/ms724421.aspx /ms724949.aspx /ms725485.aspx
Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME ' "Helper Function" for Public subs (below)
On Error GoTo Err_utc_DateToSystemTime
 
    With utc_DateToSystemTime
        .utc_wYear = Year(utc_Value): .utc_wMonth = Month(utc_Value): .utc_wDay = Day(utc_Value)
        .utc_wHour = Hour(utc_Value): .utc_wMinute = Minute(utc_Value): .utc_wSecond = Second(utc_Value): .utc_wMilliseconds = 0
    End With

Exit_utc_DateToSystemTime:
    Exit Function

Err_utc_DateToSystemTime:
    MsgBox "Error in basUTCTime:utc_DateToSystemTime: " & Err.Description
    Resume Exit_utc_DateToSystemTime

End Function
Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date ' "Helper Function" for Public Functions (below)
On Error GoTo Err_utc_SystemTimeToDate
    
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)

Exit_utc_SystemTimeToDate:
    Exit Function

Err_utc_SystemTimeToDate:
    MsgBox "Error in basUTCTime:utc_SystemTimeToDate: " & Err.Description
    Resume Exit_utc_SystemTimeToDate

End Function
Public Function TimestampToLocal(st As String) As Date
On Error GoTo Err_TimestampToLocal

    TimestampToLocal = UTCtoLocal((Val(st) / 86400) + 25569)
    
Exit_TimestampToLocal:
    Exit Function

Err_TimestampToLocal:
    MsgBox "Error in basUTCTime:TimestampToLocal: " & Err.Description
    Resume Exit_TimestampToLocal

End Function
Public Function LocalToTimestamp(dt As Date) As String
On Error GoTo Err_LocalToTimestamp

    LocalToTimestamp = (LocalToUTC(dt) - 25569) * 86400

Exit_LocalToTimestamp:
    Exit Function

Err_LocalToTimestamp:
    MsgBox "Error in basUTCTime:LocalToTimestamp: " & Err.Description
    Resume Exit_LocalToTimestamp

End Function
Public Function UTCtoLocal(utc_UtcDate As Date) As Date
On Error GoTo Err_UTCtoLocal

    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION, utc_LocalDate As utc_SYSTEMTIME
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate
    UTCtoLocal = utc_SystemTimeToDate(utc_LocalDate)

Exit_UTCtoLocal:
    Exit Function

Err_UTCtoLocal:
    MsgBox "Error in basUTCTime:UTCtoLocal: " & Err.Number & " - " & Err.Description
    Resume Exit_UTCtoLocal

End Function
Public Function LocalToUTC(utc_LocalDate As Date) As Date
On Error GoTo Err_LocalToUTC

    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION, utc_UtcDate As utc_SYSTEMTIME
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate
    LocalToUTC = utc_SystemTimeToDate(utc_UtcDate)

Exit_LocalToUTC:
    Exit Function

Err_LocalToUTC:
    MsgBox "Error in basUTCTime:LocalToUTC: " & Err.Number & " - " & Err.Description
    Resume Exit_LocalToUTC

End Function

