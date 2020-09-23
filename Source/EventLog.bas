Attribute VB_Name = "EventLog"

Option Explicit

Public Type EventRecord
   EventTimeWritten  As Variant
   EventSourceName   As String
   EventUserName     As String
   EventComputerName As String
   EventType         As String
   EventDescription  As String
   EventData         As String
   EventCategory     As Integer
   EventRecordNum    As Long
   EventID           As Long
End Type

Global colEventRecord() As EventRecord

Public Const EVNT_SYSTEM = "System"
Public Const EVNT_APP = "Application"
Public Const EVNT_SECURITY = "Security"

Public Const Event_Type_Info = "Information"
Public Const Event_Type_Warning = "Warning"
Public Const Event_Type_Error = "Error"
Public Const Event_Type_Success_Audit = "Audit Success"
Public Const Event_Type_Failure_Audit = "Audit Failure"

Public Const EVENTLOG_ERROR_TYPE = &H1
Public Const EVENTLOG_WARNING_TYPE = &H2
Public Const EVENTLOG_INFORMATION_TYPE = &H4
Public Const EVENTLOG_AUDIT_SUCCESS = &H8
Public Const EVENTLOG_AUDIT_FAILURE = &H10

Public Const Filter_Type_None = 0
Public Const Filter_Type_TimeBefore = 1
Public Const Filter_Type_TimeAfter = 2
Public Const Filter_Type_EventType = 3
Public Const Filter_Type_Source = 4
Public Const Filter_Type_Category = 5
Public Const Filter_Type_Computer = 6
Public Const Filter_Type_EventID = 7

Public Const ERR_LOGTYPE_NOT_SET = 1011 + vbObjectError
Public Const ERR_SOURCENAME_NOT_SET = 1012 + vbObjectError
Public Const ERR_BAD_INDEX = 1013 + vbObjectError
Public Const ERR_FAILED_OPEN_REGISTRY_KEY = 1014 + vbObjectError
Public Const ERR_FAILED_READ_REGISTRY_KEY = 1015 + vbObjectError
Public Const ERR_RESOURCE_DATA_NOT_FOUND = 1016 + vbObjectError
Public Const ERR_READING_EVENT_LOG = 1017 + vbObjectError
Public Const ERR_LOG_NOT_OPENED = 1018 + vbObjectError
Public Const ERR_FAILED_SET_LOG_TYPE = 1019 + vbObjectError

Public Function ReadEventLog(Server As String, _
                            Drive As String, _
                            logType As String, _
                            logFilter As Date) As Boolean
Dim plngRtn       As Long
Dim I             As Long
Dim plngEventCnt  As Long
Dim errAdd        As ListItem
Dim pEventLog     As EventLogs
Dim accessError   As String
Dim messageError  As String
Dim openLog       As Boolean
Dim readLog       As Boolean
   
On Error Resume Next
Set pEventLog = New EventLogs
If GetTimeZoneInformation(TZInfo) = 2 Then logFilter = logFilter - (1 / 24)
pEventLog.EventFilter(2) = logFilter
pEventLog.EventReadLogForward = True
pEventLog.EventDataReturnHex = False
pEventLog.EventTypeLog = logType

openLog = pEventLog.OpenAnyEventLog(Server)
readLog = pEventLog.ReadEventEntries
If Abs(openLog + readLog) <> 2 Then
    If Not openLog Then
        accessError = pEventLog.LastEventErrorDescription
        If accessError = "" Then
            accessError = logType & " Log Opening Error: Possibly evntlog2.dll is not registered."
        Else: accessError = logType & " Log Opening Error: " & accessError
        End If
    ElseIf Not readLog Then
        accessError = pEventLog.LastEventErrorDescription
        If accessError = "" Then
            ReadEventLog = False
            Set pEventLog = Nothing
            Exit Function
        Else: accessError = logType & " Log Read Error: " & accessError
        End If
    End If
    Call LogErrors(Mid(Server, 3), Drive, accessError)
    Set errAdd = frmSrvmtr.statusLW.ListItems.Add(, , Server, , "noAccess")
    errAdd.SubItems(1) = Now
    errAdd.SubItems(2) = accessError
    errAdd.SubItems(3) = "Not Sent - See Error Message."
    ReadEventLog = False
    Set pEventLog = Nothing
    Exit Function
End If

plngEventCnt = pEventLog.CountEventRecords
If plngEventCnt = 0 Then
      Set pEventLog = Nothing
      Exit Function
End If
   
If plngEventCnt > 0 Then
    ReDim colEventRecord(1 To plngEventCnt) As EventRecord
End If
  
  
If plngEventCnt > 0 Then
    For I = 1 To plngEventCnt

         colEventRecord(I).EventTimeWritten = pEventLog.EventTimeWritten(I)
         colEventRecord(I).EventType = pEventLog.EventType(I)
         colEventRecord(I).EventSourceName = pEventLog.EventSourceName(I)
         colEventRecord(I).EventID = pEventLog.EventID(I)
        

        If colEventRecord(I).EventType = Event_Type_Error Then
            ReadEventLog = True
            Set errAdd = frmSrvmtr.statusLW.ListItems.Add(, , Mid(Server, 3), , "StopEvent")
            errAdd.SubItems(1) = Now
            messageError = logType & " Log Error Type Event. Time: " & _
                                colEventRecord(I).EventTimeWritten + (1 / 24) & ", Source: " & _
                                colEventRecord(I).EventSourceName & ", Event ID: " & _
                                colEventRecord(I).EventID & ". Check Event Log for details."
            errAdd.SubItems(2) = messageError
            Call LogErrors(Mid(Server, 3), Drive, messageError)
        End If
      Next I
End If

Set pEventLog = Nothing

End Function
