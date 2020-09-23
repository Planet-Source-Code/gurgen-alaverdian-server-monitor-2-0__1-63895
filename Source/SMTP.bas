Attribute VB_Name = "SMTP"

Public Sub SendSMTP()
    Dim I As Integer
    Dim strServer As String, colonPos As Integer, lngPort As Long

    
    strServer = Trim(txtHost)
    'find out if the sender is using a Proxy server
    colonPos = InStr(strServer, ":")
    If colonPos = 0 Then
        'no proxy so use standard SMTP port
        Winsock1.Connect strServer, 25
    Else
        'Proxy, so get proxy port number and parse out the server name or IP address
        lngPort = CLng(Right$(strServer, Len(strServer) - colonPos))
        strServer = Left$(strServer, colonPos - 1)
        Winsock1.Connect strServer, lngPort
    End If
    m_State = MAIL_CONNECT

End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    '
    'Retrive data from winsock buffer
    '
    Winsock1.GetData strServerResponse
    '
    Debug.Print strServerResponse
    '
    'Get server response code (first three symbols)
    '
    strResponseCode = Left(strServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
       
        Select Case m_State
            Case MAIL_CONNECT
                'Change current state of the session
                m_State = MAIL_HELO
                '
                'Remove blank spaces
                strDataToSend = Trim$(txtSender)
                '
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                Winsock1.SendData "HELO " & strDataToSend & vbCrLf
                '
                Debug.Print "HELO " & strDataToSend
                '
            Case MAIL_HELO
                '
                'Change current state of the session
                m_State = MAIL_FROM
                '
                'Send MAIL FROM command to the server
                Winsock1.SendData "MAIL FROM:" & Trim$(txtSender) & vbCrLf
                '
                Debug.Print "MAIL FROM:" & Trim$(txtSender)
                '
            Case MAIL_FROM
                '
                'Change current state of the session
                m_State = MAIL_RCPTTO
                '
                'Send RCPT TO command to the server
                Winsock1.SendData "RCPT TO:" & Trim$(txtRecipient) & vbCrLf
                '
                Debug.Print "RCPT TO:" & Trim$(txtRecipient)
                '
            Case MAIL_RCPTTO
                '
                'Change current state of the session
                m_State = MAIL_DATA
                '
                'Send DATA command to the server
                Winsock1.SendData "DATA" & vbCrLf
                '
                Debug.Print "DATA"
                '
            Case MAIL_DATA
                '
                'Change current state of the session
                m_State = MAIL_DOT
                '
                'So now we are sending a message body
                'Each line of text must be completed with
                'linefeed symbol (Chr$(10) or vbLf) not with vbCrLf - This is wrong, it should be vbCrLf
                'see   http://cr.yp.to/docs/smtplf.html       for details
                '
                'Send Subject line
                Winsock1.SendData "From:" & txtSenderName & " <" & txtSender & ">" & vbCrLf
                Winsock1.SendData "To:" & txtRecipientName & " <" & txtRecipient & ">" & vbCrLf
                
                '
                Debug.Print "Subject: " & txtSubject
                '
                If Len(txtReplyTo.Text) > 0 Then
                    Winsock1.SendData "Subject:" & txtSubject & vbCrLf
                    Winsock1.SendData "Reply-To:" & txtReplyToName & " <" & txtReplyTo & ">" & vbCrLf & vbCrLf
                Else
                    Winsock1.SendData "Subject:" & txtSubject & vbCrLf & vbCrLf
                End If
                'Dim varLines() As String
                'Dim varLine As String
                Dim strMessage As String
                'Dim i
                '
                'Add atacchments
                strMessage = txtMessage & vbCrLf & vbCrLf & m_strEncodedFiles
                'clear memory
                m_strEncodedFiles = ""
                'Debug.Print Len(strMessage)
                'These lines aren't needed, see
                '
                'http://cr.yp.to/docs/smtplf.html for details
                '
                '*****************************************
                'Parse message to get lines (for VB6 only)
                'varLines() = Split(strMessage, vbNewLine)
                'Parse message to get lines (for VB5 and lower)
                'SplitMessage strMessage, varLines()
                'clear memory
                'strMessage = ""
                '
                'Send each line of the message
                'For i = LBound(varLines()) To UBound(varLines())
                '    Winsock1.SendData varLines(i) & vbCrLf
                '    '
                '    Debug.Print varLines(i)
                'Next
                '
                '******************************************
                Winsock1.SendData strMessage & vbCrLf
                strMessage = ""
                '
                'Send a dot symbol to inform server
                'that sending of message comleted
                Winsock1.SendData "." & vbCrLf
                '
                Debug.Print "."
                '
            Case MAIL_DOT
                'Change current state of the session
                m_State = MAIL_QUIT
                '
                'Send QUIT command to the server
                Winsock1.SendData "QUIT" & vbCrLf
                '
                Debug.Print "QUIT"
            Case MAIL_QUIT
                '
                'Close connection
                Winsock1.Close
                '
        End Select
       
    Else
        '
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        '
        Winsock1.Close
        '
        If Not m_State = MAIL_QUIT Then
            MsgBox "SMTP Error: " & strServerResponse, _
                    vbInformation, "SMTP Error"
        Else
            MsgBox "Message sent successfuly.", vbInformation
        End If
        '
    End If
    
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox "Winsock Error number " & Number & vbCrLf & _
            Description, vbExclamation, "Winsock Error"

End Sub
Public Sub SplitMessage(strMessage As String, strlines() As String)
Dim intAccs As Long
Dim I
Dim lngSpacePos As Long, lngStart As Long
    strMessage = Trim$(strMessage)
    lngSpacePos = 1
    lngSpacePos = InStr(lngSpacePos, strMessage, vbNewLine)
    Do While lngSpacePos
        intAccs = intAccs + 1
        lngSpacePos = InStr(lngSpacePos + 1, strMessage, vbNewLine)
    Loop
    ReDim strlines(intAccs)
    lngStart = 1
    For I = 0 To intAccs
        lngSpacePos = InStr(lngStart, strMessage, vbNewLine)
        If lngSpacePos Then
            strlines(I) = Mid(strMessage, lngStart, lngSpacePos - lngStart)
            lngStart = lngSpacePos + Len(vbNewLine)
        Else
            strlines(I) = Right(strMessage, Len(strMessage) - lngStart + 1)
        End If
    Next
End Sub

