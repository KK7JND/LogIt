Attribute VB_Name = "basUtilites"
' declaration for sleep function
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'declarations for working with INI files
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias _
    "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFilename As String) As Long
 
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFilename As String) As Long
 
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias _
    "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, _
    ByVal lpFilename As String) As Long
 
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFilename As String) As Long
Public Sub getSettings()
On Error GoTo Err_getSettings

    Dim RetBuff As String
    RetBuff = Space$(255)

    ' Get settings from logit.ini file and update frmMain
    varFileName = App.Path & "\logit.ini"

    With frmMain
        ' Host settings
        nSize = GetPrivateProfileString("Monitor", "host", "", RetBuff, 255, varFileName)
        .lblHost.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Monitor", "port", "", RetBuff, 255, varFileName)
        .lblHost.Tag = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Monitor", "token", "", RetBuff, 255, varFileName)
        .lblToken.Tag = Mid(RetBuff, 1, nSize)
    
        ' Default settings
        nSize = GetPrivateProfileString("Default", "mode", "", RetBuff, 255, varFileName)
        .cboMode.Tag = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Default", "band", "", RetBuff, 255, varFileName)
        .cboBand.Tag = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Default", "snr_s", "", RetBuff, 255, varFileName)
        .txtSrTx.Tag = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Default", "snr_r", "", RetBuff, 255, varFileName)
        .txtSrRx.Tag = Mid(RetBuff, 1, nSize)
   
        ' Button Settings
        nSize = GetPrivateProfileString("Button 1", "label", "", RetBuff, 255, varFileName)
        .cmdButton1.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Button 1", "mode", "", RetBuff, 255, varFileName)
        .cmdButton1.Tag = Mid(RetBuff, 1, nSize)
    
        nSize = GetPrivateProfileString("Button 2", "label", "", RetBuff, 255, varFileName)
        .cmdButton2.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Button 2", "mode", "", RetBuff, 255, varFileName)
        .cmdButton2.Tag = Mid(RetBuff, 1, nSize)
    
        nSize = GetPrivateProfileString("Button 3", "label", "", RetBuff, 255, varFileName)
        .cmdButton3.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Button 3", "mode", "", RetBuff, 255, varFileName)
        .cmdButton3.Tag = Mid(RetBuff, 1, nSize)
    
        nSize = GetPrivateProfileString("Button 4", "label", "", RetBuff, 255, varFileName)
        .cmdButton4.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Button 4", "mode", "", RetBuff, 255, varFileName)
        .cmdButton4.Tag = Mid(RetBuff, 1, nSize)
    
        nSize = GetPrivateProfileString("Button 5", "label", "", RetBuff, 255, varFileName)
        .cmdButton5.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Button 5", "mode", "", RetBuff, 255, varFileName)
        .cmdButton5.Tag = Mid(RetBuff, 1, nSize)
    
        nSize = GetPrivateProfileString("Button 6", "label", "", RetBuff, 255, varFileName)
        .cmdButton6.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Button 6", "mode", "", RetBuff, 255, varFileName)
        .cmdButton6.Tag = Mid(RetBuff, 1, nSize)
    
        nSize = GetPrivateProfileString("Button 7", "label", "", RetBuff, 255, varFileName)
        .cmdButton7.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Button 7", "mode", "", RetBuff, 255, varFileName)
        .cmdButton7.Tag = Mid(RetBuff, 1, nSize)
    
        nSize = GetPrivateProfileString("Button 8", "label", "", RetBuff, 255, varFileName)
        .cmdButton8.Caption = Mid(RetBuff, 1, nSize)
        nSize = GetPrivateProfileString("Button 8", "mode", "", RetBuff, 255, varFileName)
        .cmdButton8.Tag = Mid(RetBuff, 1, nSize)
    End With
    
Exit_getSettings:
    Exit Sub
    
Err_getSettings:
    MsgBox "Error in basUtilites:getSettings: " & Err.Description
    Resume Exit_getSettings

End Sub
Public Sub getDate()
On Error GoTo Err_getDate

    ' Adjusts local time to UTC
    utcDate = LocalToUTC(Now())
    frmMain.txtDate = Format(utcDate, "yyyy-mm-dd hh:mm:ss")
    
Exit_getDate:
    Exit Sub
    
Err_getDate:
    MsgBox "Error in basUtilites:getDate: " & Err.Description
    Resume Exit_getDate

End Sub
Public Sub sendLog()
On Error GoTo Err_sendLog

    Dim varJASON As String
    Dim RetBuff As String
    RetBuff = Space$(255)

    'Raw and escaped data for reference
    'varJASON = {"params":{"CALL":"W4FAKE","COMMENTS":"Test Comment","EXTRA":{"MODE":""},"FREQ":7112000,"GRID":"EM73","MODE":"MFSK","NAME":"Test Name","RPT.RECV":"588","RPT.SENT":"599","STATION.CALL":"W4FAKE","STATION.GRID":"EM73","STATION.OP":"W4FAKE","SUBMODE":"JS8","UTC.OFF":1699221591236,"UTC.ON":1699221591236,"_ID":-1},"type":"LOG.QSO","value":"<call:6>W4FAKE <gridsquare:4>EM73 <mode:4>MFSK <submode:3>JS8 <rst_sent:3>599 <rst_rcvd:3>588 <qso_date:8>20231105 <time_on:6>145951 <qso_date_off:8>20231105 <time_off:6>145951 <band:3>40m <freq:8>7.112000 <station_callsign:6>W4FAKE <my_gridsquare:4>EM73 <comment:12>Test Comment <name:9>Test Name <operator:6>W4FAKE <MODE:0>"}
    'varJASON = "{""params"":{""CALL"":""W4FAKE"",""COMMENTS"":""Test Comment"",""EXTRA"":{""MODE"":""""},""FREQ"":7112000,""GRID"":""EM73"",""MODE"":""MFSK"",""NAME"":""Test Name"",""RPT.RECV"":""588"",""RPT.SENT"":""599"",""STATION.CALL"":""W4FAKE"",""STATION.GRID"":""EM73"",""STATION.OP"":""W4FAKE"",""SUBMODE"":""JS8"",""UTC.OFF"":1699221591236,""UTC.ON"":1699221591236,""_ID"":-1},""type"":""LOG.QSO"",""value"":""<call:6>W4FAKE <gridsquare:4>EM73 <mode:4>MFSK <submode:3>JS8 <rst_sent:3>599 <rst_rcvd:3>588 <qso_date:8>20231105 <time_on:6>145951 <qso_date_off:8>20231105 <time_off:6>145951 <band:3>40m <freq:8>7.112000 <station_callsign:6>W4FAKE <my_gridsquare:4>EM73 <comment:12>Test Comment <name:9>Test Name <operator:6>W4FAKE <MODE:0>""}"
    
    ' Build the base JSON string
    varJASON = "{""params"":{""STATION.CALL"":""W4FAKE"",""AUTH"":"""
    varJASON = varJASON & frmMain.lblToken.Tag
    varJASON = varJASON & """},""type"":""LOG.QSO"",""value"":"""
    
    ' Fill in data from the form
    With frmMain
        If Not .txtCall.Text = vbNullString Then
            varJASON = varJASON & "<call:" & Len(.txtCall.Text) & ">" & .txtCall.Text
        End If
        
        If Not .txtDate.Text = vbNullString Then
            'calculate dates
            varCdate = Format(.txtDate.Text, "yyyymmdd")
            varJASON = varJASON & "<qso_date:" & Len(varCdate) & ">" & varCdate
            varJASON = varJASON & "<qso_date_off:" & Len(varCdate) & ">" & varCdate
            
            ' calculate times
            varCtime = Format(.txtDate.Text, "hhmmss")
            varJASON = varJASON & "<time_on:" & Len(varCtime) & ">" & varCtime
            varJASON = varJASON & "<time_off:" & Len(varCtime) & ">" & varCtime
        End If
        
        If Not .cboMode.Text = vbNullString Then
            ' this mimics the JS8Call GUI behavior so the Monitor script works as expected
            varJASON = varJASON & "<mode:4>MFSK<submode:3>JS8"
            varJASON = varJASON & "<MODE:" & Len(.cboMode.Text) & ">" & .cboMode.Text
        End If
        
        If Not .cboBand.Text = vbNullString Then
            varJASON = varJASON & "<band:" & Len(.cboBand.Text) & ">" & .cboBand.Text
        End If
        
        If Not .txtSrTx.Text = vbNullString Then
            varJASON = varJASON & "<rst_sent:" & Len(.txtSrTx.Text) & ">" & .txtSrTx.Text
        End If
        
        If Not .txtSrRx.Text = vbNullString Then
            varJASON = varJASON & "<rst_rcvd:" & Len(.txtSrRx.Text) & ">" & .txtSrRx.Text
        End If
        
        If Not .txtGrid.Text = vbNullString Then
            varJASON = varJASON & "<gridsquare:" & Len(.txtGrid.Text) & ">" & .txtGrid.Text
        End If
        
        If Not .txtName.Text = vbNullString Then
            varJASON = varJASON & "<name:" & Len(.txtName.Text) & ">" & .txtName.Text
        End If
        
        If Not .txtComments.Text = vbNullString Then
            varJASON = varJASON & "<comment:" & Len(.txtComments.Text) & ">" & .txtComments.Text
        End If
    
    End With
    
    'Tags sent from JS8Call GUI but not sent from this app
    '<freq:8>7.112000 <station_callsign:6>W4FAKE <my_gridsquare:4>EM73 <operator:6>W4FAKE
    
    ' Finish off the JSON string
    varJASON = varJASON & """" & "}"
    
    ' Fetch host settings
    With frmMain.lblHost
        varMhost = .Caption
        varMport = .Tag
    End With
    
    ' Setup Winsock control and send data
    With frmMain.ctlWinsock
        .Close
        .Protocol = sckUDPProtocol
        .RemoteHost = varMhost
        .RemotePort = varMport
        .SendData varJASON
    End With
    
Exit_sendLog:
    Exit Sub
    
Err_sendLog:
    MsgBox "Error in basUtilites:sendLog: " & Err.Description
    Resume Exit_sendLog

End Sub


