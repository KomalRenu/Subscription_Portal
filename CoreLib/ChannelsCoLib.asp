<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

Function co_AddChannel(sSiteId, sChannelId, sChannelXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
CONST PROCEDURE_NAME = "co_AddChannel"
Dim lErr
Dim sErrDesc
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturn = oSiteInfo.createChannel(sSiteId, sChannelId, sChannelXML)
        lErr = checkReturnValue(sReturn, sErrDesc)
        If lErr <> 0 Then Call LogErrorXML(aConnectionInfo, lErr, sErrDesc, Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME,  aWizardInfo(CHANNEL_ACTION), "Error calling " & aWizardInfo(CHANNEL_ACTION), LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_AddChannel = lErr
    Err.Clear

End Function


Function co_DeleteChannel(sSiteID, sChannelID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
CONST PROCEDURE_NAME = "co_DeleteChannel"
Dim lErr
Dim sErr
Dim sReturn
Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturn = oSiteInfo.deleteChannel(sSiteID, sChannelID)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "deleteChannel", "Error calling deleteChannel", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_DeleteChannel = lErr
    Err.Clear

End Function

Function co_UpdateChannel(sSiteID, sChannelID, sChannelXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
CONST PROCEDURE_NAME = "co_UpdateChannel"
Dim lErr
Dim sErr
Dim sReturn
Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturn = oSiteInfo.updateChannelProperties(sSiteID, sChannelID, sChannelXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "deleteChannel", "Error calling deleteChannel", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_UpdateChannel = lErr
    Err.Clear

End Function

%>