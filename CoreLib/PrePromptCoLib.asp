<%'** Copyright © 2000-2012 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_getUserDefaultPersonalization(sSessionID, sUserDefaultPersonalizationXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID
'*Outputs: sGetUserSecurityObjectsXML
'********************************************************
    On Error Resume Next
	const PROCEDURE_NAME = "co_getUserDefaultPersonalization"
	Dim oPersonalizationInfo
	Dim sErr

	lErrNumber = NO_ERR

    Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sUserDefaultPersonalizationXML = oPersonalizationInfo.getUserDefaultPersonalization(sSessionID)
        lErrNumber = checkReturnValue(sUserDefaultPersonalizationXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getUserDefaultPersonalization", "Error calling getUserDefaultPersonalization", LogLevelTrace)
        End If
    End If

    Set oPersonalizationInfo = Nothing

    co_getUserDefaultPersonalization = lErrNumber
    Err.Clear
End Function
%>