<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_DeleteProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, bForceDelete, sDeleteProfileXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sQuestionObjectID, sPreferenceObjectID
'*Outputs: sDeleteProfileXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_DeleteProfile"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeleteProfileCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sDeleteProfileXML = oPersonalizationInfo.deleteProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, bForceDelete)
		lErrNumber = checkReturnValue(sDeleteProfileXML, sErr)
		If lErrNumber <> NO_ERR Then
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "DeleteProfileCoLib.asp", PROCEDURE_NAME, "deleteProfile", "Error calling deleteProfile", LogLevelError)
		End If
	End If

	Set oPersonalizationInfo = Nothing

	co_DeleteProfile = lErrNumber
	Err.Clear
End Function
%>