<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_CreateProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sPreferenceObjectID, sQuestionObjectID, sProfileName, sProfileDesc
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_CreateProfile"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr
	Dim sCreateProfileXML

	lErrNumber = NO_ERR

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sCreateProfileXML = oPersonalizationInfo.createProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault)
        lErrNumber = checkReturnValue(sCreateProfileXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PostPromptCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.createProfile", "Error while calling createProfile", LogLevelError)
        End If
	End If

	Set oPersonalizationInfo = Nothing

	co_CreateProfile = lErrNumber
	Err.Clear
End Function

Function co_UpdateProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault)
'********************************************************
'*Purpose:
'*Inputs:	sSessionID, sPreferenceObjectID, sQuestionObjectID, sProfileName, sProfileDesc, bIsDefault
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_UpdateProfile"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr
	Dim sUpdateProfileXML

	lErrNumber = NO_ERR

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sUpdateProfileXML = oPersonalizationInfo.updateProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault)
        lErrNumber = checkReturnValue(sUpdateProfileXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PostPromptCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.updateProfile", "Error while calling updateProfile", LogLevelError)
        End If
	End If

	Set oPersonalizationInfo = Nothing

	co_UpdateProfile = lErrNumber
	Err.Clear
End Function
%>