<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_GetQuestionsAndProfilesForSubscriptionSet(sSessionID, sSubSetID, sServiceID, sGQAPFSSXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sSubSetID
'*Outputs: sGQAPFSSXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetQuestionsAndProfilesForSubscriptionSet"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sGQAPFSSXML = oPersonalizationInfo.getQuestionsAndProfilesForSubscriptionSet(sSessionID, sServiceID, sSubSetID)
        lErrNumber = checkReturnValue(sGQAPFSSXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PersonalizeCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getQuestionsAndProfilesForSubscriptionSet", "Error while calling getQuestionsAndProfilesForSubscriptionSet", LogLevelError)
        End If
	End If

	Set oPersonalizationInfo = Nothing

	co_GetQuestionsAndProfilesForSubscriptionSet = lErrNumber
	Err.Clear
End Function

Function co_GetSubscription(sSessionID, sSubscriptionGUID, sGetSubscriptionXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sSubscriptionGUID
'*Outputs: sGetSubscriptionXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetSubscription"
	Dim oSubscription
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSubscription = Server.CreateObject(PROGID_SUBSCRIPTION)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SUBSCRIPTION, LogLevelError)
    Else
        sGetSubscriptionXML = oSubscription.getSubscription(sSessionID, sSubscriptionGUID)
        lErrNumber = checkReturnValue(sGetSubscriptionXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PersonalizeCoLib.asp", PROCEDURE_NAME, "Subscription.getSubscription", "Error while calling getSubscription", LogLevelError)
        End If
	End If

	Set oSubscription = Nothing

	co_GetSubscription = lErrNumber
	Err.Clear
End Function

Function co_GetDetailsForQuestions(sSessionID, asQuestionObjectID, sGetDetailsForQuestionsXML)
'********************************************************
'*Purpose: Given a asQuestionObjectID, returns details of that question
'*Inputs: sSessionID, asQuestionObjectID
'*Outputs: sGetDetailsForQuestionsXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetDetailsForQuestions"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sGetDetailsForQuestionsXML = oPersonalizationInfo.getDetailsForQuestions(sSessionID, asQuestionObjectID)
        lErrNumber = checkReturnValue(sGetDetailsForQuestionsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getDetailsForQuestions", "Error while calling getDetailsForQuestions", LogLevelError)
        End If
	End If

	Set oPersonalizationInfo = Nothing

	co_GetDetailsForQuestions = lErrNumber
	Err.Clear
End Function

Function co_GetUserSecurityObjects(sSessionID, sGetUserSecurityObjectsXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID
'*Outputs: sGetUserSecurityObjectsXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetUserSecurityObjects"
    Dim oUser
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sGetUserSecurityObjectsXML = oUser.getUserSecurityObjects(sSessionID)
        lErrNumber = checkReturnValue(sGetUserSecurityObjectsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "User.getUserSecurityObjects", "Error calling getUserSecurityObjects", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_GetUserSecurityObjects = lErrNumber
    Err.Clear
End Function

%>
