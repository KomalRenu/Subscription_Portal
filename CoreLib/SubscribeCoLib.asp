<%'** Copyright © 2000-2012 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_GetUserAddressesForService(sSessionID, sServiceID, sGetUserAddressesForServiceXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sServiceID
'*Outputs: sGetUserAddressesForServiceXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetUserAddressesForService"
	Dim oAddresses
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oAddresses = Server.CreateObject(PROGID_ADDRESS)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADDRESS, LogLevelError)
	Else
	    sGetUserAddressesForServiceXML = oAddresses.getUserAddressesForService(sSessionID, sServiceID)
	    lErrNumber = checkReturnValue(sGetUserAddressesForServiceXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "Addresses.getUserAddressesForService", "Error while calling getUserAddressesForService", LogLevelError)
	    End If
	End If

	Set oAddresses = Nothing

	co_GetUserAddressesForService = lErrNumber
	Err.Clear
End Function

Function co_GetNamedSchedulesForService(sSessionID, sServiceID, bFlagValid, sGetNamedSchedulesForServiceXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sServiceID, bFlagValid
'*Outputs: sGetNamedSchedulesForServiceXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetNamedSchedulesForService"
	Dim oSystemInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSystemInfo = Server.CreateObject(PROGID_SYSTEM_INFO)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SYSTEM_INFO, LogLevelError)
	Else
	    sGetNamedSchedulesForServiceXML = oSystemInfo.getNamedSchedulesForService(sSessionID, sServiceID, bFlagValid)
	    lErrNumber = checkReturnValue(sGetNamedSchedulesForServiceXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "SystemInfo.getNamedSchedulesForService", "Error while calling getNamedSchedulesForService", LogLevelError)
	    End If
	End If

	Set oSystemInfo = Nothing

	co_GetNamedSchedulesForService = lErrNumber
	Err.Clear
End Function

Function co_DeleteSubscriptions(sDeleteSubscriptionDataXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_DeleteSubscriptions"
	Dim oSubscription
	Dim lErrNumber
	Dim sErr
	Dim sResultXML

	lErrNumber = NO_ERR

	Set oSubscription = Server.CreateObject(PROGID_SUBSCRIPTION)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SUBSCRIPTION, LogLevelError)
	Else
	    sResultXML = oSubscription.deleteSubscriptions(sDeleteSubscriptionDataXML)
	    lErrNumber = checkReturnValue(sResultXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "Subscription.deleteSubscriptions", "Error while calling deleteSubscriptions", LogLevelError)
	    End If
	End If

	Set oSubscription = Nothing

	co_DeleteSubscriptions = lErrNumber
	Err.Clear
End Function

Function co_DeleteAllSubscriptions(sSessionID, sChannelID, sDeleteAllSubscriptionsXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sChannelID
'*Outputs: sDeleteAllSubscriptionsXML
'* QUESTION: Is this API implemented?
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_DeleteAllSubscriptions"
	Dim oSubscription
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSubscription = Server.CreateObject(PROGID_SUBSCRIPTION)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SUBSCRIPTION, LogLevelError)
	Else
	    sDeleteAllSubscriptionsXML = oSubscription.deleteAllSubscriptions(sSessionID, sChannelID)
	    lErrNumber = checkReturnValue(sDeleteAllSubscriptionsXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "Subscription.deleteAllSubscriptions", "Error while calling deleteAllSubscriptions", LogLevelError)
	    End If
	End If

	Set oSubscription = Nothing

	co_DeleteAllSubscriptions = lErrNumber
	Err.Clear
End Function

Function co_EditSubscription(sSubscriptionDataXML)
'********************************************************
'*Purpose:
'*Inputs: sSubscriptionDataXML
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_EditSubscription"
	Dim oSubscription
	Dim lErrNumber
	Dim sErr
	Dim sResultXML

	lErrNumber = NO_ERR

	Set oSubscription = Server.CreateObject(PROGID_SUBSCRIPTION)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SUBSCRIPTION, LogLevelError)
	Else
	    sResultXML = oSubscription.editSubscription(sSubscriptionDataXML)
	    lErrNumber = checkReturnValue(sResultXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "Subscription.editSubscription", "Error while calling editSubscription", LogLevelError)
	    End If
	End If

	Set oSubscription = Nothing

	co_EditSubscription = lErrNumber
	Err.Clear
End Function

Function co_AddSubscription(sSubscriptionDataXML)
'********************************************************
'*Purpose:
'*Inputs: sSubscriptionDataXML
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_AddSubscription"
	Dim oSubscription
	Dim lErrNumber
	Dim sErr
	Dim sResultXML

	lErrNumber = NO_ERR

	Set oSubscription = Server.CreateObject(PROGID_SUBSCRIPTION)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SUBSCRIPTION, LogLevelError)
	Else
	    sResultXML = oSubscription.addSubscription(sSubscriptionDataXML)
	    lErrNumber = checkReturnValue(sResultXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "Subscription.addSubscription", "Error while calling addSubscription", LogLevelError)
	    End If
	End If

	Set oSubscription = Nothing

	co_AddSubscription = lErrNumber
	Err.Clear
End Function

Function co_UpdatePreferenceObjects(sPreferenceDataXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_UpdatePreferenceObjects"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr
	Dim sResultXML

	lErrNumber = NO_ERR

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
	Else
	    sResultXML = oPersonalizationInfo.updatePreferenceObjects(sPreferenceDataXML)
	    lErrNumber = checkReturnValue(sResultXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "Subscription.updatePreferenceObjects", "Error while calling updatePreferenceObjects", LogLevelError)
	    End If
	End If

	Set oPersonalizationInfo = Nothing

	co_UpdatePreferenceObjects = lErrNumber
	Err.Clear
End Function

Function co_ExpireSubscription(sSessionID, asSubscriptionGUID)
'********************************************************
'*Purpose: Given a subscriptionGUID, initiates a transaction to expire
'		   that subscription from the document repository.  The transaction's ID is returned.
'*Inputs: sAuthToken, sSiteID, asSubscriptionGUID (an array of subscriptionsGUIDs)
'*Outputs: sReqExpireSubscription
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_ExpireSubscription"
	Dim oDocRepository
	Dim lErrNumber
	Dim sErr
	Dim sResults

	lErrNumber = NO_ERR

	Set oDocRepository  = Server.CreateObject(PROGID_DOC_REPOSITORY)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_DOC_REPOSITORY, LogLevelError)
	Else
		sResults = oDocRepository.expireSubscriptions(sSessionID, asSubscriptionGUID, False)
		'lErrNumber = checkReturnValue(sResultXML, sErr)
	    If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "SubscribeCoLib.asp", PROCEDURE_NAME, "", "Error when expiring the subscription ", LogLevelError)
		End If
	End If

	Set oDocRepository = Nothing

	co_ExpireSubscription = lErrNumber
	Err.Clear
End Function
%>