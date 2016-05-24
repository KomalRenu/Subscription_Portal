<%'** Copyright © 2000-2012 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_GetUserSubscriptions(sSessionID, sChannelID, sGetUserSubscriptionsXML)
'********************************************************
'*Purpose: Get user's subscriptions for a given channel.
'*Inputs: sSessionID, sChannelID
'*Outputs: sGetUserSubscriptionsXML
'********************************************************
	On Error Resume Next
    Const PROCEDURE_NAME = "co_GetUserSubscriptions"
	Dim oSubscription
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSubscription = Server.CreateObject(PROGID_SUBSCRIPTION)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscriptionsCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SUBSCRIPTION, LogLevelError)
	Else
	    sGetUserSubscriptionsXML = oSubscription.getUserSubscriptions(sSessionID, sChannelID)
	    lErrNumber = checkReturnValue(sGetUserSubscriptionsXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "SubscriptionsCoLib.asp", PROCEDURE_NAME, "Subscription.getUserSubscriptions", "Error while calling getUserSubscriptions", LogLevelError)
	    End If
	End If

	Set oSubscription = Nothing

	co_GetUserSubscriptions = lErrNumber
	Err.Clear
End Function
%>