<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_CreateUser(sSiteID, asUserProperties)
'********************************************************
'*Purpose: Given user properties, returns XML with results of request.
'*Inputs: sSiteID, asUserProperties (an array of user properties)
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_CreateUser"
	Dim oUser
	Dim lErrNumber
	Dim sErr
	Dim sCreateUserXML

	lErrNumber = NO_ERR

	Set oUser = Server.CreateObject(PROGID_USER)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sCreateUserXML = oUser.createUser(sSiteID, asUserProperties)
        lErrNumber = checkReturnValue(sCreateUserXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.createUser", "Error while calling createUser", LogLevelError)
        End If
	End If

	Set oUser = Nothing

	co_CreateUser = lErrNumber
	Err.Clear
End Function

Function co_CreateSession(sSiteID, sUserName, sPassword, bEncryptedFlag, sCreateSessionXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, sUserName, sPassword, bEncryptedFlag
'*Outputs: sCreateSessionXML
'*NOTES: This function doesn't call checkReturnValue, because the XML
'*       needs to be parsed to get the session ID anyway.
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_CreateSession"
    Dim oUser
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sCreateSessionXML = oUser.createSession(sSiteID, sUserName, sPassword, bEncryptedFlag)
        lErrNumber = checkReturnValue(sCreateSessionXML, sErr)
        If lErrNumber <> NO_ERR Then
           Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.createSession", "Error calling createSession", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_CreateSession = lErrNumber
    Err.Clear
End Function

Function co_GetUserHint(sSiteID, sUsername, sGetUserHintXML)
'********************************************************
'*Purpose: Given a username, returns the XML from the transaction.
'*Inputs: sSiteID, sUsername
'*Outputs: sGetUserHintXML
'*NOTES: This function doesn't call checkReturnValue, because the XML
'*       needs to be parsed to get the user hint anyway.
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetUserHint"
	Dim oUser
	Dim lErrNumber

	lErrNumber = NO_ERR

	Set oUser = Server.CreateObject(PROGID_USER)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
	Else
	    sGetUserHintXML = oUser.getUserHint(sSiteID, sUsername)
	    lErrNumber = checkReturnValue(sCreateSessionXML, sErr)
        If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.getUserHint", "Error calling getUserHint", LogLevelError)
	    End If
	End If

	Set oUser = Nothing

	co_GetUserHint = lErrNumber
	Err.Clear
End Function

Function co_DeleteUser(sSessionID)
'********************************************************
'*Purpose:
'*Inputs: sSessionID
'*Outputs: sDeleteUserXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_DeleteUser"
    Dim oUser
    Dim lErrNumber
    Dim sErr
    Dim sDeleteUserXML

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sDeleteUserXML = oUser.deleteUser(sSessionID)
        lErrNumber = checkReturnValue(sDeleteUserXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.deleteUser", "Error calling deleteUser", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_DeleteUser = lErrNumber
    Err.Clear
End Function

Function co_DeactivateUser(sSessionID, sPassword)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sPassword
'*Outputs: sDeactivateUserXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_DeactivateUser"
    Dim oUser
    Dim lErrNumber
    Dim sErr
    Dim sDeactivateUserXML

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sDeactivateUserXML = oUser.deactivateUser(sSessionID, sPassword)
        lErrNumber = checkReturnValue(sDeactivateUserXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.deactivateUser", "Error calling deactivateUser", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_DeactivateUser = lErrNumber
    Err.Clear
End Function

Function co_CloseSession(sSessionID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_CloseSession"
    Dim oUser
    Dim lErrNumber
    Dim sErr
    Dim sCloseSessionXML

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sCloseSessionXML = oUser.closeSession(sSessionID)
        lErrNumber = checkReturnValue(sCloseSessionXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.closeSession", "Error calling closeSession", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_CloseSession = lErrNumber
    Err.Clear
End Function

Function co_GetUserProperties(sSessionID, sGetUserPropertiesXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetUserProperties"
    Dim oUser
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sGetUserPropertiesXML = oUser.getUserProperties(sSessionID)
        lErrNumber = checkReturnValue(sGetUserPropertiesXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.getUserProperties", "Error calling getUserProperties", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_GetUserProperties = lErrNumber
    Err.Clear
End Function

Function co_UpdateUserProperties(sSessionID, asUserProperties)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_UpdateUserProperties"
    Dim oUser
    Dim lErrNumber
    Dim sErr
    Dim sUpdateUserPropertiesXML

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sUpdateUserPropertiesXML = oUser.updateUserProperties(sSessionID, asUserProperties)
        lErrNumber = checkReturnValue(sUpdateUserPropertiesXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.updateUserProperties", "Error calling updateUserProperties", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_UpdateUserProperties = lErrNumber
    Err.Clear
End Function

Function co_SaveUserAuthenticationObjects(sSessionID, asInformationSourceID, asXMLAuthenticationString)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, asInformationSourceID, asXMLAuthenticationString
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_SaveUserAuthenticationObjects"
    Dim oUser
    Dim lErrNumber
    Dim sErr
    Dim sSaveUserAuthenticationObjectsXML

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sSaveUserAuthenticationObjectsXML = oUser.saveUserAuthenticationObjects(sSessionID, asInformationSourceID, asXMLAuthenticationString)
        lErrNumber = checkReturnValue(sSaveUserAuthenticationObjectsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.saveUserAuthenticationObjects", "Error calling saveUserAuthenticationObjects", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_SaveUserAuthenticationObjects = lErrNumber
    Err.Clear
End Function

Function co_GetInformationSourcesForSite(sSiteID, sGetInformationSourcesForSiteXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID
'*Outputs: sGetInformationSourcesForSiteXML
'*NOTES: This function doesn't call checkReturnValue, because the XML
'*       needs to be parsed to determine if there are projects.
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetInformationSourcesForSite"
    Dim oSiteInfo
    Dim lErrNumber

    lErrNumber = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sGetInformationSourcesForSiteXML = oSiteInfo.getInformationSourcesForSite(sSiteID)
        lErrNumber = checkReturnValue(sSaveUserAuthenticationObjectsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "SiteInfo.getInformationSourcesForSite", "Error calling getInformationSourcesForSite", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_GetInformationSourcesForSite = lErrNumber
    Err.Clear
End Function
%>