<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_UpdateUserPassword(sSessionID, sOldPassword, sNewPassword)
'********************************************************
'*Purpose: Changes the user's password
'*Inputs: sSessionID, sOldPassword, sNewPassword
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_UpdateUserPassword"
    Dim oUser
    Dim lErrNumber
    Dim sErr
    Dim sUpdateUserPasswordXML

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sUpdateUserPasswordXML = oUser.updateUserPassword(sSessionID, sOldPassword, sNewPassword)
        lErrNumber = checkReturnValue(sUpdateUserPasswordXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.updateUserPassword", "Error calling updateUserPassword", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_UpdateUserPassword = lErrNumber
    Err.Clear
End Function

Function co_UpdateUserAuthenticationObjects(sSessionID, asInformationSourceID, asXMLAuthenticationString)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, asInformationSourceID, asXMLAuthenticationString
'*Outputs: sUpdateUserAuthenticationObjectsXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_UpdateUserAuthenticationObjects"
    Dim oUser
    Dim lErrNumber
    Dim sErr
    Dim sUpdateUserAuthenticationObjectsXML

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sUpdateUserAuthenticationObjectsXML = oUser.updateUserAuthenticationObjects(sSessionID, asInformationSourceID, asXMLAuthenticationString)
        lErrNumber = checkReturnValue(sUpdateUserAuthenticationObjectsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "LoginCoLib.asp", PROCEDURE_NAME, "User.updateUserAuthenticationObjects", "Error calling updateUserAuthenticationObjects", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_UpdateUserAuthenticationObjects = lErrNumber
    Err.Clear
End Function
%>