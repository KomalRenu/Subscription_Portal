<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_GetFolderContents(sSessionID, sFolderID, sGetFolderContentsXML)
'********************************************************
'*Purpose: Retrieves the contents of a folder
'*Inputs: sSessionID, sFolderID
'*Outputs: sGetFolderContentsXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetFolderContents"
	Dim oSystemInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSystemInfo = Server.CreateObject(PROGID_SYSTEM_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "FoldersCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SYSTEM_INFO, LogLevelError)
    Else
        sGetFolderContentsXML = oSystemInfo.getFolderContents(sSessionID, sFolderID)
        lErrNumber = checkReturnValue(sGetFolderContentsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "FoldersCoLib.asp", PROCEDURE_NAME, "SystemInfo.getFolderContents", "Error while calling getFolderContents", LogLevelError)
        End If
	End If

	Set oSystemInfo = Nothing

	co_GetFolderContents = lErrNumber
	Err.Clear
End Function
%>