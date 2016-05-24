<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/FoldersCoLib.asp" -->
<%
Function cu_GetFolderContents(sFolderID, sGetFolderContentsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetFolderContents"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
    sSessionID = GetSessionID()

	lErrNumber = co_GetFolderContents(sSessionID, sFolderID, sGetFolderContentsXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "FoldersCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetFolderContents", LogLevelTrace)
	End If

	cu_GetFolderContents = lErrNumber
	Err.Clear
End Function
%>