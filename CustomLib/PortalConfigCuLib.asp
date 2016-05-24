<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>

<!-- #include file="../CoreLib/PortalConfigCoLib.asp" -->

<%

Function CreateNewPortal(sPortalName)
'********************************************************
'*Purpose: Creates a new VirtualDirectory in IIS and adds a definition to subscriptionPortal.properties file
'*Inputs: Portal Name
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "CreateNewPortal"
Dim lErr
Dim sInstallPath
Dim Start
Dim ASPDir
Dim sConfigName
Dim sConfigXML

    On Error Resume Next
    lErr = NO_ERR

	If lErr = NO_ERR Then
		ASPDir = Server.MapPath("../")
		sConfigName = LCase(sPortalName)

		lErr = co_CreateNewPortal(sConfigName, ASPDir)
		If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "PortalConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateNewPortal", LogLevelTrace)
	End If

	If lErr = NO_ERR Then
		sConfigXML = "<mi><oi id="""" n=""" & Server.HTMLEncode(sConfigName) & """><prs>"
		sConfigXML = sConfigXML & "<pr id=""DISPLAY_NAME""  v=""" & Server.HTMLEncode(sPortalName) & """ />"
		sConfigXML = sConfigXML & "</prs></oi></mi>"

		lErrNumber = co_SetDisplayName(sConfigXML)
		If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetDisplayName", LogLevelTrace)
	End If

    CreateNewPortal = lErr
    Err.Clear

End Function

Function DeletePortal(sPortalName)
'********************************************************
'*Purpose: Deletes a Portal from IIS and from subscriptionPortal.properties file
'*Inputs: Portal Name
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "DeletePortal"
    Dim lErr

    lErr = NO_ERR

    lErr = co_DeletePortal(sPortalName)
    If lErr <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "PortalConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_DeletePortal", LogLevelTrace)
    End If

    DeletePortal = lErr
    Err.Clear
End Function


%>