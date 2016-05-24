<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

Function co_GetInstallDirectory(sInstallPath)
'********************************************************
'*Purpose: Gets the Hydra Install Directory from Properties File
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetInstallDirectory"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oInstallPath
    Dim oDOM
    Dim oInstall
    Dim i

    lErr = NO_ERR
    Set oInstallPath = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn = oInstallPath.getInstallationPath()
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "getInstallationPath", "Error calling getInstallationPath", LogLevelError)
        End If
    End If

    If lErr = NO_ERR Then
        Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
        oDOM.async = False
        If oDOM.loadXML(sReturn) = False Then
            lErr = ERR_XML_LOAD_FAILED
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "loadXML", "Error loading XML from co_GetInstallDirectory", LogLevelError)
        End If
    End If

    If lErr = NO_ERR Then
        Set oInstall = oDOM.selectNodes("//oi[@tp='1015']")

        If (Not (oInstall Is Nothing)) Then
            For i = 0 to (oInstall.length)
		sInstallPath = oInstall(i).Attributes.getNamedItem("n").Text
	    Next
        End If

    End If

    Set oInstallPath = Nothing
    Set oDOM = Nothing
    Set oInstall = Nothing

    co_CreateNEwPortal = lErr
    Err.Clear

End Function


Function co_CreateNewPortal(sPortalName,ASPDir)
'********************************************************
'*Purpose: Creates a new VirtualDirectory in IIS and adds a definition to subscriptionPortal.properties file
'*Inputs: Portal Name,Path for ASP Directory
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_CreateNewPortal"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oPortalInfo
    Dim oVirtualdirectory
    Dim sMachineName
    Dim sDefaultWebIndex

    sMachineName = "localhost"
    sDefaultWebIndex = "1"

    lErr = NO_ERR
    Set oPortalInfo = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn = oPortalInfo.addVirtualDirectory(sPortalName)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "addVirtualDirectory", "Error calling addVirtualDirectory", LogLevelError)
        End If
    End If

    If lErr = NO_ERR Then
	    Set oVirtualDirectory = Server.CreateObject(PROGID_PORTAL_VB_ADMIN)
	    If Err.number <> NO_ERR Then
	        lErr = Err.number
	        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PORTAL_VB_ADMIN, LogLevelError)
	    Else
	        lErr = oVirtualDirectory.CreateSubscriptionVirtualDirectory(sMachineName,sDefaultWebIndex,sPortalName,ASPDir,sErr)
	        If lErr <> NO_ERR Then
        	    Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "CreateSubscriptionVirtualDirectory", "Error calling CreateSubscriptionVirtualDirectory", LogLevelError)
	        End If
	    End If

	    'If an error happened, remove PortalName from available list.
	    If lErr <> NO_ERR Then
			Call oPortalInfo.deleteVirtualDirectory(sPortalName)
		End If

    End If

	Set oVirtualDirectory = Nothing
	Set oPortalInfo = Nothing

    co_CreateNEwPortal = lErr
    Err.Clear
End Function

Function co_DeletePortal(sPortalName)
'********************************************************
'*Purpose:  Deletes the given Portal.
'*Inputs:   Portal Name
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_DeletePortal"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oPortalInfo
    Dim oVirtualdirectory
    Dim sMachineName
    Dim sDefaultWebIndex

    sMachineName = "localhost"
    sDefaultWebIndex = "1"

    lErr = NO_ERR
    Set oPortalInfo = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn = oPortalInfo.deleteVirtualDirectory(sPortalName)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "deleteVirtualDirectory", "Error calling deleteVirtualDirectory", LogLevelError)
        End If
    End If
    Set oPortalInfo = Nothing

    If lErr = NO_ERR Then
	    Set oVirtualDirectory = Server.CreateObject(PROGID_PORTAL_VB_ADMIN)
	    If Err.number <> NO_ERR Then
	        lErr = Err.number
	        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PORTAL_VB_ADMIN, LogLevelError)
	    Else
	        lErr = oVirtualDirectory.RemoveSubscriptionVirtualDirectory(sMachineName,sDefaultWebIndex,sPortalName,sErr)
	        If lErr <> NO_ERR Then
        	    Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "RemoveSubscriptionVirtualDirectory", "Error calling RemoveSubscriptionVirtualDirectory", LogLevelError)
	        End If
	    End If
	    Set oVirtualDirectory = Nothing
    End If

    co_DeletePortal = lErr
    Err.Clear
End Function

Function co_SetDisplayName(sConfigXML)
  '********************************************************
'*Purpose:  Sets the name in the properties file
'*Inputs:   sConfigXML: XML with the properties necessary to set the display name
'*Outputs:  None
'********************************************************
CONST PROCEDURE_NAME = "co_SetDisplayName"
Dim lErr
Dim sErr
Dim sReturn
Dim oAdmin

    On Error Resume Next
    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn =  oAdmin.saveSiteConfigurationProperties(sConfigXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "PortalConfigCoLib.asp", PROCEDURE_NAME, "saveSiteConfigurationProperties", "Error calling saveSiteConfigurationProperties", LogLevelError)
    End If

    Set oAdmin = Nothing

    co_SetDisplayName = lErr
    Err.Clear

End Function



%>